# Run as 32-bit powershell script. If you need to run under 64-bit, then make sure 64-bit database drivers are installed. 
# Todos: run this as a service, how to monitor to make sure this script is always running?
# Config file is separated to save disk space, otherwise each file will have same settings throughout.. 
# Need error handling if case id not found, etc. move to error folder with error description.. 

#if ($env:Processor_Architecture -ne "x86")   
#{ write-warning 'Launching x86 PowerShell'
#&"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noninteractive -noprofile -file $myinvocation.Mycommand.path -executionpolicy bypass
#exit
#}
#"Always running in 32bit PowerShell at this point."
#$env:Processor_Architecture
#[IntPtr]::Size


function WriteLog ($logFile, $message)
{
  $lineMessage = (Get-Date).ToString() + " - " + $message
  add-content -Path $logFile -Value $lineMessage -Force
}

# -------------------------------------------------------------
# get configuration information (i.e. where to put files after 
# success or error
# -------------------------------------------------------------

# This needs to change to the script folder or else it will default to c:\Users\UserName...
# It then can't find the configuration file..  TODO: Pass it from C#.. 
cd C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Scripts

[xml]$config = Get-Content config.daxml
$folderError = $config.configuration.folderOutputError
$folderSuccess = $config.configuration.folderOutputSuccess
$folderInputWork = $config.configuration.folderInputWork
$folderInput = $config.configuration.folderInput

$MainLogFile = $folderError + "\documentagent.log"

WriteLog $MainLogFile "Script started."

$LogMessage = "Processor = " + $env:Processor_Architecture
WriteLog $MainLogFile $LogMessage


# -------------------------------------------------------------
# Open OleDB connection.. 
# -------------------------------------------------------------
WriteLog $MainLogFile "Opening DB Connection"

$con = New-Object System.Data.OleDb.OleDbConnection
$con.ConnectionString = "Provider=OraOLEDB.oracle;Data Source=orcl12c;PLSQLRSet=1;UseSessionFormat=true;User Id=claimsassistant_dc8;Password=maximus;"
$con.Open()

$MessageLine = "ConnString=" + $con.ConnectionString
WriteLog $MainLogFile $MessageLine

$MessageLine = "State=" + $con.State
WriteLog $MainLogFile $MessageLine

# -----------------------------------------------------------
# poll folder for files for processing. 
# -----------------------------------------------------------

#cd C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Folder_Input_Work
#cd $folderInputWork

Get-ChildItem  $folderInputWork -Filter *.daxml | 
Foreach-Object{
    $_.FullName
    # -----------------------------------------------------------
    # Retrieve bar code values from file via metadata in XML file
    # -----------------------------------------------------------
    [xml]$content = Get-Content $_.FullName
    $barcode1 =$content.document.filename.barcode1
    $barcode2 = $content.document.filename.barcode2
    $thread = $content.document.filename.thread
    $fileName = $content.document.filename.'#text'
    $fileNameDAXML = $_.FullName
    $fileNameOnly = $_.BaseName

    # $barcode1 # Form Type
    # $barcode2 # case ID 
    # $thread   #

    # determine barcode value based on length, or some other identifier. 
    if ($barcode1.length -gt 6) {
        $formType = $barcode2
        $caseNumber = $barcode1
    }
    else
    {
        $formType = $barcode1
        $caseNumber = $barcode2
    }
    $hasError = $false

    try
    {
        # get case ID based on case number
        $cmd = $con.createcommand()
        $cmd.CommandType = [System.Data.CommandType]::Text
        $cmd.Parameters.Clear()
        $cmd.CommandText = "SELECT case_id FROM case where case_nbr = :ParamCaseNbr"
        $ParamCaseNumber = New-Object System.Data.OleDb.OleDbParameter

        $ParamCaseNumber.OleDbType = 'Varchar'
        $ParamCaseNumber.Direction = 'Input'
        $ParamCaseNumber.ParameterName = ':ParamCaseNbr'
        $ParamCaseNumber.Value =  $caseNumber
        $cmd.Parameters.Add($ParamCaseNumber)
        $CaseID = $cmd.ExecuteScalar()
        if ($CaseID -le 0 ) { 
            # Can't find case, do not process. 
            $hasError = $true 
           $LogMessage = "Thread = " + $thread + ", FileName = " + $fileNameOnly + ", Case Number (barcode2) = " + $caseNumber + ", Form Type (barcode1) = " + $formType + ", Case Number Not Found, Error = " + $hasError
            WriteLog $MainLogFile $LogMessage
        }
        else
        {
            # ---------------------------------------------------------------
            # Perform database function for inserting file into the database
            # ---------------------------------------------------------------

            [Byte[] ]$FileContent  = [System.IO.File]::ReadAllBytes($fileName)

            $Sql = "  insert into image(image_data,image_description) values(:Param1,:Param2) "

            #Set up the parameters for use with the Sql command.
            $Param1 = New-Object System.Data.OleDb.OleDbParameter
            $Param2 = New-Object System.Data.OleDb.OleDbParameter

            $Param1.OleDbType = 'Binary'
            $Param1.Direction = 'Input'
            $Param1.ParameterName = ':Param1'
            $Param1.Value =  $FileContent

            $Param2.OleDbType = 'Varchar'
            $Param2.Direction = 'Input'
            $Param2.ParameterName = ':Param2'
            $Param2.Value = $fileNameOnly # for description, just store filename ..

            #$cmd = $con.createcommand()
            $cmd.CommandType = [System.Data.CommandType]::Text
            $cmd.commandtext = $Sql

            # put [void] to not show anything.. 
            $cmd.Parameters.Clear()
            $cmd.Parameters.Add($Param1)
            $cmd.Parameters.Add($Param2)
            $cmd.ExecuteNonQuery()

            # get image ID back....
            $cmd.Parameters.Clear()
            $cmd.CommandText = "SELECT IMAGE_SEQ.currval AS IdentValue FROM dual"
            $ImageID = $cmd.ExecuteScalar()
            $ImageID

            # Insert record into Supplement....
            $Sql = "INSERT INTO SUPPLEMENT (case_id, image_id,sup_document,sup_text, sup_added_dt, sup_last_update_dt, sup_added_by_user_id, sup_last_updated_by_user_id) "
            $Sql += "values (:CaseID, :ImageID, :SupDocument, :SupText, SYSDATE, SYSDATE, :SupAddByUserID, :SupLastUpdateByUserID)"

            $ParamCaseID = New-Object System.Data.OleDb.OleDbParameter
            $ParamImageID = New-Object System.Data.OleDb.OleDbParameter
            $ParamSupDocument = New-Object System.Data.OleDb.OleDbParameter
            $ParamSupText = New-Object System.Data.OleDb.OleDbParameter
            $ParamAddByUserID = New-Object System.Data.OleDb.OleDbParameter
            $ParamLastUpdateByUserID = New-Object System.Data.OleDb.OleDbParameter

            $ParamCaseID.OleDbType = 'Numeric'
            $ParamCaseID.Direction = 'Input'
            $ParamCaseID.ParameterName = ':CaseID'
            $ParamCaseID.Value =  $CaseID

            $ParamImageID.OleDbType = 'Numeric'
            $ParamImageID.Direction = 'Input'
            $ParamImageID.ParameterName = ':ImageID'
            $ParamImageID.Value = $ImageID

            # file name
            $ParamSupDocument.OleDbType = 'Varchar'
            $ParamSupDocument.Direction = 'Input'
            $ParamSupDocument.ParameterName = ':SupDocument'
            $ParamSupDocument.Value = $fileNameOnly
            #$FileContent

            # description
            $ParamSupText.OleDbType = 'Varchar'
            $ParamSupText.Direction = 'Input'
            $ParamSupText.ParameterName = ':SupText'
            $ParamSupText.Value = $formType

            $ParamAddByUserID.OleDbType = 'Numeric'
            $ParamAddByUserID.Direction = 'Input'
            $ParamAddByUserID.ParameterName = ':SupAddByUserID'
            $ParamAddByUserID.Value =  467 # hard coded User ID

            $ParamLastUpdateByUserID.OleDbType = 'Numeric'
            $ParamLastUpdateByUserID.Direction = 'Input'
            $ParamLastUpdateByUserID.ParameterName = ':SupLastUpdateByUserID'
            $ParamLastUpdateByUserID.Value = 467 # hard coded User ID

            #$cmd = $con.createcommand()
            #$cmd.CommandType = [System.Data.CommandType]::Text
            $cmd.commandtext = $Sql

            # put [void] to not show anything.. 
            $cmd.Parameters.Clear()
            $cmd.Parameters.Add($ParamCaseID)
            $cmd.Parameters.Add($ParamImageID)
            $cmd.Parameters.Add($ParamSupDocument)
            $cmd.Parameters.Add($ParamSupText)
            $cmd.Parameters.Add($ParamAddByUserID)
            $cmd.Parameters.Add($ParamLastUpdateByUserID)
            $cmd.ExecuteNonQuery()

            $LogMessage = "Thread = " + $thread + ", FileName = " + $fileNameOnly + ", Case Number (barcode2) = " + $caseNumber + ", Form Type (barcode1) = " + $formType + ", CaseID = " + $CaseID + ", ImageID = " + $ImageID + ", Error = " + $hasError
            WriteLog $MainLogFile $LogMessage

        }

    }
    catch {
        $LogMessage = "Exception Thrown.  FileName = " + $fileNameOnly + "," + $_.Exception.Message
        WriteLog $MainLogFile $LogMessage
        $hasError -eq $false
    }
    finally {
        # -----------------------------------------------------------
        # Move file to success folder if processed
        # -----------------------------------------------------------
        if ($hasError -eq $false) { 
         move-item $fileName $folderSuccess -Force
         move-item $fileNameDAXML $folderSuccess -Force
         }
        else { 
         move-item $fileName $folderError  -Force
         move-item $fileNameDAXML $folderError  -Force
         }
    }

}

$con.Close()

WriteLog $MainLogFile "Script ended."

