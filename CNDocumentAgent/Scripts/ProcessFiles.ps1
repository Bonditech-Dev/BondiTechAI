#  -----------------------------------------------------------------------------------
# CASENAMICS BARCODE ROUTER Script
#
# This assumes you have working knowledge of Powershell programming.
# This script will run as 32-bit powershell script. 
#
# Modify this script to process the documents that were scanned. 
#
# The BarCode Router will create a config file called config.daxml in the "Scripts" folder. 
# This config file is read so we can get the input and output folders for processing the files. 
#
# This configuration file contains the location of the following:
# Input (where documents are scanned to),
# InputWork (where documents are split and the metadata file is created with the barcode values), 
# OutputError (where documents that were not processed successfully moved to)
# and OutputSuccess (where documents processed successfully are moved to).
#
# In this script, files processed successfully are moved to OutputSuccess.
# Files with errors are moved to OutputError. 
#
# Professional Services are available through Casenamics for customization services.
# Contact us at sales@casenamics.com.
# -----------------------------------------------------------------------------------

function WriteLog ($logFile, $message)
{
  $lineMessage = (Get-Date).ToString() + " - " + $message
  add-content -Path $logFile -Value $lineMessage -Force
}


# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# TODO: The folder needs to change to the folder this script is saved to
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
cd C:\Casenamics\Scanning\Scripts

# -------------------------------------------------------------
# get configuration information (i.e. where to put files after 
# success or error. Config.daxml should be located in the Scripts folder above.
# -------------------------------------------------------------

[xml]$config = Get-Content config.daxml
$folderError = $config.configuration.folderOutputError
$folderSuccess = $config.configuration.folderOutputSuccess
$folderInputWork = $config.configuration.folderInputWork
$folderInput = $config.configuration.folderInput

$MainLogFile = $folderError + "\ProcessFiles.log"

WriteLog $MainLogFile "Script started."

$LogMessage = "Processor = " + $env:Processor_Architecture
WriteLog $MainLogFile $LogMessage


WriteLog $MainLogFile "Get Files.."

# -----------------------------------------------------------
# poll folder for files for processing. 
# -----------------------------------------------------------
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

    $hasError = $false
    try
    {
		$LogMessage = "Processing FileName = " + $fileNameOnly + ", barcode1 = " + $barcode1 + ", barcode2 = " + $barcode2
        WriteLog $MainLogFile $LogMessage
		# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
		# TODO:
		# Add your code here to process the file based on the barcode
		# If you need assistance writing this script, contact support@casenamics.com
		# for professional services. 
		# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
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

WriteLog $MainLogFile "Script ended."

