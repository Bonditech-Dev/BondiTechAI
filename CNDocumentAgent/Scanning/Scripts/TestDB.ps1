# using 32-bit oracle driver.. 

function WriteLog ($logFile, $message)
{
  $lineMessage = (Get-Date).ToString() + " - " + $message
  add-content -Path $logFile -Value $lineMessage -Force
}


$MainLogFile = $folderError + "C:\Workbnch\Casenamics\Source\CNDocumentAgent\Scanning\Scripts\testdb.log"

WriteLog $MainLogFile "Script started."

$LogMessage = "Processor = " + $env:Processor_Architecture
WriteLog $MainLogFile $LogMessage


# -------------------------------------------------------------
# Open OleDB connection.. 
# -------------------------------------------------------------
WriteLog $MainLogFile "Opening DB Connection"

$con = New-Object System.Data.OleDb.OleDbConnection
$con.ConnectionString = "Provider=OraOLEDB.oracle;Data Source=orcl;PLSQLRSet=1;UseSessionFormat=true;User Id=claimsassistant_ny;Password=maximus;"
$con.Open()

$MessageLine = "ConnString=" + $con.ConnectionString
WriteLog $MainLogFile $MessageLine

$MessageLine = "State=" + $con.State
WriteLog $MainLogFile $MessageLine


$con.Close()

WriteLog $MainLogFile "Script ended."

