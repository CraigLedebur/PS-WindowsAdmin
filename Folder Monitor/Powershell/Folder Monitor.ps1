$ErrorActionPreference="STOP"

# Customizable parameters below
$MonitorFolder = "E:\FTP" # Folder to monitor
$filter = '*.*' # File extension to only look for - by default *.* = all files
$logFile = "E:\FTP\Upload Log.txt" # Path to log file where all transactions will be recorded
$loopWaitTime = 5 # Integer in seconds - time between each iteration of the loop
$MaxFileSize = '8388608' # Maximum attachment size in bytes
# End customizable parameters

function Write-Upload-Log {
    param([string]$logMessage)
    $LogStamp = Get-Date -Format [yyyy-MM-dd` HH:mm:ss]
    Out-File -FilePath $logFile -Append -InputObject "$LogStamp - $logMessage"
    Write-Host "$logMessage" -ForegroundColor green

}

$fileIncomplete = $true

Write-Host "`r`n"
Write-Upload-Log "[ MONITOR STARTING ]"

$ErrorActionPreference="Continue"
$fsMonitorFolder = New-Object IO.FileSystemWatcher $MonitorFolder, $filter -Property @{IncludeSubdirectories = $true; NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'}
$ErrorActionPreference="Stop"

Register-ObjectEvent $fsMonitorFolder Created -SourceIdentifier FolderMonitor -Action {
    $name = $Event.SourceEventArgs.Name
    $timeStamp = $Event.TimeGenerated
    $attachment = "$MonitorFolder\$name"

    Write-Upload-Log "The file '$name' was uploaded at $timeStamp"

    $filePrevSize = 0 # Preps the variable before the start of the loop

    # IO.FileSystemWatcher has an annoying tendency to declare that a file has been created or changed when it is still being uploaded. So the following while loop is a hack to make sure an upload has completed before processing that file
    # 'course, if the connection stalls for more than five seconds or drops entirely, the script will just think that the file has been completed and email accordingly.

    while($fileIncomplete) {

        Start-Sleep -s $loopWaitTime # Pauses script for $loopWaitTime seconds to give the file a chance to upload before checking the size again
        $fileCurSize = Get-Item "$MonitorFolder\$name" # Pulls file info
        $fileCurSize = $fileCurSize.length # Simplifies file info so that it only has the size in bytes

        if([int]$fileCurSize -eq [int]$filePrevSize) {

            $fileIncomplete = $false # Gets us out of the loop once it shows the file has not changed since the last iteration
        }
        $filePrevSize = $fileCurSize # If the loop isn't yet broken, the variables are prepared for the next iteration
    }

    Write-Upload-Log "File $name size is $fileCurSize bytes"

    Send-File-Email $attachment $timeStamp $fileCurSize

    Write-Upload-Log "End of processing for file $name`r`n"

    # Clears variables for the next file
    $fileIncomplete = $true
    $fileCurSize = 0
    $filePrevSize = 0
}

function Send-File-Email {
    param([string]$UploadedFile, $timeStamp, $audioFileLength)
    
    [string]$emailBody = "" # Feel free to make this the email body. It can have basic HTML markups. ie: <b>, <br>, <h1>, <p>, etc.
    
    Write-Upload-Log "File attachment is $UploadedFile"

    $ErrorActionPreference = "Continue"

    Send-MailMessage `
    -From "scriptwizard@foo.com" `
    -To 'bar@foo.com' `
    -Bcc 'foo@bar.com' `
    -Subject "File Uploaded" `
    -Attachments $UploadedFile `
    -BodyAsHtml $emailBody `
    -SmtpServer 'exch-2010.foo.com' 

    $ErrorActionPreference = "Stop"

    Write-Upload-Log "Email sent"
    
}