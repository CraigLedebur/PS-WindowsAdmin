################################################################
### Fortinet FortiGate (and other devices) automated backups ###
### Created by Craig Ledebur                                 ###
### https://github.com/CraigLedebur                          ###
### Committed to GitHub 16 January 2017                      ###
################################################################

# Requirements - Posh-SSH module
# Usage - Run in regular PowerShell
# Edit variables below to suit your environment

# BEGIN EDITABLE VARIABLES
$TranscriptPath = "C:\Scripts\FGBackup.log"
$FGListPath = "C:\Scripts\FGList.txt" # List of firewalls with host names, usernames and passwords, separated by a space, with a newline between entries

$FTPServer = "ftp.foo.bar"
$FTPUsername = "fgbackup"
$FTPPassword = "foobackupbar"

$SMTPServer = "127.0.0.1"
$SMTPFromAddress = "fgbackup@foo.bar"
$SMTPToAddress = "team@foo.bar"
# END EDITABLE VARIABLES

Import-Module Posh-SSH
$ErrorActionPreference = "Stop"
$isError = $false
[string]$HostAction = "<b>Fortigate Backups completed:</b><br><br>"
$HostError = "Fortigate Errors:<br>"

# Stops the transcription process if it wasn't actually stopped during the last run, then starts a new one
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript
$ErrorActionPreference = "Stop"
Start-Transcript -Path $TranscriptPath

# Tests to ensure it can read the list of firewalls
try {
$FGList = Get-Content $FGListPath # Reads entire file into the variable
} catch [System.Management.Automation.ItemNotFoundException] {
    Write-Host "Error: Could not find list of firewalls. Exiting."
    exit
}

ForEach ($FGLine in $FGList) {

    $Fortigate = $FGLine.Split(' ') # Splits each line into segments, separated by a space

    $FGHost = $Fortigate[0] # Hostname
    $FGUser = $Fortigate[1] # Username
    $FGPassword =  ConvertTo-SecureString -String $Fortigate[2] -AsPlainText -Force # Password, encrypted
    $FGDay = Get-Date -Format yyyy-MM-dd # Day the job is run
    $isHostError = $false # Trips if there was an error situation with this host

    # Prepares the credentials to pass onto the SSH server
    $FGCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $FGUser,$FGPassword

    # This will attempt to first connect to the SSH Server on the forti, and then execute a command
    # It will catch multiple exceptions and amend to the error log accordingly, helping
    # us to figure out the cause.
    try {
        New-SSHSession -ComputerName $FGHost -Credential $FGCredentials -Force -ErrorAction Stop
        Invoke-SSHCommand -Command "exec backup full-config ftp $FGHost.$FGDay.cfg $FTPServer $FTPUsername $FTPPassword" -SessionId 0 -EnsureConnection -ErrorAction Stop
    } catch [Renci.SshNet.Common.SshOperationTimeoutException] { 
        # The host is unreachable
        $HostError += "$FGHost - timed out<br>"
        $isHostError = $true
    } catch [Renci.SshNet.Common.SshAuthenticationException] {
        # Credentials are wrong
        $HostError += "$FGHost - incorrect credentials<br>"
        $isHostError = $true
    } catch [System.InvalidOperationException] {
        # Sometimes when it's unreachable, it gives this exception
        $HostError += "$FGHost - timed out<br>"
        $isHostError = $true
    } catch [System.Net.Sockets.SocketException] {
        # A catch-all exception for TCP/IP related errors
        $ErrorMessage = $_.Exception.Message
        $HostError += "$FGHost - $ErrorMessage<br>"
        $isHostError = $true
    } catch {
        # A catch-all for any other exception, will print a complete error message
        $ErrorMessage = $_.Exception
        $HostError += "$FGHost<br>Detailed exception message: $ErrorMessage"
        $isHostError = $true
    } finally {
        Remove-SSHSession 0
        if (!$isHostError) { $HostAction += "$FGHost - successful<br>"}
        elseif ($isHostError) { $isError = $true }
    }

}

Stop-Transcript -ErrorAction SilentlyContinue

$CompleteDate = Get-Date -Format [yyyy-MM-dd` HH:mm:ss]

if ($isError) { 
    # Since an error occurred, it will attach the log for debugging purposes
    Send-MailMessage -SmtpServer $SMTPServer -From $SMTPFromAddress -To $SMTPToAddress -Subject "Fortigate Backup - [ERROR] $CompleteDate" -BodyAsHtml "$HostError<br>$HostAction" -Attachments $TranscriptPath
} else {
    Send-MailMessage -SmtpServer $SMTPServer -From $SMTPFromAddress -To $SMTPToAddress -Subject "Fortigate Backup - [SUCCESS] $CompleteDate" -BodyAsHtml $HostAction
}