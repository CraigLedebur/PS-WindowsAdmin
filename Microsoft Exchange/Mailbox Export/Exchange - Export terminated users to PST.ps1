$ErrorActionPreference="SilentlyContinue"
# Loads Powershell modules needed for this operation, will take a while
# You may need to change the path if you're not using Exchange 2010
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
Connect-ExchangeServer -auto
Import-Module ActiveDirectory

# Change variables below
$ou = "OU=Terminated,OU=Companies,DC=company,DC=LOCAL" # OU on Active Directory to process
$PSTDestination = "\\COMPANY-FS-01\Archives\Mailboxes\" # Path to export the PSTs to
$logFolder = "C:\Mailbox Export Logs\" # Change to the folder you wish to export the logs to. Please don't forget the trailing slash
# End variables

# Makes sure no transcript is already running
Stop-Transcript | out-null

$ErrorActionPreference = "Continue"

# Sets the base timestamp each log file will be in
$startDate = Get-Date -Format "dd-MM-yyyy HHmm"

# Sets up the various log files
$logfile = $logFolder + "MailboxExport-" + $startDate + ".log"
$logQueued = $logfolder + "MailboxExport-" + $startDate + ".Queued.log"
$logExported = $logfolder + "MailboxExport-" + $startDate + ".AlreadyExported.log"
$logInProgress = $logfolder + "MailboxExport-" + $startDate + ".InProgress.log"

#Runs a transcript of the export job
Start-Transcript -path $logfile


# Gets a list of users from that OU along with the required columns
$user = Get-ADUser -SearchBase $ou -Filter {EmailAddress -like "*"} -Properties * | Select-Object SamAccountName,GivenName,Surname,Name

# Loops through every element in $user, exporting each element to $currentUser
foreach($currentUser in $user)
{
	$userName = ($currentUser.GivenName) + "." + ($currentUser.Surname)
	$exportReq = Get-MailboxExportRequest -Mailbox $currentUser.Name
	
	# Destination folder for PSTs
	$pst = "$PSTDestination\$userName.pst"
	
		# Checks to make sure the export request isn't already in the queue
	if ($exportReq.Status -eq "Queued") {
		Write-Host "Already in queue: " + $username + " at " + $pst
		echo $userName >> $logQueued
		
		# Checks if export request is already in progress - if so, skips
	} elseif ($exportReq.Status -eq "InProgress") {
		Write-Host "Already in progress: " + $username + " at " + $pst
		echo $userName >> $logInProgress
		
		# Checks if destination file exists (which implies request has completed) - if so, skips
	} elseif (Test-Path $pst) {
		  Write-Host "Already exported: " + $username + " at " + $pst
		  echo $userName >> $logExported
	}Else{
		  New-MailboxExportRequest -Mailbox $currentUser.Name -FilePath $pst
	}
}

Write-Host "End of export requests"
Stop-Transcript