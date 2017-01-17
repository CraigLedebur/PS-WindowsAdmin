################################################################
### Exchange 201x - Email Alias Search the Fast and Easy Way ###
### Created by Craig Ledebur                                 ###
### https://github.com/CraigLedebur                          ###
### Committed to GitHub 16 January 2017                      ###
################################################################

# Requirements - Exchange 2010 SP3 or newer
# Usage - Run in Exchange Management Shell
# to output to console:
# > & '.EmailAliasSearch.ps1'
# to output to text file and console:
# > & '.EmailAliasSearch.ps1 -OutFile <filename.txt>'

param (
    [string]$OutFile = ''
)

$searchTerm = Read-Host -Prompt "`r`nPlease enter search string (wildcards accepted)"
$searchScope = Read-Host -Prompt 'Search (M)ailboxes, (C)ontacts, (D)istribution Groups or (A)ll?'
$LogPath = $OutFile


# Sanity check - makes sure that the file is writable. If not, it disables logging to file.
if($LogPath.Length -ne 0) { 
        
    try {
        Write-Host "Exchange 201x - Email Alias Search`r`n" | Out-File -FilePath $LogPath -Append
    } catch [System.Management.Automation.DriveNotFoundException],[System.IO.IOException] {
        Write-Host("Invalid log file path. Will only output to console.") -ForegroundColor Red
        $LogPath = ''
    }
}


function Display-Output([string]$LineOutput) {
    Write-Host($LineOutput)

    if($LogPath.Length -ne 0) { 
        $LineOutput | Out-File -FilePath $LogPath -Append
    }
}

try {
    if (($searchScope -eq 'm') -or ($searchScope -eq 'M') -or ($searchScope -eq 'a') -or ($searchScope -eq 'A')) { 
    
       Display-Output("`r`nSearching mailboxes:")
       $Mailboxes = Get-Mailbox -result unlimited
       $Mailboxes | foreach {

       for ($i=0;$i -lt $_.EmailAddresses.Count; $i++) {
           $currentUser = $_.Alias
           $address = $_.EmailAddresses[$i]

           if ($address.SmtpAddress -like $searchTerm ) {
               Display-Output('User: ' + $currentUser + " has address: " + $address.AddressString.ToString())
               }
           }
       }
    } elseif (($searchScope -eq 'c') -or ($searchScope -eq 'C') -or ($searchScope -eq 'a') -or ($searchScope -eq 'A')) { 

       Display-Output("`r`nSearching contacts:")
       $Mailboxes = Get-MailContact -result unlimited
       $Mailboxes | foreach {

       for ($i=0;$i -lt $_.EmailAddresses.Count; $i++) {
           $currentUser = $_.Alias
           $address = $_.EmailAddresses[$i]

           if ($address.SmtpAddress -like $searchTerm ) {
               Display-Output('Contact: ' + $currentUser + " has address: " + $address.AddressString.ToString())
               }
           }
       }
    } elseif (($searchScope -eq 'd') -or ($searchScope -eq 'D') -or ($searchScope -eq 'a') -or ($searchScope -eq 'A')) { 

       Display-Output("`r`nSearching distribution groups:")
       $Mailboxes = Get-DistributionGroup -result unlimited
       $Mailboxes | foreach {

       for ($i=0;$i -lt $_.EmailAddresses.Count; $i++) {
           $currentUser = $_.Alias
           $address = $_.EmailAddresses[$i]

           if ($address.SmtpAddress -like $searchTerm ) {
               Display-Output('Distribution group: ' + $currentUser + " has address: " + $address.AddressString.ToString())
               }
           }
       }

    } else {
        Display-Output("Unrecognized command. Exiting.")
    }

# 
} catch [System.Management.Automation.CommandNotFoundException] {
    Display-Output("Could not execute Get-Mailbox command. Did you import the Exchange modules for Powershell?")
}