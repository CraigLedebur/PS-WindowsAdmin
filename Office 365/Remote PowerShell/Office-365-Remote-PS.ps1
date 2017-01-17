# A very short script to make connecting to Office365 via PowerShell a bit less annoying :-)
# Craig Ledebur

try {
    $UserCredential = Get-Credential
} catch [System.Management.Automation.ParameterBindingException] {
    Write-Host "Credentials are required to log in. Aborting." -ForegroundColor Red
    break
}

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
Connect-MsolService -Credential $UserCredential

Write-Host "Please do 'Remove-PSSession $Session' when you are finished." -ForegroundColor Green