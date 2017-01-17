Write-Host "Automated Exchange 2016 Installer"
Write-Host "For Windows 2016 and Exchange 2016 CU3"
Write-Host "The purpose of this script is to make it easier to install the prerequisites"
Write-Host "**************************************"
$ExchangeInstallPath = Read-Host "Please enter the full path to the Exchange 2016 CU3 installation folder"

$ExchangeSetup = "$ExchangeInstallPath\setup.exe"

try {
    Get-ChildItem $ExchangeSetup
} catch {
    Write-Host $_.Error
    exit
}

Write-Host "Did you install this first? http://go.microsoft.com/fwlink/?LinkId=260990"
Write-Host "If you didn't, go ahead. I'll wait!"
pause

# Installs Exchange 2016 Prerequisites for Windows 2016
Install-WindowsFeature NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering,RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation, RSAT-ADDS
# Prepares forest
.\setup.exe /PrepareAD /IAcceptExchangeServerLicenseTerms
# Prepares domain
.\setup.exe /PrepareDomain /IAcceptExchangeServerLicenseTerms