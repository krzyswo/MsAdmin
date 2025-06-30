<#
=============================================================================================
Name:           Install and Connect to Exchange Online PowerShell Module
Version:        3.0
Website:        https://github.com/krzyswo/MsAdmin
For detailed Script execution: https://github.com/krzyswo/MsAdmin
============================================================================================
#>

#Due to RPS and Basic Auth retirement in Exchange Online, we need the EXO V3 module
$Module = (Get-Module ExchangeOnlineManagement -ListAvailable) | where {$_.Version.major -ge 3}
if($Module.count -eq 0)
{
 Write-Host 'Exchange Online PowerShell V3 module is not available' -ForegroundColor yellow
 $Confirm= Read-Host 'Are you sure you want to install module? [Y] Yes [N] No'
 if($Confirm -match "[yY]")
 {
 Write-host "Installing Exchange Online PowerShell module"
 Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
 Import-Module ExchangeOnlineManagement
 }
 else
 {
 Write-Host 'EXO V3 module is required to connect to Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet.'
 Exit
 }
}
 
Write-Host 'Connecting to Exchange Online...'
Connect-ExchangeOnline
