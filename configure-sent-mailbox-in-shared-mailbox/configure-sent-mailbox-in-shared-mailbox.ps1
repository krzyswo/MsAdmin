<#
=============================================================================================
Name:          Configure Sent Items in Shared Mailbox using PowerShell
Description:   This script configures shared mailboxes to save copies of emails in the mailbox's sent items folder instead of the user's default folder.
Version:       1.0
Website:       https://github.com/krzyswo/MsAdmin

For details: https://github.com/krzyswo/MsAdmin
=============================================================================================
#>

 #Check for EXO module installation
 $Module = Get-Module ExchangeOnlineManagement -ListAvailable
 if($Module.count -eq 0) 
 { 
  Write-Host "Exchange Online PowerShell module is not available"  -ForegroundColor yellow  
  $Confirm= Read-Host "Are you sure you want to install module? [Y] Yes [N] No" 
  if($Confirm -match "[yY]") 
  { 
   Write-host "Installing Exchange Online PowerShell module"
   Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
   Import-Module ExchangeOnlineManagement
  } 
  else 
  { 
   Write-Host "EXO module is required to connect Exchange Online. Please install module using Install-Module ExchangeOnlineManagement cmdlet." 
   Exit
  }
 } 
 Write-Host "Connecting to Exchange Online..."
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($UserName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline
 }

#Retrieve all the shared mailbox and configure to save email copies in sent items
Get-Mailbox –ResultSize Unlimited -RecipientTypeDetails Sharedmailbox | foreach { 
 Set-mailbox -Identity $_.UserPrincipalName -MessageCopyForSendOnBehalfEnabled $true -MessageCopyForSentAsEnabled $true
}