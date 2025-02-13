Connect-ExchangeOnline
$mailbox = "PLACEHOLDER"
(Get-MailboxPermission $mailbox -ReadFromDomainController)[0].DelegateListLink



# Replace these variables with the target mailbox and user
$Mailbox2 = "PLACEHOLDER"  # The shared or target mailbox
$User = "PLACEHOLDER"      # The user you are checking

# Get permissions for the mailbox
$permissions = Get-MailboxPermission -Identity $Mailbox2 | Where-Object { $_.User -like $User }

if ($permissions) {
    # AutoMapping is enabled by default; check if explicitly disabled
    Write-Host "Permissions exist for $User on $Mailbox2. Check AutoMapping status via delegation assignment or manual configuration."
    Write-Host "To verify if AutoMapping is working, ask the user if the mailbox automatically shows in their Outlook."
} else {
    Write-Host "No permissions found for $User on $Mailbox."
}
