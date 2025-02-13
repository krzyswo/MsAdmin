# Connect to Exchange Online
Connect-ExchangeOnline

# Define Variables
$Mailbox = "user@domain.com"
$DelegateUser = "delegate@domain.com"

#  Get Calendar Processing Settings
Get-CalendarProcessing -Identity $Mailbox | Format-List

#  Set Calendar Processing (Example: Enable AutoAccept)
Set-CalendarProcessing -Identity $Mailbox -ResourceDelegates $DelegateUser



#########################################################################################


#  Get Mailbox Folder Permissions
Get-MailboxFolderPermission $Mailbox:\Calendar

#  Add Mailbox Folder Permission (Example: Reviewer)
Add-MailboxFolderPermission -Identity $Mailbox:\Calendar -User $DelegateUser -AccessRights Reviewer

#  Set Mailbox Folder Permission (Modify Existing User Permissions)
Set-MailboxFolderPermission -Identity $Mailbox:\Calendar -User $DelegateUser -AccessRights Editor

#  Remove Mailbox Folder Permission for a Specific User
Remove-MailboxFolderPermission -Identity $Mailbox:\Calendar -User $DelegateUser
