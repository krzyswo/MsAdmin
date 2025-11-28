# Mailbox Check Script (Safe/Read-Only Version)  
# ---------------------------------------------  
&nbsp;  
# Define mailbox and trustee  
$mailbox = "user@example.com"  
$trustee = "delegate@example.com"  
&nbsp;  
# Quota check  
Get-Mailbox -Identity $mailbox | Format-List *send*  
Get-Mailbox -Identity $mailbox | Format-List *quota*  
Get-Mailbox -Identity $mailbox -Archive | Format-List *quota*  
Get-Mailbox -Identity $mailbox | Format-List *send*, *receive*  
&nbsp;  
# Size check  
Get-MailboxStatistics -Identity $mailbox | Format-List *size*  
Get-MailboxStatistics -Identity $mailbox -Archive | Format-List *size*  
&nbsp;  
# Folder check  
Get-MailboxFolderStatistics -Identity $mailbox -ResultSize Unlimited |  
Select-Object Name, FolderPath, FolderType, StorageQuota, StorageWarningQuota, VisibleItemsInFolder, HiddenItemsInFolder, ItemsInFolderAndSubfolders, FolderSize, ItemsInFolderAndSubfolderSize |  
Out-GridView  
&nbsp;  
Get-MailboxFolderStatistics -Identity $mailbox -Archive -ResultSize Unlimited |  
Select-Object Name, FolderPath, FolderType, StorageQuota, StorageWarningQuota, VisibleItemsInFolder, HiddenItemsInFolder, ItemsInFolderAndSubfolders, FolderSize, ItemsInFolderAndSubfolderSize |  
Out-GridView  
&nbsp;  
(Get-MailboxFolderStatistics -Identity $mailbox -Archive -ResultSize Unlimited | Select-Object Name).count  
&nbsp;  
# Forwarding check  
Get-Mailbox -Identity $mailbox | Format-List *forward*  
&nbsp;  
# Inbox rules check (non-destructive)  
# Get-InboxRule -Mailbox $mailbox -Identity "Doorsturen naar privé" -IncludeHidden | Remove-InboxRule  
Get-InboxRule -Mailbox $mailbox -IncludeHidden | Format-List * | Out-GridView  
Get-InboxRule -Mailbox $mailbox -IncludeHidden | Format-List  
Get-InboxRule -Mailbox $mailbox -IncludeHidden | Select Enabled, Name, Description | Out-GridView  
&nbsp;  
# Permissions - FullAccess (read only)  
Get-MailboxPermission -Identity $mailbox | Select-Object Identity, AccessRights, User  
# Add-MailboxPermission -Identity $mailbox -User $trustee -AccessRights FullAccess -AutoMapping $false  
# Remove-MailboxPermission -Identity $mailbox -User $trustee -AccessRights FullAccess  
&nbsp;  
# Permissions - SendAs (read only)  
$trustees = Get-RecipientPermission -Identity $mailbox |  
    Where-Object { $_.trustee -ne "NT AUTHORITY\SELF" } |  
    Select-Object Identity, Trustee, AccessRights  
&nbsp;  
# Permissions - SendOnBehalf (read only)  
Get-Mailbox -Identity $mailbox | Select-Object GrantSendOnBehalfTo  
# Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @($remove = $trustee)  
# Set-Mailbox -Identity $mailbox -GrantSendOnBehalfTo @($add = $trustee)  
&nbsp;  
# Permissions - Calendar (read only)  
Get-MailboxFolderPermission -Identity "$mailbox:\calendar" | Select User  
&nbsp;  
# Add-MailboxFolderPermission -Identity "$mailbox:\calendar" -User $trustee -AccessRights Editor -SharingPermissionFlags Delegate  
# Add-MailboxFolderPermission -Identity "$mailbox:\calendar" -User $trustee -AccessRights LimitedDetails  
# Set-MailboxFolderPermission -Identity "$mailbox:\calendar" -User Default -AccessRights AvailabilityOnly  
