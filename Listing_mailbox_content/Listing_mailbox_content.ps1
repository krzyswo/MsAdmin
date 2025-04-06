# Connect to Exchange Online
Connect-ExchangeOnline

$mailbox="XXX"

Get-MailboxFolderStatistics -Identity $mailbox |Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy |out-gridview


# Connect to Exchange Online PowerShell
# Make sure you have installed the Exchange Online PowerShell module
# If not installed, run: Install-Module -Name ExchangeOnlineManagement
# Then connect using: Connect-ExchangeOnline
# Provide your Office 365 admin credentials when prompted
Connect-ExchangeOnline

get-mailbox -Identity $userEmailAddress

# Replace "user@example.com" with the actual user's email address
$userEmailAddress = "XXX@contoso.com"

# Replace "Inbox" with the folder you want to check (e.g., "SentItems", "DeletedItems", etc.)
$mailboxFolder = "Inbox"

# Get the mailbox folder content
$folderContent = Get-MailboxFolderStatistics -Identity $userEmailAddress -FolderScope $mailboxFolder

# Display the folder content in Out-GridView
$folderContent | Select-Object ItemsInFolder, FolderSize, Name
$folderContent | Select-Object ItemsInFolder, FolderSize, Name | Out-GridView

# Display whole mailbox content in Out-GridView

Get-MailboxFolderStatistics -Identity $userEmailAddress |Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy |out-gridview


Search-MailboxAuditLog -Identity "XXX@contoso.com" -Operations Move -ShowDetails | fl

###################################################################################################################
$mailbox = "XXX@contoso.com"
#checking size of folders in the mailbox
Get-MailboxFolderStatistics -Identity $mailbox  |Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy,DeletePolicy,CompliancePolicy |out-gridview #export-csv -path 'C:\Users\turekw\OneDrive - FUJITSU\Desktop\arnoudmailbox_folder_statistics.csv' -Delimiter ";" -NoTypeInformation #Out-GridView
#archive
Get-MailboxFolderStatistics -Identity $mailbox -Archive|Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy |out-gridview  #export-csv -path 'C:\Users\turekw\OneDrive - FUJITSU\Desktop\archive_mailbox_folder_statistics.csv' -Delimiter ";" -NoTypeInformation #Out-GridView
 
#checking configuration of LH,Archive,Retention on the mailbox
Get-Mailbox -Identity $mailbox | select-object retentionpolicy,retaindeleteditemsfor, litigationholdenabled, autoexpandingarchiveenabled, archiveguid,archivestatus,retentionholdenabled,`
ElcProcessingDisabled, DelayHoldApplied |Out-GridView #export-csv -Path 'C:\Users\turekw\OneDrive - FUJITSU\Desktop\Robeco_Moraal\retention_details.csv' -Delimiter ";" -NoTypeInformation
 
#enabling autoarchiving and running Folder Assistant
Enable-Mailbox $mailbox -AutoExpandingArchive
Start-ManagedFolderAssistant -Identity $mailbox
 
#if versions folder is high
Start-ManagedFolderAssistant -Identity $mailbox -HoldCleanup
Start-ManagedFolderAssistant -Identity $mailbox -FullCrawl


Get-MailboxCalendarConfiguration -Identity "XXX@contoso.com" | FL
