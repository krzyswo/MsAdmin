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
