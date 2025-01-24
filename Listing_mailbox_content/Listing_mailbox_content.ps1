# Connect to Exchange Online
Connect-ExchangeOnline

$mailbox="XXX"

Get-MailboxFolderStatistics -Identity $mailbox |Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy |out-gridview
