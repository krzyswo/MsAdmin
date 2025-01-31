Connect-ExchangeOnline

$mailbox = "jan.kowlaski@conotos.com"

Get-MailboxFolderStatistics -Identity $mailbox -ResultSize unlimited |Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy,DeletePolicy,CompliancePolicy |out-gridview
