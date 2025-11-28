# Connect to Exchange Online PowerShell
# Make sure you have installed the Exchange Online PowerShell module
# If not installed, run: Install-Module -Name ExchangeOnlineManagement
# Then connect using: Connect-ExchangeOnline
# Provide your Office 365 admin credentials when prompted
Connect-ExchangeOnline

$mailbox = "XXX@contoso.com"
#checking size of folders in the mailbox
Get-MailboxFolderStatistics -Identity $mailbox  |Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy,DeletePolicy,CompliancePolicy |out-gridview #export-csv -path 'C:\Users\turekw\OneDrive - FUJITSU\Desktop\arnoudmailbox_folder_statistics.csv' -Delimiter ";" -NoTypeInformation #Out-GridView
#checking size of folders in the mailbox archive
Get-MailboxFolderStatistics -Identity $mailbox -Archive|Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder,`
ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy |out-gridview  #export-csv -path 'C:\Users\turekw\OneDrive - FUJITSU\Desktop\archive_mailbox_folder_statistics.csv' -Delimiter ";" -NoTypeInformation #Out-GridView
 
#checking configuration of LH,Archive,Retention on the mailbox
Get-Mailbox -Identity $mailbox | select-object retentionpolicy,retaindeleteditemsfor, litigationholdenabled, autoexpandingarchiveenabled, archiveguid,archivestatus,retentionholdenabled,`
ElcProcessingDisabled, DelayHoldApplied |Out-GridView #export-csv -Path 'C:\Users\turekw\OneDrive - FUJITSU\Desktop\Robeco_Moraal\retention_details.csv' -Delimiter ";" -NoTypeInformation
