PowerShell script connecting to Exchange Online to retrieve mailbox folder statistics, exporting results to a grid view for analysis.

Requirements:

Microsoft Exchange Online PowerShell Module (Install-Module -Name ExchangeOnlineManagement)
Global Reader/Mailbox Admin permissions in Microsoft 365
Usage:

Run script in PowerShell
Authenticate with Exchange Online (Modern Auth)
Statistics for specified mailbox will display in interactive grid
Output Fields:
Name, FolderPath, FolderType, StorageQuota, ItemsInFolder, FolderSize, CompliancePolicy, and 10+ additional metrics

License:
MIT License
