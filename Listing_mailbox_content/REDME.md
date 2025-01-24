# README: Exchange Online Mailbox Folder Statistics Script

## Overview
This PowerShell script connects to Exchange Online and retrieves detailed statistics for all folders within a specified mailbox. The output includes various folder-related metrics and is displayed in an interactive grid view for easy analysis.

## Prerequisites
1. **Microsoft Exchange Online Module**: Ensure the Exchange Online PowerShell module is installed. You can install it using:
   ```powershell
   Install-Module -Name ExchangeOnlineManagement
   ```

2. **Permissions**: You need appropriate permissions to access mailbox statistics. For example, you must be assigned the "Mailbox Import Export" role.

3. **PowerShell Version**: Ensure you are running PowerShell 5.1 or later.

## Script Usage

### Parameters
- **$mailbox**: Replace `XXX` in the script with the email address of the mailbox you want to analyze.

### Instructions
1. Open PowerShell.
2. Run the script with the following steps:
   - Connect to Exchange Online:
     ```powershell
     Connect-ExchangeOnline
     ```
   - Execute the script.

3. The script retrieves mailbox folder statistics, including:
   - **Folder Name**
   - **Folder Path**
   - **Folder Type**
   - **Storage Quota**
   - **Storage Warning Quota**
   - **Visible Items in Folder**
   - **Hidden Items in Folder**
   - **Total Items in Folder**
   - **Folder Size**
   - **Items in Folder and Subfolders**
   - **Folder and Subfolder Size**
   - **Archive Policy**



## Troubleshooting
1. **Access Denied**: Verify your permissions to view mailbox statistics.
2. **Module Not Found**: Reinstall the Exchange Online module if missing.
   ```powershell
   Install-Module -Name ExchangeOnlineManagement -Force
   ```

## License
This script is provided "as-is" without warranty of any kind. Use it at your own risk.

---

For further assistance, refer to the [Microsoft Exchange Online documentation](https://learn.microsoft.com/en-us/powershell/exchange/).

