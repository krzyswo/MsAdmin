# Exchange Mailbox Capacity Management Tool

## Description
This PowerShell script provides a graphical user interface (GUI) for managing Exchange Online mailboxes. It allows administrators to perform various mailbox-related tasks, such as checking retention policies, enabling In-Place Archives, managing folder statistics, and troubleshooting archive settings.

## Features
- Check mailbox properties such as Litigation Hold, Retention Policy, and Auto Expanding Archive status.
- View mailbox folder statistics and archive folder statistics.
- Enable or verify In-Place Archive and Auto Expanding Archive settings.
- Start Managed Folder Assistant for mailboxes, with options for Hold Cleanup or Full Crawl.
- Displays detailed warnings for misconfigurations (e.g., missing archives).
- Provides an option to view results in `OutGridView` for better analysis.

## Prerequisites
1. **PowerShell Environment**: Ensure you have PowerShell installed and updated.
2. **Modules**: Install and import the `ExchangeOnlineManagement` module:
   ```powershell
   Install-Module -Name ExchangeOnlineManagement -Force
Usage
1. Launch the Tool
Run the script in a PowerShell console

powershell
Copy
Edit
.\ExchangeMailboxTool.ps1
2. Interface Overview
Text Box: Enter the mailbox identity (e.g., user@example.com).
Buttons: Execute specific actions or checks:
Check LH, Archive, Retention: Displays mailbox properties related to litigation hold, retention policies, and archive status.
Check Folder Statistics: Displays folder-level statistics for the selected mailbox.
Check Archive Folder Statistics: Displays statistics for the archive folder of the mailbox.
Check if In-Place Archive is Enabled: Checks and confirms the status of In-Place Archive.
Enable Auto Expanding Archive: Enables Auto Expanding Archive if In-Place Archive is active.
Start Managed Folder Assistant: Initiates the Managed Folder Assistant for the mailbox.
Enable In-Place Archive: Enables In-Place Archive if not already active.
3. Troubleshooting and Warnings
The tool provides:

Red warnings for:
Missing archives.
Auto Expanding Archive or In-Place Archive not being enabled.
Suggestions for corrective actions in the output fields.
Known Issues
The message for enabling Auto Expanding Archive currently displays only the mailbox name instead of the full success message. Fix is required.
Ensure email addresses with leading/trailing spaces do not break the script.
Planned Enhancements
Add static fields for checks to prevent refreshing after every click.
Provide a toggle for viewing OutGridView results.
Verify the necessity of all current checks.
Contact
For any issues or enhancements, please create an issue in the repository or contact the script maintainer.

Disclaimer: Use this tool at your own risk. Ensure you understand the impact of all operations performed using this script.

vbnet
Copy
Edit

Let me know if further adjustments are needed!
