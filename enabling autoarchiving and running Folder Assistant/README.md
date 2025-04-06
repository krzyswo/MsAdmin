This PowerShell script automates mailbox management tasks in Exchange Online. It enables autoarchiving, runs the Folder Assistant, and performs cleanup actions if the versions folder is high for a specified mailbox.

Prerequisites
Access to Exchange Online PowerShell module.
Permissions to manage mailboxes in Exchange Online.
PowerShell script execution policy set to allow running scripts.
Output
Autoexpanding archive enabled for the specified mailbox.
Folder Assistant run for the specified mailbox.
If the versions folder is high:
Hold cleanup action initiated.
Full crawl action initiated.
Notes
Ensure that you have the necessary permissions to perform these actions on the specified mailbox.
Verify the mailbox email address before running the script to avoid unintended changes.
Monitor the script execution for any errors or warnings.
License
This script is released under the MIT license.
