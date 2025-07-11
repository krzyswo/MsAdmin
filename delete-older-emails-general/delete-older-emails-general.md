Name: delete-older-emails-general
Description: This PowerShell script deletes older emails in Outlook after a specified number of days using the Search-Mailbox and DeleteContent commands.
Prerequisites: PowerShell, Outlook, appropriate permissions to run Search-Mailbox and DeleteContent commands.
Usage: Execute the script with required parameters to specify the number of days after which emails should be deleted. Ensure Outlook is configured correctly.
Output: The script exports a CSV file detailing the emails that were deleted.
Notes: Ensure to backup important emails before running this script as the deletion process is irreversible.
License: MIT License