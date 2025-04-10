Mailbox Check Script (Safe/Read-Only Version)

Description

This PowerShell script performs various checks on a specified mailbox in a read-only manner. It checks mailbox quotas, sizes, folders, forwarding settings, inbox rules, and permissions related to FullAccess, SendAs, SendOnBehalf, and Calendar access.

Prerequisites

Access to Exchange Online PowerShell session.
Permission to read mailbox properties and permissions.
PowerShell environment with Exchange Online module installed.

Output

The script retrieves and displays information about the specified mailbox, including quota details, size statistics, folder statistics, forwarding settings, inbox rules, and various permissions related to the mailbox.

Notes

Replace "user@example.com" and "delegate@example.com" with the actual mailbox and trustee email addresses.
This script is designed to be read-only and does not make any changes to the mailbox or its settings.
Some sections of the script are commented out to prevent accidental modifications. Uncomment with caution if needed.
Ensure proper permissions are in place to access and retrieve mailbox information.

License

This script is released under the MIT license.
