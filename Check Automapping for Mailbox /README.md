Check Automapping for Mailbox in Exchange Online
This script uses the Get-MailboxPermission cmdlet to check if automapping is enabled for a specified mailbox in Exchange Online.

Example Command
powershell
Kopiuj
Edytuj
(Get-MailboxPermission shared1@off365pro4.onmicrosoft.com -ReadFromDomainController)[0].DelegateListLink
What It Does
Command Explanation:
Retrieves mailbox permissions for the specified mailbox (shared1@off365pro4.onmicrosoft.com).
Checks the DelegateListLink property of the first result, which helps identify if automapping is enabled.
Requirements
Exchange Online PowerShell Module.
Appropriate permissions to access mailbox settings.
Notes
Replace shared1@off365pro4.onmicrosoft.com with the mailbox you want to check.
This script assumes basic familiarity with PowerShell and mailbox permissions in Exchange Online.
