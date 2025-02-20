Grant Teams App Permission Policy to Users in Azure AD Group

Description

This PowerShell script connects to Azure AD and Microsoft Teams to grant a specified Teams app permission policy to users in a specific Azure AD group.

Prerequisites

Access to Azure AD and Microsoft Teams PowerShell modules.
Permission to manage app permission policies in Teams.
PowerShell environment with required modules installed.

Output

The script retrieves user UPNs from a designated Azure AD group, loops through each user, and assigns the specified Teams app permission policy to them.

Notes

Update the $policyName variable with the desired policy name.
Replace the Azure AD group ObjectId "cc830891-5585-4889-bf4a-61c4c55ad51c" with the actual group ObjectId.
Verify user permissions and policy settings before executing the script.

License

This script is released under the MIT license.
