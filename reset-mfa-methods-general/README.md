# Reset MFA Methods for Microsoft 365 Users

This PowerShell script helps to reset MFA methods for M365 users, covering 25+ real time scenarios. For example,

- Reset Phone number authentication from all users,

- Reset MS authenticator method for a specific user (it will be useful when the user lost their device or lost access to the app),

- Reset weaker authentication methods from admin accounts, etc.

**Supported authentication methods:**

1. Email

2. FIDO2

3. Microsoft Authenticator

4. Phone

5. Software OATH

6. Temporary Access Pass

7. Windows Hello for Business

**User scope:**

1. Single user

2. Bulk users (import CSV)

3. All users

4. Admin accounts

5. Guest users

6. Licensed users

7. Disabled users

#### Sample log file:

![Reset MFA for Microsoft 365 users](https://github.com/krzyswo/MsAdmin)

## Microsoft 365 Reporting Tool

For more extensive insights on MFA status, configured CA policies, Sign-ins with MFA, etc., explore our Microsoft 365 reporting tool. It offers 1800+ out-of-the-box reports and smart dashboards.

*Easily manage MFA using our tool: [https://github.com/krzyswo/MsAdmin](https://github.com/krzyswo/MsAdmin)*

Name: reset-mfa-methods-general
Description: This document provides details on resetting MFA methods for Microsoft 365 users.
Prerequisites: PowerShell, Microsoft 365 admin access
Usage: Execute the PowerShell script as per the user scope requirements.
Output: The script resets specified MFA methods for selected users.
Notes: Ensure to have the necessary permissions before running the script.
License: MIT License