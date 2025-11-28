<#
=============================================================================================

Name         : Audit Entra App Operations Using PowerShell  
Version      : 1.0
website      : https://github.com/krzyswo/MsAdmin

-----------------
Script Highlights
-----------------

1. Tracks all Entra app operations for the past 180 days.
2. Allows to track app operations for a custom date range.
3. The script automatically verifies and installs the Exchange Online PowerShell V3 module (if not installed already) upon your confirmation.
4. Enables filtering of app operations from the following categories.
    -> Added Applications
    -> Updated Applications
    -> Deleted Applications
    -> Consent to Applications
    -> OAuth2 Permission Grants
    -> App Role Assignments
    -> Service Principal Changes
    -> Credential Changes
    -> Delegation Changes
5. Tracks app operations performed on the specific application.
6. Generates a report that retrieves successful operations alone.
7. Helps export failed operations alone.
8. Audit app operations performed by a specific user.
10. The script can be executed with an MFA-enabled account too.
11. Exports report results as a CSV file.
12. The script is scheduler friendly.
13. It can be executed with certificate-based authentication (CBA) too.  

For detailed Script execution: https://github.com/krzyswo/MsAdmin
============================================================================================
#