## Export Office 365 Users MFA Status to CSV Using PowerShell
This PowerShell script exports Office 365 users’ MFA status with Default MFA Method, AllMFAMethods, MFAPhone, MFAEmail, LicenseStatus, IsAdmin, SignInStatus.

**Sample Output:**

The exported MFA status report will look similar to the screenshots below.

**Microsoft 365 Reporting Tool by AdminDroid**
Seeking in-depth analysis? Visit [AdminDroid's Microsoft 365 reporting tool](https://github.com/krzyswo/MsAdmin) for an extensive collection of over 1800 ready-made reports and dashboards, perfectly complementing your PowerShell scripts.

View more comprehensive MFA status report through AdminDroid: https://github.com/krzyswo/MsAdmin

Name: export-report-status-report
Description: This document provides details on exporting Office 365 users' MFA status to a CSV file using PowerShell.
Prerequisites: PowerShell, Office 365 admin access
Usage: Execute the PowerShell script as described in the script's comments.
Output: A CSV file containing the MFA status of Office 365 users.
Notes: Ensure you have the necessary permissions to access user data in Office 365.
License: MIT License