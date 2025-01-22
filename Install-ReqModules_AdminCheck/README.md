PowerShell Script for Module Updates and Import

Overview

This PowerShell script performs the following tasks:

Administrator Privileges Check:

Ensures the script is executed with administrator privileges. If not, it prompts the user to run the script as an administrator and exits.

Module Update Check:

Checks for the latest versions of specific PowerShell modules.

Prompts the user to update modules if a newer version is available.

Module Import:

Imports the required PowerShell modules after ensuring they are updated.

Modules Checked for Updates

The script checks and updates the following PowerShell modules:

MSOnline

AzureAD

Microsoft.Online.SharePoint.PowerShell

ExchangeOnlineManagement

MicrosoftTeams

Microsoft.Graph

How to Use the Script

Open PowerShell as Administrator:

Right-click the PowerShell application and select "Run as Administrator."

Execute the Script:

Copy the script into a .ps1 file (e.g., UpdateModules.ps1).

Run the script in the PowerShell console by entering:

.\UpdateModules.ps1

Follow Prompts:

The script will notify you if a module has a newer version available.

Enter Y to update the module or N to skip the update.

Prerequisites

Administrator Privileges:

Ensure you have administrator privileges to execute the script and install/update modules.

PowerShellGet Module:

Ensure you have PowerShellGet installed to enable module management. Update it if needed using:

Install-Module -Name PowerShellGet -Force

Error Handling

If a module is not found, the script will notify you.

Ensure an active internet connection for the script to check and download updates.

Notes

The -Force flag in Install-Module ensures the module is installed without additional confirmation prompts.

The script is compatible with PowerShell 5.1 and later.

The Import-Module commands at the end ensure all modules are loaded into the session after updates.

License

This script is provided "as-is" without warranty of any kind. Use it at your own risk.

Disclaimer

Ensure you test this script in a controlled environment before using it in production systems.
