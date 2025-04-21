# PowerShell Script to Check and Update Modules

## Overview
This PowerShell script checks if the script is run with administrator privileges, lists several Microsoft modules, and compares the current installed version with the latest available version. If a newer version is found, the script prompts the user to install the update. It also imports the required modules for further use.

## Prerequisites
- **Administrator Privileges:** The script requires administrator privileges to function properly.
- **PowerShell Version:** Ensure that you are running a compatible version of PowerShell.
- **Installed Modules:** This script checks and updates the following modules:
  - MSOnline
  - AzureAD
  - Microsoft.Online.SharePoint.PowerShell
  - ExchangeOnlineManagement
  - MicrosoftTeams
  - Microsoft.Graph

## Script Workflow
1. **Check for Administrator Privileges:**  
   The script will check if it is running with administrator privileges. If not, a warning is shown, and the script exits.

2. **Check for Available Module Updates:**  
   For each module in the list, the script checks if a newer version is available compared to the current version installed on the system. If an update is available:
   - The current and latest version numbers are displayed.
   - You are prompted to confirm if you want to install the update (Y/N).

3. **Module Installation:**  
   If you choose to update a module, it is installed using the `Install-Module` cmdlet.

4. **Import Modules:**  
   After the updates (if any), the script imports all the necessary modules to ensure they are available for further use.
