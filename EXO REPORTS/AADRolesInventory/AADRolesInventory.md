Azure AD Role Inventory

Description
This PowerShell script collects a list of all active roles in the Azure Active Directory (Azure AD) tenant and gathers information about users and service principals assigned to each role. It then exports this information to a CSV file for inventory purposes.

Prerequisites
AzureAD PowerShell module must be installed. If not installed, download it from AzureAD PowerShell Gallery.
Permissions to access and retrieve information from Azure AD.
Internet connectivity to download the AzureAD module if not already installed.
Output
The script will generate an Azure AD role inventory CSV file that includes the following information for each user or service principal assigned to a role:

User Principal Name
Display Name
Assigned Role(s)
Notes
Ensure that you have the necessary permissions to access Azure AD and retrieve role information.
If the AzureAD module is not available, download it from the provided link before running the script.
The script will attempt to establish connectivity to Azure AD if not already connected.
The generated CSV file will be named with a timestamp for reference.
License
