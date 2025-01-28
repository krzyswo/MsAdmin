Transport Rules Checker
 

Description
The Transport Rules Checker is a PowerShell script that allows you to check the configuration and status of transport rules in your Exchange Online environment. It connects to Exchange Online using the provided user principal name (UPN) and displays the progress during the connection process. The script retrieves all transport rules and provides a detailed report containing their names, descriptions, priorities, enabled status, actions, conditions, exceptions, scopes, and comments.

Prerequisites
Before using this script, ensure that you have the following:

Windows PowerShell 5.1 or later
Exchange Online subscription
User principal name (UPN) for the account with Exchange Online administrative access
Usage
Follow the steps below to execute the script:

Open a PowerShell console.
Run the following command to connect to Exchange Online:
Connect-ExchangeOnline -UserPrincipalName "jankowlski@contoso.nl" -ShowProgress $true  
Replace "jankowlski@contoso.nl" with the actual user principal name (UPN) of the account you want to use for the connection.
The script will establish a connection to Exchange Online and display the progress during the connection process.
The script will retrieve all transport rules and display their details in a grid view.
Report Generation
To generate a report for the transport rules, including client and creation information, as well as change logs for subsequent reports, follow these steps:

Create a separate folder for each client in a common location to store the reports.
Modify the script to include the necessary logic for retrieving client and creation information, as well as change logs.
Use the Export-Csv cmdlet to save the report as a CSV file in the respective client folder.
Notes
This script requires administrative access to Exchange Online.
Make sure you have the necessary permissions to run PowerShell scripts on your system.
The script may take some time to execute, depending on the number of transport rules in your Exchange Online environment.
License
This script is released under the MIT License.
