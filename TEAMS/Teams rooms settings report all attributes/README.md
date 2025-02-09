This script retrieves and exports the full attributes of resource mailboxes in Exchange Online.

Prerequisites
PowerShell with Exchange Online Management module installed
Permissions to connect to Exchange Online and retrieve mailbox settings
Usage
Connect to Exchange Online:
Connect-ExchangeOnline  
 
2. Run the script to get resource mailbox settings:

./Get-ResourceMailboxSettings.ps1  
 

Parameters
No input parameters required.
Output
The script will generate a CSV file containing the full attributes of resource mailboxes, including settings like AutomateProcessing, BookingType, AllowRecurringMeetings, and more.

Notes
Ensure you have the necessary permissions to access and retrieve mailbox settings.
The CSV file will be saved at the location: C:\PLACEHOLDER\ResourceMailboxSettingsFullattributes.csv
License
This script is released under the MIT license.
