Title
Microsoft Graph Calendar Events Organizer Verification

Description
This PowerShell script connects to Microsoft Graph using application credentials to retrieve calendar events for users within a specified date range. It then identifies the organizers of these events and verifies their status within the tenant.

Prerequisites
PowerShell environment
Installation of the Microsoft.Graph module using the command: Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
Importing the Microsoft.Graph.Authentication module
Valid Microsoft Graph application credentials (tenant ID, client ID, and client secret)
Access to Microsoft Graph API

Output
The script performs the following actions:

Connects to Microsoft Graph using the provided application credentials.
Retrieves calendar events for users within the specified date range.
Identifies meeting events and collects the organizer email addresses.
Verifies the status of each organizer within the tenant.
Outputs the status of each organizer, indicating whether they are active or not in the tenant.

Notes
Ensure that the correct tenant ID, client ID, and client secret are provided in the script.
The script assumes that events with attendees are considered meetings.
Any errors encountered during event retrieval will be displayed.
The script collects unique organizer email addresses and verifies their status within the tenant.

License
This script is not licensed under any specific license and is provided as-is without any warranty.
