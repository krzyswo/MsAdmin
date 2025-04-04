Exchange Online Mailbox Settings Checker
 
Description
This PowerShell script retrieves and displays various settings for a list of specified mailboxes in Exchange Online. It provides details such as display name, user principal name, primary SMTP address, and a comprehensive list of mailbox processing settings.

Prerequisites
Access to Exchange Online
Permission to retrieve mailbox and calendar processing settings
Output
The script will generate a grid view displaying the following mailbox settings for each mailbox in the provided list:

Display Name
User Principal Name
Primary SMTP Address
Automate Processing
Allow Conflicts
Allow Distribution Group
Allow Multiple Resources
Booking Type
Booking Window In Days
Maximum Duration In Minutes
Minimum Duration In Minutes
Allow Recurring Meetings
Enforce Adjacency As Overlap
Enforce Capacity
Enforce Scheduling Horizon
Schedule Only During Work Hours
Conflict Percentage Allowed
Maximum Conflict Instances
Forward Requests To Delegates
Delete Attachments
Delete Comments
And more...
Notes
Ensure that the list of mailboxes provided is accurate and accessible.
Review the displayed settings carefully to ensure they align with the intended configurations.
Make any necessary adjustments based on the retrieved settings.
License
This project is licensed under the MIT License - see the LICENSE file for details.
