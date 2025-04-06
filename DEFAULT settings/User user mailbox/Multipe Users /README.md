This script prepares a CSV file containing a list of email addresses or aliases of target mailboxes to configure their calendar settings.

Prerequisites
PowerShell environment
Access to the Exchange server
Permission to configure mailbox calendar settings

Prepare the CSV File:

Create a CSV file (e.g., MailboxList.csv) with a header named MailboxIdentity and list the email addresses or aliases of the target mailboxes. For example:

# Define the path to the CSV file and the log file
$CsvFilePath = "C:\Path\To\MailboxList.csv"
$LogFilePath = "C:\Path\To\MailboxConfigLog.txt"

MailboxIdentity
user1@domain.com
user2@domain.com
room1@domain.com


Output
The script reads the CSV file containing mailbox identities and applies the following calendar configurations to each mailbox:

Enables reminders
Sets the default reminder time to 15 minutes before the event
Specifies workdays from Monday to Friday
Defines working hours from 8:00 AM to 5:00 PM
Sets the time zone to "W. Europe Standard Time"
Specifies Monday as the start day of the week
Shows week numbers
Uses a standard calendar color theme in OWA

Notes
Ensure the CSV file (e.g., MailboxList.csv) is correctly formatted with a header named "MailboxIdentity" and lists the email addresses or aliases of the target mailboxes.
Adjust the paths for the CSV file and log file according to your system setup.
The script logs messages with timestamps to the specified log file for tracking configuration activities.
Handle any errors that may occur during the configuration process.

License
This script is released under the MIT license.
