Export Exchange Online Distribution Group Members to CSV

This PowerShell script connects to Exchange Online and exports the members of distribution groups to a CSV file.

Prerequisites

Access to Exchange Online PowerShell session.
Permission to read distribution group members.
PowerShell environment with Exchange Online module installed.

Output

The script reads a CSV file containing distribution group names, retrieves the members of each group, and exports the data to a new CSV file. The output CSV file will contain the distribution group names and their respective members in a structured format.

Notes

Ensure you have the necessary permissions to access and retrieve distribution group members.
Modify the paths for the input CSV file and output CSV file according to your environment.
Review the exported CSV file to verify the accuracy of the distribution group members' data.

License

This script is released under the MIT license.
