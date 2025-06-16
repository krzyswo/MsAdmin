## Block External Email Forwarding
Description
Learn how to identify and block external email forwarding configurations that risk your organization efficiently using PowerShell script.

Prerequisites
- PowerShell
- Appropriate permissions to manage email forwarding settings

Usage
Run the script to generate reports and optionally block external email forwarding configurations.

Output
The script generates two reports:
1. External email forwarding configuration report
2. Inbox rules with external forwarding report

After confirmation, the script will:
- Block external forwarding configuration
- Disable all the inbox rules in the output report
- Provide the respective log file

Notes
Ensure you have the necessary permissions before running this script.

License
MIT License