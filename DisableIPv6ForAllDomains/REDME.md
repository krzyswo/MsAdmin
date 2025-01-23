
This script is designed to disable IPv6 for all accepted domains and notify the user to wait for the changes to take effect.
Prerequisites
 

PowerShell: The script is written in PowerShell, so ensure that you have PowerShell installed on your system.
Usage
 

Open a PowerShell console or the PowerShell Integrated Scripting Environment (ISE).
Copy and paste the script into the console or the script editor.
Run the script by executing it.
Script Flow
 

The script starts by fetching all accepted domains using the Get-AcceptedDomain cmdlet.
It then loops through each domain and disables IPv6 for the accepted domain using the Disable-IPv6ForAcceptedDomain function.
After disabling IPv6 for each domain, a popup window is displayed using the WScript.Shell COM object, informing the user to wait for 15 minutes for the changes to propagate.
Important Note
 

It is crucial to understand the implications of disabling IPv6 before running this script. Ensure that you have a valid reason for disabling IPv6 and that it aligns with your network requirements.
Disclaimer
 

This script is provided as-is, without any warranties or guarantees. Use it at your own risk.
Contributing
 

If you find any issues or have suggestions for improvements, please open an issue or submit a pull request on the script's GitHub repository.
License
 

This script is released under the MIT License.
