## Assign Office 365 User Manager Based on User Properties
This script sets up managers for Office 365 users using Set-AzureADUserManager, targeting multiple user properties.

### Description
The PowerShell script automates the assignment of managers to Office 365 users based on specific user properties. It utilizes the Set-AzureADUserManager cmdlet to perform this task efficiently.

### Prerequisites
- PowerShell
- AzureAD module
- Appropriate permissions to manage Office 365 user properties

### Usage
Run the script in a PowerShell environment with the necessary permissions. Ensure all user properties are correctly configured before execution.

### Output
The script outputs the status of manager assignment for each user, indicating success or failure.

### Notes
Ensure that the AzureAD module is installed and that you have administrative rights to manage user properties in Office 365.

### License
This project is licensed under the MIT License - see the LICENSE file for details.