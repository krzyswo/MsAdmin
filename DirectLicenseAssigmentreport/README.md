# README: Direct License Assignment Script

## Description
This PowerShell script identifies users in Microsoft 365 who have licenses assigned directly to them (not inherited via group assignments). The script generates a report summarizing the findings.

## Prerequisites
1. **PowerShell Environment**
   - Ensure PowerShell is installed and updated on your machine.

2. **MSOnline Module**
   - Ensure the `MSOnline` PowerShell module is installed and up to date.
   - If not installed, uncomment the following lines in the script to install or update the module:
     ```powershell
     #Uninstall-Module -Name MSOnline -Force
     #Install-Module -Name MSOnline -Force
     ```

3. **Permissions**
   - You must have appropriate permissions in Microsoft 365 to fetch user license data. Typically, this requires being a Global Administrator or License Administrator.

## Usage
1. Open PowerShell as an administrator.
2. Connect to the Microsoft Online service by running:
   ```powershell
   Connect-MsolService
   ```
   - Provide your admin credentials when prompted.

3. Run the script to fetch and analyze license assignments.

4. The script will:
   - Fetch all users in the Microsoft 365 tenant.
   - Analyze each user's licenses to check if they are directly assigned.
   - Generate a report of users with direct license assignments.

## Script Details
1. **Connect to Microsoft Online Services:**
   The script ensures a valid connection using the `Connect-MsolService` cmdlet.

2. **Fetch Users:**
   Retrieves all users using:
   ```powershell
   Get-MsolUser -All -ErrorAction Stop
   ```

3. **Analyze Licenses:**
   Each user's license assignments are analyzed. A license is considered directly assigned if:
   - The user's `ObjectId` appears in the `GroupsAssigningLicense` array, or
   - The `GroupsAssigningLicense` array is empty.

4. **Generate Report:**
   A report is created as a PowerShell object (`PSCustomObject`) containing:
   - `UserPrincipalName`: The user's email or login.
   - `ObjectId`: The unique identifier of the user.
   - `AccountSkuId`: The SKU of the assigned license.
   - `DirectAssignment`: Boolean indicating direct assignment.

5. **Output Results:**
   - If direct assignments are found, they are displayed in the console.
   - If no direct assignments are found, a message is shown.

## Troubleshooting
- **Module Issues:** If the `MSOnline` module is missing or outdated, uncomment the installation lines at the top of the script and run them to ensure the module is available.
- **Connectivity Issues:** If you cannot connect using `Connect-MsolService`, ensure your credentials are correct and your account has sufficient privileges.
- **Permissions:** Ensure you have the necessary admin roles to access user and license information.

## Example Output
- **Direct Assignments Found:**
  ```plaintext
  Found 3 direct assigned license(s):
  UserPrincipalName           ObjectId                              AccountSkuId             DirectAssignment
  -----------------           --------                              ------------             ----------------
  user1@domain.com            abcdef12-3456-7890-abcd-ef1234567890  contoso:ENTERPRISEPACK  True
  user2@domain.com            12345678-abcd-ef12-3456-7890abcdef12  contoso:BUSINESSPACK    True
  user3@domain.com            7890abcd-1234-5678-efab-cdef12345678  contoso:PROPLUS         True
  ```

- **No Direct Assignments Found:**
  ```plaintext
  No direct license assignments found
  ```

