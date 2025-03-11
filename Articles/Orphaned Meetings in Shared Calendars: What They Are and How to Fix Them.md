# Orphaned Meetings in Shared Calendars: What They Are and How to Fix Them

## What Are Orphaned Meetings?
Orphaned meetings occur when a user who scheduled recurring meetings leaves the company, and their account is deactivated or deleted. Since these meetings aren’t automatically canceled, they remain in shared calendars indefinitely without an active owner to manage them.

## Why Are They a Problem?
These orphaned meetings can cause several issues, including:

- **Unnecessary Meeting Reminders** – Users continue receiving notifications for meetings that no longer exist.
- **Blocked Room Availability** – Meeting rooms remain booked, preventing new reservations.
- **Limited Administrative Control** – Without an active organizer, standard methods of editing or removing the meetings may not work.
- **Increased IT Workload** – Resolving these meetings manually can be time-consuming and inefficient.

## How to Fix Orphaned Meetings
There are a few approaches to managing and removing orphaned meetings:

### 1. Assign Calendar Management Permissions
One method is to grant specific users administrative rights to manage shared calendars. This allows them to remove orphaned meetings manually as needed. However, this requires continuous oversight and manual intervention.

### 2. Use PowerShell Commands
Admins can use PowerShell cmdlets such as `Remove-CalendarEvents` to remove calendar events in bulk. However, this approach is not always precise—it can wipe all events within a certain timeframe instead of targeting individual orphaned meetings. Additionally, PowerShell has limitations when working with room calendars.

### 3. Automate Cleanup with Microsoft Graph API
A more flexible and scalable approach is using Microsoft Graph API, which allows admins to programmatically manage calendar events, including orphaned meetings. With the right permissions, IT teams can automate the removal of unwanted meetings without relying on manual processes or requiring additional licenses.

#### Example: Removing Orphaned Meetings Using Microsoft Graph
The following PowerShell script demonstrates how to connect to Microsoft Graph and remove orphaned meetings from shared calendars based on data provided in a CSV file.

```powershell
# Connect to IPPS session
$UserPrincipalName = "<your-upn>"  # Replace with your actual UPN
Connect-IPPSSession -UserPrincipalName $UserPrincipalName

# Retrieve compliance searches and actions
Get-ComplianceSearch
Get-ComplianceSearchAction

# Purge compliance search results
$searchName = "Results full year 2022"
for ($i = 0; $i -lt 30; $i++) {
    New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false
}

# Paths to CSV files
$csvFilePathlogi = "<your-path>\logi.csv"
$csvFilePath = "<your-path>\users1.csv"

# Import Event IDs from CSV
if (Test-Path $csvFilePathlogi) {
    $eventIds = Import-Csv -Path $csvFilePathlogi | ForEach-Object { $_.eventId }
} else {
    Write-Host "Error: CSV file with event IDs not found at $csvFilePathlogi"
    exit
}

# Import User UPNs from CSV
if (Test-Path $csvFilePath) {
    $userUPNs = Import-Csv -Path $csvFilePath | ForEach-Object { $_.UPN }
} else {
    Write-Host "Error: CSV file with user UPNs not found at $csvFilePath"
    exit
}

# Loop through each user UPN and remove the specified calendar events
foreach ($eventId in $eventIds) {
    foreach ($userUPN in $userUPNs) {
        if ($userUPN -and $eventId) {
            Remove-MgUserEvent -UserId $userUPN -EventId $eventId -ErrorAction Continue
        } else {
            Write-Host "Skipping empty UPN or Event ID"
        }
    }
}

Write-Host "Script execution completed."


```` 
## Why Use Microsoft Graph?

- **More Precise Than PowerShell** – Unlike PowerShell cmdlets that remove all events within a date range, Graph API allows for targeted deletion.
- **Scalable and Automated** – Can be scheduled to run regularly without manual intervention.
- **Supports Complex Calendar Operations** – Provides more control over meeting details, allowing admins to modify, reschedule, or delete orphaned meetings based on specific criteria.

## Preventing Orphaned Meetings

To avoid dealing with orphaned meetings in the future, you can:

- **Encourage End Dates for Recurring Meetings** – Ensure all recurring meetings have a defined expiration.
- **Train Employees on Meeting Management** – Educate users on properly canceling or transferring meetings before leaving the company.

## Final Thoughts

Orphaned meetings can cause scheduling conflicts, room booking issues, and unnecessary notifications. While manual methods like calendar delegation and PowerShell scripts can help, automating the process with Microsoft Graph API is the most effective way to keep shared calendars organized. By proactively managing orphaned meetings, businesses can reduce administrative burdens and maintain a seamless scheduling system.
