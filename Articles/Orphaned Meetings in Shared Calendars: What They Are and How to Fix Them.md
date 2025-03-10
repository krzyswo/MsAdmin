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
# Connect to Microsoft Graph using application credentials
$tenantId = "<Your-Tenant-ID>"
$clientId = "<Your-Client-ID>"
$clientSecret = Get-Credential -UserName $clientId  # Securely enter the client secret

Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $clientSecret

# Import the CSV file containing meeting subjects and dates
$csvFilePath = "C:\Path\To\Meetings.csv"
$meetingData = Import-Csv -Path $csvFilePath 

# Loop through each meeting entry
foreach ($event in $meetingData) {
    $params = @{
        subject = $event.Subject
        body = @{
            contentType = "HTML"
            content = "This is an automated cleanup task."
        }
        start = @{
            dateTime = "$($event.StartDate)T00:00:00"
            timeZone = "UTC"
        }
        end = @{
            dateTime = "$($event.EndDate)T00:00:00"
            timeZone = "UTC"
        }
        isOnlineMeeting = $false
    }

    Write-Host "Processing meeting: $($params.subject)"

    # Import the CSV file containing user email addresses
    $userCsvPath = "C:\Path\To\Users.csv"
    $userUPNs = Import-Csv -Path $userCsvPath | ForEach-Object { $_.UPN }

    # Loop through each user and remove the orphaned meeting
    foreach ($userUPN in $userUPNs) {
        $userId = $userUPN
        $userCalendar = Get-MgUserCalendar -UserId $userId | Where-Object { $_.IsDefaultCalendar -eq $true }
        
        if ($userCalendar) {
            $calendarId = $userCalendar.Id
            New-MgUserCalendarEvent -UserId $userId -CalendarId $calendarId -BodyParameter $params
        }
    }
}

```` 
## Why Use Microsoft Graph?

- **More Precise Than PowerShell** – Unlike PowerShell cmdlets that remove all events within a date range, Graph API allows for targeted deletion.
- **Scalable and Automated** – Can be scheduled to run regularly without manual intervention.
- **Supports Complex Calendar Operations** – Provides more control over meeting details, allowing admins to modify, reschedule, or delete orphaned meetings based on specific criteria.

## Preventing Orphaned Meetings

To avoid dealing with orphaned meetings in the future, you can:

- **Encourage End Dates for Recurring Meetings** – Ensure all recurring meetings have a defined expiration.
- **Implement Automated Cleanup Policies** – Use API solutions to periodically review and remove abandoned meetings.
- **Train Employees on Meeting Management** – Educate users on properly canceling or transferring meetings before leaving the company.

## Final Thoughts

Orphaned meetings can cause scheduling conflicts, room booking issues, and unnecessary notifications. While manual methods like calendar delegation and PowerShell scripts can help, automating the process with Microsoft Graph API is the most effective way to keep shared calendars organized. By proactively managing orphaned meetings, businesses can reduce administrative burdens and maintain a seamless scheduling system.
