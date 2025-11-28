Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber

Import-Module Microsoft.Graph.Authentication

# Connect to Microsoft Graph using the application credentials
$tenantId = "XXX"
$clientId = "XXX"  # Make sure to provide the correct client ID
$clientSecret = Get-Credential -UserName $clientId  # Enter the client secret in the password prompt 

# Define the date range (ISO 8601 format):
$startDateTime = "2025-04-01"
$endDateTime   = "2025-04-20"

# Connect to Microsoft Graph using app-only authentication
Connect-MgGraph -TenantId $tenantId -ClientSecret $clientSecret

# Test on one user just to see if it populate fcorrect values

Get-MgUserCalendarView -UserId "xxx@contoso.com" -StartDateTime $startDateTime -EndDateTime $endDateTime -All


###########################################################################
# Retrieve all users (both enabled and disabled) in the tenant.
$users = Get-MgUser -All

# Build a dictionary mapping tenant users by email (using UserPrincipalName) to their user object.
$tenantUsers = @{}
foreach ($u in $users) {
    $emailKey = $u.UserPrincipalName.ToLower()
    $tenantUsers[$emailKey] = $u
}

# Array to collect unique organizer email addresses from meeting events.
$organizerEmails = @()

# Loop through all users, list their meetings, and collect the organizer emails.
foreach ($user in $users) {
    Write-Output "----------------------"
    Write-Output "User: $($user.DisplayName) ($($user.UserPrincipalName)) - Enabled: $($user.AccountEnabled)"
    
    try {
        # Retrieve the user's calendar events in the specified date range.
        $events = Get-MgUserCalendarView -UserId $user.Id -StartDateTime $startDateTime -EndDateTime $endDateTime -All
        
        # Filter events to include only those with one or more attendees (assumed to be meetings).
        $meetingEvents = $events | Where-Object { $_.Attendees.Count -gt 0 }
        
        foreach ($event in $meetingEvents) {
            $organizer = $event.Organizer.EmailAddress
            Write-Output "Meeting: $($event.Subject) | Organizer: $($organizer.Name) <$($organizer.Address)>"
            
            # Add the organizer email address (in lowercase) to the collection if not already present.
            $orgEmail = $organizer.Address.ToLower()
            if (-not ($organizerEmails -contains $orgEmail)) {
                $organizerEmails += $orgEmail
            }
        }
    }
    catch {
        Write-Output "Error retrieving events for $($user.UserPrincipalName): $_"
    }
}

# After processing all users, verify the organizer accounts.
Write-Output "----------------------"
Write-Output "Verifying organizer accounts..."

foreach ($orgEmail in $organizerEmails) {
    if ($tenantUsers.ContainsKey($orgEmail)) {
        $orgUser = $tenantUsers[$orgEmail]
        if ($orgUser.AccountEnabled) {
            Write-Output "Organizer $orgEmail is active."
        }
        else {
            Write-Output "Organizer $orgEmail is found in the tenant but is NOT active."
        }
    }
    else {
        Write-Output "Organizer $orgEmail is not found in the tenant."
    }
}

