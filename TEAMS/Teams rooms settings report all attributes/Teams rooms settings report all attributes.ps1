# Connect to Exchange Online 
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Get all resource mailboxes
$mailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox

# Initialize an array to store the results
$results = @()

# Loop through each mailbox and get Calendar Processing settings
foreach ($mailbox in $mailboxes) {
    $calendarProcessing = Get-CalendarProcessing -Identity $mailbox

    # Add the settings to the results array
    $results += [PSCustomObject]@{
        DisplayName                          = $mailbox.DisplayName
        Identity                             = $mailbox.Identity
        AutomateProcessing                   = $calendarProcessing.AutomateProcessing
        AllowConflicts                       = $calendarProcessing.AllowConflicts
        AllowDistributionGroup               = $calendarProcessing.AllowDistributionGroup
        AllowMultipleResources               = $calendarProcessing.AllowMultipleResources
        BookingType                          = $calendarProcessing.BookingType
        BookingWindowInDays                  = $calendarProcessing.BookingWindowInDays
        MaximumDurationInMinutes             = $calendarProcessing.MaximumDurationInMinutes
        MinimumDurationInMinutes             = $calendarProcessing.MinimumDurationInMinutes
        AllowRecurringMeetings               = $calendarProcessing.AllowRecurringMeetings
        EnforceAdjacencyAsOverlap            = $calendarProcessing.EnforceAdjacencyAsOverlap
        EnforceCapacity                      = $calendarProcessing.EnforceCapacity
        EnforceSchedulingHorizon             = $calendarProcessing.EnforceSchedulingHorizon
        ScheduleOnlyDuringWorkHours          = $calendarProcessing.ScheduleOnlyDuringWorkHours
        ConflictPercentageAllowed            = $calendarProcessing.ConflictPercentageAllowed
        MaximumConflictInstances             = $calendarProcessing.MaximumConflictInstances
        ForwardRequestsToDelegates           = $calendarProcessing.ForwardRequestsToDelegates
        DeleteAttachments                    = $calendarProcessing.DeleteAttachments
        DeleteComments                       = $calendarProcessing.DeleteComments
        RemovePrivateProperty                = $calendarProcessing.RemovePrivateProperty
        DeleteSubject                        = $calendarProcessing.DeleteSubject
        AddOrganizerToSubject                = $calendarProcessing.AddOrganizerToSubject
        DeleteNonCalendarItems               = $calendarProcessing.DeleteNonCalendarItems
        TentativePendingApproval             = $calendarProcessing.TentativePendingApproval
        EnableResponseDetails                = $calendarProcessing.EnableResponseDetails
        OrganizerInfo                        = $calendarProcessing.OrganizerInfo
        ResourceDelegates                    = ($calendarProcessing.ResourceDelegates -join ", ")
        RequestOutOfPolicy                   = ($calendarProcessing.RequestOutOfPolicy -join ", ")
        AllRequestOutOfPolicy                = $calendarProcessing.AllRequestOutOfPolicy
        BookInPolicy                         = ($calendarProcessing.BookInPolicy -join ", ")
        AllBookInPolicy                      = $calendarProcessing.AllBookInPolicy
        RequestInPolicy                      = ($calendarProcessing.RequestInPolicy -join ", ")
        AllRequestInPolicy                   = $calendarProcessing.AllRequestInPolicy
        AddAdditionalResponse                = $calendarProcessing.AddAdditionalResponse
        AdditionalResponse                   = $calendarProcessing.AdditionalResponse
        RemoveOldMeetingMessages             = $calendarProcessing.RemoveOldMeetingMessages
        AddNewRequestsTentatively            = $calendarProcessing.AddNewRequestsTentatively
        ProcessExternalMeetingMessages       = $calendarProcessing.ProcessExternalMeetingMessages
        RemoveForwardedMeetingNotifications  = $calendarProcessing.RemoveForwardedMeetingNotifications
        AutoRSVPConfiguration                = $calendarProcessing.AutoRSVPConfiguration
        RemoveCanceledMeetings               = $calendarProcessing.RemoveCanceledMeetings
        EnableAutoRelease                    = $calendarProcessing.EnableAutoRelease
        PostReservationMaxClaimTimeInMinutes = $calendarProcessing.PostReservationMaxClaimTimeInMinutes
        MailboxOwnerId                       = $calendarProcessing.MailboxOwnerId
        IsValid                              = $calendarProcessing.IsValid
        ObjectState                          = $calendarProcessing.ObjectState
    }
}


# Export the results to a CSV file
$results | Export-Csv -Path "C:\PLACEHOLDER\ResourceMailboxSettingsFullattributes.csv" -NoTypeInformation -Encoding UTF8

Write-Output "Report generated at C:\PLACEHOLDER\ResourceMailboxSettingsFullattributes.csv"
