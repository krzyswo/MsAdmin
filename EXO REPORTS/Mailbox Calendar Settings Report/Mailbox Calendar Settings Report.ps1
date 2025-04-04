# Connect to Exchange Online
Connect-ExchangeOnline

# Define the list of mailboxes to check
# Replace with your actual list of resource mailboxes or user mailboxes
$list = @(
    "room1@yourdomain.com",
    "room2@yourdomain.com",
    "room3@yourdomain.com"
)

$settings = @()
 
foreach ($obj in $list)
{
    $mailbox = get-mailbox -Identity $obj |select DisplayName,UserPrincipalName,PrimarySmtpAddress
    $processing = get-calendarprocessing -Identity $obj
 
    $settings += [PSCustomObject]@{
       
        DisplayName                           = $mailbox.DisplayName
        UserPrincipalName                     = $mailbox.UserPrincipalName
        PrimarySmtpAddress                    = $mailbox.PrimarySmtpAddress
        AutomateProcessing                    = $processing.AutomateProcessing
        AllowConflicts                        = $processing.AllowConflicts
        AllowDistributionGroup                = $processing.AllowDistributionGroup
        AllowMultipleResources                = $processing.AllowMultipleResources
        BookingType                           = $processing.BookingType
        BookingWindowInDays                   = $processing.BookingWindowInDays
        MaximumDurationInMinutes              = $processing.MaximumDurationInMinutes
        MinimumDurationInMinutes              = $processing.MinimumDurationInMinutes
        AllowRecurringMeetings                = $processing.AllowRecurringMeetings
        EnforceAdjacencyAsOverlap             = $processing.EnforceAdjacencyAsOverlap
        EnforceCapacity                       = $processing.EnforceCapacity
        EnforceSchedulingHorizon              = $processing.EnforceSchedulingHorizon
        ScheduleOnlyDuringWorkHours           = $processing.ScheduleOnlyDuringWorkHours
        ConflictPercentageAllowed             = $processing.ConflictPercentageAllowed
        MaximumConflictInstances              = $processing.MaximumConflictInstances
        ForwardRequestsToDelegates            = $processing.ForwardRequestsToDelegates
        DeleteAttachments                     = $processing.DeleteAttachments
        DeleteComments                        = $processing.DeleteComments
        RemovePrivateProperty                 = $processing.RemovePrivateProperty
        DeleteSubject                         = $processing.DeleteSubject
        AddOrganizerToSubject                 = $processing.AddOrganizerToSubject
        DeleteNonCalendarItems                = $processing.DeleteNonCalendarItems
        TentativePendingApproval              = $processing.TentativePendingApproval
        EnableResponseDetails                 = $processing.EnableResponseDetails
        OrganizerInfo                         = $processing.OrganizerInfo
        ResourceDelegates                     = $processing.ResourceDelegates
        RequestOutOfPolicy                    = $processing.RequestOutOfPolicy
        AllRequestOutOfPolicy                 = $processing.AllRequestOutOfPolicy
        BookInPolicy                          = $processing.BookInPolicy
        AllBookInPolicy                       = $processing.AllBookInPolicy
        RequestInPolicy                       = $processing.RequestInPolicy
        AllRequestInPolicy                    = $processing.AllRequestInPolicy
        AddAdditionalResponse                 = $processing.AddAdditionalResponse
        AdditionalResponse                    = $processing.AdditionalResponse
        RemoveOldMeetingMessages              = $processing.RemoveOldMeetingMessages
        AddNewRequestsTentatively             = $processing.AddNewRequestsTentatively
        ProcessExternalMeetingMessages        = $processing.ProcessExternalMeetingMessages
        RemoveForwardedMeetingNotifications   = $processing.RemoveForwardedMeetingNotifications
        AutoRSVPConfiguration                 = $processing.AutoRSVPConfiguration
        RemoveCanceledMeetings                = $processing.RemoveCanceledMeetings
        EnableAutoRelease                     = $processing.EnableAutoRelease
        PostReservationMaxClaimTimeInMinutes  = $processing.PostReservationMaxClaimTimeInMinutes
        MailboxOwnerId                        = $processing.MailboxOwnerId
        Identity                              = $processing.Identity
        IsValid                               = $processing.IsValid
        ObjectState                           = $processing.ObjectState
       
       }
}
 
$settings |out-gridview
