# Connect to Exchange Online 
# Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# Get all resource mailboxes
$mailboxes = Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox

# Initialize an array to store the results
$results = @()

# Loop through each mailbox and get Calendar Processing settings
foreach ($mailbox in $mailboxes) {
    $calendarProcessing = Get-CalendarProcessing -Identity $mailbox |fl
    $results += [PSCustomObject]@{
        DisplayName            = $mailbox.DisplayName
        Identity               = $mailbox.Identity
        AutomateProcessing     = $calendarProcessing.AutomateProcessing
        AllowConflicts         = $calendarProcessing.AllowConflicts
        RequestOutOfPolicy     = $calendarProcessing.RequestOutOfPolicy
        BookingWindowInDays    = $calendarProcessing.BookingWindowInDays
        AllBookInPolicy        = $calendarProcessing.AllBookInPolicy
        BookInPolicy           = ($calendarProcessing.BookInPolicy -join ", ")
    }
}
