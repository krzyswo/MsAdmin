a. Prepare the CSV File:

Create a CSV file (e.g., MailboxList.csv) with a header named MailboxIdentity and list the email addresses or aliases of the target mailboxes. For example:

# Define the path to the CSV file and the log file
$CsvFilePath = "C:\Path\To\MailboxList.csv"
$LogFilePath = "C:\Path\To\MailboxConfigLog.txt"

MailboxIdentity
user1@domain.com
user2@domain.com
room1@domain.com


# Function to log messages with timestamps
function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFilePath -Value "$Timestamp - $Message"
}

# Import the CSV file
try {
    $Mailboxes = Import-Csv -Path $CsvFilePath
    Write-Log "Successfully imported CSV file: $CsvFilePath"
} catch {
    Write-Log "Failed to import CSV file: $CsvFilePath. Error: $_"
    exit
}

# Loop through each mailbox and apply configurations
foreach ($Mailbox in $Mailboxes) {
    $Identity = $Mailbox.MailboxIdentity
    try {
        Set-MailboxCalendarConfiguration -Identity $Identity `
            -RemindersEnabled $true `
            -DefaultReminderTime "00:15:00" `
            -WorkDays Monday,Tuesday,Wednesday,Thursday,Friday `
            -WorkingHoursStartTime "08:00:00" `
            -WorkingHoursEndTime "17:00:00" `
            -WorkingHoursTimeZone "W. Europe Standard Time" `
            -WeekStartDay Monday `
            -ShowWeekNumbers $true `
            -UseBrightCalendarColorThemeInOwa $false
        Write-Log "Successfully configured calendar settings for mailbox: $Identity"
    } catch {
        Write-Log "Failed to configure calendar settings for mailbox: $Identity. Error: $_"
    }
}
