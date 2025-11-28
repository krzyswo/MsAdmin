Connect-Exchangeonline

# Configure calendar settings for a user mailbox
Set-MailboxCalendarConfiguration -Identity "<MailboxIdentity>" `
    -RemindersEnabled $true `
    -DefaultReminderTime "00:15:00" `
    -WorkDays Monday,Tuesday,Wednesday,Thursday,Friday `
    -WorkingHoursStartTime "08:00:00" `
    -WorkingHoursEndTime "17:00:00" `
    -WorkingHoursTimeZone "W. Europe Standard Time" `
    -WeekStartDay Monday `
    -ShowWeekNumbers $true `
    -UseBrightCalendarColorThemeInOwa $false
