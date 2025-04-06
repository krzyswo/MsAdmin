Connect-ExchangeOnline

# Configure calendar settings for a resource mailbox
Set-MailboxCalendarConfiguration -Identity "<MailboxIdentity>" `
    -RemindersEnabled $false `
    -DefaultReminderTime "00:05:00" `
    -WorkDays Monday,Tuesday,Wednesday,Thursday,Friday `
    -WorkingHoursStartTime "09:00:00" `
    -WorkingHoursEndTime "18:00:00" `
    -WorkingHoursTimeZone "W. Europe Standard Time" `
    -WeekStartDay Monday `
    -ShowWeekNumbers $false `
    -UseBrightCalendarColorThemeInOwa $true
