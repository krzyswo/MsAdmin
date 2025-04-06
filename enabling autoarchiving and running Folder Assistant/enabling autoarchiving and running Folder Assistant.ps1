Connect-exchanegonline

$mailbox = "xxx@contoso.com"

#enabling autoarchiving and running Folder Assistant
Enable-Mailbox $mailbox -AutoExpandingArchive
Start-ManagedFolderAssistant -Identity $mailbox
 
#if versions folder is high
Start-ManagedFolderAssistant -Identity $mailbox -HoldCleanup
Start-ManagedFolderAssistant -Identity $mailbox -FullCrawl
