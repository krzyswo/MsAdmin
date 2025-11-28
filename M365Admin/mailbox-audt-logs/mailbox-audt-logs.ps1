# Connect to Exchange Online
Connect-ExchangeOnline

# Define the mailbox
$mailbox = "EMAIL_PLACEHOLDER"

# Get recipient details
Get-Recipient -Identity "EMAIL_PLACEHOLDER" | fl

# Search the mailbox audit log for a specific date range and display selected details in a grid view
Search-MailboxAuditLog -Identity $mailbox -StartDate "01/01/2025" -EndDate "01/12/2025" -ResultSize 10000 -ShowDetails |
Select-Object lastaccessed,operation,logontype,destfolderpathname,folderpathname,internallogontype,itemsubject,sourceitemsubjectslist,sourceitemfolderpathnameslist,logonuserdisplayname,clientinfostring,appid,clientappid,dirtyproperties | out-gridview

# Get recoverable items from the mailbox that match a specific subject and date range
Get-RecoverableItems -Identity "EMAIL_PLACEHOLDER" -SubjectContains "place holder" -FilterItemType IPM.Note -FilterStartTime "11/11/2024 12:00:00 AM" -FilterEndTime "11/18/2024 11:59:59 PM"
