
Exchange Online Mailbox Audit Log
 
This PowerShell script connects to Exchange Online and retrieves mailbox audit log information for a specific mailbox. It also searches for recoverable items in the mailbox based on a specific subject and date range.

Prerequisites
 

PowerShell
Exchange Online PowerShell module
Usage
 

Open PowerShell.
Connect to Exchange Online by running the following command:
Connect-ExchangeOnline  
 
3. Define the mailbox by replacing "EMAIL_PLACEHOLDER" with the actual email address:

$mailbox = "user@example.com"  
 
4. Get recipient details by running the following command:

Get-Recipient -Identity $mailbox | fl  
 
5. Search the mailbox audit log for a specific date range and display selected details in a grid view:

Search-MailboxAuditLog -Identity $mailbox -StartDate "01/01/2025" -EndDate "01/12/2025" -ResultSize 10000 -ShowDetails |  
Select-Object lastaccessed, operation, logontype, destfolderpathname, folderpathname, internallogontype, itemsubject, sourceitemsubjectslist, sourceitemfolderpathnameslist, logonuserdisplayname, clientinfostring, appid, clientappid, dirtyproperties | out-gridview  
 
6. Get recoverable items from the mailbox that match a specific subject and date range:

Get-RecoverableItems -Identity $mailbox -SubjectContains "place holder" -FilterItemType IPM.Note -FilterStartTime "11/11/2024 12:00:00 AM" -FilterEndTime "11/18/2024 11:59:59 PM"  
 

Parameters
 

$mailbox: The email address of the mailbox to retrieve audit log information from.
Output
 
The script retrieves mailbox audit log information and recoverable items from the specified mailbox. The output includes details such as the last accessed date, operation type, logon type, folder path, item subject, source item details, logon user display name, client information, app ID, client app ID, and dirty properties.

Notes
 

Make sure you have the necessary permissions to access the mailbox audit log and recoverable items.
Adjust the date ranges and search criteria as needed.
The script uses the Exchange Online PowerShell module to connect to Exchange Online. Make sure you have the module installed and imported before running the script.
License
 
This script is provided under the MIT License.
