# Connect to Exchange Online
Connect-ExchangeOnline

# Get a list of all mailboxes
Get-Mailbox

# Get a list of mailboxes that have been inactive for more than 90 days but it is depreciated needs to be done on GRAPH now.
# Get-StaleMailboxReport -InactiveMailboxDays 90

# Get a list of mailboxes that have not been logged into in the past 30 days
Get-InactiveMailboxReport -InactiveDays 30

# Get a list of mailboxes with large items (over 10 MB)
Get-LargeItemMailboxReport -SizeLimit 10MB

# Get a list of mailboxes with a forwarding address configured
Get-Mailbox | Where-Object {$_.ForwardingSmtpAddress -ne $null}

# Get a list of mailboxes that have delegates
Get-Mailbox | Where-Object {$_.DelegateType -ne "None"}

# Get a list of mailboxes with full access permissions
Get-Mailbox | Where-Object {$_.FullAccess -ne $null}

# Get a list of mailboxes with send-as permissions
Get-Mailbox | Where-Object {$_.GrantSendOnBehalfTo -ne $null}

# Get a list of mailboxes with Send on Behalf permissions
Get-Mailbox | Where-Object {$_.GrantSendOnBehalfTo -ne $null}

# Get a list of mailboxes with specific email addresses
Get-Mailbox | Where-Object {$_.EmailAddresses -like "*@example.com"}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline
