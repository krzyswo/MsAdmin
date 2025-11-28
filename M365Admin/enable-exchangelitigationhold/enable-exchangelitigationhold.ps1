# Connect to the Exchange Online PowerShell module
Connect-ExchangeOnline  

# Enable Litigation Hold for generic example users
$users = @(  
    "user1@example.com",  
    "user2@example.com",  
    "user3@example.com"  
)  

foreach ($user in $users) {  
    Set-Mailbox -Identity $user -LitigationHoldEnabled $true  
}  

# Verify litigation hold status  
foreach ($user in $users) {  
    Get-Mailbox -Identity $user | FL *LitigationHoldEnabled*  
} 
