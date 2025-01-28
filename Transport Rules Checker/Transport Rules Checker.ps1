# Connect to Exchange Online
Connect-ExchangeOnline -UserPrincipalName "janek.poranek@contoso.pl" -ShowProgress $true


# Get all transport rules
$transportRules = Get-TransportRule

# Create an array to store rule details
$ruleDetails = @()

# Iterate through the transport rules and add their details to the array
$transportRules | ForEach-Object {
    $rule = $_
    $comments = ($rule.Comments | Out-String).Trim()
    $comments = $comments -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
    $comments = $comments -join ", "
    
    $ruleDetail = [PSCustomObject]@{
        'Rule Name' = $rule.Name
        'Description' = $rule.Description
        'Priority' = $rule.Priority
        'Enabled' = $rule.Enabled
        'Actions' = ($rule.Actions | Out-String).Trim()
        'Conditions' = ($rule.Conditions | Out-String).Trim()
        'Exceptions' = ($rule.Exceptions | Out-String).Trim()
        'Scopes' = ($rule.Scopes | Out-String).Trim()
        'Comments' = $comments
        
    }
    $ruleDetails += $ruleDetail
}

# Display the rule details in a grid view
$ruleDetails | Out-GridView 

# Createa report for te Transsport Rules includign for what clinent wand when it has been created, next teposts for the same client need to have change logs save reports in separate folder for each client in the same common location 
