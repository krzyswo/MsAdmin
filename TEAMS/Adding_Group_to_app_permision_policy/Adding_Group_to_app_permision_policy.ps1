# Connect to Azure AD
Connect-AzureAD

# Connect to Teams
Connect-MicrosoftTeams

# Define the policy name
$policyName = "New Chatbot"

# Get an array of user UPNs from the "PZC-SSO-Sophie" group
$userUPNs = Get-AzureADGroupMember -ObjectId "cc830891-5585-4889-bf4a-61c4c55ad51d" | Where-Object {$_.objectType -eq "User"} | Select-Object -ExpandProperty UserPrincipalName

# Loop through the array of user UPNs and grant them the Teams app permission policy
foreach ($userUPN in $userUPNs) {
    # Get the user object using their UPN
    $user = Get-AzureADUser -ObjectId $userUPN

    # Grant the Teams app permission policy to the current user
    Grant-CSTeamsAppPermissionPolicy -PolicyName $policyName -Identity $userUPN
}
