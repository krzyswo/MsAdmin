#Add one member
Add-DistributionGroupMember -Identity "DL_PLACEHOLDER" -Member "DL_PLACEHOLDER"

########################################################################################################################################

#Add members from list
# Define the members as an array
$members = "DL_PLACEHOLDER", 
           "DL_PLACEHOLDER", 
           "DL_PLACEHOLDER", 
           "DL_PLACEHOLDER", 
           "DL_PLACEHOLDER", 
           "DL_PLACEHOLDER", 
           "DL_PLACEHOLDER"

# Loop through the members and add them to the distribution group
foreach ($member in $members) {
    Add-DistributionGroupMember -Identity "DL_PLACEHOLDER" -Member $member
}


########################################################################################################################################

# check if members were added

Get-DistributionGroupMember -Identity "DL_PLACEHOLDER" | select PrimarySmtpAddress, DisplayName
