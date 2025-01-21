# If not working check that:
# Ensure MSOnline module is up to date
#Uninstall-Module -Name MSOnline -Force
#Install-Module -Name MSOnline -Force
# Import the MSOnline module
#Import-Module MSOnline

# Connect to Microsoft Online services
Connect-MsolService


[CmdletBinding()]
param ()

# Fetch all users
$allUsers = Get-MsolUser -All -ErrorAction Stop

# Little report
$directLicenseAssignmentReport = @()
$directLicenseAssignmentCount = 0

foreach ($user in $allUsers) {
    # Processing all licenses per user
    foreach ($license in $user.Licenses) {
        <#
            The "GroupsAssigningLicense" array contains objectId's of groups which inherit licenses.
            If the array contains an entry with the user's own objectId, the license was assigned directly to the user.
            If the array contains no entries and the user has a license assigned, it is also a direct license assignment.
        #>
        if ($license.GroupsAssigningLicense -contains $user.ObjectId -or $license.GroupsAssigningLicense.Count -lt 1) {
            $directLicenseAssignmentCount++

            # Add details to the report
            $directLicenseAssignmentReport += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                ObjectId = $user.ObjectId
                AccountSkuId = $license.AccountSkuId
                DirectAssignment = $true
            }
        }
    }
}

if ($directLicenseAssignmentCount -gt 0) {
    Write-Output "nFound $directLicenseAssignmentCount direct assigned license(s):"
    Write-Output $directLicenseAssignmentReport
} else {
    Write-Output "No direct license assignments found"
}
