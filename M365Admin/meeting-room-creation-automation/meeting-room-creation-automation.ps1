# Connect to Exchange Online
Connect-ExchangeOnline

# Define the alias for the mailbox
$Alias = Read-Host "Enter the alias for the mailbox eq. meetingroom-nl-xxx-projectroom"

# Define the domain for the mailbox
$Domain = Read-Host "Enter the domain for the mailbox eq. @example.com"

# Combine the alias and domain to create the Microsoft Online Services ID
$MicrosoftOnlineServicesID = $Alias + $Domain

# Prompt for the room name (up to 64 characters) and validate its length
do {
    $Name = Read-Host "Enter the name of the room (up to 64 characters)"
    if ($Name.Length -gt 64) {
        Write-Host -ForeGroundColor RED "Error: The name must be up to 64 characters. Your name has" $Name.Length
    }
} while ($Name.Length -gt 64)

# Define various room properties
$Audio = Read-Host "Enter the name of the audio device (e.g., Teams)"
$Video = Read-Host "Enter the name of the video device (e.g., Teams)"
$Display = Read-Host "Enter the name of the display device (e.g., Flatscreen)"
$Floor = Read-Host "Enter the floor number"
$Building = Read-Host "Enter the building number"
$Street = Read-Host "Enter the street address"
$Capacity = Read-Host "Enter the room capacity"
$City = Read-Host "Enter the city"
$Country = Read-Host "Enter the two-letter country code (e.g., NL)"
$Company = Read-Host "Enter the company name"
$Department = Read-Host "Enter the department name"

# Create the new mailbox for the room (Password must be securely handled)
$RoomMailboxPassword = ConvertTo-SecureString -String "<PLACEHOLDER_PASSWORD>" -AsPlainText -Force
New-Mailbox -MicrosoftOnlineServicesID $MicrosoftOnlineServicesID -Alias $Alias -Name $Name -Room -EnableRoomMailboxAccount $true -RoomMailboxPassword $RoomMailboxPassword

# Wait for 180 seconds before continuing
Start-Sleep -Seconds 180

# Set calendar processing for the room
Set-CalendarProcessing -Identity $Alias -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -AllowConflicts $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -AddAdditionalResponse $false -BookingWindowInDays 180 -ProcessExternalMeetingMessages $true -AllBookInPolicy $true

# Set place properties for the room
Set-Place -Identity $Alias -AudioDeviceName $Audio -VideoDeviceName $Video -DisplayDeviceName $Display -IsWheelChairAccessible $true -Building $Building -Street $Street -Capacity $Capacity -City $City -CountryOrRegion $Country -Floor $Floor

# Set user properties for the room
$RoomNumber = Read-Host "Enter the room number"
Set-User -Identity $MicrosoftOnlineServicesID -Company $Company -Department $Department -Office $RoomNumber

# Ask if you want to add yourself as an owner
$DistributionGroup = Read-Host "Enter the email address of the distribution group"
$AddOwner = Read-Host "Do you want to add yourself as an owner of the distribution group? (Y/N)"
if ($AddOwner.ToUpper() -eq "Y") {
    Add-DistributionGroupOwner -Identity $DistributionGroup -Member "<PLACEHOLDER_OWNER_EMAIL>"
}

# Add the mailbox to a distribution group
$DistributionGroup = Read-Host "Enter the email address of the distribution group"
Add-DistributionGroupMember -Identity $DistributionGroup -Member $MicrosoftOnlineServicesID

# Check the place, calendar processing, mailbox, and user properties
Get-Place -Identity "$Alias" | Format-List
Get-CalendarProcessing -Identity $MicrosoftOnlineServicesID | Select -ExpandProperty BookInPolicy | Out-File -FilePath "C:\Users\<PLACEHOLDER_USER>\Downloads\bookinpolicy1.csv"
Get-Place -Identity $MicrosoftOnlineServicesID | Select AudioDeviceName,VideoDeviceName,DisplayDeviceName,IsWheelChairAccessible,Building,Street,Capacity,City,CountryOrRegion,Floor
Get-User -Identity $MicrosoftOnlineServicesID | Select Company,Department,Office
Get-CalendarProcessing -Identity $MicrosoftOnlineServicesID | Select AutomateProcessing,AddOrganizerToSubject,AllowConflicts,DeleteComments,DeleteSubject,RemovePrivateProperty,AddAdditionalResponse,BookingWindowInDays,ProcessExternalMeetingMessages,AllBookInPolicy

# Example placeholders for sensitive emails
Get-CalendarProcessing -Identity "<PLACEHOLDER_MEETINGROOM_EMAIL>" | Select AutomateProcessing,AddOrganizerToSubject,AllowConflicts,DeleteComments,DeleteSubject,RemovePrivateProperty,AddAdditionalResponse,BookingWindowInDays,ProcessExternalMeetingMessages,AllBookInPolicy
Get-User -Identity "<PLACEHOLDER_MEETINGROOM_EMAIL>" | Select Company,Department,Office
Get-Mailbox -Identity "$MicrosoftOnlineServicesID" | Select *arch*
Get-Place -Identity "<PLACEHOLDER_MEETINGROOM_EMAIL>" | Select AudioDeviceName,VideoDeviceName,DisplayDeviceName,IsWheelChairAccessible,Building,Street,Capacity,City,CountryOrRegion,Floor
