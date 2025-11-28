# Connect to EXO

Connect-ExchangeOnline 

# Get all distribution groups with Room List ResourceType
$RoomFinders = Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq "RoomList" }

# Display the list of room finders
foreach ($RoomFinder in $RoomFinders) {
    Write-Host "Room Finder Name: $($RoomFinder.DisplayName)"
    Write-Host "Room Finder Email: $($RoomFinder.PrimarySmtpAddress)"
    Write-Host "-------------------------"
}
