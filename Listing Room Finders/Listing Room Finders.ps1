# Connect to EXO
Connect-ExchangeOnline

# List Room Finders
Get-DistributionGroup -RecipientTypeDetails "RoomList"
