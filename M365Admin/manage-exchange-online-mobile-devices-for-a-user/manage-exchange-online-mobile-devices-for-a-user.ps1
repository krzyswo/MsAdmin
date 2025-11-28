Connect-MsolService

Connect-ExchangeOnline

$user = Get-MsolUser -UserPrincipalName PLACEHOLDER

$user

Get-MobileDevice -Mailbox "PLACEHOLDER"

Get-MobileDeviceStatistics -Mailbox $user.ObjectId

Remove-MobileDevice -Identity "PLACEHOLDER\ExchangeActiveSyncDevices\Hx§Outlook§E9EFCC65EB1A42CB9307F4EE165A460C"
