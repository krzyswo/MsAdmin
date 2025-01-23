#Connect to Exchange 
Connect-ExchangeOnline

#Get all accepted domains
$domains = Get-AcceptedDomain

# Loop through each domain and get the IPv6 status
foreach ($domain in $domains) {
    $domainName = $domain.DomainName
    
    # Get IPv6 for the current domain
    $ipv6Status = Get-IPv6StatusForAcceptedDomain -Domain $domainName
    $status = $ipv6Status.Status

    # Output the result
    Write-Host "IPv6 Status for ${domainName}: $status" -ForegroundColor Cyan
   # Write-Host "${domainName}" -ForegroundColor Cyan
}
