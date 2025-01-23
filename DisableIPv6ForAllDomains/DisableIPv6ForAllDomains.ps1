#CHANGES WILL TAKE PLACE AFTER MINIMUM 10 MINUTES!!!!!!!

# Get all accepted domains  
$domains = Get-AcceptedDomain  
  
# Loop through each domain and disable IPv6  
foreach ($domain in $domains) {  
    $domainName = $domain.DomainName  
  
    # Disable IPv6 for the current accepted domain  
    Disable-IPv6ForAcceptedDomain -Domain $domainName  
  
    # Display a popup window informing the user about the waiting time  
    $wshell = New-Object -ComObject WScript.Shell  
    $wshell.Popup("Please wait for 15 minutes for the changes to propagate.", 0, "IPv6 Status Update", 64)  
}  
