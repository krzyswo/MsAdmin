<# Script to block external email forwarding in Exchange Online #>

param (
    [string] $CertificateThumbPrint,
    [string] $ClientId,
    [string] $Organization,
    [string] $UserName,
    [string] $Password,
    [Switch] $ExcludeGuests,
    [Switch] $ExcludeInternalGuests,
    [String] $MailboxNames,
    [String] $RemoveEmailForwardingFromCSV,
    [String] $DisableInboxRuleFromCSV
)

# Function definitions and script logic here

# Connect to Exchange Online
Function ConnectEXO{
    # Check for EXO installation
    $Module=Get-Module ExchangeOnlineManagement -ListAvailable
    if($Module.count -eq 0)
    {
        Write-Host "Exchange online powershell is not available" -ForegroundColor Yellow
        $Confirm = Read-Host "Are you sure want to install module? [Y] Yes [N] No"
        if($Confirm -match "[yY]")
        {
            Write-Host "Installing Exchange Online Powershell module"
            Install-Module -Name ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -scope CurrentUser
            Write-Host "ExchangeOnlineManagement installed successfully..."
        }
        else
        {
            Write-Host "EXO module is required to connect Exchange Online. Please Install-module ExchangeOnlineManagement."
            Exit
        }
    }
    Write-Host "\nConnecting to Exchange Online..."
    try{
        #connect to Exchange Online 
        if(($Organization -ne "") -and ($ClientId -ne "") -and ($CertificateThumbPrint -ne ""))
        {
            #Connect Exchange online using Certificate based Authentication 
            Connect-ExchangeOnline -CertificateThumbprint $CertificateThumbPrint -AppId $ClientId -Organization $Organization -ErrorAction stop -ShowBanner:$false
        }
        elseif(($UserName -ne "") -and ($Password -ne ""))
        {
            #Connect Exchange online using username and password
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
            Connect-ExchangeOnline -Credential $Credential -ErrorAction stop -ShowBanner:$false
        }
        else
        {
            Connect-ExchangeOnline -ErrorAction stop -ShowBanner:$false
        }
    }
    catch
    {
        Write-Host "Error occurred: $($_.Exception.Message )" -ForegroundColor Red
        Exit
    }
    Write-Host "\nExchangeOnline connected successfully" 
}

# Additional functions and script execution logic

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -confirm:$false
Write-Host `n~~ Script prepared by https://github.com/krzyswo/MsAdmin ~~`n -ForegroundColor Green
