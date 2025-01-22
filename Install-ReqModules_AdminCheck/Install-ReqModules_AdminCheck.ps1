# Check if the script is being run with administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning 'This script requires administrator privileges. Please run it as an administrator.'
    Exit
}

# List of modules to check for updates
$moduleList = @(
    'MSOnline',
    'AzureAD',
    'Microsoft.Online.SharePoint.PowerShell',
    'ExchangeOnlineManagement',
    'MicrosoftTeams',
    'Microsoft.Graph'
)

# Check for newest version of each module and prompt to update if newer version found
foreach ($module in $moduleList) {
    $currentVersion = (Get-Module -ListAvailable $module).Version
    $newestVersion = (Find-Module $module | Select-Object -First 1).Version

    if ($newestVersion -gt $currentVersion) {
        Write-Host "A new version of module '$module' is available. Current version is '$currentVersion', and the newest version is '$newestVersion'."

        # Prompt to install the update
        $response = Read-Host "Do you want to install the update? (Y/N)"

        if ($response -eq "Y") {
            Install-Module -Name $module -Force
            Write-Host "Module '$module' has been updated to version '$newestVersion'."
        }
    }
    else {
        Write-Host "Module '$module' is up-to-date with version '$currentVersion'."
    }
}

# Import all required modules
Import-Module MSOnline
Import-Module AzureAD
Import-Module Microsoft.Online.SharePoint.PowerShell
Import-Module ExchangeOnlineManagement
Import-Module MicrosoftTeams
Import-Module Microsoft.Graph
