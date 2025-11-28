<#----------------------------------------------
Name: Export SharePoint Document Libraries Report
Version: 1.0
Website: https://github.com/krzyswo/MsAdmin

Script Highlights:
1. Automatically verifies and installs the PnP module upon confirmation.
2. Retrieves document libraries in SharePoint Online with details.
3. Supports single or multiple sites.
4. Compatible with MFA and Certificate-based authentication.
5. Exports results to a CSV file.
6. Scheduler friendly.
----------------------------------------------#>
Param
(
   [Parameter(Mandatory = $false)]
   [String] $UserName,
   [String] $Password,
   [String] $ClientId,
   [String] $CertificateThumbprint,
   [String] $TenantName,
   [String] $SiteAddress,
   [String] $SitesCsv
)
$PnPOnline = (Get-Module PnP.PowerShell -ListAvailable).Name
if($PnPOnline -eq $null)
{
  Write-Host "Important: SharePoint PnP PowerShell module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
  $Confirm= Read-Host 'Are you sure you want to install module? [Y] Yes [N] No'
  if($Confirm -match "[yY]")
  {
    Write-Host "Installing SharePoint PnP PowerShell module..." -ForegroundColor Magenta
    Install-Module PnP.Powershell -Repository PsGallery -Force -AllowClobber
    Import-Module PnP.Powershell -Force
    Register-PnPManagementShellAccess
  }
  else
  {
    Write-Host "Exiting. Note: SharePoint PnP PowerShell module must be available in your system to run the script"
    Exit
  }
}
Write-Host "Connecting to SharePoint PnPPowerShellOnline module..." -ForegroundColor Cyan
function Connect_SharePoint
{
    param
    (
        [Parameter(Mandatory = $true)]
        [String] $Url
    )
    try
    {
        if(($UserName -ne "") -and ($Password -ne "") -and ($TenantName -ne ""))
        {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential  = New-Object System.Management.Automation.PSCredential $UserName,$SecuredPassword
            Connect-PnPOnline -Url $Url -Credential $Credential
        }
        elseif($TenantName -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
        {
            Connect-PnPOnline -Url $Url -ClientId $ClientId -Thumbprint $CertificateThumbprint  -Tenant "$TenantName.onmicrosoft.com"
        }
        else
        {
            Connect-PnPOnline -Url $Url -Interactive
        }
    }
    catch
    {
        Write-Host "Error occurred $($Url) : $_.Exception.Message"   -Foreground Red;
    }
}
if($TenantName -eq "")
{
    $TenantName = Read-Host "Enter your Tenant Name to Connect to SharePoint Online (Example : If your tenant name is 'contoso.com', then enter 'contoso' as a tenant name )  "
}

$AdminUrl = "https://$TenantName.sharepoint.com"
connect_sharepoint -Url $AdminUrl
$Location=Get-Location
$OutputCSV = "$Location\SPO Document Library Report " + ((Get-Date -format "MMM-dd hh-mm-ss tt").ToString()) + ".csv"

function Convert_ToNearestUnit {
    param (
        [long]$LibrarySizeInBytes
    )
    if ($LibrarySizeInBytes -eq 0) {
        return "0 Bytes"
    }
    $Units = ("Bytes", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    $NearestIndex = [Math]::Min([Math]::Floor([Math]::Log($LibrarySizeInBytes, 1024)), $Units.Count - 1)
    $SizeInNearestUnit = [Math]::Round($LibrarySizeInBytes / [Math]::Pow(1024, $NearestIndex), 2)
    $Unit = $Units[$NearestIndex]
    return "$SizeInNearestUnit $Unit"
}
function Get_Statistics
{
    param
    (
        [String] $SiteUrl,
        [String] $SiteTitle
    )
    Get-PnPList  | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false} | ForEach-Object{
        if($_.Title -ne "Form Templates" -and $_.Title -ne "Style Library")
        {
            $LibrarySize = Get-PnPFolderStorageMetric -List $_.Title | Select TotalSize,TotalFileCount
            $FolderCount = $_.ItemCount - $LibrarySize.TotalFileCount
            $FilesCount = $LibrarySize.TotalFileCount
            $LibrarySizeInBytes = $LibrarySize.TotalSize
            $LibrarySize = Convert_ToNearestUnit -LibrarySizeInBytes $LibrarySizeInBytes
            $ExportResult = @{
                "Document Library Name" = $_.Title;
                "Document Library Url" = $AdminUrl+$_.DefaultViewUrl;
                "Created On" = $_.Created;
                "Site Url" =  $SiteUrl;
                "Site Name" = if ($SiteTitle) {$SiteTitle} else { "-" };
                "Library Size(Bytes)" = $LibrarySizeInBytes;
                "Folders Count" = $FolderCount;
                "Files Count" = $FilesCount;
                "Library Size" = $LibrarySize;
            }
            $ExportResult = New-Object PSObject -Property $ExportResult
            $ExportResult | Select-Object "Site Name","Site Url","Document Library Name","Document Library Url","Created On","Library Size","Library Size(Bytes)","Folders Count","Files Count" | Export-Csv -path $OutputCSV -Append -NoTypeInformation
        }
    }
}

if($SiteAddress -ne "")
{
    Connect_SharePoint -Url $SiteAddress
    $Web = Get-PnPWeb | Select Title,Url
    Get_Statistics -SiteUrl $Web.Url -SiteTitle $Web.Title
}
elseif($SitesCsv -ne "")
{
    try
    {
        Import-Csv -path $SitesCsv | ForEach-Object{
            Write-Progress -activity "Processing $($_.SitesUrl)" 
            Connect_Sharepoint -Url $_.SitesUrl 
            $Web = Get-PnPWeb | Select Url,Title 
            Get_Statistics -Objecttype $ObjectType -SiteUrl $Web.Url -SiteTitle $Web.Title
        }
    }
    catch
    {
        Write-Host "Error occurred : $_"   -Foreground Red;
    }
}
else
{
    Get-PnPTenantSite | Select Url,Title | ForEach-Object{
        Write-Progress -activity "Processing $($_.Url)" 
        Connect_SharePoint -Url $_.Url
        Get_Statistics -SiteUrl $_.Url -SiteTitle $_.Title
    }
}

if((Test-Path -Path $OutputCSV) -eq "True")
{
    Write-Host `n "The Output file available in:" -NoNewline -ForegroundColor Yellow; Write-Host "$OutputCSV" `n
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File", 4)

    If ($UserInput -eq 6)
    {
        Invoke-Item $OutputCSV
    }
}

Disconnect-PnPOnline