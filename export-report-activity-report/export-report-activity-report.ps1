<#
=============================================================================================
Name:           Export User Activity Report to CSV using PowerShell 
Version:        2.0
website:        https://github.com/krzyswo/MsAdmin

Script Highlights:
~~~~~~~~~~~~~~~~~
1.The script uses modern authentication to connect to Exchange Online.  
2.The script can be executed with MFA enabled account too.  
3.Exports report results to CSV file.  
4.Allows you to generate a user activity report for a custom period.  
5.Automatically installs the EXO V2 module (if not installed already) upon your confirmation. 
6.The script is scheduler friendly. I.e., Credential can be passed as a parameter instead of saving inside the script. 

Change Log:
~~~~~~~~~~~
  V1.0 (Jan 07, 2021)  - File created
  V1.1 (Dec 17, 2021)  - Minor usabilities
  V2.0 (May 14, 2025)  - Removed MS Online module dependency, added support for certificate-based authentication, and extended audit log retrieval from 90 to 180 days.

For detailed Script execution: https://github.com/krzyswo/MsAdmin
============================================================================================
#>

Param
(
    [Parameter(Mandatory = $false)]
    [switch]$MFA,
    [switch]$Default,
    [Nullable[DateTime]]$StartDate,
    [Nullable[DateTime]]$EndDate,
    [string]$UserID,
    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$AdminName,
    [string]$Password
)

#Check for EXO module installation
$Module = Get-Module ExchangeOnlineManagement -ListAvailable
if($Module.count -eq 0) 
{ 
 Write-Host "Exchange Online PowerShell module is not available"  -ForegroundColor yellow  
 $Confirm= Read-Host "Are you sure you want to install module? [Y] Yes [N] No" 
 if($Confirm -match "[yY]") 
 { 
  Write-host "Installing Exchange Online PowerShell module"
  Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force
 } 
 else 
 { 
  Write-Host "EXO module is required to connect Exchange Online.Please install module using Install-Module ExchangeOnlineManagement cmdlet." 
  Exit
 }
} 

Write-Host "Connecting to Exchange Online..."
 #Storing credential in script for scheduling purpose/ Passing credential as parameter - Authentication using non-MFA account
 if(($AdminName -ne "") -and ($Password -ne ""))
 {
  $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
  $Credential  = New-Object System.Management.Automation.PSCredential $AdminName,$SecuredPassword
  Connect-ExchangeOnline -Credential $Credential
 }
 elseif($Organization -ne "" -and $ClientId -ne "" -and $CertificateThumbprint -ne "")
 {
   Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint  -Organization $Organization -ShowBanner:$false
 }
 else
 {
  Connect-ExchangeOnline
 }

 $MaxStartDate=((Get-Date).AddDays(-179)).Date


#Retrive audit log for the past 180 days
if(($StartDate -eq $null) -and ($EndDate -eq $null))
{
 $EndDate=(Get-Date).Date
 $StartDate=$MaxStartDate
}
#Getting start date to audit export report
While($true)
{
 if ($StartDate -eq $null)
 {
  $StartDate=Read-Host "Enter start time for report generation '(Eg:12/15/2023)'"
 }
 Try
 {
  $Date=[DateTime]$StartDate
  if($Date -ge $MaxStartDate)
  { 
   break
  }
  else
  {
   Write-Host `
Audit can be retrieved only for the past 180 days. Please select a date after $MaxStartDate -ForegroundColor Red
   return
  }
 }
 Catch
 {
  Write-Host `
Not a valid date -ForegroundColor Red
 }
}


#Getting end date to export audit report
While($true)
{
 if ($EndDate -eq $null)
 {
  $EndDate=Read-Host "Enter End time for report generation '(Eg: 12/15/2023)'"
 }
 Try
 {
  $Date=[DateTime]$EndDate
  if($EndDate -lt ($StartDate))
  {
   Write-Host "End time should be later than start time" -ForegroundColor Red
   return
  }
  break
 }
 Catch
 {
  Write-Host `
Not a valid date -ForegroundColor Red
 }
}

$Location=Get-Location
$OutputCSV="$Location\UserActivityReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv" 
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host "Enter interval time period '(in minutes)'"
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)

#Check whether CurrentEnd exceeds EndDate
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

$AggregateResults = @()
$CurrentResult= @()
$CurrentResultCount=0
$AggregateResultCount=0
Write-Host `
Retrieving user activity log from $StartDate to $EndDate... -ForegroundColor Yellow
$i=0
$ExportResult=""   
$ExportResults=@()  

#Getting user name
if($UserID -eq "")
{ 
 $UserID=Read-Host "Enter user UPN '(eg:John@contoso.com)'"
}

while($true)
{ 
 #Write-Host "Retrieving user activity log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes"
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host "Start and end time are same.Please enter different time range" -ForegroundColor Red
  Exit
 }
 
 #Getting audit log for given time range
 $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -UserIds $UserID -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 $ResultCount=($Results | Measure-Object).count
 $AllAuditData=@()
 foreach($Result in $Results)
 {
  $i++
  $MoreInfo=$Result.auditdata
  $AuditData=$Result.auditdata | ConvertFrom-Json
  $ActivityTime=Get-Date($AuditData.CreationTime) -format g
  $UserID=$AuditData.userId
  $Operation=$AuditData.Operation
  $ResultStatus=$AuditData.ResultStatus
  $Workload=$AuditData.Workload

  #Export result to csv
  $ExportResult=@{'Activity Time'=$ActivityTime;'User Name'=$UserID;'Operation'=$Operation;'Result'=$ResultStatus;'Workload'=$Workload;'More Info'=$MoreInfo}
  $ExportResults= New-Object PSObject -Property $ExportResult  
  $ExportResults | Select-Object 'Activity Time','User Name','Operation','Result','Workload','More Info' | Export-Csv -Path $OutputCSV -Notype -Append 
 }
 Write-Progress -Activity "`
     Retrieving audit log from $StartDate to $EndDate.."`
" Processed audit record count: $i"
 $currentResultCount=$CurrentResultCount+$ResultCount
 if($CurrentResultCount -eq 50000)
 {
  Write-Host "Retrieved max record for current range.Proceeding further may cause data loss or rerun the script with reduced time interval." -ForegroundColor Red
  $Confirm=Read-Host `
Are you sure you want to continue? [Y] Yes [N] No
  if($Confirm -match "[Y]")
  {
   Write-Host "Agg $AggregateResultCount CurrentResu $CurrentResultCount"
   $AggregateResultCount +=$CurrentResultCount
   Write-Host "Proceeding audit log collection with data loss"
   [DateTime]$CurrentStart=$CurrentEnd
   [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
   $CurrentResultCount=0
   $CurrentResult = @()
   if($CurrentEnd -gt $EndDate)
   {
    $CurrentEnd=$EndDate
   }
  }
  else
  {
   Write-Host "Please rerun the script with reduced time interval" -ForegroundColor Red
   Exit
  }
 }
 
 
 if($Results.count -lt 5000)
 {
  #$AggregateResults +=$CurrentResult
  $AggregateResultCount +=$CurrentResultCount
  if($CurrentEnd -eq $EndDate)
  {
   break
  }
  $CurrentStart=$CurrentEnd 
  if($CurrentStart -gt (Get-Date))
  {
   break
  }
  $CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
  $CurrentResultCount=0
  $CurrentResult = @()
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }
 }
}

Write-Host `
~~ Script prepared by AdminDroid Community ~~`
 -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green; Write-Host "https://github.com/krzyswo/MsAdmin" -ForegroundColor Yellow -NoNewline; Write-Host " to get access to 1800+ Microsoft 365 reports. ~~" -ForegroundColor Green `
`
 
 
If($AggregateResultCount -eq 0)
{
 Write-Host "No records found"
}
else
{
 Write-Host `
The output file contains $AggregateResultCount audit records `

 if((Test-Path -Path $OutputCSV) -eq "True") 
 {
  Write-Host " The Output file available in:" -NoNewline -ForegroundColor Yellow
  Write-Host $OutputCSV 
   $Prompt = New-Object -ComObject wscript.shell   
  $UserInput = $Prompt.popup("Do you want to open output file?",`   
 0,"Open Output File",4)   
  If ($UserInput -eq 6)   
  {   
   Invoke-Item "$OutputCSV"   
  } 
 }
}

#Disconnect Exchange Online session
 Disconnect-ExchangeOnline -Confirm:$false | Out-Null
