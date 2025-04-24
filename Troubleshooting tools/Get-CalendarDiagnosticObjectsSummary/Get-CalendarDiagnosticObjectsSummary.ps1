<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 25.04.17.1622

<#
.DESCRIPTION
This Exchange Online script runs the Get-CalendarDiagnosticObjects script and returns a summarized timeline of actions in clear English
as well as the Calendar Diagnostic Objects in Excel.

.PARAMETER Identity
One or more SMTP Address of EXO User Mailbox to query.

.PARAMETER Subject
Subject of the meeting to query, only valid if Identity is a single user.

.PARAMETER MeetingID
The MeetingID of the meeting to query.

.PARAMETER TrackingLogs
Include specific tracking logs in the output. Only usable with the MeetingID parameter.

.PARAMETER Exceptions
Include Exception objects in the output. Only usable with the MeetingID parameter. (Default)

.PARAMETER ExportToExcel
Export the output to an Excel file with formatting.  Running the scrip for multiple users will create multiple tabs in the Excel file. (Default)

.PARAMETER ExportToCSV
Export the output to 3 CSV files per user.

.PARAMETER CaseNumber
Case Number to include in the Filename of the output.

.PARAMETER ShortLogs
Limit Logs to 500 instead of the default 2000, in case the server has trouble responding with the full logs.

.PARAMETER MaxLogs
Increase log limit to 12,000 in case the default 2000 does not contain the needed information. Note this can be time consuming, and it does not contain all the logs such as User Responses.

.PARAMETER CustomProperty
Advanced users can add custom properties to the output in the RAW output. This is not recommended unless you know what you are doing. The properties must be in the format of "PropertyName1, PropertyName2, PropertyName3".  The properties will be added to the RAW output and not the Timeline output.  The properties must be in the format of "PropertyName1, PropertyName2, PropertyName3".  The properties will only be added to the RAW output.

.PARAMETER ExceptionDate
Date of the Exception Meeting to collect logs for.  Fastest way to get Exceptions for a meeting.

.PARAMETER NoExceptions
Do not collect Exception Meetings.  This was the default behavior of the script, now exceptions are collected by default.

.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity someuser@microsoft.com -MeetingID 040000008200E00074C5B7101A82E008000000008063B5677577D9010000000000000000100000002FCDF04279AF6940A5BFB94F9B9F73CD
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity someuser@microsoft.com -Subject "Test One Meeting Subject"
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity User1, User2, Delegate -MeetingID $MeetingID
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -TrackingLogs -NoExceptions
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -TrackingLogs -Exceptions -ExportToExcel -CaseNumber 123456
.EXAMPLE
Get-CalendarDiagnosticObjectsSummary.ps1 -Identity $Users -MeetingID $MeetingID -TrackingLogs -ExceptionDate "01/28/2024" -CaseNumber 123456

.SYNOPSIS
Used to collect easy to read Calendar Logs.

.LINK
    https://aka.ms/callogformatter
#>

[CmdletBinding(DefaultParameterSetName = 'Subject',
    SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory, Position = 0, HelpMessage = "Enter the Identity of the mailbox(es) to query. Press <Enter> again when done.")]
    [string[]]$Identity,
    [Parameter(HelpMessage = "Export all Logs to Excel (Default).")]
    [switch]$ExportToExcel,
    [Parameter(HelpMessage = "Export all Logs to CSV files.")]
    [switch]$ExportToCSV,
    [Parameter(HelpMessage = "Case Number to include in the Filename of the output.")]
    [string]$CaseNumber,
    [Parameter(HelpMessage = "Limit Logs to 500 instead of the default 2000, in case the server has trouble responding with the full logs.")]
    [switch]$ShortLogs,
    [Parameter(HelpMessage = "Limit Logs to 12000 instead of the default 2000, in case the server has trouble responding with the full logs.")]
    [switch]$MaxLogs,
    [Parameter(HelpMessage = "Custom Property to add to the RAW output.")]
    [string[]]$CustomProperty,

    [Parameter(Mandatory, ParameterSetName = 'MeetingID', Position = 1, HelpMessage = "Enter the MeetingID of the meeting to query. Recommended way to search for CalLogs.")]
    [string]$MeetingID,
    [Parameter(HelpMessage = "Include specific tracking logs in the output. Only usable with the MeetingID parameter.")]
    [switch]$TrackingLogs,
    [Parameter(HelpMessage = "Include Exception objects in the output. Only usable with the MeetingID parameter.")]
    [switch]$Exceptions,
    [Parameter(HelpMessage = "Date of the Exception to collect the logs for.")]
    [DateTime]$ExceptionDate,
    [Parameter(HelpMessage = "Do Not collect Exception Meetings.")]
    [switch]$NoExceptions,

    [Parameter(Mandatory, ParameterSetName = 'Subject', Position = 1, HelpMessage = "Enter the Subject of the meeting. Do not include the RE:, FW:, etc.,  No wild cards (* or ?)")]
    [string]$Subject
)

# ===================================================================================================
# Auto update script
# ===================================================================================================
$BuildVersion = "25.04.17.1622"




function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebProxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            Write-Verbose "Proxy server configuration detected"
            Write-Verbose $proxyObject.OriginalString
            return $true
        } else {
            Write-Verbose "No proxy server configuration detected"
            return $false
        }
    } catch {
        Write-Verbose "Unable to check for proxy server configuration"
        return $false
    }
}

function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.InvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.InvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Remote Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = "Default")]
        [string]
        $Uri,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $true, ParameterSetName = "ParametersObject")]
        [hashtable]
        $ParametersObject,

        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [string]
        $OutFile
    )

    Write-Verbose "Calling $($MyInvocation.MyCommand)"
    if ([System.String]::IsNullOrEmpty($Uri)) {
        $Uri = $ParametersObject.Uri
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    if ($null -eq $ParametersObject) {
        $params = @{
            Uri     = $Uri
            OutFile = $OutFile
        }

        if ($UseBasicParsing) {
            $params.UseBasicParsing = $true
        }
    } else {
        $params = $ParametersObject
    }

    try {
        Invoke-WebRequest @params
    } catch {
        Write-VerboseErrorInformation
    }
}

<#
    Determines if the script has an update available.
#>
function Get-ScriptUpdateAvailable {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $BuildVersion = "25.04.17.1622"

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $result = [PSCustomObject]@{
        ScriptName     = $scriptName
        CurrentVersion = $BuildVersion
        LatestVersion  = ""
        UpdateFound    = $false
        Error          = $null
    }

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
    } else {
        try {
            $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequestWithProxyDetection -Uri $VersionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
            $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
            $result.LatestVersion = $latestVersion
            if ($null -ne $latestVersion) {
                $result.UpdateFound = ($latestVersion -ne $BuildVersion)
            } else {
                Write-Warning ("Unable to check for a script update as no script with the same name was found." +
                    "`r`nThis can happen if the script has been renamed. Please check manually if there is a newer version of the script.")
            }

            Write-Verbose "Current version: $($result.CurrentVersion) Latest version: $($result.LatestVersion) Update found: $($result.UpdateFound)"
        } catch {
            Write-Verbose "Unable to check for updates: $($_.Exception)"
            $result.Error = $_
        }
    }

    return $result
}


function Confirm-Signature {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $File
    )

    $IsValid = $false
    $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
    $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

    try {
        $sig = Get-AuthenticodeSignature -FilePath $File

        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        if (-not $chain.Build($sig.SignerCertificate)) {
            Write-Warning "Signer certificate doesn't chain correctly."
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            Write-Warning "Certificate Chain shorter than expected."
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            Write-Warning "Top-level certificate in chain is not a root certificate."
            throw
        }

        if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
            Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
            throw
        }

        Write-Host "File signed by $($sig.SignerCertificate.Subject)"

        $IsValid = $true
    } catch {
        $IsValid = $false
    }

    $IsValid
}

<#
.SYNOPSIS
    Overwrites the current running script file with the latest version from the repository.
.NOTES
    This function always overwrites the current file with the latest file, which might be
    the same. Get-ScriptUpdateAvailable should be called first to determine if an update is
    needed.

    In many situations, updates are expected to fail, because the server running the script
    does not have internet access. This function writes out failures as warnings, because we
    expect that Get-ScriptUpdateAvailable was already called and it successfully reached out
    to the internet.
#>
function Invoke-ScriptUpdate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([boolean])]
    param ()

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)
    $tempFullName = (Join-Path ((Get-Item $env:TEMP).FullName) $scriptName)

    if ($PSCmdlet.ShouldProcess("$scriptName", "Update script to latest version")) {
        try {
            Invoke-WebRequestWithProxyDetection -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName
        } catch {
            Write-Warning "AutoUpdate: Failed to download update: $($_.Exception.Message)"
            return $false
        }

        try {
            if (Confirm-Signature -File $tempFullName) {
                Write-Host "AutoUpdate: Signature validated."
                if (Test-Path $oldFullName) {
                    Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Move-Item $scriptFullName $oldFullName
                Move-Item $tempFullName $scriptFullName
                Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                Write-Host "AutoUpdate: Succeeded."
                return $true
            } else {
                Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                Write-Warning "AutoUpdate: Update was not applied."
            }
        } catch {
            Write-Warning "AutoUpdate: Failed to apply update: $($_.Exception.Message)"
        }
    }

    return $false
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Pass -Confirm:$false
    to update without prompting the user. Pass -Verbose for additional
    diagnostic output.

    Returns $true if an update was downloaded, $false otherwise. The
    result will always be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Need to pass through ShouldProcess settings to Invoke-ScriptUpdate')]
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $false)]
        [switch]
        $AutoUpdate,
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $updateInfo = Get-ScriptUpdateAvailable $VersionsUrl
    if ($updateInfo.UpdateFound) {
        if ($AutoUpdate) {
            return Invoke-ScriptUpdate
        } else {
            Write-Warning "$($updateInfo.ScriptName) $BuildVersion is outdated. Please download the latest, version $($updateInfo.LatestVersion)."
        }
    }

    return $false
}
if (Test-ScriptVersion -AutoUpdate -VersionsUrl "https://aka.ms/CL-VersionsUrl" -Confirm:$false) {
    # Update was downloaded, so stop here.
    Write-Host -ForegroundColor Red "Script was updated. Please rerun the command." -ForegroundColor Yellow
    return
}

$script:command = $MyInvocation
Write-Verbose "The script was started with the following command line:"
Write-Verbose "Name:  $($script:command.MyCommand.name)"
Write-Verbose "Command Line:  $($script:command.line)"
Write-Verbose "Script Version: $BuildVersion"
$script:BuildVersion = $BuildVersion

# ===================================================================================================
# Support scripts
# ===================================================================================================

# ===================================================================================================
# Constants to support the script
# ===================================================================================================

$script:CalendarItemTypes = @{
    'IPM.Schedule.Meeting.Request.AttendeeListReplication' = "AttendeeList"
    'IPM.Schedule.Meeting.Canceled'                        = "Cancellation"
    'IPM.OLE.CLASS.{00061055-0000-0000-C000-000000000046}' = "Exception"
    'IPM.Schedule.Meeting.Notification.Forward'            = "Forward.Notification"
    'IPM.Appointment'                                      = "Ipm.Appointment"
    'IPM.Appointment.MP'                                   = "Ipm.Appointment"
    'IPM.Schedule.Meeting.Request'                         = "Meeting.Request"
    'IPM.CalendarSharing.EventUpdate'                      = "SharingCFM"
    'IPM.CalendarSharing.EventDelete'                      = "SharingDelete"
    'IPM.Schedule.Meeting.Resp'                            = "Resp.Any"
    'IPM.Schedule.Meeting.Resp.Neg'                        = "Resp.Neg"
    'IPM.Schedule.Meeting.Resp.Tent'                       = "Resp.Tent"
    'IPM.Schedule.Meeting.Resp.Pos'                        = "Resp.Pos"
    '(Occurrence Deleted)'                                 = "Exception.Deleted"
}

# ===================================================================================================
# Functions to support the script
# ===================================================================================================

<#
.SYNOPSIS
Looks to see if there is a Mapping of ExternalMasterID to FolderName
#>
function MapSharedFolder {
    param(
        $ExternalMasterID
    )
    if ($ExternalMasterID -eq "NotFound") {
        return "Not Shared"
    } else {
        $SharedFolders[$ExternalMasterID]
    }
}

<#
.SYNOPSIS
Replaces a value of NotFound with a blank string.
#>
function ReplaceNotFound {
    param (
        $Value
    )
    if ($Value -eq "NotFound") {
        return ""
    } else {
        return $Value
    }
}

<#
.SYNOPSIS
Creates a Mapping of ExternalMasterID to FolderName
#>
function CreateExternalMasterIDMap {
    # This function will create a Map of the log folder to ExternalMasterID
    $script:SharedFolders = [System.Collections.SortedList]::new()
    Write-Verbose "Starting CreateExternalMasterIDMap"

    foreach ($ExternalID in $script:GCDO.ExternalSharingMasterId | Select-Object -Unique) {
        if ($ExternalID -eq "NotFound") {
            continue
        }

        $AllFolderNames = @($script:GCDO | Where-Object { $_.ExternalSharingMasterId -eq $ExternalID } | Select-Object -ExpandProperty OriginalParentDisplayName | Select-Object -Unique)

        if ($AllFolderNames.count -gt 1) {
            # We have 2+ FolderNames, Need to find the best one. Remove 'Calendar' from possible names
            $AllFolderNames = $AllFolderNames | Where-Object { $_ -notmatch 'Calendar' } # Need a better way to do this for other languages...
        }

        if ($AllFolderNames.Count -eq 0) {
            $SharedFolders[$ExternalID] = "UnknownSharedCalendarCopy"
            Write-Host -ForegroundColor red "Found Zero to map to."
        }

        if ($AllFolderNames.Count -eq 1) {
            $SharedFolders[$ExternalID] = $AllFolderNames
            Write-Verbose "Found map: [$AllFolderNames] is for $ExternalID"
        } else {
            # we still have multiple possible Folder Names, need to chose one or combine
            Write-Host -ForegroundColor Red "Unable to Get Exact Folder for $ExternalID"
            Write-Host -ForegroundColor Red "Found $($AllFolderNames.count) possible folders"

            if ($AllFolderNames.Count -eq 2) {
                $SharedFolders[$ExternalID] = $AllFolderNames[0] + $AllFolderNames[1]
            } else {
                $SharedFolders[$ExternalID] = "UnknownSharedCalendarCopy"
            }
        }
    }

    Write-Host -ForegroundColor Green "Created the following Shared Calendar Mapping:"
    foreach ($Key in $SharedFolders.Keys) {
        Write-Host -ForegroundColor Green "$Key : $($SharedFolders[$Key])"
    }
    # ToDo: Need to check for multiple ExternalSharingMasterId pointing to the same FolderName
    Write-Verbose "Created the following Mapping :"
    Write-Verbose $SharedFolders
}

<#
.SYNOPSIS
Convert a csv value to multiLine.
#>
function MultiLineFormat {
    param(
        $PassedString
    )
    $PassedString = $PassedString -replace "},", "},`n"
    return $PassedString.Trim()
}

# ===================================================================================================
# Build CSV to output
# ===================================================================================================

<#
.SYNOPSIS
Builds the CSV output from the Calendar Diagnostic Objects
#>
function BuildCSV {

    Write-Host "Starting to Process Calendar Logs..."
    $GCDOResults = @()
    $script:MailboxList = @{}
    Write-Host "Creating Map of Mailboxes to CNs..."
    CreateExternalMasterIDMap
    ConvertCNtoSMTP
    FixCalendarItemType($script:GCDO)

    Write-Host "Making Calendar Logs more readable..."
    $Index = 0
    foreach ($CalLog in $script:GCDO) {
        $Index++
        $ItemType = $CalendarItemTypes.($CalLog.ItemClass)

        # CleanNotFounds
        $PropsToClean = "FreeBusyStatus", "ClientIntent", "AppointmentSequenceNumber", "AppointmentLastSequenceNumber", "RecurrencePattern", "AppointmentAuxiliaryFlags", "EventEmailReminderTimer", "IsSeriesCancelled", "AppointmentCounterProposal", "MeetingRequestType", "SendMeetingMessagesDiagnostics", "AttendeeCollection"
        foreach ($Prop in $PropsToClean) {
            # Exception objects, etc. don't have these properties.
            if ($null -ne $CalLog.$Prop) {
                $CalLog.$Prop = ReplaceNotFound($CalLog.$Prop)
            }
        }

        # Record one row
        $GCDOResults += [PSCustomObject]@{
            'LogRow'                         = $Index
            'LogTimestamp'                   = ConvertDateTime($CalLog.LogTimestamp)
            'LogRowType'                     = $CalLog.LogRowType.ToString()
            'SubjectProperty'                = $CalLog.SubjectProperty
            'Client'                         = $CalLog.ShortClientInfoString
            'LogClientInfoString'            = $CalLog.LogClientInfoString
            'TriggerAction'                  = $CalLog.CalendarLogTriggerAction
            'ItemClass'                      = $ItemType
            'Seq:Exp:ItemVersion'            = CompressVersionInfo($CalLog)
            'Organizer'                      = GetDisplayName($CalLog.From)
            'From'                           = GetSMTPAddress($CalLog.From)
            'FreeBusy'                       = $CalLog.FreeBusyStatus.ToString()
            'ResponsibleUser'                = GetSMTPAddress($CalLog.ResponsibleUserName)
            'Sender'                         = GetSMTPAddress($CalLog.Sender)
            'LogFolder'                      = $CalLog.ParentDisplayName
            'OriginalLogFolder'              = $CalLog.OriginalParentDisplayName
            'SharedFolderName'               = MapSharedFolder($CalLog.ExternalSharingMasterId)
            'ReceivedRepresenting'           = GetSMTPAddress($CalLog.ReceivedRepresenting)
            'MeetingRequestType'             = $CalLog.MeetingRequestType.ToString()
            'StartTime'                      = ConvertDateTime($CalLog.StartTime)
            'EndTime'                        = ConvertDateTime($CalLog.EndTime)
            'OriginalStartDate'              = ConvertDateTime($CalLog.OriginalStartDate)
            'Location'                       = $CalLog.Location
            'CalendarItemType'               = $CalLog.CalendarItemType.ToString()
            'RecurrencePattern'              = $CalLog.RecurrencePattern
            'AppointmentAuxiliaryFlags'      = $CalLog.AppointmentAuxiliaryFlags.ToString()
            'DisplayAttendeesAll'            = $CalLog.DisplayAttendeesAll
            'AttendeeCount'                  = GetAttendeeCount($CalLog.DisplayAttendeesAll)
            'AppointmentState'               = $CalLog.AppointmentState.ToString()
            'ResponseType'                   = $CalLog.ResponseType.ToString()
            'ClientIntent'                   = $CalLog.ClientIntent.ToString()
            'AppointmentRecurring'           = $CalLog.AppointmentRecurring
            'HasAttachment'                  = $CalLog.HasAttachment
            'IsCancelled'                    = $CalLog.IsCancelled
            'IsAllDayEvent'                  = $CalLog.IsAllDayEvent
            'IsSeriesCancelled'              = $CalLog.IsSeriesCancelled
            'SendMeetingMessagesDiagnostics' = $CalLog.SendMeetingMessagesDiagnostics
            'AttendeeCollection'             = MultiLineFormat($CalLog.AttendeeCollection)
            'CalendarLogRequestId'           = $CalLog.CalendarLogRequestId.ToString()    # Move to front.../ Format in groups???
        }
    }
    $script:EnhancedCalLogs = $GCDOResults

    Write-Host -ForegroundColor Green "Calendar Logs have been processed, Exporting logs to file..."
    Export-CalLog
}

function ConvertDateTime {
    param(
        [string] $DateTime
    )
    if ([string]::IsNullOrEmpty($DateTime) -or
        $DateTime -eq "N/A" -or
        $DateTime -eq "NotFound") {
        return ""
    }
    return [DateTime]$DateTime
}

function GetAttendeeCount {
    param(
        [string] $AttendeesAll
    )
    if ($AttendeesAll -ne "NotFound") {
        return ($AttendeesAll -split ';').Count
    } else {
        return "-"
    }
}

<#
.SYNOPSIS
Corrects the CalenderItemType column
#>
function FixCalendarItemType {
    param(
        $CalLogs
    )
    foreach ($CalLog in $CalLogs) {
        if ($CalLog.OriginalStartDate -ne "NotFound" -and ![string]::IsNullOrEmpty($CalLog.OriginalStartDate)) {
            $CalLog.CalendarItemType = "Exception"
            $CalLog.isException = $true
        }
    }
}

function CompressVersionInfo {
    param(
        $CalLog
    )
    [string] $CompressedString = ""
    if ($CalLog.AppointmentSequenceNumber -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.AppointmentSequenceNumber)) {
        $CompressedString = "-:"
    } else {
        $CompressedString = $CalLog.AppointmentSequenceNumber.ToString() + ":"
    }
    if ($CalLog.AppointmentLastSequenceNumber -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.AppointmentLastSequenceNumber)) {
        $CompressedString += "-:"
    } else {
        $CompressedString += $CalLog.AppointmentLastSequenceNumber.ToString() + ":"
    }
    if ($CalLog.ItemVersion -eq "NotFound" -or [string]::IsNullOrEmpty($CalLog.ItemVersion)) {
        $CompressedString += "-"
    } else {
        $CompressedString += $CalLog.ItemVersion.ToString()
    }

    return $CompressedString
}

# ===================================================================================================
# BuildTimeline
# ===================================================================================================

<#
.SYNOPSIS
    Tries to builds a timeline of the history of the meeting based on the diagnostic objects.

.DESCRIPTION
    By using the time sorted diagnostic objects for one user on one meeting, we try to give a high level
    overview of what happened to the meeting. This can be use to get a quick overview of the meeting and
    then you can look into the CalLog in Excel to get more details.

    The timeline will skip a lot of the noise (isIgnorable) in the CalLogs. It skips EBA (Event Based Assistants),
    and other EXO internal processes, which are (99% of the time) not interesting to the end user and just setting
    hidden internal properties (i.e. things like HasBeenIndex, etc.)

    It also skips items from Shared Calendars, which are calendars that have a Modern Sharing relationship setup,
    which creates a replicated copy of another users. If you want to look at the actions this user took on
    another users calendar, you can look at that users Calendar Logs.

.NOTES
    The timeline will never be perfect, but if you see a way to make it more understandable, readable, etc.,
    please let me know or fix it yourself on GitHub.
    I use a iterative approach to building this, so it will get better over time.
#>

function FindOrganizer {
    param (
        $CalLog
    )
    $Script:Organizer = "Unknown"
    if ($null -ne $CalLog.From) {
        if ($null -ne $CalLog.From.SmtpEmailAddress) {
            $Script:Organizer = $($CalLog.From.SmtpEmailAddress)
        } elseif ($null -ne $CalLog.From.DisplayName) {
            $Script:Organizer = $($CalLog.From.DisplayName)
        } elseif ($calLog.From -match "^\s*<") {
            $Script:Organizer = $($CalLog.From -split "<")[1] -replace ">", ""
        } else {
            $Script:Organizer = $($CalLog.From)
        }
    }
    Write-Host -ForegroundColor Green "Setting Organizer to : [$Script:Organizer]"
}

function FindFirstMeeting {
    [array]$IpmAppointments = $script:GCDO | Where-Object { $_.ItemClass -eq "IPM.Appointment" -and $_.ExternalSharingMasterId -eq "NotFound" }
    if ($IpmAppointments.count -eq 0) {
        Write-Host "All CalLogs are from Shared Calendar, getting values from first IPM.Appointment."
        $IpmAppointments = $script:GCDO | Where-Object { $_.ItemClass -eq "IPM.Appointment" }
    }
    if ($IpmAppointments.count -eq 0) {
        Write-Host -ForegroundColor Red "Warning: Cannot find any IPM.Appointments, if this is the Organizer, check for the Outlook Bifurcation issue."
        Write-Host -ForegroundColor Red "Warning: No IPM.Appointment found. CalLogs start to expire after 31 days."
        return $null
    } else {
        return $IpmAppointments[0]
    }
}

function BuildTimeline {
    $script:TimeLineOutput = @()

    $script:FirstLog = FindFirstMeeting
    FindOrganizer($script:FirstLog)

    # Ignorable and items from Shared Calendars are not included in the TimeLine.
    [array]$InterestingCalLogs = $script:EnhancedCalLogs | Where-Object { $_.LogRowType -eq "Interesting" -and $_.SharedFolderName -eq "Not Shared" }

    if ($InterestingCalLogs.count -eq 0) {
        Write-Host "All CalLogs are Ignorable, nothing to create a timeline with, displaying initial values."
    } else {
        Write-Host "Found $($script:EnhancedCalLogs.count) Log entries, only the $($InterestingCalLogs.count) Non-Ignorable entries will be analyzed in the TimeLine. `n"
    }

    if ($script:CalLogsDisabled) {
        Write-Host -ForegroundColor Red "Warning: CalLogs are disabled for this user, Timeline / CalLogs will be incomplete."
        return
    }

    Write-DashLineBoxColor "  TimeLine for: [$Identity]",
    "CollectionDate: $($(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))",
    "ScriptVersion: $ScriptVersion",
    "  Subject: $($script:GCDO[0].NormalizedSubject)",
    "  Organizer: $Script:Organizer",
    "  MeetingID: $($script:GCDO[0].CleanGlobalObjectId)"
    [array]$Header = "MeetingID: "+ ($script:GCDO[0].CleanGlobalObjectId)

    CreateMeetingSummary -Time "Calendar Timeline for Meeting" -MeetingChanges $Header
    if ($null -ne $FirstLog) {
        CreateMeetingSummary -Time "Initial Message Values" -Entry $script:FirstLog -LongVersion
    }

    # Look at each CalLog and build the Timeline
    foreach ($CalLog in $InterestingCalLogs) {
        [bool] $script:MeetingSummaryNeeded = $False
        [bool] $script:AddChangedProperties = $False

        $MeetingChanges = CreateTimelineRow
        # Create the Timeline by adding to Time to the generated MeetingChanges
        $Time = "$($($CalLog.LogRow).toString().PadRight(5)) -- $(ConvertDateTime($CalLog.LogTimestamp))"

        if ($MeetingChanges) {
            if ($script:MeetingSummaryNeeded) {
                CreateMeetingSummary -Time $Time -MeetingChanges $MeetingChanges
                CreateMeetingSummary -Time " " -ShortVersion -Entry $CalLog
            } else {
                CreateMeetingSummary -Time $Time -MeetingChanges $MeetingChanges
                if ($script:AddChangedProperties) {
                    FindChangedProperties
                }
            }
        }

        # Setup Previous log (if current logs is an IPM.Appointment)
        if ($CalendarItemTypes.($CalLog.ItemClass) -eq "Ipm.Appointment" -or $CalendarItemTypes.($CalLog.ItemClass) -eq "Exception") {
            $script:PreviousCalLog = $CalLog
        }
    }

    Export-Timeline
}
function Convert-Data {
    param(
        [Parameter(Mandatory = $True)]
        [string[]] $ArrayNames,
        [switch ] $NoWarnings = $False
    )
    $ValidArrays = @()
    $ItemCounts = @()
    $VariableLookup = @{}
    foreach ($Array in $ArrayNames) {
        try {
            $VariableData = Get-Variable -Name $Array -ErrorAction Stop
            $VariableLookup[$Array] = $VariableData.Value
            $ValidArrays += $Array
            $ItemCounts += ($VariableData.Value | Measure-Object).Count
        } catch {
            if (!$NoWarnings) {
                Write-Warning -Message "No variable found for [$Array]"
            }
        }
    }
    $MaxItemCount = ($ItemCounts | Measure-Object -Maximum).Maximum
    $FinalArray = @()
    for ($Inc = 0; $Inc -lt $MaxItemCount; $Inc++) {
        $FinalObj = New-Object PsObject
        foreach ($Item in $ValidArrays) {
            $FinalObj | Add-Member -MemberType NoteProperty -Name $Item -Value $VariableLookup[$Item][$Inc]
        }
        $FinalArray += $FinalObj
    }

    return $FinalArray
    $FinalArray = @()
}

# ===================================================================================================
# Write Out one line of the Meeting Summary (Time + Meeting Changes)
# ===================================================================================================
function CreateMeetingSummary {
    param(
        [array] $Time,
        [array] $MeetingChanges,
        $Entry,
        [switch] $LongVersion,
        [switch] $ShortVersion
    )

    $InitialSubject = "Subject: " + $Entry.NormalizedSubject
    $InitialOrganizer = "Organizer: " + $Entry.SentRepresentingDisplayName
    $InitialSender = "Sender: " + $Entry.SentRepresentingDisplayName
    $InitialToList = "To List: " + $Entry.DisplayAttendeesAll
    $InitialLocation = "Location: " + $Entry.Location

    if ($ShortVersion -or $LongVersion) {
        $InitialStartTime = "StartTime: " + $Entry.StartTime.ToString()
        $InitialEndTime = "EndTime: " + $Entry.EndTime.ToString()
    }

    if ($longVersion -and ($Entry.Timezone -ne "")) {
        $InitialTimeZone = "Time Zone: " + $Entry.Timezone
    } else {
        $InitialTimeZone = "Time Zone: Not Populated"
    }

    if ($Entry.AppointmentRecurring) {
        $InitialRecurring = "Recurring: Yes - Recurring"
    } else {
        $InitialRecurring = "Recurring: No - Single instance"
    }

    if ($longVersion -and $Entry.AppointmentRecurring) {
        $InitialRecurrencePattern = "RecurrencePattern: " + $Entry.RecurrencePattern
        $InitialSeriesStartTime = "Series StartTime: " + $Entry.StartTime.ToString() + "Z"
        $InitialSeriesEndTime = "Series EndTime: " + $Entry.StartTime.ToString() + "Z"
        if (!$Entry.ViewEndTime) {
            $InitialEndDate = "Meeting Series does not have an End Date."
        }
    }

    if (!$Time) {
        $Time = $Entry.LogTimestamp
    }

    if (!$MeetingChanges) {
        $MeetingChanges = @()
        $MeetingChanges += $InitialSubject, $InitialOrganizer, $InitialSender, $InitialToList, $InitialLocation, $InitialStartTime, $InitialEndTime, $InitialTimeZone, $InitialRecurring, $InitialRecurrencePattern, $InitialSeriesStartTime , $InitialSeriesEndTime , $InitialEndDate
    }

    if ($ShortVersion) {
        $MeetingChanges = @()
        $MeetingChanges += $InitialToList, $InitialLocation, $InitialStartTime, $InitialEndTime, $InitialRecurring
    }

    $script:TimeLineOutput += Convert-Data -ArrayNames "Time", "MeetingChanges"
}

$WellKnownCN_CA = "MICROSOFT SYSTEM ATTENDANT"
$CalAttendant = "Calendar Assistant"
$WellKnownCN_Trans = "MicrosoftExchange"
$Transport = "Transport Service"
<#
.SYNOPSIS
Get the Mailbox for the Passed in Identity.
Might want to extend to do 'Get-MailUser' as well.
.PARAMETER CN of the Mailbox
    The mailbox for which to retrieve properties.
.PARAMETER Organization
    [Optional] Organization to search for the mailbox in.
#>
function GetMailbox {
    param(
        [string]$Identity,
        [string]$Organization,
        [bool]$UseGetMailbox
    )

    $params = @{Identity = $Identity
        ErrorAction      = "SilentlyContinue"
    }

    if ($UseGetMailbox) {
        $Cmdlet = "Get-Mailbox"
        $params.Add("IncludeInactiveMailbox", $true)
    } else {
        $Cmdlet = "Get-Recipient"
    }

    try {
        Write-Verbose "Searching $Cmdlet $(if (-not ([string]::IsNullOrEmpty($Organization))) {"with Org: $Organization"}) for $Identity."

        if (-not ([string]::IsNullOrEmpty($Organization)) -and $script:MSSupport) {
            Write-Verbose "Using Organization parameter"
            $params.Add("Organization", $Organization)
        } elseif (-not ([string]::IsNullOrEmpty($Organization))) {
            Write-Verbose "Using -OrganizationalUnit parameter with $Organization."
            $params.Add("Organization", $Organization)
        }

        Write-Verbose "Running $Cmdlet with params: $($params.Values)"
        $RecipientOutput = & $Cmdlet @params
        Write-Verbose "RecipientOutput: $RecipientOutput"

        if (!$RecipientOutput) {
            Write-Host "Unable to find [$Identity]$(if ($Organization -ne `"`" ) {" in Organization:[$Organization]"})."
            Write-Host "Trying to find a Group Mailbox for [$Identity]..."
            $RecipientOutput = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue -GroupMailbox
            if (!$RecipientOutput) {
                Write-Host "Unable to find a Group Mailbox for [$Identity] either."
                return $null
            } else {
                Write-Verbose "Found GroupMailbox [$($RecipientOutput.DisplayName)]"
            }
        }

        if ($null -eq $script:PIIAccess) {
            [bool]$script:PIIAccess = CheckForPIIAccess($RecipientOutput.DisplayName)
        }

        if ($script:PIIAccess) {
            Write-Verbose "Found [$($RecipientOutput.DisplayName)]"
        } else {
            Write-Verbose "No PII Access for [$Identity]"
        }

        return $RecipientOutput
    } catch {
        Write-Error "An error occurred while running ${Cmdlet}: [$_]"
    }
}

<#
.SYNOPSIS
Checks the identities are EXO Mailboxes.
#>
function CheckIdentities {
    if (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue) {
        Write-Host "Validated connection to Exchange Online..."
    } else {
        Write-Error "Get-Mailbox cmdlet not found. Please validate that you are running this script from an Exchange Management Shell and try again."
        Write-Host "Look at Import-Module ExchangeOnlineManagement and Connect-ExchangeOnline."
        exit
    }

    # See if it is a Customer Tenant running the cmdlet. (They will not have access to Organization parameter)
    $script:MSSupport = [Bool](Get-Help Get-Mailbox -Parameter Organization -ErrorAction SilentlyContinue)
    Write-Verbose "MSSupport: $script:MSSupport"

    Write-Host "Checking for at least one valid mailbox..."
    $IdentityList = @()

    Write-Host "Preparing to check $($Identity.count) Mailbox(es)..."

    foreach ($Id in $Identity) {
        $Account = GetMailbox -Identity $Id -UseGetMailbox $true
        if ($null -eq $Account) {
            # -or $script:MB.GetType().FullName -ne "Microsoft.Exchange.Data.Directory.Management.Mailbox") {
            Write-DashLineBoxColor "`n Error: Mailbox [$Id] not found on Exchange Online.  Please validate the mailbox name and try again.`n" -Color Red
            continue
        }
        if (-not (CheckForPIIAccess($Account.DisplayName))) {
            Write-Host -ForegroundColor DarkRed "No PII access for Mailbox [$Id]. Falling back to SMTP Address."
            $IdentityList += $ID
            if ($null -eq $script:MB) {
                $script:MB = $Account
            }
        } else {
            Write-Host "Mailbox [$Id] found as : $($Account.DisplayName)"
            $IdentityList += $Account.PrimarySmtpAddress.ToString()
            if ($null -eq $script:MB) {
                $script:MB = $Account
            }
        }
        if ($Account.CalendarVersionStoreDisabled -eq $true) {
            [bool]$script:CalLogsDisabled = $true
            Write-Host -ForegroundColor DarkRed "Mailbox [$Id] has CalendarVersionStoreDisabled set to True.  This mailbox will not have Calendar Logs."
            Write-Host -ForegroundColor DarkRed "Some logs will be available for Mailbox [$Id] but they will not be complete."
        }
        if ($Account.RecipientTypeDetails -eq "RoomMailbox" -or $Account.RecipientTypeDetails -eq "EquipmentMailbox") {
            if ($script:PIIAccess -eq $true) {
                $script:Rooms += $Account.PrimarySmtpAddress.ToString()
            } else {
                $script:Rooms += $Id
            }
            Write-Host -ForegroundColor Green "[$Id] is a Room / Equipment Mailbox."
        }
    }

    Write-Verbose "IdentityList: $IdentityList"

    if ($IdentityList.count -eq 0) {
        Write-DashLineBoxColor "`n No valid mailboxes found.  Please validate the mailbox name and try again. `n" Red
        exit
    }

    return $IdentityList
}

<#
.SYNOPSIS
Creates a list of CN that are used in the Calendar Logs, Looks up the Mailboxes and stores them in the MailboxList.
#>
function ConvertCNtoSMTP {
    # Creates a list of CN's that we will do MB look up on
    $CNEntries = @()
    $CNEntries += ($script:GCDO.ResponsibleUserName.ToUpper() | Select-Object -Unique)
    $CNEntries += ($script:GCDO.SentRepresentingEmailAddress.ToUpper() | Select-Object -Unique)
    $CNEntries += ($script:GCDO.Sender.ToUpper() | Select-Object -Unique)
    # $CNEntries += ($script:GCDO.SenderEmailAddress.ToUpper() | Select-Object -Unique)
    $CNEntries = $CNEntries | Select-Object -Unique
    Write-Verbose "`t Have $($CNEntries.count) CN Entries to look for..."
    Write-Verbose "CNEntries: "; foreach ($CN in $CNEntries) { Write-Verbose `t`t$CN }

    $Org = $script:MB.OrganizationalUnit.split('/')[-1]

    # Creates a Dictionary of MB's that we will use to look up the CN's
    Write-Verbose "Converting CN entries into SMTP Addresses..."

    foreach ($CNEntry in $CNEntries) {
        if ($CNEntry -match 'cn=([\w,\s.@-]*[^/])$') {
            if ($CNEntry -match $WellKnownCN_CA) {
                $script:MailboxList[$CNEntry] = $CalAttendant
            } elseif ($CNEntry -match $WellKnownCN_Trans) {
                $script:MailboxList[$CNEntry] = $Transport
            } else {
                $script:MailboxList[$CNEntry] = (GetMailbox -Identity $CNEntry -Organization $Org)
            }
        }
        # New more readable format!
        else {
            if ( $CNEntry -match "<*@*>") {
                $script:MailboxList[$CNEntry] = GetSMTPAddress -PassedCN $CNEntry
            } else {
                Write-Verbose "GetSMTPAddress: Passed in Value does not look like a CN or SMTP Address: [$CNEntry]"
            }
            $script:MailboxList[$CNEntry] = $CNEntry
        }
    }

    foreach ($key in $script:MailboxList.Keys) {
        $value = $script:MailboxList[$key]
        Write-Verbose "$key :: $($value.DisplayName)"
    }
}

<#
.SYNOPSIS
Gets DisplayName from a passed in CN that matches an entry in the MailboxList
#>
function GetDisplayName {
    param(
        $PassedCN
    )
    Write-Verbose "GetDisplayName:: Working on [$PassedCN]"
    if ($PassedCN.Properties.Name -contains 'DisplayName') {
        return GetMailboxProp -PassedCN $PassedCN -Prop "DisplayName"
    } elseif ($PassedCN -match '<' ) {
        return $PassedCN.ToString().split("<")[0].replace('"', '')
    } elseif ($PassedCN -match '\[OneOff') {
        return $PassedCN.ToString().split('"')[1]
    } else {
        Write-Verbose "Unable to get the DisplayName for [$PassedCN]"
        return $PassedCN
    }
}

<#
.SYNOPSIS
Gets SMTP Address from a passed in CN that matches an entry in the MailboxList
#>
function GetSMTPAddress {
    param(
        $PassedCN
    )

    if ($PassedCN -match $WellKnownCN_Trans) {
        return $Transport
    } elseif ($PassedCN -match $WellKnownCN_CA) {
        return $CalAttendant
    } elseif ($PassedCN -match "<*@*>") {
        # This is a new format that we are seeing in the Calendar Logs.
        # Example: '"Jon Doe" <Jon.Doe@Contoso.com>'
        $SMTPAddress = $($PassedCN -split ("<")[-1] -split (">")[0])[1].Trim()
        return $SMTPAddress
    } elseif ($PassedCN -match '<O=') {
        #Matching "Users Name" </O=...>
        $pattern = "<([^>]*)>"
        #$matches = [regex]::Matches($PassedCN, $pattern)
        $MailboxOU = ([regex]::Matches($PassedCN, $pattern)).groups[1].value
        Write-Verbose "Using /OU format to look up mailbox for [$PassedCN]"
        return GetMailboxProp -PassedCN $MailboxOU -Prop "PrimarySmtpAddress"
    } elseif ($PassedCN -match 'cn=([\w,\s.@-]*[^/])$') {
        return GetMailboxProp -PassedCN $PassedCN -Prop "PrimarySmtpAddress"
    } elseif (($PassedCN -match "NotFound") -or ([string]::IsNullOrEmpty($PassedCN))) {
        return ""
    } elseif ($PassedCN -match "InvalidSchemaPropertyName") {
        Write-Verbose "GetSMTPAddress: Passed in Value is empty or not Valid: [$PassedCN]."
        return ""
    } elseif ($PassedCN -match "@") {
        Write-Verbose "GetSMTPAddress: Looks like we have an SMTP Address already: [$PassedCN]."
        $PassedCN.SMTPAddress
        return $PassedCN
    } else {
        # We have a problem, we don't have a CN or an SMTP Address
        Write-Warning "GetSMTPAddress: Passed in Value does not look like a CN or SMTP Address: [$PassedCN]."
        return $PassedCN
    }
}

<#
.SYNOPSIS
    This function gets a more readable Name from a CN or the Calendar Assistant.
.PARAMETER PassedCN
    The common name (CN) of the mailbox user or the Calendar Assistant.
.OUTPUTS
    Returns the last part of the CN so that it is more readable
#>
function BetterThanNothingCNConversion {
    param (
        $PassedCN
    )
    if ($PassedCN -match $WellKnownCN_CA) {
        return $CalAttendant
    }

    if ($PassedCN -match $WellKnownCN_Trans) {
        return $Transport
    }

    if ($PassedCN -match 'cn=([\w,\s.@-]*[^/])$') {
        $cNameMatch = $PassedCN -split "cn="

        # Normally a readable name is sectioned off with a "-" at the end.
        # Example /o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=d61149258ba04404adda42f336b504ed-Delegate
        if ($cNameMatch[-1] -match "-[\w* -.]*") {
            Write-Verbose "BetterThanNothingCNConversion: Matched : [$($cNameMatch[-1])]"
            $cNameSplit = $cNameMatch.split('-')[-1]
            # Sometimes we have a more than one "-" in the name, so we end up with only 1-4 chars which is too little.
            # Example: .../CN=RECIPIENTS/CN=83DAA772E6A94DA19402AA6B41770486-4DB5F0EB-4A
            if ($cNameSplit.length -lt 5) {
                Write-Verbose "BetterThanNothingCNConversion: [$cNameSplit] is too short"
                $cNameSplit= $cNameMatch.split('-')[-2] + '-' + $cNameMatch.split('-')[-1]
                Write-Verbose "BetterThanNothingCNConversion: Returning Lengthened : [$cNameSplit]"
            }
            return $cNameSplit
        }
        # Sometimes we do not have the "-" in front of the Name.
        # Example: "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=user123"
        if ($cNameMatch[-1] -match "[\w* -.]*") {
            Write-Verbose "BetterThanNothingCNConversion: Returning : [$($cNameMatch[-1])]"
            return $cNameMatch.split('-')[-1]
        }
    }
}

<#
.SYNOPSIS
Checks if an entries is Redacted to protect PII.
#>
function CheckForPIIAccess {
    param(
        $PassedString
    )
    if ($PassedString -match "REDACTED-") {
        return $false
    } else {
        return $true
    }
}

<#
.SYNOPSIS
    Retrieves mailbox properties for a given mailbox.
.DESCRIPTION
    This function retrieves mailbox properties for a given mailbox using Exchange Web Services (EWS).
.PARAMETER CN of the Mailbox
    The mailbox for which to retrieve properties.
.PARAMETER PropertySet
    The set of properties to retrieve.
#>
function GetMailboxProp {
    param(
        $PassedCN,
        $Prop
    )

    Write-Debug "GetMailboxProp: [$Prop]: Searching for:[$PassedCN]..."

    if (($Prop -ne "PrimarySmtpAddress") -and ($Prop -ne "DisplayName")) {
        Write-Error "GetMailboxProp:Invalid Property: [$Prop]"
        return "Invalid Property"
    }

    if ($script:MailboxList.count -gt 0) {
        switch -Regex ($PassedCN) {
            $WellKnownCN_CA {
                return $CalAttendant
            }
            $WellKnownCN_Trans {
                return $Transport
            }
            default {
                if ($null -ne $script:MailboxList[$PassedCN]) {
                    $ReturnValue = $script:MailboxList[$PassedCN].$Prop

                    if ($null -eq $ReturnValue) {
                        Write-Error "`t GetMailboxProp:$Prop :NotFound for ::[$PassedCN]"
                        return BetterThanNothingCNConversion($PassedCN)
                    }

                    Write-Verbose "`t GetMailboxProp:[$Prop] :Found::[$ReturnValue]"
                    if (-not (CheckForPIIAccess($ReturnValue))) {
                        Write-Verbose "No PII Access for [$ReturnValue]"
                        return BetterThanNothingCNConversion($PassedCN)
                    }
                    return $ReturnValue
                } else {
                    Write-Verbose "`t GetMailboxProp:$Prop :NotFound::$PassedCN"
                    return BetterThanNothingCNConversion($PassedCN)
                }
            }
        }
    } else {
        Write-Host -ForegroundColor Red "$script:MailboxList is empty, unable to do CN to SMTP mapping."
        return BetterThanNothingCNConversion($PassedCN)
    }
}

# ===================================================================================================
# Constants to support the script
# ===================================================================================================

$script:CustomPropertyNameList =
"AppointmentCounterProposal",
"AppointmentLastSequenceNumber",
"AppointmentRecurring",
"CalendarItemType",
"CalendarLogTriggerAction",
"CalendarProcessed",
"ChangeList",
"ClientBuildVersion",
"ClientIntent",
"ClientProcessName",
"CreationTime",
"DisplayAttendeesCc",
"DisplayAttendeesTo",
"EventEmailReminderTimer",
"ExternalSharingMasterId",
"FreeBusyStatus",
"From",
"HasAttachment",
"InternetMessageId",
"IsAllDayEvent",
"IsCancelled",
"IsException",
"IsMeeting",
"IsOrganizerProperty",
"IsSharedInEvent",
"ItemID",
"LogBodyStats",
"LogClientInfoString",
"LogRowType",
"LogTimestamp",
"NormalizedSubject",
"OriginalStartDate",
"ReminderDueByInternal",
"ReminderIsSetInternal",
"ReminderMinutesBeforeStartInternal",
"SendMeetingMessagesDiagnostics",
"Sensitivity",
"SentRepresentingDisplayName",
"SentRepresentingEmailAddress",
"ShortClientInfoString",
"TimeZone"

$LogLimit = 2000

if ($ShortLogs.IsPresent) {
    $LogLimit = 500
}

if ($MaxLogs.IsPresent) {
    $LogLimit = 12000
}

$LimitedItemClasses = @(
    "IPM.Appointment",
    "IPM.Schedule.Meeting.Request",
    "IPM.Schedule.Meeting.Canceled",
    "IPM.Schedule.Meeting.Forwarded"
)

<#
.SYNOPSIS
Run Get-CalendarDiagnosticObjects for passed in User with Subject or MeetingID.
#>
function GetCalendarDiagnosticObjects {
    param(
        [string]$Identity,
        [string]$Subject,
        [string]$MeetingID
    )

    $params = @{
        Identity           = $Identity
        CustomPropertyName = $script:CustomPropertyNameList
        WarningAction      = "Ignore"
        MaxResults         = $LogLimit
        ResultSize         = $LogLimit
        ShouldBindToItem   = $true
        ShouldDecodeEnums  = $true
    }

    if ($TrackingLogs.IsPresent) {
        Write-Host -ForegroundColor Yellow "Including Tracking Logs in the output."
        $script:CustomPropertyNameList += "AttendeeListDetails", "AttendeeCollection"
        $params.Add("ShouldFetchAttendeeCollection", $true)
        $params.Remove("CustomPropertyName")
        $params.Add("CustomPropertyName", $script:CustomPropertyNameList)
    }

    if (-not [string]::IsNullOrEmpty($ExceptionDate)) {
        Write-Host -ForegroundColor Green "---------------------------------------"
        Write-Host -ForegroundColor Green "Pulling all the Exceptions for [$ExceptionDate] and adding them to the output."
        Write-Host -ForegroundColor Green "---------------------------------------"
        $params.Add("AnalyzeExceptionWithOriginalStartDate", $ExceptionDate)
    }

    if ($MaxLogs.IsPresent) {
        Write-Host -ForegroundColor Yellow "Limiting the number of logs to $LogLimit, and limiting the number of Item Classes retrieved."
        $params.Add("ItemClass", $LimitedItemClasses)
    }

    if ($null -ne $CustomProperty) {
        Write-Host -ForegroundColor Yellow "Adding custom properties to the RAW output."
        $params.Remove("CustomPropertyName")
        $script:CustomPropertyNameList += $CustomProperty
        Write-Host -ForegroundColor Yellow "Adding extra CustomProperty: [$CustomProperty]"
        $params.Add("CustomPropertyName", $script:CustomPropertyNameList)
    }

    if ($Identity -and $MeetingID) {
        Write-Verbose "Getting CalLogs for [$Identity] with MeetingID [$MeetingID]."
        if ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent) {
            Write-Host -ForegroundColor Yellow ($params.GetEnumerator() | ForEach-Object { "`t$($_.Key) = $($_.Value)`n" })
        }
        $CalLogs = Get-CalendarDiagnosticObjects @params -MeetingID $MeetingID
    } elseif ($Identity -and $Subject ) {
        Write-Verbose "Getting CalLogs for [$Identity] with Subject [$Subject]."
        $CalLogs = Get-CalendarDiagnosticObjects @params -Subject $Subject

        # No Results, do a Deep search with ExactMatch.
        if ($CalLogs.count -lt 1) {
            $CalLogs = Get-CalendarDiagnosticObjects @Params -Subject $Subject -ExactMatch $true
        }
    }

    Write-Host "Found $($CalLogs.count) Calendar Logs for [$Identity]"
    return $CalLogs
}

<#
.SYNOPSIS
This function retrieves calendar logs from the specified source with a subject that matches the provided criteria.
.PARAMETER Identity
The Identity of the mailbox to get calendar logs from.
.PARAMETER Subject
The subject of the calendar logs to retrieve.
#>
function GetCalLogsWithSubject {
    param (
        [string] $Identity,
        [string] $Subject
    )
    Write-Host "Getting CalLogs from [$Identity] with subject [$Subject]]"

    $InitialCDOs = GetCalendarDiagnosticObjects -Identity $Identity -Subject $Subject
    $GlobalObjectIds = @()

    # Find all the unique Global Object IDs
    foreach ($ObjectId in $InitialCDOs.CleanGlobalObjectId) {
        if (![string]::IsNullOrEmpty($ObjectId) -and
            $ObjectId -ne "NotFound" -and
            $ObjectId -ne "InvalidSchemaPropertyName" -and
            $ObjectId.Length -ge 90) {
            $GlobalObjectIds += $ObjectId
        }
    }

    $GlobalObjectIds = $GlobalObjectIds | Select-Object -Unique
    Write-Host "Found $($GlobalObjectIds.count) unique GlobalObjectIds."
    Write-Host "Getting the set of CalLogs for each GlobalObjectID."

    if ($GlobalObjectIds.count -eq 1) {
        $script:GCDO = $InitialCDOs; # use the CalLogs that we already have, since there is only one.
        BuildCSV
        BuildTimeline
    } elseif ($GlobalObjectIds.count -gt 1) {
        Write-Host "Found multiple GlobalObjectIds: $($GlobalObjectIds.Count)."
        foreach ($MID in $GlobalObjectIds) {
            Write-DashLineBoxColor "Processing MeetingID: [$MID]"
            $script:GCDO = GetCalendarDiagnosticObjects -Identity $Identity -MeetingID $MID
            Write-Verbose "Found $($GCDO.count) CalLogs with MeetingID[$MID] ."
            BuildCSV
            BuildTimeline
        }
    } else {
        Write-Warning "No CalLogs were found."
    }
}

<#
.SYNOPSIS
Checks if a set of Calendar Logs is from the Organizer.
#>
function SetIsOrganizer {
    param(
        $CalLogs
    )
    [bool] $IsOrganizer = $false

    foreach ($CalLog in $CalLogs) {
        if ($CalLog.ItemClass -eq "Ipm.Appointment" -and
            $CalLog.ExternalSharingMasterId -eq "NotFound" -and
            ($CalLog.ResponseType -eq "1" -or $CalLog.ResponseType -eq "Organizer")) {
            $IsOrganizer = $true
            Write-Host -ForegroundColor Green "IsOrganizer: [$IsOrganizer]"
            return $IsOrganizer
        }
    }
    Write-Verbose "IsOrganizer: [$IsOrganizer]"
    return $IsOrganizer
}

<#
.SYNOPSIS
Checks if a set of Calendar Logs is from a Resource Mailbox.
#>
function SetIsRoom {
    param(
        $CalLogs
    )

    # See if we have already determined this is a Room MB.
    if ($script:Rooms -contains $Identity) {
        return $true
    }

    # Simple logic is if RBA is running on the MB, it is a Room MB, otherwise it is not.
    foreach ($CalLog in $CalLogs) {
        Write-Verbose "Checking if this is a Room Mailbox. [$($CalLog.ItemClass)] [$($CalLog.ExternalSharingMasterId)] [$($CalLog.LogClientInfoString)]"
        if ($CalLog.ItemClass -eq "IPM.Appointment" -and
            $CalLog.ExternalSharingMasterId -eq "NotFound" -and
            $CalLog.LogClientInfoString -like "*ResourceBookingAssistant*" ) {
            return $true
        }
    }
    return $false
}

<#
.SYNOPSIS
Checks if a set of Calendar Logs is from a Recurring Meeting.
#>
function SetIsRecurring {
    param(
        $CalLogs
    )
    Write-Host -ForegroundColor Yellow "Looking for signs of a recurring meeting."
    [bool] $IsRecurring = $false
    # See if this is a recurring meeting
    foreach ($CalLog in $CalLogs) {
        if ($CalendarItemTypes.($CalLog.ItemClass) -eq "Ipm.Appointment" -and
            # Commenting this out will get all the updates for shared calendars, which is important with Delegates.
            #      $CalLog.ExternalSharingMasterId -eq "NotFound" -and
            $CalLog.CalendarItemType.ToString() -eq "RecurringMaster") {
            $IsRecurring = $true
            Write-Verbose "Found recurring meeting."
            return $IsRecurring
        }
    }
    Write-Verbose "Did not find signs of recurring meeting."
    return $IsRecurring
}

<#
.SYNOPSIS
Check for Bifurcation issue
#>
function CheckForBifurcation {
    param (
        $CalLogs
    )
    Write-Verbose  "Looking for signs of the Bifurcation Issue."
    [bool] $IsBifurcated = $false
    # See if there is an IPM.Appointment in the CalLogs.
    foreach ($CalLog in $CalLogs) {
        if ($CalLog.ItemClass -eq "IPM.Appointment" -and
            $CalLog.ExternalSharingMasterId -eq "NotFound") {
            $IsBifurcated = $false
            Write-Verbose "Found Ipm.Appointment, likely not a bifurcation issue."
            return $IsBifurcated
        }
    }
    Write-Host -ForegroundColor Red "Did not find any Ipm.Appointments in the CalLogs. If this is the Organizer of the meeting, this could the the Outlook Bifurcation issue."
    Write-Host -ForegroundColor Yellow "`t This could be the Outlook Bifurcation issue, where Outlook saves to the Organizers Mailbox on one thread and send to the attendee via transport on another thread.  If the save to Organizers mailbox failed, we get into the Bifurcated State, where the Organizer does not have the meeting but the Attendees do."
    Write-Host -ForegroundColor Yellow "`t See https://support.microsoft.com/en-us/office/meeting-request-is-missing-from-organizers-calendar-c13c47cd-18f9-4ef0-b9d0-d9e174912c4a"
    return $IsBifurcated
}

# ===================================================================================================
# FileNames
# ===================================================================================================
function Get-FileName {
    Write-Host -ForegroundColor Cyan "Creating FileName for $Identity..."

    $ThisMeetingID = $script:GCDO.CleanGlobalObjectId | Select-Object -Unique
    $ShortMeetingID = $ThisMeetingID.Substring($ThisMeetingID.length - 6)

    if ($script:Identity -like "*@*") {
        $script:ShortId = $script:Identity.split('@')[0]
    } else {
        $script:ShortId = $script:Identity
    }
    $script:ShortId = $ShortId.Substring(0, [System.Math]::Min(20, $ShortId.Length))

    if (($null -eq $CaseNumber) -or
        ([string]::IsNullOrEmpty($CaseNumber))) {
        $Case = ""
    } else {
        $Case = $CaseNumber + "_"
    }

    if ($ExportToExcel.IsPresent) {
        $script:FileName = "$($Case)CalLogSummary_$($ShortMeetingID).xlsx"
        Write-Host -ForegroundColor Blue -NoNewline "All Calendar Logs for meetings ending in ID [$ShortMeetingID] will be saved to : "
        Write-Host -ForegroundColor Yellow "$Filename"
    } else {
        $script:Filename = "$($Case)$($ShortId)_$ShortMeetingID.csv"
        $script:FilenameRaw = "$($Case)$($ShortId)_RAW_$($ShortMeetingID).csv"
        $Script:TimeLineFilename = "$($Case)$($ShortId)_TimeLine_$ShortMeetingID.csv"

        Write-Host -ForegroundColor Cyan -NoNewline "Enhanced Calendar Logs for [$Identity] has been saved to : "
        Write-Host -ForegroundColor Yellow "$Filename"

        Write-Host -ForegroundColor Cyan -NoNewline "Raw Calendar Logs for [$Identity] has been saved to : "
        Write-Host -ForegroundColor Yellow "$FilenameRaw"

        Write-Host -ForegroundColor Cyan -NoNewline "TimeLine for [$Identity] has been saved to : "
        Write-Host -ForegroundColor Yellow "$TimeLineFilename"
    }
}

function Export-CalLog {
    Get-FileName

    if ($ExportToExcel.IsPresent) {
        Export-CalLogExcel
    } else {
        Export-CalLogCSV
    }
}

function Export-CalLogCSV {
    $GCDOResults | Export-Csv -Path $Filename -NoTypeInformation -Encoding UTF8
    $script:GCDO | Export-Csv -Path $FilenameRaw -NoTypeInformation -Encoding UTF8
}

function Export-Timeline {
    Write-Verbose "Export to Excel is : $ExportToExcel"

    # Display Timeline to screen:
    Write-Host -ForegroundColor Cyan "Timeline for [$Identity]..."
    $script:TimeLineOutput

    if ($ExportToExcel.IsPresent) {
        Export-TimelineExcel
    } else {
        $script:TimeLineOutput | Export-Csv -Path $script:TimeLineFilename -NoTypeInformation -Encoding UTF8 -Append
    }
}

<#
.SYNOPSIS
    This is the part that generates the heart of the timeline, a Giant Switch statement based on the ItemClass.
#>
function CreateTimelineRow {
    switch -Wildcard ($CalLog.ItemClass) {
        Meeting.Request {
            switch ($CalLog.TriggerAction) {
                Create {
                    if ($IsOrganizer) {
                        if ($CalLog.IsException -eq $True) {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Created an Exception Meeting Request with $($CalLog.Client) for [$($CalLog.StartTime)]."
                        } else {
                            [array] $Output  = "[$($CalLog.ResponsibleUser)] Created a Meeting Request with $($CalLog.Client)"
                        }
                    } else {
                        if ($CalLog.DisplayAttendeesTo -ne $script:PreviousCalLog.DisplayAttendeesTo -or $CalLog.DisplayAttendeesCc -ne $script:PreviousCalLog.DisplayAttendeesCc) {
                            [array] $Output = "The user Forwarded a Meeting Request with $($CalLog.Client)."
                        } else {
                            if ($CalLog.Client -eq "Transport") {
                                if ($CalLog.IsException -eq $True) {
                                    [array] $Output = "Transport delivered a new Meeting Request from [$($CalLog.From)] for an exception starting on [$($CalLog.StartTime)]" + $(if ($null -ne $($CalLog.ReceivedRepresenting)) { " for user [$($CalLog.ReceivedRepresenting)]" }) + "."
                                    $script:MeetingSummaryNeeded = $True
                                } else {
                                    [Array] $Output = "Transport delivered a new Meeting Request from [$($CalLog.From)]" +
                                    $(if ($null -ne $($CalLog.ReceivedRepresenting) -and $CalLog.ReceivedRepresenting -ne $CalLog.ReceivedBy)
                                        { " for user [$($CalLog.ReceivedRepresenting)]" }) + "."
                                }
                            } elseif ($calLog.client -eq "ResourceBookingAssistant") {
                                [array] $Output  = "ResourceBookingAssistant Forwarded a Meeting Request to a Resource Delegate."
                            } elseif ($CalLog.Client -eq "CalendarRepairAssistant") {
                                if ($CalLog.IsException -eq $True) {
                                    [array] $Output = "CalendarRepairAssistant Created a new Meeting Request to repair an inconsistency with an exception starting on [$($CalLog.StartTime)]."
                                } else {
                                    [array] $Output = "CalendarRepairAssistant Created a new Meeting Request to repair an inconsistency."
                                }
                            } else {
                                if ($CalLog.IsException -eq $True) {
                                    [array] $Output = "[$($CalLog.ResponsibleUser)] Created a new Meeting Request with $($CalLog.Client) for an exception starting on [$($CalLog.StartTime)]."
                                } else {
                                    [array] $Output = "[$($CalLog.ResponsibleUser)] Created a new Meeting Request with $($CalLog.Client)."
                                }
                            }
                        }
                    }
                }
                Update {
                    if ($calLog.client -eq "ResourceBookingAssistant") {
                        [array] $Output  = "ResourceBookingAssistant Updated the Meeting Request."
                    } else {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Updated the $($CalLog.MeetingRequestType.Value) Meeting Request with $($CalLog.Client)."
                    }
                }
                MoveToDeletedItems {
                    if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                        [array] $Output = "$($CalLog.Client) Deleted the Meeting Request."
                    } else {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Deleted the Meeting Request with $($CalLog.Client)."
                    }
                }
                default {
                    [array] $Output = "[$($CalLog.ResponsibleUser)] Deleted the $($CalLog.MeetingRequestType.Value) Meeting Request with $($CalLog.Client)."
                }
            }
        }
        Resp* {
            switch ($CalLog.ItemClass) {
                "Resp.Tent" { $MeetingRespType = "Tentative" }
                "Resp.Neg" { $MeetingRespType = "DECLINE" }
                "Resp.Pos" { $MeetingRespType = "ACCEPT" }
            }

            if ($CalLog.AppointmentCounterProposal -eq "True") {
                [array] $Output = "[$($CalLog.Organizer)] send a $($MeetingRespType) response message with a New Time Proposal: $($CalLog.StartTime) to $($CalLog.EndTime)"
            } else {
                switch -Wildcard ($CalLog.TriggerAction) {
                    "Update" { $Action = "Updated" }
                    "Create" { $Action = "Sent" }
                    "*Delete*" { $Action = "Deleted" }
                    default {
                        $Action = "Updated"
                    }
                }

                $Extra = ""
                if ($CalLog.CalendarItemType -eq "Exception") {
                    $Extra = " to the meeting starting $($CalLog.StartTime)"
                } elseif ($CalLog.AppointmentRecurring) {
                    $Extra = " to the meeting series"
                }

                if ($IsOrganizer) {
                    [array] $Output = "[$($CalLog.Organizer)] $($Action) a $($MeetingRespType) meeting Response message$($Extra)."
                } else {
                    switch ($CalLog.Client) {
                        ResourceBookingAssistant {
                            [array] $Output = "ResourceBookingAssistant $($Action) a $($MeetingRespType) Meeting Response message$($Extra)."
                        }
                        Transport {
                            [array] $Output = "[$($CalLog.From)] $($Action) $($MeetingRespType) Meeting Response message$($Extra)."
                        }
                        default {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] $($Action) [$($CalLog.Organizer)]'s $($MeetingRespType) Meeting Response with $($CalLog.Client)$($Extra)."
                        }
                    }
                }
            }
        }
        Forward.Notification {
            [array] $Output = "The meeting was FORWARDED by [$($CalLog.Organizer)]."
        }
        Exception {
            if ($CalLog.ResponsibleUser -ne "Calendar Assistant") {
                [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d Exception starting $($CalLog.StartTime) to the meeting series with $($CalLog.Client)."
            }
        }
        Ipm.Appointment {
            switch ($CalLog.TriggerAction) {
                Create {
                    if ($IsOrganizer) {
                        if ($CalLog.Client -eq "Transport") {
                            [array] $Output = "Transport Created a new meeting."
                        } else {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Created a new Meeting with $($CalLog.Client)."
                        }
                    } else {
                        switch ($CalLog.Client) {
                            Transport {
                                [array] $Output = "Transport Created a new Meeting on the calendar from [$($CalLog.Organizer)] and marked it Tentative."
                            }
                            ResourceBookingAssistant {
                                [array] $Output = "ResourceBookingAssistant Created a new Meeting on the calendar from [$($CalLog.Organizer)] and marked it Tentative."
                            }
                            default {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] Created the Meeting with $($CalLog.Client)."
                            }
                        }
                    }
                }
                Update {
                    switch ($CalLog.Client) {
                        Transport {
                            if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                                [array] $Output = "Transport Updated the meeting based on changes made to the meeting on [$($CalLog.Sender)] calendar."
                            } else {
                                [array] $Output = "Transport $($CalLog.TriggerAction)d the meeting based on changes made by [$($CalLog.ResponsibleUser)]."
                            }
                        }
                        LocationProcessor {
                            [array] $Output = ""
                        }
                        ResourceBookingAssistant {
                            [array] $Output = "ResourceBookingAssistant $($CalLog.TriggerAction)d the Meeting."
                        }
                        CalendarRepairAssistant {
                            [array] $Output = "CalendarRepairAssistant $($CalLog.TriggerAction)d the Meeting to repair an inconsistency."
                        }
                        default {
                            if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                                [array] $Output = "The Exchange System $($CalLog.TriggerAction)d the meeting via the Calendar Assistant."
                            } else {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d the Meeting with $($CalLog.Client)."
                                $script:AddChangedProperties = $True
                            }
                        }
                    }

                    if ($CalLog.FreeBusyStatus -eq 2 -and $script:PreviousCalLog.FreeBusyStatus -ne 2) {
                        if ($CalLog.ResponsibleUserName -eq "Calendar Assistant") {
                            [array] $Output = "$($CalLog.Client) Accepted the meeting."
                        } else {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Accepted the meeting with $($CalLog.Client)."
                        }
                        $script:AddChangedProperties = $False
                    } elseif ($CalLog.FreeBusyStatus -ne 2 -and $script:PreviousCalLog.FreeBusyStatus -eq 2) {
                        if ($IsOrganizer) {
                            [array] $Output = "[$($CalLog.ResponsibleUser)] Cancelled the Meeting with $($CalLog.Client)."
                        } else {
                            if ($CalLog.ResponsibleUser -ne "Calendar Assistant") {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] Declined the meeting with $($CalLog.Client)."
                            }
                        }
                        $script:AddChangedProperties = $False
                    }
                }
                SoftDelete {
                    switch ($CalLog.Client) {
                        Transport {
                            [array] $Output = "Transport $($CalLog.TriggerAction)d the meeting based on changes by [$($CalLog.Organizer)]."
                        }
                        LocationProcessor {
                            [array] $Output = ""
                        }
                        ResourceBookingAssistant {
                            [array] $Output = "ResourceBookingAssistant $($CalLog.TriggerAction)d the Meeting."
                        }
                        default {
                            if ($CalLog.ResponsibleUser -eq "Calendar Assistant") {
                                [array] $Output = "The Exchange System $($CalLog.TriggerAction)d the meeting via the Calendar Assistant."
                            } else {
                                [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d the meeting with $($CalLog.Client)."
                                $script:AddChangedProperties = $True
                            }
                        }
                    }

                    if ($CalLog.FreeBusyStatus -eq 2 -and $script:PreviousCalLog.FreeBusyStatus -ne 2) {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Accepted the Meeting with $($CalLog.Client)."
                        $script:AddChangedProperties = $False
                    } elseif ($CalLog.FreeBusyStatus -ne 2 -and $script:PreviousCalLog.FreeBusyStatus -eq 2) {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] Declined the Meeting with $($CalLog.Client)."
                        $script:AddChangedProperties = $False
                    }
                }
                MoveToDeletedItems {
                    [array] $Output = "[$($CalLog.ResponsibleUser)] Deleted the Meeting with $($CalLog.Client) (Moved the Meeting to the Deleted Items)."
                }
                default {
                    [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction) the Meeting with $($CalLog.Client)."
                    $script:MeetingSummaryNeeded = $False
                }
            }
        }
        Cancellation {
            switch ($CalLog.Client) {
                Transport {
                    if ($CalLog.IsException -eq $True) {
                        [array] $Output = "Transport $($CalLog.TriggerAction)d a Meeting Cancellation based on changes by [$($CalLog.SenderSMTPAddress)] for the exception starting on [$($CalLog.StartTime)]"
                    } else {
                        [array] $Output = "Transport $($CalLog.TriggerAction)d a Meeting Cancellation based on changes by [$($CalLog.SenderSMTPAddress)]."
                    }
                }
                ResourceBookingAssistant {
                    if ($CalLog.TriggerAction -eq "MoveToDeletedItems") {
                        [array] $Output = "ResourceBookingAssistant Deleted the Cancellation."
                    } else {
                        [array] $Output = "ResourceBookingAssistant $($CalLog.TriggerAction)d the Cancellation."
                    }
                }
                default {
                    if ($CalLog.IsException -eq $True) {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d a Cancellation with $($CalLog.Client) for the exception starting on [$($CalLog.StartTime)]."
                    } elseif ($CalLog.CalendarItemType -eq "RecurringMaster") {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d a Cancellation for the Series with $($CalLog.Client)."
                    } else {
                        [array] $Output = "[$($CalLog.ResponsibleUser)] $($CalLog.TriggerAction)d the Cancellation with $($CalLog.Client)."
                    }
                }
            }
        }
        default {
            if ($CalLog.TriggerAction -eq "Create") {
                $Action = "New "
            } else {
                $Action = "$($CalLog.TriggerAction)"
            }
            [array] $Output = "[$($CalLog.ResponsibleUser)] performed a $($Action) on the $($CalLog.ItemClass) with $($CalLog.Client)."
        }
    }

    return $Output
}

<#
.SYNOPSIS
    Determines if key properties of the calendar log have changed.
.DESCRIPTION
    This function checks if the properties of the calendar log have changed by comparing the current
    Calendar log to the Previous calendar log (where it was an IPM.Appointment - i.e. the meeting)

    Changed properties will be added to the Timeline.
#>
function FindChangedProperties {
    if ($CalLog.Client -ne "LocationProcessor" -or $CalLog.Client -notlike "*EBA*" -or $CalLog.Client -notlike "*TBA*") {
        if ($script:PreviousCalLog -and $script:AddChangedProperties) {
            if ($CalLog.StartTime.ToString() -ne $script:PreviousCalLog.StartTime.ToString()) {
                [Array]$TimeLineText = "The StartTime changed from [$($script:PreviousCalLog.StartTime)] to: [$($CalLog.StartTime)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.EndTime.ToString() -ne $script:PreviousCalLog.EndTime.ToString()) {
                [Array]$TimeLineText = "The EndTime changed from [$($script:PreviousCalLog.EndTime)] to: [$($CalLog.EndTime)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.SubjectProperty -ne $script:PreviousCalLog.SubjectProperty) {
                [Array]$TimeLineText = "The SubjectProperty changed from [$($script:PreviousCalLog.SubjectProperty)] to: [$($CalLog.SubjectProperty)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.NormalizedSubject -ne $script:PreviousCalLog.NormalizedSubject) {
                [Array]$TimeLineText = "The NormalizedSubject changed from [$($script:PreviousCalLog.NormalizedSubject)] to: [$($CalLog.NormalizedSubject)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.Location -ne $script:PreviousCalLog.Location) {
                [Array]$TimeLineText = "The Location changed from [$($script:PreviousCalLog.Location)] to: [$($CalLog.Location)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.TimeZone -ne $script:PreviousCalLog.TimeZone) {
                [Array]$TimeLineText = "The TimeZone changed from [$($script:PreviousCalLog.TimeZone)] to: [$($CalLog.TimeZone)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.DisplayAttendeesAll -ne $script:PreviousCalLog.DisplayAttendeesAll) {
                [Array]$TimeLineText = "The All Attendees changed from [$($script:PreviousCalLog.DisplayAttendeesAll)] to: [$($CalLog.DisplayAttendeesAll)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.AppointmentRecurring -ne $script:PreviousCalLog.AppointmentRecurring) {
                [Array]$TimeLineText = "The Appointment Recurrence changed from [$($script:PreviousCalLog.AppointmentRecurring)] to: [$($CalLog.AppointmentRecurring)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.HasAttachment -ne $script:PreviousCalLog.HasAttachment) {
                [Array]$TimeLineText = "The Meeting has Attachment changed from [$($script:PreviousCalLog.HasAttachment)] to: [$($CalLog.HasAttachment)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.IsCancelled -ne $script:PreviousCalLog.IsCancelled) {
                [Array]$TimeLineText = "The Meeting is Cancelled changed from [$($script:PreviousCalLog.IsCancelled)] to: [$($CalLog.IsCancelled)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.IsAllDayEvent -ne $script:PreviousCalLog.IsAllDayEvent) {
                [Array]$TimeLineText = "The Meeting is an All Day Event changed from [$($script:PreviousCalLog.IsAllDayEvent)] to: [$($CalLog.IsAllDayEvent)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.IsException -ne $script:PreviousCalLog.IsException) {
                [Array]$TimeLineText = "The Meeting Is Exception changed from [$($script:PreviousCalLog.IsException)] to: [$($CalLog.IsException)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.IsSeriesCancelled -ne $script:PreviousCalLog.IsSeriesCancelled) {
                [Array]$TimeLineText = "The Is Series Cancelled changed from [$($script:PreviousCalLog.IsSeriesCancelled)] to: [$($CalLog.IsSeriesCancelled)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.EventEmailReminderTimer -ne $script:PreviousCalLog.EventEmailReminderTimer) {
                [Array]$TimeLineText = "The Email Reminder changed from [$($script:PreviousCalLog.EventEmailReminderTimer)] to: [$($CalLog.EventEmailReminderTimer)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.FreeBusyStatus -ne $script:PreviousCalLog.FreeBusyStatus) {
                [Array]$TimeLineText = "The FreeBusy Status changed from [$($script:PreviousCalLog.FreeBusyStatus)] to: [$($CalLog.FreeBusyStatus)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.AppointmentState -ne $script:PreviousCalLog.AppointmentState) {
                [Array]$TimeLineText = "The Appointment State changed from [$($script:PreviousCalLog.AppointmentState)] to: [$($CalLog.AppointmentState)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.MeetingRequestType -ne $script:PreviousCalLog.MeetingRequestType) {
                [Array]$TimeLineText = "The Meeting Request Type changed from [$($script:PreviousCalLog.MeetingRequestType.Value)] to: [$($CalLog.MeetingRequestType.Value)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.CalendarItemType -ne $script:PreviousCalLog.CalendarItemType) {
                [Array]$TimeLineText = "The Calendar Item Type changed from [$($script:PreviousCalLog.CalendarItemType)] to: [$($CalLog.CalendarItemType)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.ResponseType -ne $script:PreviousCalLog.ResponseType) {
                [Array]$TimeLineText = "The ResponseType changed from [$($script:PreviousCalLog.ResponseType)] to: [$($CalLog.ResponseType)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.SenderSMTPAddress -ne $script:PreviousCalLog.SenderSMTPAddress) {
                [Array]$TimeLineText = "The Sender Email Address changed from [$($script:PreviousCalLog.SenderSMTPAddress)] to: [$($CalLog.SenderSMTPAddress)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.From -ne $script:PreviousCalLog.From) {
                [Array]$TimeLineText = "The From changed from [$($script:PreviousCalLog.From)] to: [$($CalLog.From)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }

            if ($CalLog.ReceivedRepresenting -ne $script:PreviousCalLog.ReceivedRepresenting) {
                [Array]$TimeLineText = "The Received Representing changed from [$($script:PreviousCalLog.ReceivedRepresenting)] to: [$($CalLog.ReceivedRepresenting)]"
                CreateMeetingSummary -Time " " -MeetingChanges $TimeLineText
            }
        }
    }
}

<#
.SYNOPSIS
    Function to write a line of text surrounded by a dash line box.

.DESCRIPTION
    The Write-DashLineBoxColor function is used to create a quick and easy display around a line of text. It generates a box made of dash characters ("-") and displays the provided line of text inside the box.

.PARAMETER Line
    Specifies the line of text to be displayed inside the dash line box.

.PARAMETER Color
    Specifies the color of the dash line box and the text. The default value is "White".

.PARAMETER DashChar
    Specifies the character used to create the dash line. The default value is "-".

.EXAMPLE
    Write-DashLineBoxColor -Line "Hello, World!" -Color "Yellow" -DashChar "="
    Displays:
    ==============
    Hello, World!
    ==============
#>
function Write-DashLineBoxColor {
    [CmdletBinding()]
    param(
        [string[]]$Line,
        [string] $Color = "White",
        [char] $DashChar = "-"
    )
    $highLineLength = 0
    $Line | ForEach-Object { if ($_.Length -gt $highLineLength) { $highLineLength = $_.Length } }
    $dashLine = [string]::Empty
    1..$highLineLength | ForEach-Object { $dashLine += $DashChar }
    Write-Host
    Write-Host -ForegroundColor $Color $dashLine
    $Line | ForEach-Object { Write-Host -ForegroundColor $Color $_ }
    Write-Host -ForegroundColor $Color $dashLine
    Write-Host
}

# Default to Excel unless specified otherwise.
if (!$ExportToCSV.IsPresent) {
    Write-Host -ForegroundColor Yellow "Exporting to Excel."
    $script:ExportToExcel = $true

function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

    return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
}
    $script:IsAdministrator = Confirm-Administrator

# ===================================================================================================
# ImportExcel Functions
# see https://github.com/dfinke/ImportExcel for information on the module.
# ===================================================================================================
function CheckExcelModuleInstalled {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param ()

    if (Get-Command -Module ImportExcel) {
        Write-Host "ImportExcel module is already installed."
    } else {
        # This is slow, to the tune of ~10 seconds, but much more complete.
        # Check if ImportExcel module is installed
        $moduleInstalled = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'ImportExcel' }

        if ($moduleInstalled) {
            Write-Host "ImportExcel module is already installed."
        } else {
            # Check if running with administrator rights
            if (-not $script:IsAdministrator) {
                Write-Host "Please run the script as an administrator to install the ImportExcel module."
                exit
            }

            # Ask user if they want to install the module
            if ($PSCmdlet.ShouldProcess('ImportExcel module', 'Import')) {
                Write-Verbose "Installing ImportExcel module..."

                # Install ImportExcel module
                Install-Module -Name ImportExcel -Force -AllowClobber

                Write-Host -ForegroundColor Green "Done. ImportExcel module is now installed."
                Write-Host -ForegroundColor Yellow "Please rerun the Script to get your Calendar Logs."
                exit
            }
        }
    }
}

# Export to Excel
function Export-CalLogExcel {
    Write-Host -ForegroundColor Cyan "Exporting Enhanced CalLogs to Excel Tab [$ShortId]..."
    $ExcelParamsArray = GetExcelParams -path $FileName -tabName $ShortId

    $excel = $GCDOResults | Export-Excel @ExcelParamsArray -PassThru

    FormatHeader ($excel)

    Export-Excel -ExcelPackage $excel -WorksheetName $ShortId -MoveToStart

    # Export Raw Logs for Developer Analysis
    Write-Host -ForegroundColor Cyan "Exporting Raw CalLogs to Excel Tab [$($ShortId + "_Raw")]..."
    $script:GCDO | Export-Excel -Path  $FileName -WorksheetName $($ShortId + "_Raw") -AutoFilter -FreezeTopRow -BoldTopRow -MoveToEnd
    LogScriptInfo
}

function LogScriptInfo {
    # Only need to run once per script execution.
    if ($null -eq $script:CollectedCmdLine) {
        $RunInfo = @()
        $RunInfo += [PSCustomObject]@{
            Key   = "Script Name"
            Value = $($script:command.MyCommand.Name)
        }
        $RunInfo += [PSCustomObject]@{
            Key   ="RunTime"
            Value = Get-Date
        }
        $RunInfo += [PSCustomObject]@{
            Key   = "Command Line"
            Value = $($script:command.Line)
        }
        $RunInfo += [PSCustomObject]@{
            Key   = "Script Version"
            Value =  $script:BuildVersion
        }
        $RunInfo += [PSCustomObject]@{
            Key   = "User"
            Value =  whoami.exe
        }
        $RunInfo += [PSCustomObject]@{
            Key   = "PowerShell Version"
            Value = $PSVersionTable.PSVersion
        }
        $RunInfo += [PSCustomObject]@{
            Key   = "OS Version"
            Value = $(Get-CimInstance -ClassName Win32_OperatingSystem).Version
        }
        $RunInfo += [PSCustomObject]@{
            Key   = "More Info"
            Value = "https://learn.microsoft.com/en-us/exchange/troubleshoot/calendars/analyze-calendar-diagnostic-logs"
        }

        $RunInfo | Export-Excel -Path $FileName -WorksheetName "Script Info" -MoveToEnd
        $script:CollectedCmdLine = $true
    }
    # If someone runs the script the script again logs will update, but ScriptInfo does not update. Need to add new table for each run.
}

function Export-TimelineExcel {
    Write-Host -ForegroundColor Cyan "Exporting Timeline to Excel..."
    $script:TimeLineOutput | Export-Excel -Path $FileName -WorksheetName $($ShortId + "_TimeLine") -Title "Timeline for $Identity" -AutoSize -FreezeTopRow -BoldTopRow
}

function GetExcelParams($path, $tabName) {
    if ($script:IsOrganizer) {
        $TableStyle = "Light10" # Orange for Organizer
        $TitleExtra = ", Organizer"
    } elseif ($script:IsRoomMB) {
        Write-Host -ForegroundColor green "Room Mailbox Detected"
        $TableStyle = "Light11" # Green for Room Mailbox
        $TitleExtra = ", Resource"
    } else {
        $TableStyle = "Light12" # Light Blue for normal
        # Dark Blue for Delegates (once we can determine this)
    }

    if ($script:CalLogsDisabled) {
        $TitleExtra += ", WARNING: CalLogs are Turned Off for $Identity! This will be a incomplete story"
    }

    return @{
        Path                    = $path
        FreezeTopRow            = $true
        #  BoldTopRow              = $true
        Verbose                 = $false
        TableStyle              = $TableStyle
        WorksheetName           = $tabName
        TableName               = $tabName
        FreezeTopRowFirstColumn = $true
        AutoFilter              = $true
        AutoNameRange           = $true
        Append                  = $true
        Title                   = "Enhanced Calendar Logs for $Identity" + $TitleExtra + " for MeetingID [$($script:GCDO[0].CleanGlobalObjectId)]."
        TitleSize               = 14
        ConditionalText         = $ConditionalFormatting
    }
}

# Need better way of tagging cells than the Range.  Every time one is updated, you need to update all the ones after it.
$ConditionalFormatting = $(
    # Client, ShortClientInfoString and LogClientInfoString
    New-ConditionalText "Outlook" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText "OWA" -ConditionalTextColor DarkGreen -BackgroundColor $null
    New-ConditionalText "Teams" -ConditionalTextColor DarkGreen -BackgroundColor $null
    New-ConditionalText "Transport" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText "Repair" -ConditionalTextColor DarkRed -BackgroundColor LightPink
    New-ConditionalText "Other ?BA" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText "TimeService" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText "Other REST" -ConditionalTextColor DarkRed -BackgroundColor $null
    New-ConditionalText "Unknown" -ConditionalTextColor DarkRed -BackgroundColor $null
    New-ConditionalText "ResourceBookingAssistant" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText "Calendar Replication" -ConditionalTextColor Blue -BackgroundColor $null

    # LogRowType
    New-ConditionalText -Range "C:C" -ConditionalType ContainsText -Text "Interesting" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText -Range "C:C" -ConditionalType ContainsText -Text "SeriesException" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText -Range "C:C" -ConditionalType ContainsText -Text "DeletedSeriesException" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range "C:C" -ConditionalType ContainsText -Text "MeetingMessageChange" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range "C:C" -ConditionalType ContainsText -Text "SyncOrReplication" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText -Range "C:C" -ConditionalType ContainsText -Text "OtherAssistant" -ConditionalTextColor Orange -BackgroundColor $null

    # TriggerAction
    New-ConditionalText -Range "G:G" -ConditionalType ContainsText -Text "Create" -ConditionalTextColor Green -BackgroundColor $null
    New-ConditionalText -Range "G:G" -ConditionalType ContainsText -Text "Delete" -ConditionalTextColor Red -BackgroundColor $null

    # ItemClass
    New-ConditionalText -Range "H:H" -ConditionalType ContainsText -Text "IPM.Appointment" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText -Range "H:H" -ConditionalType ContainsText -Text "Cancellation" -ConditionalTextColor Black -BackgroundColor Orange
    New-ConditionalText -Range "H:H" -ConditionalType ContainsText -Text ".Request" -ConditionalTextColor DarkGreen -BackgroundColor $null
    New-ConditionalText -Range "H:H" -ConditionalType ContainsText -Text ".Resp." -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range "H:H" -ConditionalType ContainsText -Text "IPM.OLE.CLASS" -ConditionalTextColor Plum -BackgroundColor $null

    # FreeBusyStatus
    New-ConditionalText -Range "L3:L9999" -ConditionalType ContainsText -Text "Free" -ConditionalTextColor Red -BackgroundColor $null
    New-ConditionalText -Range "L3:L9999" -ConditionalType ContainsText -Text "Tentative" -ConditionalTextColor Orange -BackgroundColor $null
    New-ConditionalText -Range "L3:L9999" -ConditionalType ContainsText -Text "Busy" -ConditionalTextColor Green -BackgroundColor $null

    # Shared Calendar information
    New-ConditionalText -Range "Q3:Q9999" -ConditionalType Equal -Text "Not Shared" -ConditionalTextColor Blue -BackgroundColor $null
    New-ConditionalText -Range "Q3:Q9999" -ConditionalType Equal -Text "TRUE" -ConditionalTextColor Blue -BackgroundColor Orange

    # MeetingRequestType
    New-ConditionalText -Range "T:T" -ConditionalType ContainsText -Text "Outdated" -ConditionalTextColor DarkRed -BackgroundColor LightPink

    # CalendarItemType
    New-ConditionalText -Range "X3:X9999" -ConditionalType ContainsText -Text "RecurringMaster" -ConditionalTextColor $null -BackgroundColor Plum

    # AppointmentAuxiliaryFlags
    New-ConditionalText -Range "AB3:AB9999" -ConditionalType ContainsText -Text "Copied" -ConditionalTextColor DarkRed -BackgroundColor LightPink
    New-ConditionalText -Range "AA3:AA9999" -ConditionalType ContainsText -Text "ForwardedAppointment" -ConditionalTextColor DarkRed -BackgroundColor $null

    # ResponseType
    New-ConditionalText -Range "AD3:AD9999" -ConditionalType ContainsText -Text "Organizer" -ConditionalTextColor Orange -BackgroundColor $null
)

function FormatHeader {
    param(
        [object] $excel
    )
    $sheet = $excel.Workbook.Worksheets[$ShortId]
    $HeaderRow = 2
    $n = 0

    # Static List of Columns for now...
    $sheet.Column(++$n) | Set-ExcelRange -Width 6 -HorizontalAlignment Center         # LogRow
    Set-CellComment -Text "This is the Enhanced Calendar Logs for [$Identity] for MeetingID `n [$($script:GCDO[0].CleanGlobalObjectId)]." -Row $HeaderRow -ColumnNumber $n -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -NumberFormat "m/d/yyyy h:mm:ss" -HorizontalAlignment Center #LogTimestamp
    Set-CellComment -Text "LogTimestamp: Time when the change was recorded in the CalLogs. This and all Times are in UTC." -Row $HeaderRow -ColumnNumber $n -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # LogRowType
    Set-CellComment -Text "LogRowType: Interesting logs are what to focus on, filter all the others out to start with." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # SubjectProperty
    Set-CellComment -Text "SubjectProperty: The Subject of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # Client
    Set-CellComment -Text "Client (ShortClientInfoString): The 'friendly' Client name of the client that made the change." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 5 -HorizontalAlignment Left          # LogClientInfoString
    Set-CellComment -Text "LogClientInfoString: Full Client Info String of client that made the change." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 12 -HorizontalAlignment Center       # TriggerAction
    Set-CellComment -Text "TriggerAction (CalendarLogTriggerAction): The type of action that caused the change." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 18 -HorizontalAlignment Left         # ItemClass
    Set-CellComment -Text "ItemClass: The Class of the Calendar Item" -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center        # Seq:Exp:ItemVersion
    Set-CellComment -Text "Seq:Exp:ItemVersion (AppointmentLastSequenceNumber:AppointmentSequenceNumber:ItemVersion): The Sequence Version, the Exception Version, and the Item Version.  Each type of item has its own count." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # Organizer
    Set-CellComment -Text "Organizer (From.FriendlyDisplayName): The Organizer of the Calendar Item." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # From
    Set-CellComment -Text "From: The SMTP address of the Organizer of the Calendar Item." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 12 -HorizontalAlignment Center         # FreeBusyStatus
    Set-CellComment -Text "FreeBusy (FreeBusyStatus): The FreeBusy Status of the Calendar Item." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # ResponsibleUser
    Set-CellComment -Text "ResponsibleUser(ResponsibleUserName): The Responsible User of the change." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # Sender
    Set-CellComment -Text "Sender (SenderEmailAddress): The Sender of the change." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 16 -HorizontalAlignment Left         # LogFolder
    Set-CellComment -Text "LogFolder (ParentDisplayName): The Log Folder that the CalLog was in." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 16 -HorizontalAlignment Left         # OriginalLogFolder
    Set-CellComment -Text "OriginalLogFolder (OriginalParentDisplayName): The Original Log Folder that the item was in / delivered to." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 15 -HorizontalAlignment Left         # SharedFolderName
    Set-CellComment -Text "SharedFolderName: Was this from a Modern Sharing, and if so what Folder." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Left         # ReceivedRepresenting
    Set-CellComment -Text "ReceivedRepresenting: Who the item was Received for, of then the Delegate." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center         # MeetingRequestType
    Set-CellComment -Text "MeetingRequestType: The Meeting Request Type of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 23 -NumberFormat "m/d/yyyy h:mm:ss" -HorizontalAlignment Center         # StartTime
    Set-CellComment -Text "StartTime: The Start Time of the Meeting. This and all Times are in UTC." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 23 -NumberFormat "m/d/yyyy h:mm:ss" -HorizontalAlignment Center         # EndTime
    Set-CellComment -Text "EndTime: The End Time of the Meeting. This and all Times are in UTC." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 17 -NumberFormat "m/d/yyyy h:mm:ss"  -HorizontalAlignment Left         # OriginalStartDate
    Set-CellComment -Text "OriginalStartDate: The Original Start Date of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Left         # Location
    Set-CellComment -Text "Location: The Location of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 15 -HorizontalAlignment Center         # CalendarItemType
    Set-CellComment -Text "CalendarItemType: The Calendar Item Type of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left         # RecurrencePattern
    Set-CellComment -Text "RecurrencePattern: The Recurrence Pattern of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 30 -HorizontalAlignment Center       # AppointmentAuxiliaryFlags
    Set-CellComment -Text "AppointmentAuxiliaryFlags: The Appointment Auxiliary Flags of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 30 -HorizontalAlignment Left         # DisplayAttendeesAll
    Set-CellComment -Text "DisplayAttendeesAll: List of the Attendees of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center        # AttendeeCount
    Set-CellComment -Text "AttendeeCount: The Attendee Count." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Left          # AppointmentState
    Set-CellComment -Text "AppointmentState: The Appointment State of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center         # ResponseType
    Set-CellComment -Text "ResponseType: The Response Type of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 20 -HorizontalAlignment Center         # ClientIntent
    Set-CellComment -Text "ClientIntent: The Client Intent of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center         # AppointmentRecurring
    Set-CellComment -Text "AppointmentRecurring: Is this a Recurring Meeting?" -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center         # HasAttachment
    Set-CellComment -Text "HasAttachment: Does this Meeting have an Attachment?" -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center         # IsCancelled
    Set-CellComment -Text "IsCancelled: Is this Meeting Cancelled?" -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center         # IsAllDayEvent
    Set-CellComment -Text "IsAllDayEvent: Is this an All Day Event?" -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 10 -HorizontalAlignment Center         # IsSeriesCancelled
    Set-CellComment -Text "IsSeriesCancelled: Is this a Series Cancelled Meeting?" -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 30 -HorizontalAlignment Left           # SendMeetingMessagesDiagnostics
    Set-CellComment -Text "SendMeetingMessagesDiagnostics: Compound Property to describe why meeting was or was not sent to everyone." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 50 -HorizontalAlignment Left           # AttendeeCollection
    Set-CellComment -Text "AttendeeCollection: The Attendee Collection of the Meeting, use -TrackingLogs to get values." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet
    $sheet.Column(++$n) | Set-ExcelRange -Width 40 -HorizontalAlignment Center          # CalendarLogRequestId
    Set-CellComment -Text "CalendarLogRequestId: The Calendar Log Request ID of the Meeting." -Row $HeaderRow -ColumnNumber $n  -Worksheet $sheet

    # Update header rows after all the others have been set.
    # Title Row
    $sheet.Row(1) | Set-ExcelRange -HorizontalAlignment Left
    Set-CellComment -Text "For more information see: Https:\\aka.ms\AnalyzeCalLogs"  -Row 1 -ColumnNumber 1  -Worksheet $sheet

    # Set the Header row to be bold and left aligned
    $sheet.Row($HeaderRow) | Set-ExcelRange -Bold -HorizontalAlignment Left
}
}

# Default to Collecting Exceptions
if ((!$NoExceptions.IsPresent) -and ([string]::IsNullOrEmpty($ExceptionDate))) {
    $Exceptions=$true
    Write-Host -ForegroundColor Yellow "Collecting Exceptions."
    Write-Host -ForegroundColor Yellow "`tTo not collecting Exceptions, use the -NoExceptions switch."
} else {
    Write-Host -ForegroundColor Green "---------------------------------------"
    if ($NoExceptions.IsPresent) {
        Write-Host -ForegroundColor Green "Not Checking for Exceptions"
    } else {
        Write-Host -ForegroundColor Green "Checking for Exceptions on $ExceptionDate"
    }
    Write-Host -ForegroundColor Green "---------------------------------------"
}

# ===================================================================================================
# Main
# ===================================================================================================

$ValidatedIdentities = CheckIdentities -Identity $Identity

if ($ExportToExcel.IsPresent) {
    CheckExcelModuleInstalled
}

if (-not ([string]::IsNullOrEmpty($Subject)) ) {
    if ($ValidatedIdentities.count -gt 1) {
        Write-Warning "Multiple mailboxes were found, but only one is supported for Subject searches.  Please specify a single mailbox."
        exit
    }
    $script:Identity = $ValidatedIdentities
    GetCalLogsWithSubject -Identity $ValidatedIdentities -Subject $Subject
} elseif (-not ([string]::IsNullOrEmpty($MeetingID))) {
    # Process Logs based off Passed in MeetingID
    foreach ($ID in $ValidatedIdentities) {
        Write-DashLineBoxColor "Looking for CalLogs from [$ID] with passed in MeetingID."
        Write-Verbose "Running: Get-CalendarDiagnosticObjects -Identity [$ID] -MeetingID [$MeetingID] -CustomPropertyNames $CustomPropertyNameList -WarningAction Ignore -MaxResults $LogLimit -ResultSize $LogLimit -ShouldBindToItem $true;"
        [array] $script:GCDO = GetCalendarDiagnosticObjects -Identity $ID -MeetingID $MeetingID
        $script:Identity = $ID
        if ($script:GCDO.count -gt 0) {
            Write-Host -ForegroundColor Cyan "Found $($script:GCDO.count) CalLogs with MeetingID [$MeetingID]."
            $script:IsOrganizer = (SetIsOrganizer -CalLogs $script:GCDO)
            Write-Host -ForegroundColor Cyan "The user [$ID] $(if ($IsOrganizer) {"IS"} else {"is NOT"}) the Organizer of the meeting."

            $script:IsRoomMB = (SetIsRoom -CalLogs $script:GCDO)
            if ($script:IsRoomMB) {
                Write-Host -ForegroundColor Cyan "The user [$ID] is a Room Mailbox."
            }

            if (CheckForBifurcation($script:GCDO) -ne false) {
                Write-Host -ForegroundColor Red "Warning: No IPM.Appointment found. CalLogs start to expire after 31 days."
            }

            if ($Exceptions.IsPresent) {
                Write-Verbose "Looking for Exception Logs..."
                $IsRecurring = SetIsRecurring -CalLogs $script:GCDO
                Write-Verbose "Meeting IsRecurring: $IsRecurring"

                if ($IsRecurring) {
                    #collect Exception Logs
                    $ExceptionLogs = @()
                    $LogToExamine = @()
                    $LogToExamine = $script:GCDO | Where-Object { $_.ItemClass -like 'IPM.Appointment*' } | Sort-Object ItemVersion

                    Write-Host -ForegroundColor Cyan "Found $($LogToExamine.count) CalLogs to examine for Exception Logs."
                    if ($LogToExamine.count -gt 100) {
                        Write-Host -ForegroundColor Cyan "`t This is a large number of logs to examine, this may take a while."
                    }
                    $logLeftCount = $LogToExamine.count

                    $ExceptionLogs = $LogToExamine | ForEach-Object {
                        $logLeftCount -= 1
                        Write-Verbose "Getting Exception Logs for [$($_.ItemId.ObjectId)]"
                        Get-CalendarDiagnosticObjects -Identity $ID -ItemIds $_.ItemId.ObjectId -ShouldFetchRecurrenceExceptions $true -CustomPropertyNames $CustomPropertyNameList -ShouldBindToItem $true -WarningAction SilentlyContinue
                        if (($logLeftCount % 10 -eq 0) -and ($logLeftCount -gt 0)) {
                            Write-Host -ForegroundColor Cyan "`t [$($logLeftCount)] logs left to examine..."
                        }
                    }
                    # Remove the IPM.Appointment logs as they are already in the CalLogs.
                    $ExceptionLogs = $ExceptionLogs | Where-Object { $_.ItemClass -notlike "IPM.Appointment*" }
                    Write-Host -ForegroundColor Cyan "Found $($ExceptionLogs.count) Exception Logs, adding them into the CalLogs."

                    $script:GCDO = $script:GCDO + $ExceptionLogs | Select-Object *, @{n='OrgTime'; e= { [DateTime]::Parse($_.LogTimestamp.ToString()) } } | Sort-Object OrgTime
                    $LogToExamine = $null
                    $ExceptionLogs = $null
                } else {
                    Write-Host -ForegroundColor Cyan "No Recurring Meetings found, no Exception Logs to collect."
                }
            }

            BuildCSV
            BuildTimeline
        } else {
            Write-Warning "No CalLogs were found for [$ID] with MeetingID [$MeetingID]."
        }
    }
} else {
    Write-Warning "A valid MeetingID was not found, nor Subject. Please confirm the MeetingID or Subject and try again."
}

Write-DashLineBoxColor "Hope this script was helpful in getting and understanding the Calendar Logs.",
"More Info on Getting the logs: https://aka.ms/GetCalLogs",
"and on Analyzing the logs: https://aka.ms/AnalyzeCalLogs",
"If you have issues or suggestion for this script, please send them to: ",
"`t CalLogFormatterDevs@microsoft.com" -Color Yellow -DashChar "="

if ($ExportToExcel.IsPresent) {
    Write-Host
    Write-Host -ForegroundColor Blue -NoNewline "All Calendar Logs are saved to: "
    Write-Host -ForegroundColor Yellow ".\$Filename"
}

# SIG # Begin signature block
# MIIoDwYJKoZIhvcNAQcCoIIoADCCJ/wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA0t2/fBfzRC7KL
# ESCPS+YE1uFQlpUvkVgziKGUxEtusqCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
# Bv9XKydyAAAAAAQEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTE0WhcNMjUwOTExMjAxMTE0WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC0KDfaY50MDqsEGdlIzDHBd6CqIMRQWW9Af1LHDDTuFjfDsvna0nEuDSYJmNyz
# NB10jpbg0lhvkT1AzfX2TLITSXwS8D+mBzGCWMM/wTpciWBV/pbjSazbzoKvRrNo
# DV/u9omOM2Eawyo5JJJdNkM2d8qzkQ0bRuRd4HarmGunSouyb9NY7egWN5E5lUc3
# a2AROzAdHdYpObpCOdeAY2P5XqtJkk79aROpzw16wCjdSn8qMzCBzR7rvH2WVkvF
# HLIxZQET1yhPb6lRmpgBQNnzidHV2Ocxjc8wNiIDzgbDkmlx54QPfw7RwQi8p1fy
# 4byhBrTjv568x8NGv3gwb0RbAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU8huhNbETDU+ZWllL4DNMPCijEU4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMjkyMzAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjmD9IpQVvfB1QehvpC
# Ge7QeTQkKQ7j3bmDMjwSqFL4ri6ae9IFTdpywn5smmtSIyKYDn3/nHtaEn0X1NBj
# L5oP0BjAy1sqxD+uy35B+V8wv5GrxhMDJP8l2QjLtH/UglSTIhLqyt8bUAqVfyfp
# h4COMRvwwjTvChtCnUXXACuCXYHWalOoc0OU2oGN+mPJIJJxaNQc1sjBsMbGIWv3
# cmgSHkCEmrMv7yaidpePt6V+yPMik+eXw3IfZ5eNOiNgL1rZzgSJfTnvUqiaEQ0X
# dG1HbkDv9fv6CTq6m4Ty3IzLiwGSXYxRIXTxT4TYs5VxHy2uFjFXWVSL0J2ARTYL
# E4Oyl1wXDF1PX4bxg1yDMfKPHcE1Ijic5lx1KdK1SkaEJdto4hd++05J9Bf9TAmi
# u6EK6C9Oe5vRadroJCK26uCUI4zIjL/qG7mswW+qT0CW0gnR9JHkXCWNbo8ccMk1
# sJatmRoSAifbgzaYbUz8+lv+IXy5GFuAmLnNbGjacB3IMGpa+lbFgih57/fIhamq
# 5VhxgaEmn/UjWyr+cPiAFWuTVIpfsOjbEAww75wURNM1Imp9NJKye1O24EspEHmb
# DmqCUcq7NqkOKIG4PVm3hDDED/WQpzJDkvu4FrIbvyTGVU01vKsg4UfcdiZ0fQ+/
# V0hf8yrtq9CkB8iIuk5bBxuPMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGe8wghnrAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggZAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwLwYJKoZIhvcNAQkEMSIE
# IDwgetJ5jLfyEvRG7bIsvBEP+53PnpSYvUtzqZtLrzQZMEIGCisGAQQBgjcCAQwx
# NDAyoBSAEgBNAGkAYwByAG8AcwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29m
# dC5jb20wDQYJKoZIhvcNAQEBBQAEggEApLFx2t2RRz5velreIbr05EwrXUAZCSq0
# XzPORM9cZMDJwvD5viUwpU7W+M2ZV1u4jf5TYXBi2QoIak4V61x29CmJGxIsXB+P
# j+po0wNAejXEcnIVQvDIwKKJ19zGRPr30DXoKmdfuYbllEb/6Ckv1KPbrsQNDVwB
# pt/t+LP2oeB3Rcsd/vaJ+JYj6cQaH4b+43nirWovOAbPGebHoRlqld8abPXWvAY5
# X1FARh4mJ/o2L6CHZiAmyBnUzYVNMSBxRMO2WI+w+rrhUU4Y8RW6H5QfRGNXlmt3
# XpvBnDfAzqkp7ytnfKdaQ5RToKxGWOA29fcNSZqgnWIIy+bswpjKqqGCF5cwgheT
# BgorBgEEAYI3AwMBMYIXgzCCF38GCSqGSIb3DQEHAqCCF3AwghdsAgEDMQ8wDQYJ
# YIZIAWUDBAIBBQAwggFSBgsqhkiG9w0BCRABBKCCAUEEggE9MIIBOQIBAQYKKwYB
# BAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCDEBppqq3s/n8FSQL5Q8+fEllmup9OY
# d2eXjZcy1tId1wIGZ/gQj9psGBMyMDI1MDQxODIwMTE0OC44MTlaMASAAgH0oIHR
# pIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYD
# VQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hp
# ZWxkIFRTUyBFU046ODkwMC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFNlcnZpY2WgghHtMIIHIDCCBQigAwIBAgITMwAAAg4syyh9lSB1
# YwABAAACDjANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMDAeFw0yNTAxMzAxOTQzMDNaFw0yNjA0MjIxOTQzMDNaMIHLMQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQg
# QW1lcmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046ODkw
# MC0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCs5t7iRtXt0hbeo9ME
# 78ZYjIo3saQuWMBFQ7X4s9vooYRABTOf2poTHatx+EwnBUGB1V2t/E6MwsQNmY5X
# pM/75aCrZdxAnrV9o4Tu5sBepbbfehsrOWRBIGoJE6PtWod1CrFehm1diz3jY3H8
# iFrh7nqefniZ1SnbcWPMyNIxuGFzpQiDA+E5YS33meMqaXwhdb01Cluymh/3EKvk
# nj4dIpQZEWOPM3jxbRVAYN5J2tOrYkJcdDx0l02V/NYd1qkvUBgPxrKviq5kz7E6
# AbOifCDSMBgcn/X7RQw630Qkzqhp0kDU2qei/ao9IHmuuReXEjnjpgTsr4Ab33IC
# AKMYxOQe+n5wqEVcE9OTyhmWZJS5AnWUTniok4mgwONBWQ1DLOGFkZwXT334IPCq
# d4/3/Ld/ItizistyUZYsml/C4ZhdALbvfYwzv31Oxf8NTmV5IGxWdHnk2Hhh4bnz
# TKosEaDrJvQMiQ+loojM7f5bgdyBBnYQBm5+/iJsxw8k227zF2jbNI+Ows8HLeZG
# t8t6uJ2eVjND1B0YtgsBP0csBlnnI+4+dvLYRt0cAqw6PiYSz5FSZcbpi0xdAH/j
# d3dzyGArbyLuo69HugfGEEb/sM07rcoP1o3cZ8eWMb4+MIB8euOb5DVPDnEcFi4N
# DukYM91g1Dt/qIek+rtE88VS8QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFIVxRGlS
# EZE+1ESK6UGI7YNcEIjbMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1Gely
# MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
# cHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
# bDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBD
# QSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQB14L2TL+L8
# OXLxnGSal2h30mZ7FsBFooiYkUVOY05F9pnwPTVufEDGWEpNNy2OfaUHWIOoQ/9/
# rjwO0hS2SpB0BzMAk2gyz92NGWOpWbpBdMvrrRDpiWZi/uLS4ZGdRn3P2DccYmlk
# NP+vaRAXvnv+mp27KgI79mJ9hGyCQbvtMIjkbYoLqK7sF7Wahn9rLjX1y5QJL4lv
# Ey3QmA9KRBj56cEv/lAvzDq7eSiqRq/pCyqyc8uzmQ8SeKWyWu6DjUA9vi84QsmL
# jqPGCnH4cPyg+t95RpW+73snhew1iCV+wXu2RxMnWg7EsD5eLkJHLszUIPd+XClD
# +FTvV03GfrDDfk+45flH/eKRZc3MUZtnhLJjPwv3KoKDScW4iV6SbCRycYPkqoWB
# rHf7SvDA7GrH2UOtz1Wa1k27sdZgpG6/c9CqKI8CX5vgaa+A7oYHb4ZBj7S8u8sg
# xwWK7HgWDRByOH3CiJu4LJ8h3TiRkRArmHRp0lbNf1iAKuL886IKE912v0yq55t8
# jMxjBU7uoLsrYVIoKkzh+sAkgkpGOoZL14+dlxVM91Bavza4kODTUlwzb+SpXsSq
# Vx8nuB6qhUy7pqpgww1q4SNhAxFnFxsxiTlaoL75GNxPR605lJ2WXehtEi7/+YfJ
# qvH+vnqcpqCjyQ9hNaVzuOEHX4MyuqcjwjCCB3EwggVZoAMCAQICEzMAAAAVxedr
# ngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4
# MzIyNVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qls
# TnXIyjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLA
# EBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrE
# qv1yaa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyF
# Vk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1o
# O5pGve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg
# 3viSkR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2
# TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07B
# MzlMjgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJ
# NmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6
# r1AFemzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+
# auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3
# FQIEFgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl
# 0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUH
# AgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0
# b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMA
# dQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAW
# gBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8v
# Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRf
# MjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEw
# LTA2LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL
# /Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu
# 6WZnOlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5t
# ggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfg
# QJY4rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8s
# CXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCr
# dTDFNLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZ
# c9d/HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2
# tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8C
# wYKiexcdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9
# JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDB
# cQZqELQdVTNYs6FwZvKhggNQMIICOAIBATCB+aGB0aSBzjCByzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFt
# ZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOjg5MDAt
# MDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
# oiMKAQEwBwYFKw4DAhoDFQBK6HY/ZWLnOcMEQsjkDAoB/JZWCKCBgzCBgKR+MHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBCwUAAgUA660a
# WzAiGA8yMDI1MDQxODE4MzcxNVoYDzIwMjUwNDE5MTgzNzE1WjB3MD0GCisGAQQB
# hFkKBAExLzAtMAoCBQDrrRpbAgEAMAoCAQACAioyAgH/MAcCAQACAhSqMAoCBQDr
# rmvbAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMH
# oSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBABfWi1mC+dRpvBjp0Qe3
# thoz2g035Nz/LuNKA63so80wq0tVC1KiDwLYNFDbEgbRvxWJgA4Gtled+sPIrr1w
# rT2ONzlmpi+vNx5l1oXPB5H8uYkMz9tnRTJgM1D+PRPmBqMMiJwDhRGLbV7V3z60
# eg/skUjHCFjbi3+E1Zm3ruHpKRypYz5imG98I/R9STM5I5S+vzE3xvgb3ID9PW4/
# vBsR0s9tZu2ay3SxtAZQ1Os0xE7hdy2kNT5KB+bqNtEA4TUnmLyjsQu7CdMUC3GQ
# cj5cakkRil0bKWuP3fj4S+SqHnDAQzmoETO/xiTfdWKOGEr3jM2EU5kgPWn+mx93
# Q+MxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAIT
# MwAAAg4syyh9lSB1YwABAAACDjANBglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcN
# AQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCD1tW4oP9TDZ0fVNEwT
# vtmze9FMu1UW2CwEb9BT0ChJRzCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQwgb0E
# IAF0HXMl8OmBkK267mxobKSihwOdP0eUNXQMypPzTxKGMIGYMIGApH4wfDELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
# b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIOLMsofZUgdWMAAQAAAg4wIgQg
# 5IEymE3w/NHd2CchR3v6c6DmnvnwaW9iBWlPFCBX1agwDQYJKoZIhvcNAQELBQAE
# ggIAf5PTn2sXKr6WxMbJW5XnielzA/4VYPrGMjsAayBO8pL9q5//cX7pfmjWzRaG
# UgriOzHIWQOJlBMUhLl5jHlYnj5Sflq5V5NrMdlbISwelsD20MsX87MnEdkX+42f
# I4aqM+ICmdGLNMMiCvRv1qFPPjqagHIzrY7NE3RszKx8wzG5a2X0NspJ7AITWSNr
# ZGq/FGGvsqtLEJ6+njs9l7uvZnMK/sxx0gDboaHgh0jzvUCd3/+AGSO5WLyxeOBp
# RCgcTMACe4mgwSiQC6KcW16w6fSZpNvWuOyOF+QGcAE9kiXpkfay9MfDoiKtby6c
# dMSG7HGY0NrAxtLVZ0jaGadJY62A25DGZmxQ5JYo460SkEBagmQQ1axe0xYPGUiF
# XYtQ37Ve4hMOTL5cPafn78B+ZY84oiU1ofpiyWEjw05VP5lJTMU839VSTmDH/6MV
# kYRcKqrvhsxjUpO3mDIXpOCOf/p4WdzggjVCiOHNCb24stE4JIvnC+illSnlANcB
# yHYtdjcpzZhwTAllluU+hNGifj236q7a/q+Lkjq0em992TFyQhMWE+VJQChzSCAS
# QQOJ5QT0xZebOa7WDFynWe+X2rlzVOuFuTLLEoAtBsR+zhO5GMQfy14xVyZdbv2z
# lp6rw6vjRAedWSoXgTMKTYlgbuhmrolJa2kjBXKEpZZbHhI=
# SIG # End signature block
