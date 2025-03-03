# mhytic issue
# enable aut expandign archive as a message trow only mailbox name instead of full Auto Expanding Archive enabled for $mailbox its wring it shot show whow  sentence, confirm if part throw "In-Place Archive is not enabled for $mailbox. Enable it before enabling Auto Expanding Archive." will populate and display correctly

# co poprawic
#1 dodac osobne pola textowe do checkow zeby sie nei refreshowaly po kazdym kliku ? ale czy na pewno ma to sens
#2 przeniesc button 9 pod button 5 tak zebyb byl 6
#3 poprawic nazwe 1 buttona
#4 sprawdzic czy wszyskie checki sa i czy wszystkie sa potrzebne
#5 dac userowi wybor czy chce zobaczyc outgris view po tym jak zobaczy wynik w polu textowym i otworzyc jest user bedzie checial
#6 dodac na czewono info ze user neima archovum enabled, user nei ma in-pace archive enabled user nei ma auroexpandign archive enabled
#7 dodac instrukcje na co patrzec w czasie trubbleshotingu i podswietlic na czerwono value ktore za do zmiany
#8 spacje przed i po emailu psuja skrypt

# Add System.Windows.Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a new form
$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Exchange Mailbox Capacity Management Tool"
$Form.Size = New-Object System.Drawing.Size(520,600)

# Create buttons and other controls
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(10,10)
$Label.Size = New-Object System.Drawing.Size(480,23)

$TextBox = New-Object System.Windows.Forms.TextBox
$TextBox.Location = New-Object System.Drawing.Size(10,40)
$TextBox.Size = New-Object System.Drawing.Size(480,23)

# Create labels for button groups
$Label1 = New-Object System.Windows.Forms.Label
$Label1.Location = New-Object System.Drawing.Size(10,70)
$Label1.Size = New-Object System.Drawing.Size(480,23)
$Label1.Text = "Check Mailbox"

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Size(10,310)
$Label2.Size = New-Object System.Drawing.Size(480,23)
$Label2.Text = "Fix Mailbox"

# Create a shared runspace
$Runspace = [runspacefactory]::CreateRunspace()
$Runspace.Open()

# Create a single PowerShell instance
$PowerShell = [powershell]::Create()
$PowerShell.Runspace = $Runspace

# Login to Exchange Online
$PowerShell.AddScript({
    try {
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowProgress $true
        return "Logged in successfully"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
        return "Login failed"
    }
})
$Label.Text = $PowerShell.Invoke()[0]
$PowerShell.Commands.Clear()

# Create buttons and assign click events
# Buttons for checking the mailbox
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Location = New-Object System.Drawing.Size(10,100)
$Button1.Size = New-Object System.Drawing.Size(480,23)
$Button1.Text = "Check LH,Archive,Retention on the mailbox"
$Button1.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                $mailboxInfo = Get-Mailbox -Identity $mailbox | select-object UserPrincipalName,retentionpolicy,retaindeleteditemsfor, litigationholdenabled, archiveguid,archivestatus,retentionholdenabled,ElcProcessingDisabled, DelayHoldApplied,DelayReleaseHoldApplied, SingleItemRecoveryEnable, ArchiveDatabase, ArchiveGuid, ArchiveName, ArchiveQuota, ArchiveStatus, ArchiveWarningQuota, AutoExpandingArchiveEnabled, DisabledArchiveDatabase, DisabledArchiveGuid, EnableArchive, RetentionUrl, RetentionVersion
                $mailboxInfo | Out-GridView
                $archiveStatus = $mailboxInfo.ArchiveStatus
                $archiveQuota = $mailboxInfo.ArchiveQuota
                $autoExpandingArchiveEnabled = $mailboxInfo.AutoExpandingArchiveEnabled
                return "Archive Status: $archiveStatus, Archive Quota: $archiveQuota, Auto Expanding Archive Enabled: $autoExpandingArchiveEnabled"
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Label.Text = $PowerShell.Invoke()[0]
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})

$Button2 = New-Object System.Windows.Forms.Button
$Button2.Location = New-Object System.Drawing.Size(10,130)
$Button2.Size = New-Object System.Drawing.Size(480,23)
$Button2.Text = "Check Folder Statistics"
$Button2.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                Get-MailboxFolderStatistics -Identity $mailbox -ResultSize unlimited  |Select-Object Name, FolderPath, FolderType, StorageQuota,StorageWarningQuota,VisibleItemsInFolder,HiddenItemsInFolder, ItemsInFolder,FolderSize,ItemsInFolderAndSubfolders,FolderAndSubfolderSize,ArchivePolicy,DeletePolicy,CompliancePolicy,RetentionFlags | Out-GridView
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Result = $PowerShell.Invoke()
        $Label.Text = "Completed"
        $Result | Out-GridView
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})

$Button3 = New-Object System.Windows.Forms.Button
$Button3.Location = New-Object System.Drawing.Size(10, 160)
$Button3.Size = New-Object System.Drawing.Size(480, 23)
$Button3.Text = "Check Archive Folder Statistics"
$Button3.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                $stats = Get-MailboxFolderStatistics -Identity $mailbox -Archive -ResultSize unlimited |
                    Select-Object Name, FolderPath, FolderType, StorageQuota, StorageWarningQuota, VisibleItemsInFolder, HiddenItemsInFolder, ItemsInFolder, FolderSize, ItemsInFolderAndSubfolders, FolderAndSubfolderSize, ArchivePolicy, DeletePolicy, CompliancePolicy, RetentionFlags
                
                if ($stats) {
                    $stats
                } else {
                    throw "Archive folder doesn't exist for the mailbox: $mailbox"
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        
        $Result = $PowerShell.Invoke()
        $Label.Text = "Completed"
        $Result | Out-GridView
        $PowerShell.Commands.Clear()
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})


$Button4 = New-Object System.Windows.Forms.Button
$Button4.Location = New-Object System.Drawing.Size(10,190)
$Button4.Size = New-Object System.Drawing.Size(480,23)
$Button4.Text = "Check if In-Place Archive is Enabled"
$Button4.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                $mailboxInfo = Get-Mailbox -Identity $mailbox
                if ($mailboxInfo.ArchiveStatus -eq "Active") {
                    return "In-Place Archive is enabled for $mailbox"
                } else {
                    return "In-Place Archive is not enabled for $mailbox"
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Label.Text = $PowerShell.Invoke()[0]
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})

$Button5 = New-Object System.Windows.Forms.Button
$Button5.Location = New-Object System.Drawing.Size(10,340)
$Button5.Size = New-Object System.Drawing.Size(480,23)
$Button5.Text = "Enable Auto Expanding Archive"
$Button5.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                $mailboxInfo = Get-Mailbox -Identity $mailbox
                if ($mailboxInfo.ArchiveStatus -eq "Active") {
                    Enable-Mailbox -Identity $mailbox -AutoExpandingArchive
                    return "Auto Expanding Archive enabled for $mailbox" # to wyswietla tylko nawe mailboxa, nei wiem czemu naprawic albo wyalic calkiem
                } else {
                    return "In-Place Archive is not enabled for $mailbox. Enable it before enabling Auto Expanding Archive."
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Label.Text = $PowerShell.Invoke()[0]
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})

$Button6 = New-Object System.Windows.Forms.Button
$Button6.Location = New-Object System.Drawing.Size(10,370)
$Button6.Size = New-Object System.Drawing.Size(480,23)
$Button6.Text = "Start Managed Folder Assistant"
$Button6.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                Start-ManagedFolderAssistant -Identity $mailbox
                return "Managed Folder Assistant started for $mailbox"
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Label.Text = $PowerShell.Invoke()[0]
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})

$Button7 = New-Object System.Windows.Forms.Button
$Button7.Location = New-Object System.Drawing.Size(10,400)
$Button7.Size = New-Object System.Drawing.Size(480,23)
$Button7.Text = "Start Managed Folder Assistant with Hold Cleanup"
$Button7.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                Start-ManagedFolderAssistant -Identity $mailbox -HoldCleanup
                return "Managed Folder Assistant with Hold Cleanup started for $mailbox"
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Label.Text = $PowerShell.Invoke()[0]
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})

$Button8 = New-Object System.Windows.Forms.Button
$Button8.Location = New-Object System.Drawing.Size(10,430)
$Button8.Size = New-Object System.Drawing.Size(480,23)
$Button8.Text = "Start Managed Folder Assistant with Full Crawl"
$Button8.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                Start-ManagedFolderAssistant -Identity $mailbox -FullCrawl
                return "Managed Folder Assistant with Full Crawl started for $mailbox"
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Label.Text = $PowerShell.Invoke()[0]
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})

$Button9 = New-Object System.Windows.Forms.Button
$Button9.Location = New-Object System.Drawing.Size(10,460)
$Button9.Size = New-Object System.Drawing.Size(480,23)
$Button9.Text = "Enable In-Place Archive"
$Button9.Add_Click({
    if ($TextBox.Text -ne "") {
        $mailbox = $TextBox.Text
        $PowerShell.AddScript({
            param($mailbox)
            try {
                $mailboxInfo = Get-Mailbox -Identity $mailbox
                if ($mailboxInfo.ArchiveStatus -ne "Active") {
                    Enable-Mailbox $mailbox -Archive
                    return "In-Place Archive enabled for $mailbox"
                } else {
                    return "In-Place Archive is already enabled for $mailbox"
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error: " + $_.Exception.Message)
            }
        }).AddParameter('mailbox', $mailbox)
        $Label.Text = $PowerShell.Invoke()[0]
        $PowerShell.Commands.Clear()
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a mailbox identity.")
    }
})


# Add controls to the form
$Form.Controls.Add($Button1)
$Form.Controls.Add($Button2)
$Form.Controls.Add($Button3)
$Form.Controls.Add($Button4)
$Form.Controls.Add($Button5)
$Form.Controls.Add($Button6)
$Form.Controls.Add($Button7)
$Form.Controls.Add($Button8)
$Form.Controls.Add($Button9)
$Form.Controls.Add($TextBox)
$Form.Controls.Add($Label)

# Show the form
$Form.ShowDialog()


#Working Tested Implmented
