# Connect to MS Teams
Connect-MicrosoftTeams

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mastery of Teams Policies: The Selector Edition"
$form.Size = New-Object System.Drawing.Size(500, 320)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

# Define the label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(460, 30)
$label.Text = "Ah, my dear friend, welcome to the policy search portal. Now, pray tell me, what manner of policy tickles your fancy?"

# Add the label to the form
$form.Controls.Add($label)

# Define the list box
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 50)
$listBox.Size = New-Object System.Drawing.Size(460, 180)

# Populate the list box with policy types
$policies = @("Teams App Setup", "Teams App Permission")
$policies | ForEach-Object { [void]$listBox.Items.Add($_) }

# Add the list box to the form
$form.Controls.Add($listBox)

# Define the OK button
$buttonOK = New-Object System.Windows.Forms.Button
$buttonOK.Location = New-Object System.Drawing.Point(210, 240)
$buttonOK.Size = New-Object System.Drawing.Size(75, 23)
$buttonOK.Text = "OK"

# Add the OK button event handler
$buttonOK.Add_Click({
        $global:selectedPolicyType = $listBox.SelectedItem
        $form.Close()
        if ($selectedPolicyType) {
            [System.Windows.Forms.MessageBox]::Show("You selected '$selectedPolicyType'", "Selection", "OK", "Information")
            # Play a sound after popup appears
        (New-Object Media.SoundPlayer "C:\Windows\Media\tada.wav").Play()
            Write-Host "Selected Policy Type: $selectedPolicyType"
        }
    })

# Add the OK button to the form
$form.Controls.Add($buttonOK)

# Define the Cancel button
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(295, 240)
$buttonCancel.Size = New-Object System.Drawing.Size(75, 23)
$buttonCancel.Text = "Cancel"

# Add the Cancel button event handler
$buttonCancel.Add_Click({
        $form.Close()
        return  # Add this line to stop the script
    })

# Add the Cancel button to the form
$form.Controls.Add($buttonCancel)

# Show the form
[void]$form.ShowDialog()


if ($selectedPolicyType -eq "Teams App Permission") {

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Define the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Teams App Permission Policy Selector"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    
    # Define the list box
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 10)
    $listBox.Size = New-Object System.Drawing.Size(360, 180)
    
    # Populate the list box with policies
    $policies = Get-CsTeamsAppPermissionPolicy
    $policies | ForEach-Object { [void]$listBox.Items.Add($_.Identity.Replace('Tag:', '')) }
    
    # Add the list box to the form
    $form.Controls.Add($listBox)
    
    # Define the OK button
    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Location = New-Object System.Drawing.Point(210, 200)
    $buttonOK.Size = New-Object System.Drawing.Size(75, 23)
    $buttonOK.Text = "OK"
    
    # Add the OK button event handler
    $buttonOK.Add_Click({
            $global:selectedPolicy = $listBox.SelectedItem
            $form.Close()
            if ($selectedPolicy) {
                [System.Windows.Forms.MessageBox]::Show("You selected '$selectedPolicy'", "Selection", "OK", "Information")
                # Play a sound after popup appears
            (New-Object Media.SoundPlayer "C:\Windows\Media\tada.wav").Play()
                Write-Host "Selected Policy: $selectedPolicy"

                Write-Host "Attention, all passengers! 
    This script might take longer than a visit to the Unseen University library. 
    Yes, that's right, you might as well bring along some of Ridcully's best brandy and settle in for the long haul. 
    But don't worry, we'll keep you entertained along the way with a variety of fun sound effects, including the occasional 'Ook!' from the Librarian.
    
    And remember what Death once said: 
    'TIME IS A DRAGONE WITH A TRIPLE HEADED TAIL, SWEEPING OUT OF THE PAST, THROUGH THE PRESENT, AND INTO THE FUTURE. UNLESS YOU HAVE A VERY SHARP SWORD, IT'S DIFFICULT TO GET A GRIP ON.' 
    So sit back, relax, and let the script do its thing!"

    
            }
        })
    
    # Add the OK button to the form
    $form.Controls.Add($buttonOK)
    
    # Define the Cancel button
    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Location = New-Object System.Drawing.Point(295, 200)
    $buttonCancel.Size = New-Object System.Drawing.Size(75, 23)
    $buttonCancel.Text = "Cancel"
    
    # Add the Cancel button event handler
    $buttonCancel.Add_Click({
            $form.Close()
            return
        })
    
    # Add the Cancel button to the form
    $form.Controls.Add($buttonCancel)
    
    # Show the form
    [void]$form.ShowDialog()
    
    # Create an array to hold error messages
    $errors = @()
    
    # Get all users in the organization
    $users = Get-CsOnlineUser
    
    # Loop through each user and check their assigned policies
    foreach ($user in $users) {
        try {
            $policies = Get-CsUserPolicyAssignment -Identity $user.UserPrincipalName -ErrorAction Stop
    
            # Check if the selected policy is assigned to the user
            $hasSelectedPolicy = $policies.PolicyName -eq $global:selectedPolicy
    
            # If the user has the selected policy assigned, list their name and UPN
            if ($hasSelectedPolicy) {
                Write-Host "User:" $user.DisplayName "UPN:" $user.UserPrincipalName "has $($global:selectedPolicy) policy assigned."
                # Play a different sound after each found user
                (New-Object Media.SoundPlayer "C:\Windows\Media\chimes.wav").Play()
            }
        }
        catch {
            
            # Add error message to the errors array
            $errorMessage = New-Object PSObject -Property @{
                UserPrincipalName = $user.UserPrincipalName
                ErrorMessage      = $_.Exception.Message
            }
            $errors += $errorMessage
        }
    }
    
    # Show message box to ask if errors should be exported
    $exportErrors = [System.Windows.Forms.MessageBox]::Show("Do you want to export errors to a CSV file?", "Export Errors", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    # Play sound
    (New-Object Media.SoundPlayer "C:\Windows\Media\tada.wav").Play()
    
    if ($exportErrors -eq [System.Windows.Forms.DialogResult]::Yes) {
        $errors | Export-Csv -Path "C:\Scripts\Errors\File.csv" -NoTypeInformation
        Write-Host "Errors exported to CSV file."
    }
    else {
        Write-Host "Errors not exported."
    }
    
    # Stop wondow from closing
    Read-Host -Prompt "Press Enter to exit"
} 
elseif ($selectedPolicyType -eq "Teams App Setup") {
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Define the form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Teams App Setup Policy Selector"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    
    # Define the list box
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 10)
    $listBox.Size = New-Object System.Drawing.Size(360, 180)
    
    # Populate the list box with policies
    $policies = Get-CsTeamsAppSetupPolicy
    $policies | ForEach-Object { [void]$listBox.Items.Add($_.Identity.Replace('Tag:', '')) }
    
    # Add the list box to the form
    $form.Controls.Add($listBox)
    
    # Define the OK button
    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Location = New-Object System.Drawing.Point(210, 200)
    $buttonOK.Size = New-Object System.Drawing.Size(75, 23)
    $buttonOK.Text = "OK"
    
    # Add the OK button event handler
    $buttonOK.Add_Click({
            $global:selectedPolicy = $listBox.SelectedItem
            $form.Close()
            if ($selectedPolicy) {
                [System.Windows.Forms.MessageBox]::Show("You selected '$selectedPolicy'", "Selection", "OK", "Information")
                # Play a sound after popup appears
            (New-Object Media.SoundPlayer "C:\Windows\Media\tada.wav").Play()
                Write-Host "Selected Policy: $selectedPolicy"

                Write-Host "Attention, all Jedi Knights and Rebel scum! 
    This script might take longer than completing the Kessel Run in less than 12 parsecs. 
    That's right, you might as well grab a drink at the Mos Eisley cantina and settle in for the long haul. 
    But don't worry, we'll keep you entertained along the way with a variety of fun sound effects, including the occasional blast from a TIE fighter.
    
    And remember what Yoda once said: 
    'PATIENCE YOU MUST HAVE, my young Padawan. For the script to complete, it takes time, as does the path to the Force.' 
    So sit back, relax, and let the script do its thing. May the Force be with you!"
    
    
            }
        })
    
    # Add the OK button to the form
    $form.Controls.Add($buttonOK)
    
    # Define the Cancel button
    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Location = New-Object System.Drawing.Point(295, 200)
    $buttonCancel.Size = New-Object System.Drawing.Size(75, 23)
    $buttonCancel.Text = "Cancel"
    
    # Add the Cancel button event handler
    $buttonCancel.Add_Click({
            $form.Close()
            return
        })
    
    # Add the Cancel button to the form
    $form.Controls.Add($buttonCancel)
    
    # Show the form
    [void]$form.ShowDialog()
    
    # Create an array to hold error messages
    $errors = @()
    
    # Get all users in the organization
    $users = Get-CsOnlineUser
    
    # Loop through each user and check their assigned policies
    foreach ($user in $users) {
        try {
            $policies = Get-CsUserPolicyAssignment -Identity $user.UserPrincipalName -ErrorAction Stop
    
            # Check if the selected policy is assigned to the user
            $hasSelectedPolicy = $policies.PolicyName -eq $global:selectedPolicy
    
            # If the user has the selected policy assigned, list their name and UPN
            if ($hasSelectedPolicy) {
                Write-Host "User:" $user.DisplayName "UPN:" $user.UserPrincipalName "has $($global:selectedPolicy) policy assigned."
                # Play a different sound after each found user
                (New-Object Media.SoundPlayer "C:\Windows\Media\chimes.wav").Play()
            }
        }
        catch {
            
            # Add error message to the errors array
            $errorMessage = New-Object PSObject -Property @{
                UserPrincipalName = $user.UserPrincipalName
                ErrorMessage      = $_.Exception.Message
            }
            $errors += $errorMessage
        }
    }
    
    # Show message box to ask if errors should be exported
    $exportErrors = [System.Windows.Forms.MessageBox]::Show("Do you want to export errors to a CSV file?", "Export Errors", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    # Play sound
    (New-Object Media.SoundPlayer "C:\Windows\Media\tada.wav").Play()
    
    if ($exportErrors -eq [System.Windows.Forms.DialogResult]::Yes) {
        $errors | Export-Csv -Path "C:\Scripts\Errors\File.csv" -NoTypeInformation
        Write-Host "Errors exported to CSV file."
    }
    else {
        Write-Host "Errors not exported."
    }
    
    # Stop wondow from closing
    Read-Host -Prompt "Type Enter to exit"
} 
else {
    Write-Host "Invalid policy type. Please enter either teamsappermission or teamsapsetup."
}
