# Change Office Edition GUI
# Based on MAS version 3.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:masver = "3.0"
$script:mas = "https://massgrave.dev/"

# Create main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Change Office Edition $masver"
$mainForm.Size = New-Object System.Drawing.Size(600,400)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false

# Create main menu
$mainMenu = New-Object System.Windows.Forms.GroupBox
$mainMenu.Location = New-Object System.Drawing.Point(20,20)
$mainMenu.Size = New-Object System.Drawing.Size(540,320)
$mainMenu.Text = "Main Menu"

# Create buttons
$btnChangeAll = New-Object System.Windows.Forms.Button
$btnChangeAll.Location = New-Object System.Drawing.Point(30,40)
$btnChangeAll.Size = New-Object System.Drawing.Size(480,40)
$btnChangeAll.Text = "Change all editions"
$btnChangeAll.Add_Click({
    $script:change = 1
    Show-EditionSelection
})

$btnAddEdition = New-Object System.Windows.Forms.Button
$btnAddEdition.Location = New-Object System.Drawing.Point(30,90)
$btnAddEdition.Size = New-Object System.Drawing.Size(480,40)
$btnAddEdition.Text = "Add edition"
$btnAddEdition.Add_Click({
    $script:change = 0
    Show-EditionSelection
})

$btnRemoveEdition = New-Object System.Windows.Forms.Button
$btnRemoveEdition.Location = New-Object System.Drawing.Point(30,140)
$btnRemoveEdition.Size = New-Object System.Drawing.Size(480,40)
$btnRemoveEdition.Text = "Remove edition"
$btnRemoveEdition.Add_Click({ Show-RemoveEdition })

$btnEditApps = New-Object System.Windows.Forms.Button
$btnEditApps.Location = New-Object System.Drawing.Point(30,190)
$btnEditApps.Size = New-Object System.Drawing.Size(480,40)
$btnEditApps.Text = "Add/Remove apps"
$btnEditApps.Add_Click({ Show-EditApps })

$btnUpdateChannel = New-Object System.Windows.Forms.Button
$btnUpdateChannel.Location = New-Object System.Drawing.Point(30,240)
$btnUpdateChannel.Size = New-Object System.Drawing.Size(480,40)
$btnUpdateChannel.Text = "Change Office Update Channel"
$btnUpdateChannel.Add_Click({ Show-UpdateChannel })

# Add controls to form
$mainMenu.Controls.AddRange(@($btnChangeAll, $btnAddEdition, $btnRemoveEdition, $btnEditApps, $btnUpdateChannel))
$mainForm.Controls.Add($mainMenu)

# Helper functions
function Show-EditionSelection {
    $editionForm = New-Object System.Windows.Forms.Form
    $editionForm.Text = "Select Office Edition"
    $editionForm.Size = New-Object System.Drawing.Size(500,400)
    $editionForm.StartPosition = "CenterScreen"

    $editionList = New-Object System.Windows.Forms.ListBox
    $editionList.Location = New-Object System.Drawing.Point(20,20)
    $editionList.Size = New-Object System.Drawing.Size(440,280)

    # Add edition types
    $editionList.Items.AddRange(@(
        "Suites_Retail",
        "Suites_Volume",
        "SingleApps_Retail",
        "SingleApps_Volume"
    ))

    $btnSelect = New-Object System.Windows.Forms.Button
    $btnSelect.Location = New-Object System.Drawing.Point(20,320)
    $btnSelect.Size = New-Object System.Drawing.Size(440,30)
    $btnSelect.Text = "Select"
    $btnSelect.Add_Click({
        if ($editionList.SelectedItem) {
            $script:selectedEdition = $editionList.SelectedItem
            $editionForm.Close()
            Process-EditionChange
        }
    })

    $editionForm.Controls.AddRange(@($editionList, $btnSelect))
    $editionForm.ShowDialog()
}

function Process-EditionChange {
    # Here we would implement the actual edition change logic
    # This would involve calling the appropriate C2R commands
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Would you like to proceed with the Office edition change?`nSelected: $script:selectedEdition",
        "Confirm Change",
        [System.Windows.Forms.MessageBoxButtons]::YesNo
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Implementation of actual change logic would go here
        # This would involve running the C2R commands from the original script
        [System.Windows.Forms.MessageBox]::Show(
            "Edition change process initiated.`nPlease wait while the changes are applied.",
            "Processing"
        )
    }
}

# Helper functions for Office operations
function Get-OfficeInfo {
    # Get installed Office products and channels
    $script:_oIds = @()
    $script:_updch = ""
    $script:_version = ""
    
    # Run setup.exe /query
    $setupPath = "${env:CommonProgramFiles}\Microsoft Shared\ClickToRun\setup.exe"
    if (Test-Path $setupPath) {
        $queryResult = & $setupPath /query
        
        # Parse the output to get installed products
        $queryResult | ForEach-Object {
            if ($_ -match "Product\s+(\w+):\s+(\w+)") {
                $script:_oIds += $matches[2]
            }
            if ($_ -match "Channel:\s+(.+)") {
                $script:_updch = $matches[1]
            }
            if ($_ -match "Version:\s+(.+)") {
                $script:_version = $matches[1]
            }
        }
    }
}

function Show-RemoveEdition {
    Get-OfficeInfo
    
    if ($script:_oIds.Count -le 1) {
        [System.Windows.Forms.MessageBox]::Show(
            "Only one Office edition is installed. This option requires multiple installed editions.",
            "Cannot Remove Edition",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $removeForm = New-Object System.Windows.Forms.Form
    $removeForm.Text = "Remove Office Edition"
    $removeForm.Size = New-Object System.Drawing.Size(500,400)
    $removeForm.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,20)
    $label.Size = New-Object System.Drawing.Size(440,30)
    $label.Text = "Select an Office edition to remove:"

    $editionList = New-Object System.Windows.Forms.ListBox
    $editionList.Location = New-Object System.Drawing.Point(20,60)
    $editionList.Size = New-Object System.Drawing.Size(440,240)
    $editionList.Items.AddRange($script:_oIds)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Location = New-Object System.Drawing.Point(20,320)
    $btnRemove.Size = New-Object System.Drawing.Size(440,30)
    $btnRemove.Text = "Remove Selected Edition"
    $btnRemove.Add_Click({
        if ($editionList.SelectedItem) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to remove $($editionList.SelectedItem)?",
                "Confirm Removal",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                $setupPath = "${env:CommonProgramFiles}\Microsoft Shared\ClickToRun\setup.exe"
                $edition = $editionList.SelectedItem
                Start-Process -FilePath $setupPath -ArgumentList "/configure", "/productstoremove=$edition" -Wait -NoNewWindow
                [System.Windows.Forms.MessageBox]::Show("Edition removal process completed.", "Success")
                $removeForm.Close()
            }
        }
    })

    $removeForm.Controls.AddRange(@($label, $editionList, $btnRemove))
    $removeForm.ShowDialog()
}

function Show-EditApps {
    Get-OfficeInfo
    
    $editForm = New-Object System.Windows.Forms.Form
    $editForm.Text = "Add/Remove Office Apps"
    $editForm.Size = New-Object System.Drawing.Size(500,500)
    $editForm.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,20)
    $label.Size = New-Object System.Drawing.Size(440,30)
    $label.Text = "Select Office edition to modify:"

    $editionList = New-Object System.Windows.Forms.ComboBox
    $editionList.Location = New-Object System.Drawing.Point(20,50)
    $editionList.Size = New-Object System.Drawing.Size(440,30)
    $editionList.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $editionList.Items.AddRange($script:_oIds)
    if ($editionList.Items.Count -gt 0) {
        $editionList.SelectedIndex = 0
    }

    # Apps checkboxes
    $appsGroup = New-Object System.Windows.Forms.GroupBox
    $appsGroup.Location = New-Object System.Drawing.Point(20,100)
    $appsGroup.Size = New-Object System.Drawing.Size(440,300)
    $appsGroup.Text = "Select apps to include:"

    $apps = @{
        "Access" = New-Object System.Windows.Forms.CheckBox
        "Excel" = New-Object System.Windows.Forms.CheckBox
        "OneDrive" = New-Object System.Windows.Forms.CheckBox
        "OneNote" = New-Object System.Windows.Forms.CheckBox
        "Outlook" = New-Object System.Windows.Forms.CheckBox
        "PowerPoint" = New-Object System.Windows.Forms.CheckBox
        "Publisher" = New-Object System.Windows.Forms.CheckBox
        "SkypeForBusiness" = New-Object System.Windows.Forms.CheckBox
        "Teams" = New-Object System.Windows.Forms.CheckBox
        "Word" = New-Object System.Windows.Forms.CheckBox
    }

    $y = 30
    foreach ($app in $apps.Keys) {
        $apps[$app].Location = New-Object System.Drawing.Point(20,$y)
        $apps[$app].Size = New-Object System.Drawing.Size(400,20)
        $apps[$app].Text = $app
        $apps[$app].Checked = $true
        $appsGroup.Controls.Add($apps[$app])
        $y += 25
    }

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Location = New-Object System.Drawing.Point(20,420)
    $btnApply.Size = New-Object System.Drawing.Size(440,30)
    $btnApply.Text = "Apply Changes"
    $btnApply.Add_Click({
        if ($editionList.SelectedItem) {
            $excludedApps = @()
            foreach ($app in $apps.Keys) {
                if (-not $apps[$app].Checked) {
                    $excludedApps += $app.ToLower()
                }
            }
            
            $excludeList = if ($excludedApps.Count -gt 0) { $excludedApps -join "," } else { "" }
            $edition = $editionList.SelectedItem
            
            $setupPath = "${env:CommonProgramFiles}\Microsoft Shared\ClickToRun\setup.exe"
            $configArgs = "/configure /product=$edition"
            if ($excludeList) {
                $configArgs += " /exclude=$excludeList"
            }
            
            Start-Process -FilePath $setupPath -ArgumentList $configArgs -Wait -NoNewWindow
            [System.Windows.Forms.MessageBox]::Show("App configuration has been updated.", "Success")
            $editForm.Close()
        }
    })

    $editForm.Controls.AddRange(@($label, $editionList, $appsGroup, $btnApply))
    $editForm.ShowDialog()
}

function Show-UpdateChannel {
    Get-OfficeInfo
    
    $channelForm = New-Object System.Windows.Forms.Form
    $channelForm.Text = "Change Office Update Channel"
    $channelForm.Size = New-Object System.Drawing.Size(500,400)
    $channelForm.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,20)
    $label.Size = New-Object System.Drawing.Size(440,40)
    $label.Text = "Current channel: $script:_updch`nSelect new update channel:"

    $channels = @{
        "Current Channel" = "492350f6-3a01-4f97-b9c0-c7c6ddf67d60"
        "Monthly Enterprise Channel" = "55336b82-a18d-4dd6-b5f6-9e5095c314a6"
        "Semi-Annual Enterprise Channel" = "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
        "Semi-Annual Enterprise Channel (Preview)" = "b8f9b850-328d-4355-9145-c59439a0c4cf"
        "Beta Channel" = "5440fd1f-7ecb-4221-8110-145efaa6372f"
        "Current Channel (Preview)" = "64256afe-f5d9-4f86-8936-8840a6a4f5be"
    }

    $channelList = New-Object System.Windows.Forms.ListBox
    $channelList.Location = New-Object System.Drawing.Point(20,70)
    $channelList.Size = New-Object System.Drawing.Size(440,240)
    $channelList.Items.AddRange($channels.Keys)

    $btnInfo = New-Object System.Windows.Forms.Button
    $btnInfo.Location = New-Object System.Drawing.Point(20,320)
    $btnInfo.Size = New-Object System.Drawing.Size(210,30)
    $btnInfo.Text = "Channel Information"
    $btnInfo.Add_Click({
        Start-Process "https://learn.microsoft.com/microsoft-365-apps/updates/overview-update-channels"
    })

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Location = New-Object System.Drawing.Point(250,320)
    $btnApply.Size = New-Object System.Drawing.Size(210,30)
    $btnApply.Text = "Apply Channel"
    $btnApply.Add_Click({
        if ($channelList.SelectedItem) {
            $channelId = $channels[$channelList.SelectedItem]
            
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to change to $($channelList.SelectedItem)?",
                "Confirm Channel Change",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Update registry keys
                $regPaths = @(
                    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\PropertyBag"
                )
                
                foreach ($path in $regPaths) {
                    if (Test-Path $path) {
                        Set-ItemProperty -Path $path -Name "CDNBaseUrl" -Value "http://officecdn.microsoft.com/pr/$channelId"
                        Set-ItemProperty -Path $path -Name "UpdateChannel" -Value "http://officecdn.microsoft.com/pr/$channelId"
                    }
                }
                
                # Run Office update
                Start-Process "${env:CommonProgramFiles}\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" -ArgumentList "/update user updatepromptuser=True" -Wait -NoNewWindow
                
                [System.Windows.Forms.MessageBox]::Show(
                    "Update channel has been changed. Office will update to the new channel on next update.",
                    "Success"
                )
                $channelForm.Close()
            }
        }
    })

    $channelForm.Controls.AddRange(@($label, $channelList, $btnInfo, $btnApply))
    $channelForm.ShowDialog()
}

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [System.Windows.Forms.MessageBox]::Show(
        "This script needs to be run as Administrator.`nPlease restart with elevated privileges.",
        "Admin Rights Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    exit
}

# Show the main form
$mainForm.ShowDialog()
