# Change Windows Edition GUI with CBS Support
# Based on MAS version 3.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:masver = "3.0"
$script:mas = "https://massgrave.dev/"
$script:winbuild = [System.Environment]::OSVersion.Version.Build
$script:stageCurrentCBS = $false

# Create main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Change Windows Edition $masver"
$mainForm.Size = New-Object System.Drawing.Size(800,600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.FormBorderStyle = "FixedDialog"
$mainForm.MaximizeBox = $false

# Create main panel
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Location = New-Object System.Drawing.Point(20,20)
$mainPanel.Size = New-Object System.Drawing.Size(740,520)

# Create current edition label
$currentEditionLabel = New-Object System.Windows.Forms.Label
$currentEditionLabel.Location = New-Object System.Drawing.Point(0,0)
$currentEditionLabel.Size = New-Object System.Drawing.Size(740,40)
$currentEditionLabel.Text = "Current Windows Edition: Detecting..."

# Create method selection group
$methodGroup = New-Object System.Windows.Forms.GroupBox
$methodGroup.Location = New-Object System.Drawing.Point(0,50)
$methodGroup.Size = New-Object System.Drawing.Size(740,70)
$methodGroup.Text = "Upgrade Method"

$radioDISM = New-Object System.Windows.Forms.RadioButton
$radioDISM.Location = New-Object System.Drawing.Point(10,20)
$radioDISM.Size = New-Object System.Drawing.Size(200,30)
$radioDISM.Text = "DISM API Method"
$radioDISM.Checked = $true

$radioCBS = New-Object System.Windows.Forms.RadioButton
$radioCBS.Location = New-Object System.Drawing.Point(220,20)
$radioCBS.Size = New-Object System.Drawing.Size(200,30)
$radioCBS.Text = "CBS Upgrade Method"

# Create CBS options
$cbsOptions = New-Object System.Windows.Forms.GroupBox
$cbsOptions.Location = New-Object System.Drawing.Point(430,10)
$cbsOptions.Size = New-Object System.Drawing.Size(300,50)
$cbsOptions.Text = "CBS Options"
$cbsOptions.Visible = $false

$stageCheckbox = New-Object System.Windows.Forms.CheckBox
$stageCheckbox.Location = New-Object System.Drawing.Point(10,20)
$stageCheckbox.Size = New-Object System.Drawing.Size(280,20)
$stageCheckbox.Text = "Stage current edition while changing"
$stageCheckbox.Checked = $false

# Create target editions list
$targetEditionsList = New-Object System.Windows.Forms.ListBox
$targetEditionsList.Location = New-Object System.Drawing.Point(0,130)
$targetEditionsList.Size = New-Object System.Drawing.Size(740,300)

# Create change button
$changeButton = New-Object System.Windows.Forms.Button
$changeButton.Location = New-Object System.Drawing.Point(0,440)
$changeButton.Size = New-Object System.Drawing.Size(740,40)
$changeButton.Text = "Change Edition"
$changeButton.Enabled = $false

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(0,490)
$statusLabel.Size = New-Object System.Drawing.Size(740,30)
$statusLabel.Text = ""

# CBS related functions from original script
function Get-CBSTargetEditions {
    $installCandidates = @{}
    $removalCandidates = @()

    $packageStates = @{
        0x0 = 'NotPresent'
        0x1 = 'UninstallPending'
        0x2 = 'Staged'
        0x50 = 'Installed'
        0x70 = 'InstallPending'
    }

    $packagePattern = 'Microsoft-Windows-*Edition~*.mum'
    $editionPackagePattern = 'Microsoft-Windows-*Edition~31bf3856ad364e35~*~~*.mum'

    Get-ChildItem -Path "$env:SystemRoot\servicing\Packages" -Filter $packagePattern | ForEach-Object {
        $packageName = $_.Name
        if($packageName -notmatch $editionPackagePattern) {
            return
        }

        $state = 0
        $package = Get-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages\$packageName"
        $currentState = $package.GetValue('CurrentState')
        if($null -ne $currentState) {
            $state = $currentState
        }

        $packageEdition = ($packageName -split '~')[0] -replace 'Microsoft-Windows-|Edition$',''
        if($state -eq 0x50) {
            $installCandidates[$packageEdition] = @()
        }
    }

    return $installCandidates.Keys
}

function Find-EditionXml {
    param (
        [String]$Edition
    )

    $servicingEditions = "$env:SystemRoot\servicing\Editions"
    $editionXml = "${Edition}Edition.xml"
    $editionXmlInServicing = Join-Path $servicingEditions $editionXml

    if(Test-Path -Path $editionXmlInServicing -PathType Leaf) {
        return $editionXmlInServicing
    }

    return $null
}

function Invoke-CBSUpgrade {
    param (
        [string]$TargetEdition
    )

    $editionXml = Find-EditionXml -Edition $TargetEdition
    if(-not $editionXml) {
        throw "Edition XML not found for $TargetEdition"
    }

    # Prepare CBS session
    $session = New-Object -ComObject Microsoft.Windows.Servicing.Session
    $pkg = $session.GetPackage("Microsoft-Windows-${TargetEdition}Edition~31bf3856ad364e35~amd64~~10.0.0.0")

    if($stageCheckbox.Checked) {
        # Stage current edition
        $pkg.Stage()
    }

    # Install target edition
    $pkg.Install()

    return $true
}

# Helper function to get current Windows edition
function Get-CurrentEdition {
    try {
        $currentEditionLabel.Text = "Loading Version..."
        $mainForm.Refresh()
        
        $edition = (Get-WindowsEdition -Online).Edition
        $currentEditionLabel.Text = "Current Windows Edition: $edition (Build: $winbuild)"
        return $edition
    }
    catch {
        $currentEditionLabel.Text = "Error detecting Windows edition"
        return $null
    }
}

# Helper function to get target editions
function Get-TargetEditions {
    $targetEditionsList.Items.Clear()
    $targetEditionsList.Items.Add("Loading available editions...")
    $mainForm.Refresh()
    
    if ($radioCBS.Checked) {
        $editions = Get-CBSTargetEditions
    }
    else {
        try {
            $editions = Get-WindowsEdition -Online -Target | Select-Object -ExpandProperty Edition
        }
        catch {
            $editions = @()
        }
    }

    $targetEditionsList.Items.Clear()
    if ($editions.Count -eq 0) {
        $targetEditionsList.Items.Add("No editions found!")
    } else {
        foreach($edition in $editions) {
            $targetEditionsList.Items.Add($edition)
        }
    }
    
    $changeButton.Enabled = $targetEditionsList.Items.Count -gt 0 -and $editions.Count -gt 0
}

# DISM API upgrade function
function Invoke-DISMUpgrade {
    param (
        [string]$TargetEdition
    )

    try {
        $result = Set-WindowsEdition -Online -Target $TargetEdition -NoRestart
        return $result
    }
    catch {
        throw
    }
}

# Change edition function
function Change-WindowsEdition {
    $selectedEdition = $targetEditionsList.SelectedItem
    if (-not $selectedEdition) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a target edition first.",
            "No Edition Selected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to change to $selectedEdition?`nThe system will restart after the change.",
        "Confirm Edition Change",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            $statusLabel.Text = "Initiating edition change... Please wait."
            $mainForm.Refresh()

            if ($radioCBS.Checked) {
                $changeResult = Invoke-CBSUpgrade -TargetEdition $selectedEdition
            }
            else {
                $changeResult = Invoke-DISMUpgrade -TargetEdition $selectedEdition
            }

            $restart = [System.Windows.Forms.MessageBox]::Show(
                "Edition change initiated successfully. System needs to restart.`nRestart now?",
                "Restart Required",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($restart -eq [System.Windows.Forms.DialogResult]::Yes) {
                Restart-Computer -Force
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to change edition: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        finally {
            $statusLabel.Text = ""
        }
    }
}

# Event handlers
$changeButton.Add_Click({ Change-WindowsEdition })

$targetEditionsList.Add_SelectedIndexChanged({
    $changeButton.Enabled = $targetEditionsList.SelectedItem -ne $null
})

$radioCBS.Add_CheckedChanged({
    $cbsOptions.Visible = $radioCBS.Checked
    Get-TargetEditions
})

$radioDISM.Add_CheckedChanged({
    $cbsOptions.Visible = $false
    Get-TargetEditions
})

# Add controls to form
$methodGroup.Controls.AddRange(@($radioDISM, $radioCBS, $cbsOptions))
$cbsOptions.Controls.Add($stageCheckbox)
$mainPanel.Controls.AddRange(@(
    $currentEditionLabel,
    $methodGroup,
    $targetEditionsList,
    $changeButton,
    $statusLabel
))
$mainForm.Controls.Add($mainPanel)

# Initialize
Get-CurrentEdition
Get-TargetEditions

# Show form
$mainForm.ShowDialog()

