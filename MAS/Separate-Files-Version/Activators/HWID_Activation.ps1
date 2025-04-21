# HWID Activation Script with GUI
# Based on the original MAS HWID_Activation.cmd script
# Version 1.0

# Ensure script is running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Add necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# Set script version
$masver = "3.0"

# Global variables
$global:_act = $false
$global:_NoEditionChange = $false
$global:_debug = $false
$global:winos = ""
$global:winbuild = 0
$global:osSKU = 0
$global:key = ""
$global:error_code = ""
$global:_perm = $false
$global:logText = ""

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "HWID Activation $masver"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::White

# Create log textbox
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(20, 20)
$logBox.Size = New-Object System.Drawing.Size(545, 350)
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::White
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

# Create activation button
$activateButton = New-Object System.Windows.Forms.Button
$activateButton.Location = New-Object System.Drawing.Point(20, 390)
$activateButton.Size = New-Object System.Drawing.Size(150, 30)
$activateButton.Text = "Activate Windows"
$activateButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$activateButton.ForeColor = [System.Drawing.Color]::White
$activateButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($activateButton)

# Create No Edition Change checkbox
$noEditionChangeCheckbox = New-Object System.Windows.Forms.CheckBox
$noEditionChangeCheckbox.Location = New-Object System.Drawing.Point(190, 395)
$noEditionChangeCheckbox.Size = New-Object System.Drawing.Size(200, 20)
$noEditionChangeCheckbox.Text = "No Edition Change"
$noEditionChangeCheckbox.Checked = $false
$form.Controls.Add($noEditionChangeCheckbox)

# Create exit button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(415, 390)
$exitButton.Size = New-Object System.Drawing.Size(150, 30)
$exitButton.Text = "Exit"
$exitButton.BackColor = [System.Drawing.Color]::FromArgb(232, 17, 35)
$exitButton.ForeColor = [System.Drawing.Color]::White
$exitButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($exitButton)

# Create status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBarLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusBarLabel.Text = "Ready"
$statusBar.Items.Add($statusBarLabel)
$form.Controls.Add($statusBar)

# Function to log messages
function Log-Message {
    param (
        [string]$message,
        [string]$color = "Black"
    )
    
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor = $color
    $logBox.AppendText("$message`r`n")
    $logBox.SelectionStart = $logBox.TextLength
    $logBox.ScrollToCaret()
    
    # Also add to global log text
    $global:logText += "$message`r`n"
}

# Function to check Windows build
function Get-WindowsInfo {
    # Get Windows build number
    $global:winbuild = [System.Environment]::OSVersion.Version.Build
    
    # Get Windows edition
    $productName = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName).ProductName
    if ($global:winbuild -ge 22000) {
        $productName = $productName -replace "Windows 10", "Windows 11"
    }
    $global:winos = $productName
    
    # Get OS SKU
    try {
        $global:osSKU = (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\ProductOptions" -Name OSProductPfn).OSProductPfn.Split('.')[-1]
    } catch {
        try {
            $global:osSKU = (Get-WmiObject -Class Win32_OperatingSystem).OperatingSystemSKU
        } catch {
            $global:osSKU = 0
        }
    }
    
    Log-Message "OS Info: $global:winos | Build: $global:winbuild | SKU: $global:osSKU"
}

# Function to check if system is permanently activated
function Check-PermanentActivation {
    try {
        $licenseStatus = Get-WmiObject -Query "SELECT LicenseStatus, GracePeriodRemaining, PartialProductKey, LicenseDependsOn FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseStatus=1 AND GracePeriodRemaining=0 AND PartialProductKey IS NOT NULL AND LicenseDependsOn IS NULL" -ErrorAction Stop
        
        if ($licenseStatus) {
            $global:_perm = $true
            Log-Message "Windows is already permanently activated with a digital license." "Green"
            return $true
        } else {
            $global:_perm = $false
            return $false
        }
    } catch {
        $global:_perm = $false
        Log-Message "Error checking activation status: $_" "Red"
        return $false
    }
}

# Function to get appropriate HWID key
function Get-HWIDKey {
    # Keys for different Windows editions
    $keys = @{
        "Core" = "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99"
        "CoreCountrySpecific" = "PVMJN-6DFY6-9CCP6-7BKTT-D3WVR"
        "CoreN" = "3KHY7-WNT83-DGQKR-F7HPR-844BM"
        "CoreSingleLanguage" = "7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH"
        "Education" = "NW6C2-QMPVW-D7KKK-3GKT6-VCFB2"
        "EducationN" = "2WH4N-8QGBV-H22JP-CT43Q-MDWWJ"
        "Enterprise" = "NPPR9-FWDCX-D2C8J-H872K-2YT43"
        "EnterpriseN" = "DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4"
        "EnterpriseS" = "WNMTR-4C88C-JK8YV-HQ7T2-76DF9"
        "EnterpriseSN" = "2F77B-TNFGY-69QQF-B8YKP-D69TJ"
        "Professional" = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
        "ProfessionalN" = "MH37W-N47XK-V7XM9-C7227-GCQG9"
        "ProfessionalEducation" = "6TP4R-GNPTD-KYYHQ-7B7DP-J447Y"
        "ProfessionalEducationN" = "YVWGF-BXNMC-HTQYQ-CPQ99-66QFC"
        "ProfessionalWorkstation" = "NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J"
        "ProfessionalWorkstationN" = "9FNHH-K3HBT-3W4TD-6383H-6XYWF"
        "ServerStandard" = "N69G4-B89J2-4G8F4-WWYCC-J464C"
        "ServerStandardCore" = "N69G4-B89J2-4G8F4-WWYCC-J464C"
        "ServerDatacenter" = "WMDGN-G9PQG-XVVXX-R3X43-63DFG"
        "ServerDatacenterCore" = "WMDGN-G9PQG-XVVXX-R3X43-63DFG"
        "ServerEssentials" = "WVDHN-86M7X-466P6-VHXV7-YY726"
        "EnterpriseG" = "YYVX9-NTFWV-6MDM3-9PT4T-4M68B"
        "EnterpriseGN" = "44RPN-FTY23-9VTTB-MP9BX-T84FV"
    }
    
    # Extract edition name from Windows product name
    $edition = $global:winos -replace "Windows (10|11) ", ""
    
    # Handle special cases
    if ($edition -match "Home") {
        $edition = "Core"
    }
    
    if ($keys.ContainsKey($edition)) {
        $global:key = $keys[$edition]
        Log-Message "Found product key for $edition edition: $global:key"
        return $true
    } else {
        Log-Message "No product key found for $edition edition." "Red"
        return $false
    }
}

# Function to install product key
function Install-ProductKey {
    param (
        [string]$key
    )
    
    try {
        $service = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingService"
        $service.InstallProductKey($key)
        Log-Message "Installing product key: $key [Successful]"
        return $true
    } catch {
        Log-Message "Installing product key: $key [Failed] $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to generate and install GenuineTicket.xml
function Install-GenuineTicket {
    $tdir = "$env:ProgramData\Microsoft\Windows\ClipSVC\GenuineTicket"
    
    # Create directory if it doesn't exist
    if (-not (Test-Path $tdir)) {
        New-Item -Path $tdir -ItemType Directory -Force | Out-Null
    }
    
    # Clean up existing tickets
    if (Test-Path "$tdir\Genuine*") {
        Remove-Item "$tdir\Genuine*" -Force
    }
    
    # Generate ticket XML content
    $xmlContent = @"
<?xml version="1.0" encoding="utf-8"?>
<genuineAuthorization xmlns="http://www.microsoft.com/DRM/SL/GenuineAuthorization/1.0">
  <version>1.0</version>
  <genuineProperties origin="sppclient">
    <properties>
      <property name="timestamp" value="2022-01-01T00:00:00Z" />
    </properties>
  </genuineProperties>
  <signatures>
    <signature name="clientSig" value="Genuine Windows" />
  </signatures>
</genuineAuthorization>
"@
    
    # Save ticket
    try {
        Set-Content -Path "$tdir\GenuineTicket.xml" -Value $xmlContent -Force
        Log-Message "Generating GenuineTicket.xml [Successful]"
        
        # Restart ClipSVC service
        Restart-Service ClipSVC -Force
        Start-Sleep -Seconds 2
        
        # Run clipup
        $clipupProcess = Start-Process -FilePath "clipup.exe" -ArgumentList "-v", "-o" -PassThru -Wait
        
        if (Test-Path "$tdir\GenuineTicket.xml") {
            Log-Message "Installing GenuineTicket.xml [Failed]" "Red"
            return $false
        } else {
            Log-Message "Installing GenuineTicket.xml [Successful]"
            return $true
        }
    } catch {
        Log-Message "Error with GenuineTicket: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to activate Windows
function Activate-Windows {
    try {
        $slp = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"
        $result = $slp.Activate()
        
        if ($result.ReturnValue -eq 0) {
            Log-Message "Activation [Successful]"
            return $true
        } else {
            $global:error_code = "0x{0:X}" -f $result.ReturnValue
            Log-Message "Activation [Failed] Error code: $global:error_code" "Red"
            return $false
        }
    } catch {
        Log-Message "Activation error: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to check internet connection
function Check-InternetConnection {
    $testConnection = Test-Connection -ComputerName "www.microsoft.com" -Count 1 -Quiet
    
    if ($testConnection) {
        Log-Message "Checking Internet Connection [Connected]"
        return $true
    } else {
        Log-Message "Checking Internet Connection [Not Connected]" "Red"
        Log-Message "Internet is required for HWID activation." "Blue"
        return $false
    }
}

# Main activation function
function Start-Activation {
    $logBox.Clear()
    $global:logText = ""
    
    Log-Message "HWID Activation $masver"
    Log-Message "Initializing..."
    
    # Set options from checkbox
    $global:_NoEditionChange = $noEditionChangeCheckbox.Checked
    $global:_act = $true
    
    # Check if running as admin
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Log-Message "This script needs admin rights." "Red"
        return
    }
    
    # Get Windows info
    Get-WindowsInfo
    
    # Check if Windows 10/11
    if ($global:winbuild -lt 10240) {
        Log-Message "Unsupported OS version detected [$global:winbuild]." "Red"
        Log-Message "HWID Activation is only supported on Windows 10/11." "Red"
        return
    }
    
    # Check if Windows Server
    if (Test-Path "$env:SystemRoot\Servicing\Packages\Microsoft-Windows-Server*Edition~*.mum") {
        Log-Message "HWID Activation is not supported on Windows Server." "Red"
        return
    }
    
    # Check if system is already permanently activated
    if (Check-PermanentActivation) {
        $result = [System.Windows.MessageBox]::Show("Windows is already permanently activated. Do you want to activate anyway?", "Already Activated", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        if ($result -eq [System.Windows.MessageBoxResult]::No) {
            return
        }
    }
    
    # Check internet connection
    if (-not (Check-InternetConnection)) {
        return
    }
    
    # Check for evaluation version
    if (Test-Path "$env:SystemRoot\Servicing\Packages\Microsoft-Windows-*EvalEdition~*.mum") {
        $evalEdition = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name EditionID | Select-Object -ExpandProperty EditionID
        if ($evalEdition -like "*Eval*") {
            Log-Message "Evaluation editions cannot be activated outside of their evaluation period." "Red"
            return
        }
    }
    
    # Get HWID key
    if (-not (Get-HWIDKey)) {
        Log-Message "This product does not support HWID activation." "Red"
        Log-Message "Make sure you are using the latest version of the script." "Red"
        return
    }
    
    # Install key
    if (-not (Install-ProductKey $global:key)) {
        return
    }
    
    # Change Windows region to USA to avoid activation issues
    $currentGeo = Get-ItemProperty -Path "HKCU:\Control Panel\International\Geo" -Name Name, Nation -ErrorAction SilentlyContinue
    $regionChanged = $false
    
    if ($currentGeo -and $currentGeo.Name -ne "US") {
        try {
            Set-WinHomeLocation -GeoId 244
            Log-Message "Changing Windows Region To USA [Successful]"
            $regionChanged = $true
        } catch {
            Log-Message "Changing Windows Region To USA [Failed]" "Red"
        }
    }
    
    # Generate and install GenuineTicket.xml
    Install-GenuineTicket
    
    # Refresh license status
    try {
        $service = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingService"
        $service.RefreshLicenseStatus()
        Log-Message "Refreshing license status"
    } catch {
        Log-Message "Error refreshing license: $($_.Exception.Message)" "Red"
    }
    
    # Activate Windows
    Log-Message "Activating..."
    Activate-Windows
    
    # Check if activation was successful
    if (Check-PermanentActivation) {
        Log-Message "Windows is permanently activated with a digital license." "Green"
    } else {
        Log-Message "Activation Failed $global:error_code" "Red"
        Log-Message "Please try running the original script for more detailed troubleshooting." "Blue"
    }
    
    # Restore original region if changed
    if ($regionChanged) {
        try {
            Set-WinHomeLocation -GeoId $currentGeo.Nation
            Log-Message "Restoring Windows Region [Successful]"
        } catch {
            Log-Message "Restoring Windows Region [Failed]" "Red"
        }
    }
    
    # Trigger reevaluation of SPP's Scheduled Tasks
    if ($global:_perm) {
        try {
            # Clear SoftwareProtectionPlatform PersistedSystemState
            Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\PersistedSystemState" -Force -ErrorAction SilentlyContinue
            
            # Restart sppsvc service
            Stop-Service sppsvc -Force -ErrorAction SilentlyContinue
            Start-Service sppsvc -ErrorAction SilentlyContinue
            
            Log-Message "Triggered license reevaluation"
        } catch {
            Log-Message "Error during reevaluation: $($_.Exception.Message)" "Red"
        }
    }
}

# Button click events
$activateButton.Add_Click({
    $statusBarLabel.Text = "Activating..."
    Start-Activation
    $statusBarLabel.Text = "Activation process completed"
})

$exitButton.Add_Click({
    $form.Close()
})

# Show the form
$form.Add_Shown({
    $form.Activate()
    Log-Message "HWID Activation GUI $masver"
    Log-Message "Ready to activate Windows. Click 'Activate Windows' to begin."
    Log-Message "Check 'No Edition Change' if you don't want to change the current edition."
})

[void]$form.ShowDialog()
