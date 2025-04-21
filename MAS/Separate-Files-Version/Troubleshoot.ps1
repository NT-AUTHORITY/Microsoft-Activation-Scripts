Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "MAS Troubleshoot"
$mainForm.Size = New-Object System.Drawing.Size(800, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.BackColor = [System.Drawing.Color]::White

# Create output panel
$outputPanel = New-Object System.Windows.Forms.Panel
$outputPanel.Location = New-Object System.Drawing.Point(10, 10)
$outputPanel.Size = New-Object System.Drawing.Size(765, 400)
$outputPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

# Create rich text box for output
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(5, 5)
$outputBox.Size = New-Object System.Drawing.Size(753, 388)
$outputBox.ReadOnly = $true
$outputBox.BackColor = [System.Drawing.Color]::Black
$outputBox.ForeColor = [System.Drawing.Color]::White
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputPanel.Controls.Add($outputBox)
$mainForm.Controls.Add($outputPanel)

# Create buttons panel
$buttonsPanel = New-Object System.Windows.Forms.Panel
$buttonsPanel.Location = New-Object System.Drawing.Point(10, 420)
$buttonsPanel.Size = New-Object System.Drawing.Size(765, 130)
$buttonsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$mainForm.Controls.Add($buttonsPanel)

# Helper function to create buttons
function Create-Button {
    param (
        [string]$text,
        [int]$x,
        [int]$y,
        [int]$width = 230,
        [int]$height = 35
    )
    
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $button.Size = New-Object System.Drawing.Size($width, $height)
    $button.Text = $text
    $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $button.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonsPanel.Controls.Add($button)
    return $button
}

# Create buttons
$helpButton = Create-Button "Help" 10 10
$dismButton = Create-Button "DISM RestoreHealth" 260 10
$sfcButton = Create-Button "SFC Scannow" 510 10
$wmiButton = Create-Button "Fix WMI" 10 55
$licenseButton = Create-Button "Fix Licensing" 260 55
$wpaButton = Create-Button "Fix WPA Registry" 510 55

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::White
    )
    
    $outputBox.SelectionStart = $outputBox.TextLength
    $outputBox.SelectionLength = 0
    $outputBox.SelectionColor = $Color
    $outputBox.AppendText($Text)
    $outputBox.SelectionColor = $outputBox.ForeColor
    $outputBox.ScrollToCaret()
}

# Function to run command and capture output
function Run-Command {
    param(
        [string]$Command,
        [string]$Arguments
    )
    
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo.FileName = $Command
    $process.StartInfo.Arguments = $Arguments
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.RedirectStandardOutput = $true
    $process.StartInfo.RedirectStandardError = $true
    $process.StartInfo.CreateNoWindow = $true
    
    $outputBuilder = New-Object System.Text.StringBuilder
    $errorBuilder = New-Object System.Text.StringBuilder
    
    $outputHandler = {
        if (! [String]::IsNullOrEmpty($EventArgs.Data)) {
            $outputBuilder.AppendLine($EventArgs.Data)
        }
    }
    
    $errorHandler = {
        if (! [String]::IsNullOrEmpty($EventArgs.Data)) {
            $errorBuilder.AppendLine($EventArgs.Data)
        }
    }
    
    $process.OutputDataReceived += $outputHandler
    $process.ErrorDataReceived += $errorHandler
    
    $process.Start() | Out-Null
    $process.BeginOutputReadLine()
    $process.BeginErrorReadLine()
    $process.WaitForExit()
    
    return @{
        Output = $outputBuilder.ToString()
        Error = $errorBuilder.ToString()
        ExitCode = $process.ExitCode
    }
}

# Help button click event
$helpButton.Add_Click({
    Write-ColorOutput "Opening help page...`r`n" -Color Yellow
    Start-Process "https://massgrave.dev/troubleshoot.html"
})

# DISM button click event
$dismButton.Add_Click({
    Write-ColorOutput "`r`nRunning DISM RestoreHealth...`r`n" -Color Yellow
    $result = Run-Command "DISM.exe" "/Online /Cleanup-Image /RestoreHealth"
    Write-ColorOutput $result.Output -Color White
    if ($result.Error) {
        Write-ColorOutput $result.Error -Color Red
    }
    Write-ColorOutput "`r`nDISM RestoreHealth completed.`r`n" -Color Green
})

# SFC button click event
$sfcButton.Add_Click({
    Write-ColorOutput "`r`nRunning SFC Scannow...`r`n" -Color Yellow
    $result = Run-Command "sfc.exe" "/scannow"
    Write-ColorOutput $result.Output -Color White
    if ($result.Error) {
        Write-ColorOutput $result.Error -Color Red
    }
    Write-ColorOutput "`r`nSFC Scannow completed.`r`n" -Color Green
})

# Fix WMI button click event
$wmiButton.Add_Click({
    Write-ColorOutput "`r`nFixing WMI...`r`n" -Color Yellow
    
    # Stop WMI Service
    Stop-Service Winmgmt -Force
    
    # Rebuild WMI Repository
    $result = Run-Command "winmgmt.exe" "/salvagerepository"
    if ($result.ExitCode -eq 0) {
        $result = Run-Command "winmgmt.exe" "/resetrepository"
    }
    
    # Start WMI Service
    Start-Service Winmgmt
    
    Write-ColorOutput "`r`nWMI Fix completed.`r`n" -Color Green
})

# Fix Licensing button click event
$licenseButton.Add_Click({
    Write-ColorOutput "`r`nFixing Licensing...`r`n" -Color Yellow
    
    # Stop services
    Stop-Service -Name "ClipSVC" -Force
    Stop-Service -Name "sppsvc" -Force
    
    # Delete tokens
    Remove-Item -Path "$env:SystemRoot\System32\spp\store\2.0\tokens.dat" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:SystemRoot\System32\spp\store\2.0\cache\cache.dat" -Force -ErrorAction SilentlyContinue
    
    # Start services
    Start-Service -Name "ClipSVC"
    Start-Service -Name "sppsvc"
    
    Write-ColorOutput "`r`nLicensing Fix completed.`r`n" -Color Green
})

# Fix WPA Registry button click event
$wpaButton.Add_Click({
    Write-ColorOutput "`r`nFixing WPA Registry...`r`n" -Color Yellow
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform",
        "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
    )
    
    foreach ($path in $registryPaths) {
        if (Test-Path $path) {
            Remove-ItemProperty -Path $path -Name "BackupProductKeyDefault" -ErrorAction SilentlyContinue
            Remove-ItemProperty -Path $path -Name "ProductKeyDefault" -ErrorAction SilentlyContinue
        }
    }
    
    Write-ColorOutput "`r`nWPA Registry Fix completed.`r`n" -Color Green
})

# Initial system check
function Check-SystemStatus {
    Write-ColorOutput "Checking system status...`r`n" -Color Yellow
    
    # Check Windows version
    $osInfo = Get-WmiObject Win32_OperatingSystem
    Write-ColorOutput "Windows Version: " -Color White
    Write-ColorOutput "$($osInfo.Caption) $($osInfo.Version)`r`n" -Color Green
    
    # Check PowerShell execution policy
    $policy = Get-ExecutionPolicy
    Write-ColorOutput "PowerShell Execution Policy: " -Color White
    Write-ColorOutput "$policy`r`n" -Color Green
    
    # Check important services
    $services = @("ClipSVC", "sppsvc", "Winmgmt")
    foreach ($service in $services) {
        $status = Get-Service -Name $service -ErrorAction SilentlyContinue
        Write-ColorOutput "Service $service Status: " -Color White
        if ($status) {
            Write-ColorOutput "$($status.Status)`r`n" -Color $(if ($status.Status -eq "Running") { "Green" } else { "Red" })
        } else {
            Write-ColorOutput "Not Found`r`n" -Color Red
        }
    }
}

# Run initial check when form loads
$mainForm.Add_Shown({
    Check-SystemStatus
})

# Show the form
[void]$mainForm.ShowDialog()