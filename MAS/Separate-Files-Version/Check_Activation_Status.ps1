Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Check Activation Status"
$mainForm.Size = New-Object System.Drawing.Size(800, 600)
$mainForm.StartPosition = "CenterScreen"

# Create rich text box for output
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 10)
$outputBox.Size = New-Object System.Drawing.Size(765, 500)
$outputBox.ReadOnly = $true
$outputBox.BackColor = [System.Drawing.Color]::Black
$outputBox.ForeColor = [System.Drawing.Color]::White
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$mainForm.Controls.Add($outputBox)

# Create refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(10, 520)
$refreshButton.Size = New-Object System.Drawing.Size(100, 30)
$refreshButton.Text = "Refresh"
$mainForm.Controls.Add($refreshButton)

# Create save button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Location = New-Object System.Drawing.Point(120, 520)
$saveButton.Size = New-Object System.Drawing.Size(100, 30)
$saveButton.Text = "Save Log"
$mainForm.Controls.Add($saveButton)

# Function to append colored text
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
}

# Function to check Windows activation status
function Check-WindowsActivation {
    $outputBox.Clear()
    Write-ColorOutput "Checking Windows Activation Status...`r`n" -Color Yellow
    
    try {
        # Get Windows activation information
        $slp = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL"
        $sls = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingService"
        
        Write-ColorOutput "============================================================`r`n" -Color Cyan
        Write-ColorOutput "Windows Edition: " -Color White
        Write-ColorOutput "$((Get-WmiObject Win32_OperatingSystem).Caption)`r`n" -Color Green
        
        Write-ColorOutput "Product Key Channel: " -Color White
        Write-ColorOutput "$($slp.Description)`r`n" -Color Green
        
        Write-ColorOutput "Activation Status: " -Color White
        switch ($slp.LicenseStatus) {
            0 { Write-ColorOutput "Unlicensed`r`n" -Color Red }
            1 { Write-ColorOutput "Licensed`r`n" -Color Green }
            2 { Write-ColorOutput "Initial grace period`r`n" -Color Yellow }
            3 { Write-ColorOutput "Additional grace period`r`n" -Color Yellow }
            4 { Write-ColorOutput "Non-genuine grace period`r`n" -Color Red }
            5 { Write-ColorOutput "Notification`r`n" -Color Yellow }
            6 { Write-ColorOutput "Extended grace period`r`n" -Color Yellow }
            default { Write-ColorOutput "Unknown`r`n" -Color Red }
        }
        
        Write-ColorOutput "Product Key: " -Color White
        Write-ColorOutput "$($slp.PartialProductKey)`r`n" -Color Green
        
        Write-ColorOutput "License Expiration: " -Color White
        if ($slp.LicenseStatus -eq 1) {
            Write-ColorOutput "Permanent`r`n" -Color Green
        } else {
            $graceMinutes = $slp.GracePeriodRemaining
            $graceDays = [math]::Round($graceMinutes / 1440)
            Write-ColorOutput "$graceDays days remaining`r`n" -Color Yellow
        }
    }
    catch {
        Write-ColorOutput "Error occurred while checking activation status:`r`n$($_.Exception.Message)`r`n" -Color Red
    }
}

# Function to check Office activation status
function Check-OfficeActivation {
    Write-ColorOutput "`r`n============================================================`r`n" -Color Cyan
    Write-ColorOutput "Checking Office Activation Status...`r`n" -Color Yellow
    
    try {
        $officePaths = @(
            "C:\Program Files\Microsoft Office\Office16",
            "C:\Program Files (x86)\Microsoft Office\Office16"
        )
        
        $ospp = $null
        foreach ($path in $officePaths) {
            if (Test-Path $path) {
                $ospp = Get-WmiObject -Query "SELECT * FROM OfficeSoftwareProtectionProduct WHERE PartialProductKey IS NOT NULL"
                break
            }
        }
        
        if ($ospp) {
            foreach ($product in $ospp) {
                Write-ColorOutput "Product: " -Color White
                Write-ColorOutput "$($product.Name)`r`n" -Color Green
                
                Write-ColorOutput "Activation Status: " -Color White
                switch ($product.LicenseStatus) {
                    0 { Write-ColorOutput "Unlicensed`r`n" -Color Red }
                    1 { Write-ColorOutput "Licensed`r`n" -Color Green }
                    2 { Write-ColorOutput "Initial grace period`r`n" -Color Yellow }
                    3 { Write-ColorOutput "Additional grace period`r`n" -Color Yellow }
                    4 { Write-ColorOutput "Non-genuine grace period`r`n" -Color Red }
                    5 { Write-ColorOutput "Notification`r`n" -Color Yellow }
                    default { Write-ColorOutput "Unknown`r`n" -Color Red }
                }
            }
        } else {
            Write-ColorOutput "No Microsoft Office installation detected.`r`n" -Color Yellow
        }
    }
    catch {
        Write-ColorOutput "Error occurred while checking Office activation status:`r`n$($_.Exception.Message)`r`n" -Color Red
    }
}

# Refresh button click event
$refreshButton.Add_Click({
    Check-WindowsActivation
    Check-OfficeActivation
})

# Save button click event
$saveButton.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $saveDialog.DefaultExt = "txt"
    $saveDialog.AddExtension = $true
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputBox.Text | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
    }
})

# Initial check when form loads
$mainForm.Add_Shown({
    Check-WindowsActivation
    Check-OfficeActivation
})

# Show the form
$mainForm.ShowDialog()