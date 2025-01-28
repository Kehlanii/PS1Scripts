# MKV Audio Track Switcher by kehlanii
# Requires PowerShell 5.1 or newer
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to get/set the saved MKVToolNix path
function Get-SavedMkvToolPath {
    $regPath = "HKCU:\Software\MkvAudioSwitcher"
    if (Test-Path $regPath) {
        return (Get-ItemProperty -Path $regPath -Name "MkvToolPath" -ErrorAction SilentlyContinue).MkvToolPath
    }
    return $null
}

function Save-MkvToolPath {
    param($path)
    $regPath = "HKCU:\Software\MkvAudioSwitcher"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "MkvToolPath" -Value $path
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MKV Audio Track Switcher"
$form.Size = New-Object System.Drawing.Size(600,450)
$form.StartPosition = "CenterScreen"

# Create MKVToolNix path section
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Location = New-Object System.Drawing.Point(10,20)
$pathLabel.Size = New-Object System.Drawing.Size(560,20)
$pathLabel.Text = "MKVToolNix Path (select mkvpropedit.exe):"
$form.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(10,45)
$pathTextBox.Size = New-Object System.Drawing.Size(460,20)
$pathTextBox.Text = Get-SavedMkvToolPath
$form.Controls.Add($pathTextBox)

$pathButton = New-Object System.Windows.Forms.Button
$pathButton.Location = New-Object System.Drawing.Point(480,43)
$pathButton.Size = New-Object System.Drawing.Size(90,25)
$pathButton.Text = "Browse"
$pathButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "mkvpropedit.exe|mkvpropedit.exe|All files (*.*)|*.*"
    $openFileDialog.InitialDirectory = "C:\Program Files\MKVToolNix"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $pathTextBox.Text = $openFileDialog.FileName
        Save-MkvToolPath $openFileDialog.FileName
    }
})
$form.Controls.Add($pathButton)

# Create a label for MKV files
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,80)
$label.Size = New-Object System.Drawing.Size(560,20)
$label.Text = "Select MKV files to modify:"
$form.Controls.Add($label)

# Create a listbox for files
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,110)
$listBox.Size = New-Object System.Drawing.Size(460,250)
$listBox.SelectionMode = "MultiExtended"
$form.Controls.Add($listBox)

# Create Browse button for MKV files
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(480,110)
$browseButton.Size = New-Object System.Drawing.Size(90,30)
$browseButton.Text = "Browse"
$browseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "MKV files (*.mkv)|*.mkv"
    $openFileDialog.Multiselect = $true
    if ($openFileDialog.ShowDialog() -eq "OK") {
        foreach ($file in $openFileDialog.FileNames) {
            $listBox.Items.Add($file)
        }
    }
})
$form.Controls.Add($browseButton)

# Create Clear button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(480,150)
$clearButton.Size = New-Object System.Drawing.Size(90,30)
$clearButton.Text = "Clear List"
$clearButton.Add_Click({
    $listBox.Items.Clear()
    $progressBar.Value = 0
    $statusLabel.Text = "Ready"
})
$form.Controls.Add($clearButton)

# Create Process button
$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(200,370)
$processButton.Size = New-Object System.Drawing.Size(200,30)
$processButton.Text = "Make Second Track Default"
$processButton.Add_Click({
    if (-not (Test-Path $pathTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select the correct path to mkvpropedit.exe first!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    if ($listBox.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one MKV file to process!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $progressBar.Value = 0
    $progressBar.Maximum = $listBox.Items.Count
    
    foreach ($file in $listBox.Items) {
        try {
            # Use mkvpropedit to change the default track
            $mkvpropedit = $pathTextBox.Text
            $command = "& `"$mkvpropedit`" `"$file`" --edit track:a1 --set flag-default=0 --edit track:a2 --set flag-default=1"
            Invoke-Expression $command
            
            $progressBar.Value += 1
            $statusLabel.Text = "Processing: $([Math]::Round(($progressBar.Value / $progressBar.Maximum) * 100))%"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error processing file: $file`n$($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
    
    if ($progressBar.Value -eq $progressBar.Maximum) {
        $statusLabel.Text = "Complete! All files processed."
        [System.Windows.Forms.MessageBox]::Show("Processing complete!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($processButton)

# Create progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,410)
$progressBar.Size = New-Object System.Drawing.Size(560,20)
$form.Controls.Add($progressBar)

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10,370)
$statusLabel.Size = New-Object System.Drawing.Size(180,20)
$statusLabel.Text = "Ready"
$form.Controls.Add($statusLabel)

# Show the form
$form.ShowDialog()