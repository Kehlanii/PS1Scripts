# MKVAudioExtractor by Kehlanii
# Requires PowerShell 5.1 or newer

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form with a modern size and style
$form = New-Object System.Windows.Forms.Form
$form.Text = "Modern MKV Audio Extractor"
$form.Size = New-Object System.Drawing.Size(1000,700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)

# Create a modern panel for the top section
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Location = New-Object System.Drawing.Point(0,0)
$topPanel.Size = New-Object System.Drawing.Size(1000,60)
$topPanel.BackColor = [System.Drawing.Color]::White
$topPanel.BorderStyle = "None"

# Create a modern-styled MKV files selection button
$btnSelectMKV = New-Object System.Windows.Forms.Button
$btnSelectMKV.Location = New-Object System.Drawing.Point(20,15)
$btnSelectMKV.Size = New-Object System.Drawing.Size(150,30)
$btnSelectMKV.Text = "Browse MKV Files"
$btnSelectMKV.FlatStyle = "Flat"
$btnSelectMKV.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$btnSelectMKV.ForeColor = [System.Drawing.Color]::White
$btnSelectMKV.Cursor = "Hand"

# Create a modern textbox to display selected files count
$txtFilePath = New-Object System.Windows.Forms.TextBox
$txtFilePath.Location = New-Object System.Drawing.Point(190,15)
$txtFilePath.Size = New-Object System.Drawing.Size(770,30)
$txtFilePath.ReadOnly = $true
$txtFilePath.BorderStyle = "FixedSingle"
$txtFilePath.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Create a modern ListView for files and their tracks
$listFiles = New-Object System.Windows.Forms.ListView
$listFiles.Location = New-Object System.Drawing.Point(20,80)
$listFiles.Size = New-Object System.Drawing.Size(940,480)
$listFiles.View = [System.Windows.Forms.View]::Details
$listFiles.FullRowSelect = $true
$listFiles.GridLines = $true
$listFiles.CheckBoxes = $true
$listFiles.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Add columns to ListView with modern widths
$listFiles.Columns.Add("File Name", 300)
$listFiles.Columns.Add("Track ID", 80)
$listFiles.Columns.Add("Language", 100)
$listFiles.Columns.Add("Codec", 120)
$listFiles.Columns.Add("Properties", 320)

# Create a modern bottom panel
$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Location = New-Object System.Drawing.Point(0,580)
$bottomPanel.Size = New-Object System.Drawing.Size(1000,120)
$bottomPanel.BackColor = [System.Drawing.Color]::White
$bottomPanel.BorderStyle = "None"

# Create a modern extract button
$btnExtract = New-Object System.Windows.Forms.Button
$btnExtract.Location = New-Object System.Drawing.Point(20,15)
$btnExtract.Size = New-Object System.Drawing.Size(150,35)
$btnExtract.Text = "Extract Selected"
$btnExtract.FlatStyle = "Flat"
$btnExtract.BackColor = [System.Drawing.Color]::FromArgb(0,150,0)
$btnExtract.ForeColor = [System.Drawing.Color]::White
$btnExtract.Enabled = $false
$btnExtract.Cursor = "Hand"

# Create a modern status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(190,20)
$lblStatus.Size = New-Object System.Drawing.Size(770,25)
$lblStatus.Text = "Ready"
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Create a modern progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,65)
$progressBar.Size = New-Object System.Drawing.Size(940,35)
$progressBar.Style = "Continuous"

# Global variable to store selected files
$global:selectedFiles = @()

# Function to get audio tracks from MKV file
function Get-MKVAudioTracks {
    param($mkvPath)
    
    $tracks = @()
    $mkvInfo = & mkvmerge -J $mkvPath | ConvertFrom-Json
    
    foreach ($track in $mkvInfo.tracks) {
        if ($track.type -eq "audio") {
            $channelsText = "Channels: " + $track.properties.audio_channels
            $samplingText = "Sampling rate: " + $track.properties.audio_sampling_frequency + " Hz"
            $properties = $channelsText + ", " + $samplingText
            
            $trackInfo = @{
                ID = $track.id
                Language = if ($track.properties.language) { $track.properties.language } else { "und" }
                Codec = $track.codec
                Properties = $properties
            }
            $tracks += $trackInfo
        }
    }
    return $tracks
}

# MKV files selection handler
$btnSelectMKV.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "MKV files (*.mkv)|*.mkv"
    $openFileDialog.Multiselect = $true
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $listFiles.Items.Clear()
        $global:selectedFiles = $openFileDialog.FileNames
        $fileCount = $global:selectedFiles.Count
        $txtFilePath.Text = "Selected " + $fileCount + " files"
        
        $totalFiles = $global:selectedFiles.Count
        $currentFile = 0
        
        foreach ($file in $global:selectedFiles) {
            $currentFile++
            $fileName = [System.IO.Path]::GetFileName($file)
            $lblStatus.Text = "Reading file " + $currentFile + " of " + $totalFiles + ": " + $fileName
            $progressBar.Value = ($currentFile / $totalFiles) * 100
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                $tracks = Get-MKVAudioTracks $file
                foreach ($track in $tracks) {
                    $item = New-Object System.Windows.Forms.ListViewItem($fileName)
                    $item.SubItems.Add($track.ID)
                    $item.SubItems.Add($track.Language)
                    $item.SubItems.Add($track.Codec)
                    $item.SubItems.Add($track.Properties)
                    $item.Tag = @{
                        FilePath = $file
                        TrackID = $track.ID
                    }
                    $item.Checked = $true
                    $listFiles.Items.Add($item)
                }
            }
            catch {
                $lblStatus.Text = "Error reading file " + $fileName + ": " + $_
            }
        }
        
        $btnExtract.Enabled = $true
        $progressBar.Value = 0
        $trackCount = $listFiles.Items.Count
        $lblStatus.Text = "Ready to extract audio from " + $trackCount + " tracks"
    }
})

# Extract button handler
$btnExtract.Add_Click({
    $checkedItems = $listFiles.Items | Where-Object { $_.Checked }
    if ($checkedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one track to extract", "No Tracks Selected")
        return
    }
    
    $outputFolder = New-Object System.Windows.Forms.FolderBrowserDialog
    $outputFolder.Description = "Select output folder for extracted audio"
    
    if ($outputFolder.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $totalTracks = $checkedItems.Count
        $currentTrack = 0
        
        foreach ($item in $checkedItems) {
            $currentTrack++
            $fileName = [System.IO.Path]::GetFileNameWithoutExtension($item.Text)
            $language = $item.SubItems[2].Text
            
            # Construct output file path
            $outputFile = Join-Path $outputFolder.SelectedPath ($fileName + "-" + $language + ".mka")
            $lblStatus.Text = "Extracting track " + $currentTrack + " of " + $totalTracks + ": " + $fileName
            $progressBar.Value = ($currentTrack / $totalTracks) * 100
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                # Correcting the command for mkvextract
                $trackArg = $item.Tag.TrackID.ToString() + ":`"$outputFile`""
                $extractCommand = "mkvextract tracks `"$($item.Tag.FilePath)`" $trackArg"
                Write-Host "Executing: $extractCommand"  # This will log the command for debugging
                
                # Run the extraction command
                Invoke-Expression $extractCommand
            }
            catch {
                $lblStatus.Text = "Error extracting from " + $fileName + ": " + $_
                continue
            }
        }
        
        $progressBar.Value = 100
        $lblStatus.Text = "Extraction completed! Processed " + $totalTracks + " tracks"
        [System.Windows.Forms.MessageBox]::Show("Audio extraction completed!", "Success")
        $progressBar.Value = 0
    }
})


# Add controls to panels and form
$topPanel.Controls.Add($btnSelectMKV)
$topPanel.Controls.Add($txtFilePath)
$bottomPanel.Controls.Add($btnExtract)
$bottomPanel.Controls.Add($lblStatus)
$bottomPanel.Controls.Add($progressBar)

$form.Controls.Add($topPanel)
$form.Controls.Add($listFiles)
$form.Controls.Add($bottomPanel)
# Show the form
$form.ShowDialog()