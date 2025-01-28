# MKVConverter by Kehlanii
# Requires PowerShell 5.1 or newer
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MKV Converter"
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Create MKVMerge path input
$lblMkvPath = New-Object System.Windows.Forms.Label
$lblMkvPath.Location = New-Object System.Drawing.Point(10,20)
$lblMkvPath.Size = New-Object System.Drawing.Size(120,20)
$lblMkvPath.Text = "MKVMerge Path:"
$form.Controls.Add($lblMkvPath)

$txtMkvPath = New-Object System.Windows.Forms.TextBox
$txtMkvPath.Location = New-Object System.Drawing.Point(130,20)
$txtMkvPath.Size = New-Object System.Drawing.Size(350,20)
$form.Controls.Add($txtMkvPath)

$btnBrowseMkv = New-Object System.Windows.Forms.Button
$btnBrowseMkv.Location = New-Object System.Drawing.Point(490,20)
$btnBrowseMkv.Size = New-Object System.Drawing.Size(80,20)
$btnBrowseMkv.Text = "Browse"
$btnBrowseMkv.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "mkvmerge.exe|mkvmerge.exe|All files (*.*)|*.*"
    if($openFileDialog.ShowDialog() -eq "OK") {
        $txtMkvPath.Text = $openFileDialog.FileName
    }
})
$form.Controls.Add($btnBrowseMkv)

# Create source folder input
$lblSourcePath = New-Object System.Windows.Forms.Label
$lblSourcePath.Location = New-Object System.Drawing.Point(10,50)
$lblSourcePath.Size = New-Object System.Drawing.Size(120,20)
$lblSourcePath.Text = "Source Folder:"
$form.Controls.Add($lblSourcePath)

$txtSourcePath = New-Object System.Windows.Forms.TextBox
$txtSourcePath.Location = New-Object System.Drawing.Point(130,50)
$txtSourcePath.Size = New-Object System.Drawing.Size(350,20)
$form.Controls.Add($txtSourcePath)

$btnBrowseSource = New-Object System.Windows.Forms.Button
$btnBrowseSource.Location = New-Object System.Drawing.Point(490,50)
$btnBrowseSource.Size = New-Object System.Drawing.Size(80,20)
$btnBrowseSource.Text = "Browse"
$btnBrowseSource.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Source File"
    if($openFileDialog.ShowDialog() -eq "OK") {
        # Get the directory path from the selected file
        $txtSourcePath.Text = [System.IO.Path]::GetDirectoryName($openFileDialog.FileName)
    }
})
$form.Controls.Add($btnBrowseSource)

# Create file extensions input
$lblExtensions = New-Object System.Windows.Forms.Label
$lblExtensions.Location = New-Object System.Drawing.Point(10,80)
$lblExtensions.Size = New-Object System.Drawing.Size(120,20)
$lblExtensions.Text = "File Extensions:"
$form.Controls.Add($lblExtensions)

$txtExtensions = New-Object System.Windows.Forms.TextBox
$txtExtensions.Location = New-Object System.Drawing.Point(130,80)
$txtExtensions.Size = New-Object System.Drawing.Size(350,20)
$txtExtensions.Text = "*.avi,*.mp4,*.wmv,*.ts"
$form.Controls.Add($txtExtensions)

# Create progress display
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,110)
$progressBar.Size = New-Object System.Drawing.Size(560,20)
$form.Controls.Add($progressBar)

# Create status listbox
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,140)
$listBox.Size = New-Object System.Drawing.Size(560,180)
$form.Controls.Add($listBox)

# Create convert button
$btnConvert = New-Object System.Windows.Forms.Button
$btnConvert.Location = New-Object System.Drawing.Point(10,330)
$btnConvert.Size = New-Object System.Drawing.Size(560,30)
$btnConvert.Text = "Convert to MKV"
$btnConvert.Add_Click({
    if(![System.IO.File]::Exists($txtMkvPath.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select valid mkvmerge.exe path!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if(![System.IO.Directory]::Exists($txtSourcePath.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select valid source folder!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $extensions = $txtExtensions.Text.Split(',')
    $files = @()
    foreach($ext in $extensions) {
        $files += Get-ChildItem -Path $txtSourcePath.Text -Filter $ext.Trim()
    }

    if($files.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No files found with specified extensions!", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $progressBar.Maximum = $files.Count
    $progressBar.Value = 0
    $listBox.Items.Clear()

    foreach($file in $files) {
        $outputPath = [System.IO.Path]::ChangeExtension($file.FullName, "mkv")
        $listBox.Items.Add("Converting: $($file.Name)")
        $listBox.SelectedIndex = $listBox.Items.Count - 1
        
        try {
            $process = Start-Process -FilePath $txtMkvPath.Text -ArgumentList "`"$($file.FullName)`" -o `"$outputPath`"" -NoNewWindow -Wait -PassThru
            if($process.ExitCode -eq 0) {
                $listBox.Items.Add("Successfully converted: $($file.Name)")
            } else {
                $listBox.Items.Add("Error converting: $($file.Name)")
            }
        } catch {
            $listBox.Items.Add("Error: $_")
        }
        
        $progressBar.Value++
        $listBox.SelectedIndex = $listBox.Items.Count - 1
    }

    [System.Windows.Forms.MessageBox]::Show("Conversion completed!", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})
$form.Controls.Add($btnConvert)

# Show the form
$form.ShowDialog()