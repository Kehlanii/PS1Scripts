# Universal Bulk Archive Extractor by Kehlanii
# Requires PowerShell 5.1 or newer

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Universal Archive Extractor" Height="500" Width="800"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="100"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Selected Archives:" Margin="0,0,0,5"/>
        <ListBox Grid.Row="1" Name="FileList" Margin="0,0,0,10"/>
        
        <Grid Grid.Row="2" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button Name="BrowseFiles" Content="Select Archives" Padding="10,5" Margin="0,0,10,0" Grid.Column="0"/>
            <TextBox Name="OutputPath" IsReadOnly="True" Grid.Column="1" Height="28" Background="#FFF0F0F0"/>
            <Button Name="BrowseFolder" Content="Select Output" Padding="10,5" Margin="10,0,0,0" Grid.Column="2"/>
        </Grid>
        
        <CheckBox Grid.Row="3" Name="ExtractToSingleFolder" Content="Extract all archives to a single folder" Margin="0,0,0,10"/>
        
        <Grid Grid.Row="4" Margin="0,0,0,10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock Name="StatusText" Text="Ready" Margin="0,0,0,5"/>
            <ProgressBar Grid.Row="1" Name="ProgressBar" Height="20"/>
        </Grid>

        <Button Grid.Row="5" Name="ExtractButton" Content="Extract Archives" Padding="20,5" HorizontalAlignment="Left"/>
        <TextBox Grid.Row="6" Name="LogBox" IsReadOnly="True" 
                 VerticalScrollBarVisibility="Visible" 
                 TextWrapping="Wrap"
                 Height="100"
                 Margin="0,10,0,0"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$controls = @{}
"FileList", "BrowseFiles", "OutputPath", "BrowseFolder", "ProgressBar", "ExtractButton", "LogBox", "ExtractToSingleFolder", "StatusText" | ForEach-Object {
    $controls[$_] = $window.FindName($_)
}

# Initialize selected files array
$script:selectedFiles = @()

function Write-ExtractorLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "$timestamp - [$Type] $Message"
    
    Write-Host $(switch ($Type) {
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
        default { Write-Host $logMessage }
    })
    
    $controls.LogBox.Dispatcher.Invoke([Action]{
        $controls.LogBox.AppendText("$logMessage`n")
        $controls.LogBox.ScrollToEnd()
    })
}

function Test-PathIsValid {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        $null = [System.IO.Path]::GetFullPath($Path)
        return $true
    }
    catch {
        return $false
    }
}

function New-ExtractionDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        if (Test-Path -LiteralPath $Path) {
            Get-ChildItem -Path $Path -Recurse -Force | Remove-Item -Force -Recurse
            Remove-Item -LiteralPath $Path -Force
            Start-Sleep -Milliseconds 500
        }
        
        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Start-Sleep -Milliseconds 500
        Write-ExtractorLog "Created clean directory: $Path"
        return $true
    }
    catch {
        Write-ExtractorLog "Failed to create directory: $Path - $($_.Exception.Message)" -Type 'Error'
        throw
    }
}

function Update-ExtractionProgress {
    param(
        [Parameter(Mandatory = $true)]
        [double]$PercentComplete,
        [string]$CurrentOperation
    )
    
    $controls.ProgressBar.Dispatcher.Invoke([Action]{
        $controls.ProgressBar.Value = $PercentComplete
        $controls.StatusText.Text = "$CurrentOperation - $([math]::Round($PercentComplete))% Complete"
        $window.Title = "Universal Archive Extractor - $([math]::Round($PercentComplete))% Complete"
    })
    Write-ExtractorLog $CurrentOperation
}

function Invoke-ArchiveExtraction {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$ArchiveFiles,
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        [bool]$SingleFolder
    )
    
    Write-ExtractorLog "Starting archive extraction process"
    
    if (-not (Test-PathIsValid $DestinationPath)) {
        throw "Invalid destination path: $DestinationPath"
    }
    
    if (-not (Test-Path -LiteralPath $DestinationPath)) {
        New-ExtractionDirectory $DestinationPath
    }
    elseif ($SingleFolder) {
        New-ExtractionDirectory $DestinationPath
    }
    
    $total = $ArchiveFiles.Count
    $current = 0
    $errors = @()
    
    foreach ($archiveFile in $ArchiveFiles) {
        $current++
        $filePath = $archiveFile.FullName
        $currentProgress = [math]::Round((($current - 1) / $total) * 100)
        
        try {
            Update-ExtractionProgress -PercentComplete $currentProgress -CurrentOperation "Processing $($archiveFile.Name) ($current of $total)"
            
            if (-not (Test-Path -LiteralPath $filePath)) {
                throw "Archive file not found: $filePath"
            }
            
            $extension = $archiveFile.Extension.ToLower()
            $baseName = $archiveFile.BaseName
            $extractPath = if ($SingleFolder) { $DestinationPath } else { Join-Path -Path $DestinationPath -ChildPath $baseName }
            
            if (-not $SingleFolder) {
                New-ExtractionDirectory $extractPath
            }
            
            switch ($extension) {
                ".zip" {
                    try {
                        Add-Type -AssemblyName System.IO.Compression.FileSystem
                        $zip = [System.IO.Compression.ZipFile]::OpenRead($filePath)
                        $totalEntries = $zip.Entries.Count
                        $processedEntries = 0
                        
                        foreach ($entry in $zip.Entries) {
                            $processedEntries++
                            $entryProgress = ($processedEntries / $totalEntries) * (100 / $total)
                            $totalProgress = $currentProgress + $entryProgress
                            
                            if ($processedEntries % 5 -eq 0) {
                                Update-ExtractionProgress -PercentComplete $totalProgress -CurrentOperation "Extracting $($archiveFile.Name) - $processedEntries of $totalEntries entries"
                            }
                            
                            $targetPath = Join-Path $extractPath $entry.FullName
                            $targetDir = [System.IO.Path]::GetDirectoryName($targetPath)
                            
                            if (-not (Test-Path -LiteralPath $targetDir)) {
                                New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                            }
                            
                            if (-not $entry.Name -eq '') {
                                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $targetPath, $true)
                            }
                        }
                        $zip.Dispose()
                    }
                    catch {
                        throw "ZIP extraction failed: $($_.Exception.Message)"
                    }
                }
                { $_ -in ".7z", ".rar" } {
                    if (-not (Get-Command -Name Expand-7Zip -ErrorAction SilentlyContinue)) {
                        throw "7Zip4PowerShell module is required for $extension files"
                    }
                    
                    try {
                        $fileSize = (Get-Item -LiteralPath $filePath).Length
                        $startTime = Get-Date
                        
                        $job = Start-Job -ScriptBlock {
                            param($filePath, $extractPath)
                            Expand-7Zip -ArchiveFileName $filePath -TargetPath $extractPath
                        } -ArgumentList $filePath, $extractPath
                        
                        while ($job.State -eq 'Running') {
                            $elapsedTime = ((Get-Date) - $startTime).TotalSeconds
                            if ($elapsedTime -gt 0) {
                                $progressEstimate = [math]::Min(95, ($elapsedTime / ($fileSize / 1MB) * 10))
                                $totalProgress = $currentProgress + ($progressEstimate / $total)
                                Update-ExtractionProgress -PercentComplete $totalProgress -CurrentOperation "Extracting $($archiveFile.Name) ($progressEstimate% estimated)"
                            }
                            Start-Sleep -Milliseconds 100
                        }
                        
                        $job | Wait-Job | Receive-Job
                        Remove-Job -Job $job -Force
                    }
                    catch {
                        throw "7Zip extraction failed: $($_.Exception.Message)"
                    }
                }
                default {
                    throw "Unsupported archive format: $extension"
                }
            }
            
            Write-ExtractorLog "Successfully extracted: $baseName"
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-ExtractorLog $errorMsg -Type 'Error'
            $errors += $errorMsg
        }
    }
    
    Update-ExtractionProgress -PercentComplete 100 -CurrentOperation "Extraction Complete!"
    
    if ($errors.Count -gt 0) {
        throw ($errors -join "`n")
    }
}

$controls.BrowseFiles.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Multiselect = $true
    $fileDialog.Filter = "All Archives (*.zip;*.7z;*.rar)|*.zip;*.7z;*.rar|ZIP files (*.zip)|*.zip|7Z files (*.7z)|*.7z|RAR files (*.rar)|*.rar|All files (*.*)|*.*"
    
    Write-ExtractorLog "Opening file selection dialog"
    if ($fileDialog.ShowDialog() -eq 'OK') {
        $script:selectedFiles = @()
        $controls.FileList.Items.Clear()
        
        foreach($filePath in $fileDialog.FileNames) {
            try {
                $fileInfo = Get-Item -LiteralPath $filePath -ErrorAction Stop
                $script:selectedFiles += $fileInfo
                $controls.FileList.Items.Add($fileInfo.FullName)
                Write-ExtractorLog "Selected file: $($fileInfo.FullName)"
            }
            catch {
                Write-ExtractorLog "Error adding file $filePath - $($_.Exception.Message)" -Type 'Error'
            }
        }
        Write-ExtractorLog "Total files selected: $($script:selectedFiles.Count)"
    }
})

$controls.BrowseFolder.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    Write-ExtractorLog "Opening folder selection dialog"
    if ($folderDialog.ShowDialog() -eq 'OK') {
        $controls.OutputPath.Text = $folderDialog.SelectedPath
        Write-ExtractorLog "Selected output folder: $($folderDialog.SelectedPath)"
    }
})

$controls.ExtractButton.Add_Click({
    if ($script:selectedFiles.Count -eq 0) {
        Write-ExtractorLog "No files selected" -Type 'Warning'
        [System.Windows.MessageBox]::Show("Please select at least one archive file.", "No Files Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($controls.OutputPath.Text)) {
        Write-ExtractorLog "No output folder selected" -Type 'Warning'
        [System.Windows.MessageBox]::Show("Please select an output folder.", "No Output Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $controls.ExtractButton.IsEnabled = $false
    $controls.ProgressBar.Value = 0
    
    try {
        Write-ExtractorLog "Starting extraction process"
        $singleFolder = $controls.ExtractToSingleFolder.IsChecked
        Invoke-ArchiveExtraction -ArchiveFiles $script:selectedFiles -DestinationPath $controls.OutputPath.Text -SingleFolder $singleFolder
        Write-ExtractorLog "Extraction completed successfully!"
        [System.Windows.MessageBox]::Show("Extraction completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
    catch {
        Write-ExtractorLog "Extraction failed: $($_.Exception.Message)" -Type 'Error'
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Extraction Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
    finally {
        $controls.ExtractButton.IsEnabled = $true
        $controls.ProgressBar.Value = 0
        $controls.StatusText.Text = "Ready"
        $window.Title = "Universal Archive Extractor"
    }
})

# Show the window
$window.ShowDialog()