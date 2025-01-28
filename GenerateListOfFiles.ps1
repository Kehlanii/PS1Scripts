Add-Type -AssemblyName System.Windows.Forms

# Stworzenie formularzu
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Generator nazw w pliku by kehlanii'
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = 'CenterScreen'

# Stworzenie przycisku do przeglądania folderów
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = 'Przeglądaj'
$browseButton.Location = New-Object System.Drawing.Point(10, 10)
$browseButton.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($browseButton)

# Stworzenie Przycisku do wygenerowania pliku tekstowego
$generateButton = New-Object System.Windows.Forms.Button
$generateButton.Text = 'Wygeneruj listę'
$generateButton.Location = New-Object System.Drawing.Point(140, 10)
$generateButton.Size = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($generateButton)

# Stworzenie przycisku do zmiany motywu
$themeButton = New-Object System.Windows.Forms.Button
$themeButton.Text = 'Przełączanie motywu'
$themeButton.Location = New-Object System.Drawing.Point(270, 10)
$themeButton.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($themeButton)

# etykieta, aby wyświetlić dane wyjściowe wyboru folderu
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(10, 50)
$outputLabel.Size = New-Object System.Drawing.Size(360, 100)
$outputLabel.Text = "Proszę wybrać folder..."
$form.Controls.Add($outputLabel)

# Zmienna do przechowywania wybranej ścieżki folderu
$script:folderPath = $null
$script:isDarkTheme = $false

# Funkcja do ustawienia ciemnego motywu
function Set-DarkTheme {
    $form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $form.ForeColor = [System.Drawing.Color]::White
    $browseButton.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $browseButton.ForeColor = [System.Drawing.Color]::White
    $generateButton.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $generateButton.ForeColor = [System.Drawing.Color]::White
    $themeButton.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $themeButton.ForeColor = [System.Drawing.Color]::White
    $outputLabel.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $outputLabel.ForeColor = [System.Drawing.Color]::White
}

# Funkcja do ustawienia Jasnego Motywu
function Set-LightTheme {
    $form.BackColor = [System.Drawing.SystemColors]::Control
    $form.ForeColor = [System.Drawing.Color]::Black
    $browseButton.BackColor = [System.Drawing.SystemColors]::Control
    $browseButton.ForeColor = [System.Drawing.Color]::Black
    $generateButton.BackColor = [System.Drawing.SystemColors]::Control
    $generateButton.ForeColor = [System.Drawing.Color]::Black
    $themeButton.BackColor = [System.Drawing.SystemColors]::Control
    $themeButton.ForeColor = [System.Drawing.Color]::Black
    $outputLabel.BackColor = [System.Drawing.SystemColors]::Control
    $outputLabel.ForeColor = [System.Drawing.Color]::Black
}

# Obsługa zdarzeń przełączania motywu
$themeButton.Add_Click({
    $script:isDarkTheme = !$script:isDarkTheme
    if ($script:isDarkTheme) {        # Removed the space after $
        Set-DarkTheme
    } else {
        Set-LightTheme
    }
})

# Obsługa zdarzeń przeglądania folderów
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $folderBrowser.ValidateNames = $false
    $folderBrowser.CheckFileExists = $false
    $folderBrowser.CheckPathExists = $true
    $folderBrowser.FileName = "Folder Selection"
    $folderBrowser.Filter = "Folders|no_files"
    $folderBrowser.Title = "Wybierz folder"
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:folderPath = [System.IO.Path]::GetDirectoryName($folderBrowser.FileName)
        $outputLabel.Text = "Wybrany folder: $script:folderPath"
    }
})

# Obsługa zdarzeń do generowania listy
$generateButton.Add_Click({
    if ($script:folderPath) {
        $files = Get-ChildItem -Path $script:folderPath -File | Select-Object -ExpandProperty Name
        $outputFilePath = Join-Path -Path $script:folderPath -ChildPath 'nazwy.txt'
        
        # Using -LiteralPath instead of -FilePath
        $files | Out-File -LiteralPath $outputFilePath
        $outputLabel.Text = "File list generated: $outputFilePath"
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a folder first.")
    }
})

# Ustawienie początkowego motywu
Set-LightTheme

# Pokaż formularz
$form.ShowDialog()