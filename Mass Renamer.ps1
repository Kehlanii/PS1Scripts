# MASS Renamer by kehlanii
# Requires PowerShell 5.1 or newer
Add-Type -AssemblyName System.Windows.Forms

# Globalna zmienna przechowująca ścieżkę do wybranego folderu
$global:folderPath = (Get-Location).Path

# Tworzenie głównego formularza
$form = New-Object System.Windows.Forms.Form
$form.Text = "Zmieniacz nazw by kehlanii"
$form.Width = 850
$form.Height = 750 
$form.MaximumSize = New-Object System.Drawing.Size(1050, 900) 
$form.MinimumSize = New-Object System.Drawing.Size(650, 500)

# Dodanie nowych komponentów do zmiany nazw wyrażeń regularnych
# Pole wyboru do włączania/wyłączania zmiany nazwy wyrażenia regularnego
$regexCheckbox = New-Object System.Windows.Forms.CheckBox
$regexCheckbox.Text = "Użyj wyrażeń regularnych"
$regexCheckbox.Location = New-Object System.Drawing.Point(650, 625)
$regexCheckbox.Width = 300
$form.Controls.Add($regexCheckbox)

# Pole tekstowe dla wzorca wyrażeń regularnych
$regexPatternTextBox = New-Object System.Windows.Forms.TextBox
$regexPatternTextBox.Location = New-Object System.Drawing.Point(10, 660)
$regexPatternTextBox.Width = 400
$regexPatternTextBox.Text = "Wzór regex"
$regexPatternTextBox.Enabled = $false
$form.Controls.Add($regexPatternTextBox)

# Pole tekstowe dla replacement pattern
$replacementTextBox = New-Object System.Windows.Forms.TextBox
$replacementTextBox.Location = New-Object System.Drawing.Point(420, 660)
$replacementTextBox.Width = 400
$replacementTextBox.Text = "Zamiennik"
$replacementTextBox.Enabled = $false
$form.Controls.Add($replacementTextBox)

# Włączanie/wyłączanie pól tekstowych regex na podstawie pola wyboru
$regexCheckbox.Add_CheckedChanged({
    if ($regexCheckbox.Checked) {
        $regexPatternTextBox.Enabled = $true
        $replacementTextBox.Enabled = $true
    } else {
        $regexPatternTextBox.Enabled = $false
        $replacementTextBox.Enabled = $false
    }
})

# Stosowanie ciemnego motywu
function Set-DarkTheme {
    $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
    $listview.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $listview.ForeColor = [System.Drawing.Color]::White
    $originalNameTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $originalNameTextBox.ForeColor = [System.Drawing.Color]::White
    $newNameTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $newNameTextBox.ForeColor = [System.Drawing.Color]::White

    $browseButton.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 66)
    $browseButton.ForeColor = [System.Drawing.Color]::White
    $refreshButton.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 66) 
    $refreshButton.ForeColor = [System.Drawing.Color]::White
    $renameButton.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 66)
    $renameButton.ForeColor = [System.Drawing.Color]::White
    $loadNamesButton.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 66)
    $loadNamesButton.ForeColor = [System.Drawing.Color]::White
    $exitButton.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 66)
    $exitButton.ForeColor = [System.Drawing.Color]::White

    $regexCheckbox.ForeColor = [System.Drawing.Color]::White
    $regexPatternTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $regexPatternTextBox.ForeColor = [System.Drawing.Color]::White
    $replacementTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $replacementTextBox.ForeColor = [System.Drawing.Color]::White
    $regexComboBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $regexComboBox.ForeColor = [System.Drawing.Color]::White
    $selectAllCheckbox.ForeColor = [System.Drawing.Color]::White
    $fileCountLabel.ForeColor = [System.Drawing.Color]::White
}

# Stosowanie jasnego motywu
function Set-LightTheme {
    $form.BackColor = [System.Drawing.Color]::White
    $listview.BackColor = [System.Drawing.Color]::White
    $listview.ForeColor = [System.Drawing.Color]::Black
    $originalNameTextBox.BackColor = [System.Drawing.Color]::White
    $originalNameTextBox.ForeColor = [System.Drawing.Color]::Black
    $newNameTextBox.BackColor = [System.Drawing.Color]::White
    $newNameTextBox.ForeColor = [System.Drawing.Color]::Black

    $browseButton.BackColor = [System.Drawing.Color]::LightGray
    $browseButton.ForeColor = [System.Drawing.Color]::Black
    $refreshButton.BackColor = [System.Drawing.Color]::LightGray
    $refreshButton.ForeColor = [System.Drawing.Color]::Black
    $renameButton.BackColor = [System.Drawing.Color]::LightGray
    $renameButton.ForeColor = [System.Drawing.Color]::Black
    $loadNamesButton.BackColor = [System.Drawing.Color]::LightGray
    $loadNamesButton.ForeColor = [System.Drawing.Color]::Black
    $exitButton.BackColor = [System.Drawing.Color]::LightGray
    $exitButton.ForeColor = [System.Drawing.Color]::Black

    $regexCheckbox.ForeColor = [System.Drawing.Color]::Black
    $regexPatternTextBox.BackColor = [System.Drawing.Color]::White
    $regexPatternTextBox.ForeColor = [System.Drawing.Color]::Black
    $replacementTextBox.BackColor = [System.Drawing.Color]::White
    $replacementTextBox.ForeColor = [System.Drawing.Color]::Black
    $regexComboBox.BackColor = [System.Drawing.Color]::White
    $regexComboBox.ForeColor = [System.Drawing.Color]::Black
    $selectAllCheckbox.ForeColor = [System.Drawing.Color]::Black
    $fileCountLabel.ForeColor = [System.Drawing.Color]::Black
}

# Utwórz widok ListView, aby wyświetlić pliki
$listview = New-Object System.Windows.Forms.ListView
$listview.Location = New-Object System.Drawing.Point(10, 10)
$listview.Width = 810
$listview.Height = 450
$listview.View = [System.Windows.Forms.View]::Details
$listview.MultiSelect = $true
$listview.Columns.Add("Original Name", 300)
$listview.Columns.Add("New Name", 300)
$listview.FullRowSelect = $true
$form.Controls.Add($listview)

# Dodanie etykiety dla licznika plików
$fileCountLabel = New-Object System.Windows.Forms.Label
$fileCountLabel.Location = New-Object System.Drawing.Point(10, 540)
$fileCountLabel.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($fileCountLabel)

# Global variables for tracking sorting state
$script:sortColumn = -1
$script:sortOrder = "None"

# Function to perform natural sorting
function Get-NaturalSort {
    param([string]$InputString)
    return [regex]::Replace($InputString, '\d+', { $args[0].Value.PadLeft(20) })
}

# Function to handle sorting
function Format-ListView {
    param (
        [System.Windows.Forms.ListView]$listView,
        [int]$columnIndex
    )

    if ($script:sortColumn -eq $columnIndex) {
        # Toggle sorting order if the same column is clicked again
        $script:sortOrder = if ($script:sortOrder -eq "Ascending") {
            "Descending"
        } else {
            "Ascending"
        }
    } else {
        # New column selected, sort ascending
        $script:sortColumn = $columnIndex
        $script:sortOrder = "Ascending"
    }

    # Sort the items
    $listView.BeginUpdate()
    try {
        $items = @($listView.Items)
        $sortedItems = $items | Sort-Object -Property {
            if ($script:sortColumn -eq 0) {
                # Sort by original name
                Get-NaturalSort $_.Text
            } else {
                # Sort by new name
                Get-NaturalSort $_.SubItems[1].Text
            }
        } -Descending:($script:sortOrder -eq "Descending")

        $listView.Items.Clear()
        $listView.Items.AddRange($sortedItems)
    }
    finally {
        $listView.EndUpdate()
    }
}

# Attach the ColumnClick event to the ListView
$listview.Add_ColumnClick({
    param ($listView, $e)
    Format-ListView -listView $listView -columnIndex $e.Column
})

# Funkcja ładowania plików do widoku listy
function Get-Files($folderPath) {
    $listview.Items.Clear()
    $files = Get-ChildItem -Path $folderPath | Where-Object { $_.Extension -in @('.mkv', '.mp4', '.avi', '.ass') }
    foreach ($file in $files) {
        $item = New-Object System.Windows.Forms.ListViewItem($file.Name)
        $item.SubItems.Add($file.Name)
        $listview.Items.Add($item)
    }
    $fileCountLabel.Text = "Liczba plików: " + $files.Count
}

# Tworzenie pola tekstowego do wyświetlania oryginalnych nazw
$originalNameTextBox = New-Object System.Windows.Forms.TextBox
$originalNameTextBox.Location = New-Object System.Drawing.Point(10, 470)
$originalNameTextBox.Width = 400
$originalNameTextBox.Height = 60
$originalNameTextBox.Multiline = $true
$originalNameTextBox.ScrollBars = "Vertical"
$originalNameTextBox.ReadOnly = $true
$form.Controls.Add($originalNameTextBox)

# Tworzenie pola tekstowego do edycji nowych nazw
$newNameTextBox = New-Object System.Windows.Forms.TextBox
$newNameTextBox.Location = New-Object System.Drawing.Point(420, 470)
$newNameTextBox.Width = 400
$newNameTextBox.Height = 60
$newNameTextBox.Multiline = $true
$newNameTextBox.ScrollBars = "Vertical"
$form.Controls.Add($newNameTextBox)

# Obsługa zdarzenia kliknięcia elementu w celu umożliwienia edycji
$listview.Add_Click({
    foreach ($item in $listview.Items) {
        $item.BackColor = [System.Drawing.Color]::White
        $item.ForeColor = [System.Drawing.Color]::Black
    }

    if ($listview.SelectedItems.Count -gt 0) {
        $originalNameTextBox.Text = ""
        $newNameTextBox.Text = ""
        foreach ($selectedItem in $listview.SelectedItems) {
            $originalNameTextBox.Text += "$($selectedItem.Text)`r`n"
            $newNameTextBox.Text += "$($selectedItem.SubItems[1].Text)`r`n"
        }
        $originalNameTextBox.Text = $originalNameTextBox.Text.TrimEnd("`r`n")
        $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
        $originalNameTextBox.Visible = $true
        $newNameTextBox.Visible = $true

        foreach ($item in $listview.SelectedItems) {
            $item.BackColor = [System.Drawing.Color]::LightBlue
            $item.ForeColor = [System.Drawing.Color]::Black
        }
    } else {
        $originalNameTextBox.Text = ""
        $newNameTextBox.Text = ""
        $originalNameTextBox.Visible = $false
        $newNameTextBox.Visible = $false
    }
})

# Obsługa zdarzenia zmiany tekstu w polu nowych nazw
$newNameTextBox.Add_TextChanged({
    $newNames = $newNameTextBox.Text -split "`r`n"
    $selectedItems = $listview.SelectedItems

    if ($newNames.Count -eq $selectedItems.Count) {
        for ($i = 0; $i -lt $selectedItems.Count; $i++) {
            $newName = $newNames[$i].Trim()
            if (-not [string]::IsNullOrWhiteSpace($newName)) {
                $selectedItems[$i].SubItems[1].Text = $newName
            }
        }
    }
})

# Obsługa zdarzenia opuszczenia TextBox w celu zapisania nowych nazw
$newNameTextBox.Add_Leave({
    $newNames = $newNameTextBox.Text -split "`r`n"
    $selectedItems = $listview.SelectedItems

    if ($newNames.Count -ne $selectedItems.Count) {
        [System.Windows.Forms.MessageBox]::Show("Liczba nowych nazw musi odpowiadać liczbie wybranych elementów.")
        return
    }

    for ($i = 0; $i -lt $selectedItems.Count; $i++) {
        $newName = $newNames[$i].Trim()
        if (-not [string]::IsNullOrWhiteSpace($newName)) {
            $selectedItems[$i].SubItems[1].Text = $newName
        }
    }
})

# Tworzenie przycisku "Przeglądaj folder"
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Przeglądaj folder"
$browseButton.Size = New-Object System.Drawing.Size(150, 30)
$browseButton.Location = New-Object System.Drawing.Point(10, 590) 
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $folderBrowser.ValidateNames = $false
    $folderBrowser.CheckFileExists = $false
    $folderBrowser.CheckPathExists = $true
    $folderBrowser.FileName = "Folder Selection"
    $folderBrowser.Filter = "Folders|no_files"
    $folderBrowser.Title = "Wybierz folder"
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:folderPath = [System.IO.Path]::GetDirectoryName($folderBrowser.FileName)
        Get-Files -folderPath $global:folderPath
    }
})
$form.Controls.Add($browseButton)

# Dodaj ComboBox dla predefiniowanych wzorców regex
$regexComboBox = New-Object System.Windows.Forms.ComboBox
$regexComboBox.Location = New-Object System.Drawing.Point(10, 630)
$regexComboBox.Width = 400
$regexComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$regexComboBox.Items.AddRange(@(
    "Zamiana cyfr na litery",
    "Usuwanie białych znaków",
    "Usuwanie specjalnych znaków",
    "Zamiana małych liter na wielkie",
    "Własny wzór"
    "S0XEY"
    "Usuń kropki (z wyjątkiem rozszerzenia)"
    "Dodaj tytuł"
    "S0XE0Y"
))
$regexComboBox.SelectedIndex = 4 # Default to "Własny wzór"
$form.Controls.Add($regexComboBox)

# Modified ComboBox SelectedIndexChanged event
$regexComboBox.Add_SelectedIndexChanged({
    switch ($regexComboBox.SelectedIndex) {
        0 { # Replace numbers with letters
            $regexPatternTextBox.Text = "\d"
            $replacementTextBox.Text = "X"  # Default replacement
            $replacementTextBox.Enabled = $true  
            
            # Auto-apply the pattern to selected items
            if ($listview.SelectedItems.Count -gt 0) {
                $newNameTextBox.Text = ""
                foreach ($item in $listview.SelectedItems) {
                    $newName = [regex]::Replace($item.Text, "\d", $replacementTextBox.Text)
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                }
                $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
            }
        }
        1 { # Remove whitespaces
            $regexPatternTextBox.Text = "\s+"
            $replacementTextBox.Text = ""
            
            # Auto-apply the pattern to selected items
            if ($listview.SelectedItems.Count -gt 0) {
                $newNameTextBox.Text = ""
                foreach ($item in $listview.SelectedItems) {
                    $newName = [regex]::Replace($item.Text, "\s+", "")
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                }
                $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
            }
        }
        2 { # Remove special characters
            $regexPatternTextBox.Text = "[^a-zA-Z0-9\.]"  # Added \. to preserve file extension
            $replacementTextBox.Text = ""
            
            # Auto-apply the pattern to selected items
            if ($listview.SelectedItems.Count -gt 0) {
                $newNameTextBox.Text = ""
                foreach ($item in $listview.SelectedItems) {
                    $newName = [regex]::Replace($item.Text, "[^a-zA-Z0-9\.]", "")
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                }
                $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
            }
        }
        3 { # Convert to uppercase
            $regexPatternTextBox.Text = ".*"
            $replacementTextBox.Text = ""
            
            # Auto-apply the pattern to selected items
            if ($listview.SelectedItems.Count -gt 0) {
                $newNameTextBox.Text = ""
                foreach ($item in $listview.SelectedItems) {
                    $newName = $item.Text.ToUpper()
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                }
                $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
            }
        }
        4 { # Custom pattern (enable textbox)
            $regexPatternTextBox.Enabled = $true
            $replacementTextBox.Enabled = $true
            $regexPatternTextBox.Text = "Wzór regex"
            $replacementTextBox.Text = "Zamiennik"
        }
        # First fix the preview logic in the switch statement (option 5):
        5 { # Title format to "S0XEY" for anime-style filenames
            # Updated regex pattern to handle anime-style filenames with brackets
            $regexPatternTextBox.Text = '^\[.*?\]\s*(.*?)\s*-\s*(\d+).*$'
            $replacementTextBox.Text = "1"  # Default season number
            $regexPatternTextBox.Enabled = $false
            $replacementTextBox.Enabled = $true
            
            if ($listview.SelectedItems.Count -gt 0) {
                $newNameTextBox.Text = ""
                
                foreach ($item in $listview.SelectedItems) {
                    $itemName = $item.Text
                    $match = [regex]::Match($itemName, $regexPatternTextBox.Text)
                    
                    if ($match.Success) {
                        # Extract components
                        $titlePart = $match.Groups[1].Value.Trim()  # Series name (Dragon Ball)
                        $episodeNum = $match.Groups[2].Value.PadLeft(2, '0')  # Episode number padded to 2 digits
                        $seasonNum = $replacementTextBox.Text.PadLeft(2, '0')  # Season number from user input
                        
                        # Get the original extension
                        $extension = [System.IO.Path]::GetExtension($itemName)
                        
                        # Construct the new name
                        $newName = "${titlePart} S${seasonNum}E${episodeNum}${extension}"
                        $item.SubItems[1].Text = $newName
                        $newNameTextBox.Text += "$newName`r`n"
                    } else {
                        # If pattern doesn't match, keep original name
                        $newNameTextBox.Text += "$($item.Text)`r`n"
                    }
                }
                $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
            }
        }

        6 { # Remove dots except for the extension dot
            $regexPatternTextBox.Text = "(?<=\w)\.(?=.*\.)" # Pattern to match dots except the last one
            $replacementTextBox.Text = " " # Replace matched dots with spaces

            # Auto-apply the pattern to selected items
            if ($listview.SelectedItems.Count -gt 0) {
                $newNameTextBox.Text = ""
                foreach ($item in $listview.SelectedItems) {
                    # Replace dots with spaces in the filename except the last one before the extension
                    $newName = [regex]::Replace($item.Text, "(?<=\w)\.(?=.*\.)", " ")
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                }
                $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
            }
        }
        7 { # Add custom title before filename
            $regexPatternTextBox.Text = "^(.*)$"  # Capture the entire filename
            $replacementTextBox.Enabled = $true
            
            # Auto-apply the pattern to selected items
            if ($listview.SelectedItems.Count -gt 0) {
                $newNameTextBox.Text = ""
                foreach ($item in $listview.SelectedItems) {
                    # Get the custom title from Zamiennik TextBox and add it before the filename
                    $customTitle = $replacementTextBox.Text.Trim()
                    $originalName = $item.Text  # Preserve the original filename
                    $newName = "$customTitle $originalName"  # Combine title with the original filename
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                }
                $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
            }
        }

        8 { # Title format to "S0XE0Y"
        $regexPatternTextBox.Text = "^(.*?)\s*(\d+)(\.[^.]+)?$"  # Updated regex pattern
        $replacementTextBox.Text = "1"  # Default season number
        $regexPatternTextBox.Enabled = $false  # Lock pattern
        $replacementTextBox.Enabled = $true   # Allow modifying the season part
    
        if ($listview.SelectedItems.Count -gt 0) {
            $newNameTextBox.Text = ""
    
            foreach ($item in $listview.SelectedItems) {
                $itemName = $item.Text
                $match = [regex]::Match($itemName, $regexPatternTextBox.Text)
                if ($match.Success) {
                    $titlePart = $match.Groups[1].Value.Trim()
                    $episodeNum = $match.Groups[2].Value.PadLeft(2, '0')
                    $seasonNum = $replacementTextBox.Text.PadLeft(2, '0')
                    $extension = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { "" }
                    
                    $newName = "S${seasonNum}E${episodeNum} ${titlePart}${extension}"
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                } else {
                    $newNameTextBox.Text += "$($item.Text)`r`n"
                }
            }
            $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
        }
    }
            } # End of switch statement
        }) # End of Add_SelectedIndexChanged event handler


# Obsługa zdarzenia zmiany tekstu w polu zamiennika
$replacementTextBox.Add_TextChanged({
    if ($regexCheckbox.Checked -and $listview.SelectedItems.Count -gt 0) {
        $newNameTextBox.Text = ""
        
        foreach ($item in $listview.SelectedItems) {
            $itemName = $item.Text
            
            if ($regexComboBox.SelectedIndex -eq 5) {
                $match = [regex]::Match($itemName, '^\[.*?\]\s*(.*?)\s*-\s*(\d+).*$')
                
                if ($match.Success) {
                    $titlePart = $match.Groups[1].Value.Trim()
                    $episodeNum = $match.Groups[2].Value.PadLeft(2, '0')
                    $extension = [System.IO.Path]::GetExtension($itemName)
                    $seasonNum = $replacementTextBox.Text.PadLeft(2, '0')
                    
                    $newName = "${titlePart} S${seasonNum}E${episodeNum}${extension}"
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                } else {
                    $newNameTextBox.Text += "$($item.Text)`r`n"
                }
            }
            if ($regexComboBox.SelectedIndex -eq 8) { # S0XE0Y pattern
                $match = [regex]::Match($itemName, "^(.*?)\s*(\d+)(\.[^.]+)?$")
                if ($match.Success) {
                    $titlePart = $match.Groups[1].Value.Trim()
                    $episodeNum = $match.Groups[2].Value.PadLeft(2, '0')
                    $seasonNum = $replacementTextBox.Text.PadLeft(2, '0')
                    $extension = if ($match.Groups[3].Success) { $match.Groups[3].Value } else { "" }
                    
                    $newName = "S${seasonNum}E${episodeNum} ${titlePart}${extension}"
                    $item.SubItems[1].Text = $newName
                    $newNameTextBox.Text += "$newName`r`n"
                } else {
                    $newNameTextBox.Text += "$($item.Text)`r`n"
                }
            }
            elseif ($regexComboBox.SelectedIndex -eq 7) {
                $customTitle = $replacementTextBox.Text.Trim()
                $newName = "$customTitle $($item.Text)"
                $item.SubItems[1].Text = $newName
                $newNameTextBox.Text += "$newName`r`n"
            }
        }
        $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
    }
})
# Upewnij się, że pola tekstowe wzorca wyrażenia regularnego są prawidłowo włączone/wyłączone
$regexCheckbox.Add_CheckedChanged({
    if ($regexCheckbox.Checked) {
        $regexPatternTextBox.Enabled = $true
        $replacementTextBox.Enabled = $true
        $regexComboBox.Enabled = $true
    } else {
        $regexPatternTextBox.Enabled = $false
        $replacementTextBox.Enabled = $false
        $regexComboBox.Enabled = $false
    }
})

# Zapewnienie domyślnego stanu składników wyrażenia regularnego
$regexComboBox.Enabled = $false
$regexPatternTextBox.Enabled = $false
$replacementTextBox.Enabled = $false

# Tworzenie przycisku "Odśwież"
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Odśwież"
$refreshButton.Size = New-Object System.Drawing.Size(150, 30)
$refreshButton.Location = New-Object System.Drawing.Point(170, 590) 
$refreshButton.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($global:folderPath)) {
        Get-Files -folderPath $global:folderPath
        $originalNameTextBox.Text = ""
        $newNameTextBox.Text = ""
        $originalNameTextBox.Visible = $false
        $newNameTextBox.Visible = $false
    } else {
        [System.Windows.Forms.MessageBox]::Show("Proszę najpierw wybrać folder.")
    }
})
$form.Controls.Add($refreshButton)

# Tworzenie przycisku Zmień nazwy plików
$renameButton = New-Object System.Windows.Forms.Button
$renameButton.Text = "Zmień nazwy plików"
$renameButton.Size = New-Object System.Drawing.Size(150, 30)
$renameButton.Location = New-Object System.Drawing.Point(330, 590)
$renameButton.Add_Click({
    $newNames = $newNameTextBox.Text -split "`r`n"
    $selectedItems = $listview.SelectedItems

    if ($newNames.Count -ne $selectedItems.Count) {
        [System.Windows.Forms.MessageBox]::Show("Liczba nowych nazw musi odpowiadać liczbie wybranych elementów.")
        return
    }

    for ($i = 0; $i -lt $selectedItems.Count; $i++) {
        $newName = $newNames[$i].Trim()

        if (-not [string]::IsNullOrWhiteSpace($newName)) {
            $originalFilePath = Join-Path -Path $global:folderPath -ChildPath $selectedItems[$i].Text
            $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName

# Sprawdź, czy wyrażenie regularne jest włączone i zastosuj je
if ($regexCheckbox.Checked) {
    $regexPattern = $regexPatternTextBox.Text
    $replacement = $replacementTextBox.Text

    if (-not [string]::IsNullOrWhiteSpace($regexPattern)) {
        try {
            if ($regexComboBox.SelectedIndex -eq 5) { # S0XEY pattern
                $match = [regex]::Match($selectedItems[$i].Text, "^\[.*?\]\s*(.*?)\s*-\s*(\d+).*?(\.[^.]+)$")
                if ($match.Success) {
                    $titlePart = $match.Groups[1].Value.Trim()
                    $episodeNum = $match.Groups[2].Value.PadLeft(2, '0')
                    $extension = $match.Groups[3].Value
                    $seasonNum = $replacement.PadLeft(2, '0')
                    
                    # Construct new name
                    $newName = "${titlePart} S${seasonNum}E${episodeNum}${extension}"
                    $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName
                }
            }
            elseif ($regexComboBox.SelectedIndex -eq 7) { # Custom title pattern
                $customTitle = $replacementTextBox.Text.Trim()
                $originalName = $selectedItems[$i].Text
                $newName = "$customTitle $originalName"
                $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName
            }
            elseif ($regexComboBox.SelectedIndex -eq 8) { # S0XE0Y pattern
                $match = [regex]::Match($selectedItems[$i].Text, "^(.*?)(?:\s*-\s*)(\d+)(\.[^.]+)$")
                if ($match.Success) {
                    $titlePart = $match.Groups[1].Value.Trim()
                    $episodeNum = $match.Groups[2].Value.PadLeft(2, '0')
                    $seasonNum = $replacement.PadLeft(2, '0')
                    $extension = $match.Groups[3].Value

                    # Construct new name
                    $newName = "S${seasonNum}E${episodeNum} ${titlePart}${extension}"
                    $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName
                }
            }
            else {
                # Regular regex pattern handling
                $newName = [regex]::Replace($selectedItems[$i].Text, $regexPattern, $replacement)
                $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Wystąpił błąd w wyrażeniu regularnym: " + $_.Exception.Message)
            return
        }
    }
}
            # Handle duplicate filenames
            $counter = 1
            $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($newName)
            $extension = [System.IO.Path]::GetExtension($newName)
            
            while (Test-Path $newFilePath) {
                $newName = "{0} ({1}){2}" -f $fileNameWithoutExt, $counter, $extension
                $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName
                $counter++
            }

            try {
                # Zmień nazwę pliku
                Rename-Item -Path $originalFilePath -NewName $newName -ErrorAction Stop
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Błąd podczas zmiany nazwy pliku: " + $_.Exception.Message)
                return
            }
        }
    }

    # Odświeżenie listy plików po zmianie nazwy
    Get-Files -folderPath $global:folderPath
})
$form.Controls.Add($renameButton)

# Function to validate filename and return detailed error message
function Test-ValidFileName {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName
    )
    
    try {
        # Invalid characters in Windows filenames
        $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
        $invalidCharsFound = [System.Collections.Generic.List[string]]::new()
        
        # Check for invalid characters
        foreach ($char in $invalidChars) {
            if ($FileName.Contains($char)) {
                $charDescription = switch ($char) {
                    "`0" { "<null>" }
                    "`a" { "<bell>" }
                    "`b" { "<backspace>" }
                    "`t" { "<tab>" }
                    "`n" { "<newline>" }
                    "`v" { "<vertical tab>" }
                    "`f" { "<formfeed>" }
                    "`r" { "<carriage return>" }
                    " " { "<space>" }
                    default { 
                        if ([char]::IsControl($char)) {
                            "<control-{0}>" -f ([int][char]$char)
                        } else {
                            [string]$char
                        }
                    }
                }
                [void]$invalidCharsFound.Add($charDescription)
            }
        }
        
        # Check for reserved Windows filenames
        $reservedNames = @('CON','PRN','AUX','NUL','COM1','COM2','COM3','COM4',
                          'COM5','COM6','COM7','COM8','COM9','LPT1','LPT2',
                          'LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9')
        
        $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        
        if ($reservedNames -contains $nameWithoutExt.ToUpper()) {
            return @{
                IsValid = $false
                Error = "Nazwa '$nameWithoutExt' jest zastrzeżona w systemie Windows."
            }
        }
        
        # Check filename length (using Windows MAX_PATH constant)
        if ($FileName.Length -gt 255) {
            return @{
                IsValid = $false
                Error = "Nazwa pliku jest za długa. Maksymalna długość to 255 znaków, aktualna długość: $($FileName.Length)"
            }
        }
        
        # Check if filename is empty or whitespace after trimming
        if ([string]::IsNullOrWhiteSpace($FileName.Trim())) {
            return @{
                IsValid = $false
                Error = "Nazwa pliku nie może być pusta."
            }
        }
        
        # Check if filename ends with period or space
        if ($FileName.EndsWith('.') -or $FileName.EndsWith(' ')) {
            return @{
                IsValid = $false
                Error = "Nazwa pliku nie może kończyć się kropką ani spacją."
            }
        }
        
        if ($invalidCharsFound.Count -gt 0) {
            return @{
                IsValid = $false
                Error = "Znaleziono niedozwolone znaki: $($invalidCharsFound -join ', ')"
            }
        }
        
        return @{
            IsValid = $true
            Error = $null
        }
    }
    catch {
        return @{
            IsValid = $false
            Error = "Wystąpił nieoczekiwany błąd podczas walidacji nazwy pliku: $($_.Exception.Message)"
        }
    }
}

# Rename button click handler
$renameButton.Add_Click({
    try {
        $newNames = $newNameTextBox.Text -split "`r`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $selectedItems = @($listview.SelectedItems)
        
        if ($selectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Proszę wybrać pliki do zmiany nazwy.",
                "Brak wybranych plików",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        if ($newNames.Count -ne $selectedItems.Count) {
            [System.Windows.Forms.MessageBox]::Show(
                "Liczba nowych nazw ($($newNames.Count)) musi odpowiadać liczbie wybranych elementów ($($selectedItems.Count)).",
                "Niezgodna liczba elementów",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # Create progress form
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Postęp zmiany nazw"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $progressForm.ControlBox = $false
        
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(360, 23)
        $progressBar.Location = New-Object System.Drawing.Point(10, 40)
        $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $progressBar.Maximum = $selectedItems.Count
        
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Size = New-Object System.Drawing.Size(360, 20)
        $progressLabel.Location = New-Object System.Drawing.Point(10, 10)
        $progressLabel.AutoSize = $false
        
        $progressForm.Controls.AddRange(@($progressBar, $progressLabel))
        
        # Initialize counters and error collection
        $errorMessages = [System.Collections.Generic.List[string]]::new()
        $successCount = 0
        
        # Show progress form
        $progressForm.Show()
        $progressForm.Refresh()

        # Process each file
        for ($i = 0; $i -lt $selectedItems.Count; $i++) {
            $newName = $newNames[$i].Trim()
            $progressBar.Value = $i
            $progressLabel.Text = "Przetwarzanie: $($i + 1) z $($selectedItems.Count)"
            $progressForm.Refresh()

            if (-not [string]::IsNullOrWhiteSpace($newName)) {
                # Validate new filename
                $validation = Test-ValidFileName -FileName $newName
                if (-not $validation.IsValid) {
                    $errorMessages.Add("Błąd dla pliku '$($selectedItems[$i].Text)': $($validation.Error)")
                    continue
                }

                $originalFilePath = Join-Path -Path $global:folderPath -ChildPath $selectedItems[$i].Text
                $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName

                try {
                    # Handle duplicate filenames
                    $counter = 1
                    $fileNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($newName)
                    $extension = [System.IO.Path]::GetExtension($newName)
                    
                    while (Test-Path -LiteralPath $newFilePath) {
                        $newName = "{0} ({1}){2}" -f $fileNameWithoutExt, $counter, $extension
                        $newFilePath = Join-Path -Path $global:folderPath -ChildPath $newName
                        $counter++
                    }

                    Rename-Item -LiteralPath $originalFilePath -NewName $newName -ErrorAction Stop
                    $successCount++
                }
                catch {
                    $errorMessages.Add("Błąd dla pliku '$($selectedItems[$i].Text)': $($_.Exception.Message)")
                }
            }
        }

        $progressForm.Close()
        $progressForm.Dispose()

        # Show summary message
        $summary = "Pomyślnie zmieniono nazwy $successCount plików.`n`n"
        if ($errorMessages.Count -gt 0) {
            $summary += "Wystąpiły błędy podczas zmiany nazw następujących plików:`n"
            $summary += $errorMessages -join "`n"
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            $summary,
            "Podsumowanie zmian nazw", 
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $(if ($errorMessages.Count -gt 0) {
                [System.Windows.Forms.MessageBoxIcon]::Warning
            } else {
                [System.Windows.Forms.MessageBoxIcon]::Information
            })
        )

        # Refresh the file list
        if (Test-Path -LiteralPath $global:folderPath) {
            Get-Files -folderPath $global:folderPath
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Wystąpił nieoczekiwany błąd: $($_.Exception.Message)",
            "Błąd",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        if ($progressForm -and -not $progressForm.IsDisposed) {
            $progressForm.Close()
            $progressForm.Dispose()
        }
    }
})

# Utwórz przycisk do wczytywania nazw z pliku .txt
$loadNamesButton = New-Object System.Windows.Forms.Button
$loadNamesButton.Text = "Załaduj nowe nazwy z pliku"
$loadNamesButton.Size = New-Object System.Drawing.Size(150, 30)
$loadNamesButton.Location = New-Object System.Drawing.Point(490, 590)  
$loadNamesButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $openFileDialog.FileName
        $newNames = Get-Content -Path $filePath
        
        # Sprawdź, czy zaznaczone są jakieś pliki
        if ($listview.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Proszę najpierw zaznaczyć pliki do zmiany nazw.")
            return
        }

        # Sprawdź, czy liczba nazw pasuje do wybranych plików
        if ($newNames.Count -ne $listview.SelectedItems.Count) {
            [System.Windows.Forms.MessageBox]::Show("Liczba nazw w pliku ($($newNames.Count)) nie odpowiada liczbie zaznaczonych plików ($($listview.SelectedItems.Count)).")
            return
        }

        # Aktualizacja elementów newNameTextBox i ListView z zachowaniem rozszerzeń
        $newNameTextBox.Text = ""
        for ($i = 0; $i -lt $listview.SelectedItems.Count; $i++) {
            $originalExtension = [System.IO.Path]::GetExtension($listview.SelectedItems[$i].Text)
            $newNameWithExtension = $newNames[$i] + $originalExtension
            
            $newNameTextBox.AppendText("$newNameWithExtension`r`n")
            $listview.SelectedItems[$i].SubItems[1].Text = $newNameWithExtension
        }
        $newNameTextBox.Text = $newNameTextBox.Text.TrimEnd("`r`n")
    }
})
$form.Controls.Add($loadNamesButton)

# Utwórz przycisk wyjścia z aplikacji
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Wyjście"
$exitButton.Size = New-Object System.Drawing.Size(150, 30)
$exitButton.Location = New-Object System.Drawing.Point(650, 590) 
$exitButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($exitButton)

# Switch-Theme function to toggle between light and dark themes
function Switch-Theme {
    $script:darkTheme = -not $script:darkTheme
    if (-not $script:darkTheme) {
        # Switch to Light Theme
        Set-LightTheme
        $lightBulbButton.Text = "💡"  # Lightbulb emoji
    } else {
        # Switch to Dark Theme
        Set-DarkTheme
        $lightBulbButton.Text = "🔆"  # Sun emoji for filled lightbulb
    }
}

# Create the Lightbulb Button
$lightBulbButton = New-Object System.Windows.Forms.Button
$lightBulbButton.Width = 30
$lightBulbButton.Height = 30
$lightBulbButton.Text = "💡"  # Unicode lightbulb emoji
$lightBulbButton.Font = New-Object System.Drawing.Font("Segoe UI Emoji", 12)
$lightBulbButton.Location = New-Object System.Drawing.Point(790, 540)  # Adjust location as needed
$lightBulbButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$lightBulbButton.FlatAppearance.BorderSize = 0
$form.Controls.Add($lightBulbButton)

# Bind the button click to the Switch-Theme function
$lightBulbButton.Add_Click({
    Switch-Theme
})

# Bind the button click to the Switch-Theme function
$lightBulbButton.Add_Click({
    Switch-Theme
})

# Bind the button click to the Switch-Theme function
$lightBulbButton.Add_Click({
    Switch-Theme
})

# Utwórz pole wyboru "Zaznacz wszystkie pliki"
$selectAllCheckbox = New-Object System.Windows.Forms.CheckBox
$selectAllCheckbox.Text = "Zaznacz wszystkie pliki"
$selectAllCheckbox.Location = New-Object System.Drawing.Point(650, 570)
$selectAllCheckbox.Width = 150
$selectAllCheckbox.Add_CheckedChanged({
    if ($selectAllCheckbox.Checked) {
        foreach ($item in $listview.Items) {
            $item.Selected = $true
            $item.BackColor = [System.Drawing.Color]::LightBlue
            $item.ForeColor = [System.Drawing.Color]::Black
        }
    } else {
        foreach ($item in $listview.Items) {
            $item.Selected = $false
            $item.BackColor = [System.Drawing.Color]::White
            $item.ForeColor = [System.Drawing.Color]::Black
        }
    }
})
$form.Controls.Add($selectAllCheckbox)

# Apply the default theme (light)
Set-LightTheme

# Show the form
$form.ShowDialog()