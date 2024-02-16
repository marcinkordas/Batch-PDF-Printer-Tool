Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'PDF Printer'
$form.Size = New-Object System.Drawing.Size(500,450)
$form.StartPosition = 'CenterScreen'

# Folder Browser
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Location = New-Object System.Drawing.Point(10,10)
$browseButton.Size = New-Object System.Drawing.Size(75,23)
$browseButton.Text = 'Browse...'
$browseButton.Add_Click({
    $result = $folderBrowser.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathTextBox.Text = $folderBrowser.SelectedPath
        Update-FileList $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButton)

# Path TextBox
$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(90,12)
$pathTextBox.Size = New-Object System.Drawing.Size(290,20)
$form.Controls.Add($pathTextBox)

# Update File List Function
function Update-FileList($folderPath) {
    $checkedListBox.Items.Clear()
    $pdfFiles = Get-ChildItem -Path $folderPath -Filter "*.pdf"
    foreach ($file in $pdfFiles) {
        $checkedListBox.Items.Add($file.Name, $true)
    }
}

# CheckedListBox for PDF files
$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
$checkedListBox.Location = New-Object System.Drawing.Point(10,70)
$checkedListBox.Size = New-Object System.Drawing.Size(470,200)
$checkedListBox.CheckOnClick = $true
$form.Controls.Add($checkedListBox)

# Select All Button
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Location = New-Object System.Drawing.Point(10,280)
$selectAllButton.Size = New-Object System.Drawing.Size(75,23)
$selectAllButton.Text = 'Select All'
$selectAllButton.Add_Click({
    for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
        $checkedListBox.SetItemChecked($i, $true)
    }
})
$form.Controls.Add($selectAllButton)

# Unselect All Button
$unselectAllButton = New-Object System.Windows.Forms.Button
$unselectAllButton.Location = New-Object System.Drawing.Point(90,280)
$unselectAllButton.Size = New-Object System.Drawing.Size(75,23)
$unselectAllButton.Text = 'Unselect All'
$unselectAllButton.Add_Click({
    for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
        $checkedListBox.SetItemChecked($i, $false)
    }
})
$form.Controls.Add($unselectAllButton)



# Remaining GUI elements (Number of Copies, Color Mode, Printer Selection, Print Button)
# Number of Copies
$copiesLabel = New-Object System.Windows.Forms.Label
$copiesLabel.Location = New-Object System.Drawing.Point(10,40)
$copiesLabel.Size = New-Object System.Drawing.Size(100,20)
$copiesLabel.Text = 'Number of Copies:'
$form.Controls.Add($copiesLabel)

$copiesTextBox = New-Object System.Windows.Forms.TextBox
$copiesTextBox.Location = New-Object System.Drawing.Point(110,40)
$copiesTextBox.Size = New-Object System.Drawing.Size(50,20)
$copiesTextBox.Text = '1'
$form.Controls.Add($copiesTextBox)


# Color Mode
$colorModeLabel = New-Object System.Windows.Forms.Label
$colorModeLabel.Location = New-Object System.Drawing.Point(10,320)
$colorModeLabel.Size = New-Object System.Drawing.Size(100,20)
$colorModeLabel.Text = 'Color Mode:'
$form.Controls.Add($colorModeLabel)

$colorModeComboBox = New-Object System.Windows.Forms.ComboBox
$colorModeComboBox.Location = New-Object System.Drawing.Point(110,317)
$colorModeComboBox.Size = New-Object System.Drawing.Size(121,21)
$colorModeComboBox.DropDownStyle = 'DropDownList'
$colorModeComboBox.Items.Add('Color')
$colorModeComboBox.Items.Add('Grayscale')
$colorModeComboBox.Items.Add('BlackOnly')
$colorModeComboBox.SelectedIndex = 0
$form.Controls.Add($colorModeComboBox)

# Add a Progress Bar to the form
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 350)
#$progressBar.Size = New-Object System.Drawing.Size(470, 23)
# Adjust the Progress Bar for padding
$progressBar.Size = New-Object System.Drawing.Size(460, 23) # Adjust width for padding
$form.Controls.Add($progressBar)

# Print Button with updated logic to print selected files only
$printButton = New-Object System.Windows.Forms.Button
$printButton.Location = New-Object System.Drawing.Point(10,380)#10)
$printButton.Size = New-Object System.Drawing.Size(75,23)
$printButton.Text = 'Print'



# Assuming other parts of the script are unchanged and the printer selection has been removed

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000 # 1 second



# Assuming initialization of $progressBar and other controls are done above

$printButton.Add_Click({
    $selectedFiles = @()
    for ($i = 0; $i -lt $checkedListBox.CheckedItems.Count; $i++) {
        $selectedFiles += $checkedListBox.CheckedItems[$i]
    }
    $folderPath = $pathTextBox.Text
    $numberOfCopies = [int]$copiesTextBox.Text

    $progressBar.Maximum = $selectedFiles.Count * $numberOfCopies
    $progressBar.Value = 0

    # Start background job for printing
    $job = Start-Job -ScriptBlock {
        param($folderPath, $selectedFiles, $numberOfCopies)
        $progress = 0
        foreach ($fileName in $selectedFiles) {
            $fullPath = Join-Path -Path $folderPath -ChildPath $fileName
            for ($copyIndex = 0; $copyIndex -lt $numberOfCopies; $copyIndex++) {
                Start-Process -FilePath $fullPath -Verb Print -PassThru -Wait
                $progress += 1
                # Ideally, update a shared variable or file with $progress
            }
        }
        # Return final progress to indicate job completion
        return $progress
    } -ArgumentList $folderPath, $selectedFiles, $numberOfCopies

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000 # 1 second
    $timer.Add_Tick({
        if ($job -ne $null -and $job.Id -ne $null) {
            $jobRefresh = Get-Job -Id $job.Id
            if ($jobRefresh.State -eq 'Completed') {
                $finalProgress = Receive-Job -Job $jobRefresh
                if ($finalProgress -eq $progressBar.Maximum) {
                    $progressBar.Value = $progressBar.Maximum
                }
                $timer.Stop()
                $timer.Dispose()
                Remove-Job -Job $jobRefresh
                # Optionally, notify the user of completion
            } else {
                # Increment the progress bar to indicate ongoing activity
                # Note: This does not reflect actual progress, adjust as needed
                $progressBar.Value = [Math]::Min($progressBar.Value + 1, $progressBar.Maximum)
            }
        }
    })
    $timer.Start()
})

# Example logic to start the timer, ensuring it's only started after being properly initialized


#$form.ShowDialog()

# Correct placement of the Print button addition
# No changes needed here based on your script, but ensure it's added before the first $form.ShowDialog()



# Ensure the Print button is added before the form is shown
$form.Controls.Add($printButton)

# Single call to show the form
$form.ShowDialog()