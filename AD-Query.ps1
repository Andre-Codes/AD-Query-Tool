#####################
####PREREQUISITES####

<#

1. Active Directory:
    1a. https://learn.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps
    1b. import-module ActiveDirectory

2. 

#>

# Get the directory where the script is located
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Change the working directory to the script directory
Set-Location $scriptPath

#Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

# Create a new form
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'AD Searcher | User Account Tools'
$form_darkOFFbackColor = "white"
$form_darkONbackColor = "DarkSlateGray"
$form.BackColor = "$form_darkOFFbackColor"
$form_DarkOnForeColor = "white"
$form_DarkOFFforeColor= "black"
$formWidth = 280
$formHeight = 600
$formWidth_Expanded = 1225
$form.Width = $formWidth
$form.Height = $formHeight
$form.StartPosition = "CenterScreen"

####################
#region Property Boxes
####################
# Create labels and text boxes for each property
$labelName = New-Object System.Windows.Forms.Label
$labelName.Text = 'Username:'
$labelName.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($labelName)

$textBoxName = New-Object System.Windows.Forms.TextBox
$textBoxName.Location = New-Object System.Drawing.Point(110, 10)
$textBoxName.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Search-AD
    }
})
$form.Controls.Add($textBoxName)


$labelDepartment = New-Object System.Windows.Forms.Label
$labelDepartment.Text = 'Department:'
$labelDepartment.Location = New-Object System.Drawing.Point(10, 40)
$form.Controls.Add($labelDepartment)

$textBoxDepartment = New-Object System.Windows.Forms.TextBox
$textBoxDepartment.Location = New-Object System.Drawing.Point(110, 40)
$textBoxDepartment.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Search-AD
    }
})
$form.Controls.Add($textBoxDepartment)

$labelEmailAddress = New-Object System.Windows.Forms.Label
$labelEmailAddress.Text = 'Email:'
$labelEmailAddress.Location = New-Object System.Drawing.Point(10, 70)
$form.Controls.Add($labelEmailAddress)

$textBoxEmailAddress = New-Object System.Windows.Forms.TextBox
$textBoxEmailAddress.Location = New-Object System.Drawing.Point(110, 70)
$textBoxEmailAddress.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Search-AD
    }
})
$form.Controls.Add($textBoxEmailAddress)

$labelDisplayName = New-Object System.Windows.Forms.Label
$labelDisplayName.Text = 'Last, First:'
$labelDisplayName.Location = New-Object System.Drawing.Point(10, 100)
$form.Controls.Add($labelDisplayName)

$textBoxDisplayName = New-Object System.Windows.Forms.TextBox
$textBoxDisplayName.Location = New-Object System.Drawing.Point(110, 100)
$textBoxDisplayName.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Search-AD
    }
})
$form.Controls.Add($textBoxDisplayName)

$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = 'Job Title:'
$labelTitle.Location = New-Object System.Drawing.Point(10, 130)
$form.Controls.Add($labelTitle)

$textBoxTitle = New-Object System.Windows.Forms.TextBox
$textBoxTitle.Location = New-Object System.Drawing.Point(110, 130)
$textBoxTitle.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Search-AD
    }
})
$form.Controls.Add($textBoxTitle)

$labelMobile = New-Object System.Windows.Forms.Label
$labelMobile.Text = 'Mobile #:'
$labelMobile.Location = New-Object System.Drawing.Point(10, 160)
$form.Controls.Add($labelMobile)

$textBoxMobile = New-Object System.Windows.Forms.TextBox
$textBoxMobile.Location = New-Object System.Drawing.Point(110, 160)
$textBoxMobile.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Search-AD
    }
})
$form.Controls.Add($textBoxMobile)

$labelTelephoneNumber = New-Object System.Windows.Forms.Label
$labelTelephoneNumber.Text = 'Office #:'
$labelTelephoneNumber.Location = New-Object System.Drawing.Point(10, 190)
$form.Controls.Add($labelTelephoneNumber)

$textBoxTelephoneNumber = New-Object System.Windows.Forms.TextBox
$textBoxTelephoneNumber.Location = New-Object System.Drawing.Point(110, 190)
$textBoxTelephoneNumber.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        Search-AD
    }
})
$form.Controls.Add($textBoxTelephoneNumber)

# Change ALL PROPERTY BOX LABEL properties
#Change font
foreach ($Control in $Form.Controls) {
    if ($Control.GetType().Name -eq "Label") {
        $Control.Font = 'Segoe UI,11'
    }
}

#endregion

####################
#region Results box
####################
# Create a text box for displaying the results
$textBoxResults = New-Object System.Windows.Forms.ListView
$textBoxResults.Location = New-Object System.Drawing.Point(220, 10)
$textBoxResults.Size = New-Object System.Drawing.Size(975, 490)
$textBoxResults.Font = New-Object System.Drawing.Font("Lucidia", 10)
$textBoxResults.BorderStyle = 'Fixed3D'
$textBoxResults_DarkONbackcolor = "DarkGray"
$textBoxResults_DarkOffbackcolor = "GhostWhite"
$textBoxResults.BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor
$textBoxResults_DarkONforeColor = "white"
$textBoxResults_DarkOFFforeColor = "black"
$textBoxResults.ForeColor = [System.Drawing.Color]::$textBoxResults_DarkOFFforeColor
$textBoxResults.View = [System.Windows.Forms.View]::Details
$textBoxResults.FullRowSelect = $true
$textBoxResults.MultiSelect = $true
$textBoxResults.Visible = $false
$textBoxResults.GridLines = $true
$textBoxResults.AllowColumnReorder = $true

$textBoxResults.ShowItemToolTips = $true

#List Box Property COLUMNS
$textBoxResults.Columns.Add("Full Name", 150) | Out-Null
$textBoxResults.Columns.Add("Username", 100) | Out-Null
$textBoxResults.Columns.Add("Department", 150) | Out-Null
$textBoxResults.Columns.Add("Email", 175) | Out-Null
$textBoxResults.Columns.Add("Job Title", 175) | Out-Null
$textBoxResults.Columns.Add("Mobile Phone", 100) | Out-Null
$textBoxResults.Columns.Add("Office Phone", 100) | Out-Null

$enabledProperty_backcolor = "LightPink"


$form.Controls.Add($textBoxResults)
#endregion

####################
#region Divider lines
####################
# Divider between properties and search/sort
$panelDivider = New-Object System.Windows.Forms.Panel
$panelDivider.Location = New-Object System.Drawing.Point(10, 250)
$panelDivider.Size = New-Object System.Drawing.Size(200, 1)
$panelDivider.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($panelDivider)

# Divider between search/sort and filter section
$panelDivider2 = New-Object System.Windows.Forms.Panel
$panelDivider2.Location = New-Object System.Drawing.Point(10, 345)
$panelDivider2.Size = New-Object System.Drawing.Size(200, 1)
$panelDivider2.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($panelDivider2)
#endregion

####################
#region Options and Controls
####################

########################
#SORY BY LIST BOX ###
########################
#Create a list box for selecting the sort property
$labelSort = New-Object System.Windows.Forms.Label
$labelSort.Text = 'Sort by:'
$labelSort.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$labelSort.Location = New-Object System.Drawing.Point(10, 260)
$form.Controls.Add($labelSort)

$comboBoxSort = New-Object System.Windows.Forms.ComboBox
$comboBoxSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxSort.Location = New-Object System.Drawing.Point(110, 260)
$comboBoxSort.Width = 100
$comboBoxSort.Items.Add('Full Name')
$comboBoxSort.Items.Add('Username')
$comboBoxSort.Items.Add('Department')
$comboBoxSort.Items.Add('Email')
$comboBoxSort.Items.Add('Job Title')
$comboBoxSort.Items.Add('Mobile Phone')
$comboBoxSort.Items.Add('Office Phone')
$comboBoxSort.Items.Add('Enabled')
$comboBoxSort.Enabled = $true

# Populate list
#Load-ComboBox $comboBoxSort 'Full Name','Username','Department','Email','Job Title','Mobile Phone','Office Phone','Enabled'
$comboBoxSort.SelectedIndex = 0
$form.Controls.Add($comboBoxSort)

#Event handlers for ComboBox
$comboBoxSort.add_SelectedIndexChanged({
    $sortProperty = $comboBoxSort.SelectedItem.ToString()
})
########################

####################
### SEARCH BUTTON ###
# Create a button for performing the search
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = 'Search'
$buttonSearch.Font             = 'Microsoft Sans Serif,10'
$buttonSearch.FlatStyle        = 'Popup'
$buttonSearch.FlatAppearance.BorderSize = 1
$buttonSearch.BackColor        = '#FFAFDAFF'
$buttonSearch.UseVisualStyleBackColor   = $false
$buttonSearch.Size = New-Object System.Drawing.Size(90, 30)
$buttonSearch.Location = New-Object System.Drawing.Point(10, 295)

$form.Controls.Add($buttonSearch)

# Create a search results message label
$buttonSearchMessage = New-Object System.Windows.Forms.Label
$buttonSearchMessage.Location = New-Object System.Drawing.Point(10, 325)
$buttonSearchMessage.Size = New-Object System.Drawing.Size(300, 20)
$buttonSearchMessage.ForeColor = "red"
$form.Controls.Add($buttonSearchMessage)

#Event handlers for Search button
$buttonSearch.add_Click({
    Search-AD
})

#############################
### CLEAR RESULTS BUTTON ###
# Create a button for performing the search
$clearResults = New-Object System.Windows.Forms.Button
$clearResults.Text = 'Clear'
$clearResults.Font             = 'Microsoft Sans Serif,10'
$clearResults.FlatStyle        = 'Popup'
$clearResults.FlatAppearance.BorderSize = 1
$clearResults.BackColor        = '#FFAFDAFF'
$clearResults.UseVisualStyleBackColor   = $false
$clearResults.Size = New-Object System.Drawing.Size(90, 30)
$clearResults.Location = New-Object System.Drawing.Point(120, 295)
$form.Controls.Add($clearResults)

$clearResults.Add_Click({
    $textBoxResults.Items.Clear()
    #Restore smaller window size
    Set-DefaultFormProperties
})

#############################
#   FILTER BOX


#############################
###       StatusStrip     ###
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = "GhostWhite"
$form.Controls.Add($statusStrip)

# Create a Label in the StatusStrip
$statusStriplabel_filter = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusStriplabel_filter.Text = "Filter:"
$statusStrip.Items.Add($statusStriplabel_filter)

$textbox_filter = New-Object System.Windows.Forms.ToolStripTextBox
$textbox_filter.Size = New-Object System.Drawing.Size(80, 20)
$textbox_filter.Enabled = $false
$statusStrip.Items.Add($textbox_filter)

# Create a Label in the StatusStrip
$statusStripLabel_totalResult = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusStripLabel_totalResult.Text = "Total:"
$statusStripLabel_totalResult.BorderSides = "Left"
$statusStripLabel_totalResult.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$statusStrip.Items.Add($statusStripLabel_totalResult)

# Create a filter results message label
$statusStriplabel_filterMessage = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusStriplabel_filterMessage.ForeColor = "darkgreen"
$statusStriplabel_filterMessage.Font = New-Object System.Drawing.Font('Consolas', 12)
$statusStriplabel_filterMessage.BorderStyle = "Raised"
$statusStriplabel_filterMessage.BorderSides = 'Right'
$statusStrip.Items.Add($statusStriplabel_filterMessage)


$statusStripCheckbox_Dark = New-Object System.Windows.Forms.ToolStripButton
$statusStripCheckbox_Dark.Text = "Dark Mode"
$script:DarkModeStatus = $false #default theme mode
$darkModeON_foreColor = "White"
$darkModeOFF_foreColor = "Black"
$darkModeON_backColor = "DarkSlateGray"
$darkModeOFF_backColor = "LightGray"
$statusStripCheckbox_Dark.BackColor = "$darkModeOFF_backColor"
$statusStripCheckbox_Dark.ForeColor = "$darkModeOFF_foreColor"
$statusStrip.Items.Add($statusStripCheckbox_Dark)

# Event handlers for the filter button/box
$textbox_filter.add_TextChanged({
    Filter-Results
    if ($textbox_filter.Text.Length -lt 1) {
            $statusStriplabel_filterMessage.Text = "$($globalResults.Count) users"
            $statusStriplabel_filterMessage.ForeColor = "darkgreen"
        }
})
# Create a TextBox in the StatusStrip


$statusStripCheckbox_Dark.add_Click({

    #if ($script:DarkModeStatus -eq $false) {$script:DarkModeStatus = $true}

    if ($script:DarkModeStatus -eq $false) {
        $script:darkModeStatus = $true
        $statusStripCheckbox_Dark.Text = "Light Mode"
        $form.BackColor = $form_darkONbackColor
        $statusStripCheckbox_Dark.BackColor = $darkModeON_backColor
        $statusStripCheckbox_Dark.ForeColor = $darkModeON_foreColor
        $textBoxResults.BackColor = $textBoxResults_DarkONbackcolor
        $script:textLabel_ForeColor = $form_DarkOnForeColor #used in separate 'user details' form

        foreach ($Control in $Form.Controls) {
            #Adjust forecolor label font
            if ($Control.GetType().Name -eq "Label") {
                $Control.ForeColor = $form_DarkOnForeColor
            }
        }
        return
    }

    if ($script:DarkModeStatus -eq $true) {
        $script:DarkModeStatus = $false
        $statusStripCheckbox_Dark.Text = "Dark Mode"
        $form.BackColor = $form_darkOFFbackColor
        $statusStripCheckbox_Dark.BackColor = $darkModeOFF_backColor
        $statusStripCheckbox_Dark.ForeColor = $darkModeOFF_foreColor
        $textBoxResults.BackColor = $textBoxResults_DarkOffbackcolor
        $script:textLabel_ForeColor = $form_DarkOFFforeColor #used in separate 'user details' form

        foreach ($Control in $Form.Controls) {
            if ($Control.GetType().Name -eq "Label") {
                $Control.ForeColor = $form_DarkOFFforeColor
            }
        }
        return
    }
    
})

#endregion



################################################

################################################
#region         CONTEXT MENUS
################################################
# Create a context menu
### MAIN CONTEXT MENU ITEMS ###
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
#Item 0
$contextMenu.Items.Add('Call Office')  | Out-Null
#Item 1
$contextMenu.Items.Add('Call Mobile')  | Out-Null
#Item 2
$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null
#Item 3
# Create variable for Email context item so it can be referenced and updated later
$emailContextItem = New-Object System.Windows.Forms.ToolStripMenuItem
$emailContextItem.Text = 'Send Email'
$contextMenu.Items.Add($emailContextItem)  | Out-Null
#Item 4
$contextMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

### COPY TOOL STRIP ITEMS ###
# Add a "Copy Property" menu item with sub-menu items for copying individual properties
$contextMenu_CopyProp = New-Object System.Windows.Forms.ToolStripMenuItem
$contextMenu_CopyProp.Text = 'Copy Property'

# Add the ItemSelectionChanged event handler to the textBoxResults control
$textBoxResults.Add_ItemSelectionChanged({
 Copy-Property -ListViewRow $textBoxResults.SelectedItems[0]
})


# Add the main context menu to the listview control
$textBoxResults.ContextMenuStrip = $contextMenu
# Add the toolstrip to the main context menu
$contextMenu.Items.Add($contextMenu_CopyProp)

#####################################
#region CALL PHONES EVENTS
#####################################

# OFFICE PHONE EVENT
$contextMenu.Items[0].Add_Click({

    # Save the selected item
    $selectedItems = $textBoxResults.SelectedItems

    $OfficeCopied = $SelectedItems.SubItems[6].Text

    if ($OfficeCopied -eq "---") {return}

    $jabberPath = "C:\Program Files (x86)\Cisco Systems\Cisco Jabber\ciscojabber.exe"
    # Check if Jabber exists, if not throw error and exit
    if (-not (Test-Path -Path $jabberPath)) {
    [System.Windows.Forms.MessageBox]::Show("Cisco Jabber is not installed", "App Not Found", `
    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
    }
    if ($selectedItems.Count -gt 1) {
        [System.Windows.Forms.MessageBox]::Show("Select only 1 user", "Selection Error", `
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    Set-Clipboard -Value $OfficeCopied

    # Start Jabber and wait for it to open
    $process = Start-Process -FilePath $jabberPath
    # Wait for the application window to appear
    while (-not $process.MainWindowHandle) {
        Start-Sleep -Milliseconds 100
    }

    # Activate and bring the window to the front
    $process.Refresh()
    $process.MainWindowHandle | Out-Null
    Start-Sleep -Seconds 1
    $process.MainWindowHandle | ForEach-Object { $_.Select() | Out-Null }

    #Send the Ctrl^0 command to select the search bar
    [Windows.Forms.SendKeys]::SendWait('^0')
    # Send the paste command
    [Windows.Forms.SendKeys]::SendWait('^v')
    # Send the Enter command
    [Windows.Forms.SendKeys]::SendWait("{ENTER}")

    $process.Dispose()

})

# MOBILE PHONE EVENT
$contextMenu.Items[1].Add_Click({

     # Save the selected item
     $selectedItems = $textBoxResults.SelectedItems

     $MobileCopied = $SelectedItems.SubItems[5].Text

     if ($MobileCopied -eq "---") {return}

    $jabberPath = "C:\Program Files (x86)\Cisco Systems\Cisco Jabber\ciscojabber.exe"
    # Check if Jabber exists, if not throw error and exit
    if (-not (Test-Path -Path $jabberPath)) {
    [System.Windows.Forms.MessageBox]::Show("Cisco Jabber is not installed", "App not found", `
    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return
    }

    if ($selectedItems.Count -gt 1) {
        [System.Windows.Forms.MessageBox]::Show("Select only 1 user", "Selection Error", `
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    Set-Clipboard -Value $MobileCopied

    # Start Jabber and wait for it to open
    Start-Process -FilePath $jabberPath
    Start-Sleep -Seconds 4

    #Send the Ctrl^0 command to select the search bar
    [Windows.Forms.SendKeys]::SendWait('^0')
    # Send the paste command
    [Windows.Forms.SendKeys]::SendWait('^v')
    # Send the Enter command
    [Windows.Forms.SendKeys]::SendWait("{ENTER}")

})
#endregion CALL PHONES EVENTS
################################################

#####################################
#region EMAIL EVENT
$contextMenu.Items[3].Add_Click({
    $selectedItems = $textBoxResults.SelectedItems
    foreach ($user in $selectedItems) {
        $emailAddress = $user.SubItems[3].Text

        $outlook = New-Object -ComObject Outlook.Application
        $mail = $Outlook.CreateItem(0)

        # Apply email address
        $mail.To = $emailAddress

        # Set the recipient and send the email
        $mail.Display()
                
    }
})
#endregion EMAIL EVENT
################################################



#############################
#region MOUSE / KEY EVENTS
#############################

$textBoxResults.Add_MouseDown({
    #Update the Send Email option to reflect num. of selected
    $selectedItemsCount = $textBoxResults.SelectedItems.Count
    if ($selectedItemsCount -lt 1) {$selectedItemsCount = "1"}
    $emailContextItem.text = "Send Email (" + $selectedItemsCount + ")"
})

# CTRL+C for copying rows
$textBoxResults.Add_KeyDown({
    Copy-EntireRow
})


$textBoxResults.Add_DoubleClick({

    CreateUserDetailsForm



})



# Add a handler for the main form's FormClosing event
$form.Add_FormClosed({
    # Iterate through all open forms and dispose of them
    foreach ($window in [System.Windows.Forms.Application]::OpenForms) {
        if ($window -ne $form) {
            $window.Dispose()
        }
    }
})

#endregion
#############################



########################################
#region         FUNCTIONS
########################################

function Copy-Property {
    param (
        $ListViewRow
    )
    # Clear the "Copy Property" drop-down button's items
    $contextMenu_CopyProp.DropDownItems.Clear()

    # Get the selected row
    $selectedRow = $ListViewRow

    # Add a menu item for each subitem in the selected row
    for ($i = 0; $i -lt $selectedRow.SubItems.Count; $i++) {
        $subItem = $selectedRow.SubItems[$i]
        $menuItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $menuItem.Text = $subItem.Text
        $menuItem.Tag = $i
        $menuItem.Add_Click({
            # Get the clicked menu item's Tag property (which is the subitem index)
            $subItemIndex = $args[0].Tag

            # Get the selected row
            $selectedRow = $ListViewRow

            # Get the selected subitem
            $selectedSubItem = $selectedRow.SubItems[$subItemIndex]

            # Copy the selected subitem to the clipboard
            Set-Clipboard -Value $selectedSubItem.Text
        }.GetNewClosure())
        $contextMenu_CopyProp.DropDownItems.Add($menuItem)
    }
}


function Copy-EntireRow {
    $selectedItems = $textBoxResults.SelectedItems
    if ($selectedItems.Count -gt 0) {
        $copiedText = ($textBoxResults.Columns.Text -join ',') + "`r`n"
        foreach ($item in $selectedItems) {
            $copiedText += $item.SubItems[0].Text + ',' + $item.SubItems[1].Text + ',' + $item.SubItems[2].Text + `
            ',' + $item.SubItems[3].Text +',' + $item.SubItems[4].Text +',' + $item.SubItems[5].Text + `
            ',' + $item.SubItems[6].Text + "`r`n"
        }
        Set-Clipboard -Value $copiedText
        #[System.Windows.Forms.MessageBox]::Show('Selected items copied to clipboard.', 'Copy Selected')
    }
}

# Define variables needed outside scope of functions
$globalResults = @()
$filterBy = @()
$resultsFiltered = @()
$sortProperty  = @()

####################
# Main form defaults
####################
function Set-DefaultFormProperties {
    $form.Width = $formWidth
    $form.Height = $formHeight
    $textBoxResults.Visible = $false
    $textbox_filter.Enabled = $false
    $textbox_filter.Clear()
    $textbox_filter.Enabled = $false
    $statusStriplabel_filterMessage.Visible = $false
    $statusStriplabel_filter.Enabled = $false
    #Clear all text boxes
    foreach ($control in $form.Controls) {
        if ($control.GetType().Name -like "TextBox*") {
            $control.Clear()
        }
    }
}

####################
# Filter function
####################
function Filter-Results {
    # Get the filter string from the filter text box
    $filterBy = $textbox_filter.Text
    # Get the sort property to be used as the object for filtering
    $sortProperty = $comboBoxSort.SelectedItem.ToString()
    # Set filter highlight color
    $filteredProperty_backcolor="LightBlue"

    # Apply the filter using the Where-Object cmdlet and update the results array
    # Using $_.psobject.Properties.Value will search ALL properties for the filter -
    #  Change it to $_."specific property" to filter by specific prop.
    $resultsFiltered = $globalResults | Where-Object { $_.psobject.Properties.Value -like "*$filterBy*" }

    # Update the display with the filtered results
    # clear original results
    $textBoxResults.Items.Clear()
    # Display new results with filter applied
    foreach ($result in $resultsFiltered) {

        if ($result.'Full Name') {
            $row = New-Object System.Windows.Forms.ListViewItem($result.'Full Name')
            if (($row.SubItems[0].Text -like "*$filterBy*") -and ($filterBy -ne "")) {
                $row.SubItems[0].BackColor = [System.Drawing.Color]::$filteredProperty_backcolor
            } else {$row.SubItems[0].ResetStyle()} #Previously set to: .BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor
        } else {$row = New-Object System.Windows.Forms.ListViewItem("---")}

        if ($result.'Username') {
            $row.SubItems.Add($result.'Username')
            if (($row.SubItems[1].Text -like "*$filterBy*") -and ($filterBy -ne "")) {
                $row.SubItems[1].BackColor = [System.Drawing.Color]::$filteredProperty_backcolor
            } else {$row.SubItems[1].ResetStyle()}
        } else {$row.SubItems.Add("---")}

        if ($result.'Department') {
            $row.SubItems.Add($result.'Department')
            if (($row.SubItems[2].Text -like "*$filterBy*") -and ($filterBy -ne "")) {
                $row.SubItems[2].BackColor = [System.Drawing.Color]::$filteredProperty_backcolor
            } else {$row.SubItems[2].ResetStyle()}
        } else {$row.SubItems.Add("---")}

        if ($result.'Email') {
            $row.SubItems.Add($result.'Email')
            if (($row.SubItems[3].Text -like "*$filterBy*") -and ($filterBy -ne "")) {
                $row.SubItems[3].BackColor = [System.Drawing.Color]::$filteredProperty_backcolor
            } else {$row.SubItems[3].ResetStyle()}
        } else {$row.SubItems.Add("---")}

        if ($result.'Job Title') {
            $row.SubItems.Add($result.'Job Title')
            if (($row.SubItems[4].Text -like "*$filterBy*") -and ($filterBy -ne "")) {
                $row.SubItems[4].BackColor = [System.Drawing.Color]::$filteredProperty_backcolor
            } else {$row.SubItems[4].ResetStyle()}
        } else {$row.SubItems.Add("---")}

        if ($result.'Mobile Phone') {
            $row.SubItems.Add($result.'Mobile Phone')
            if (($row.SubItems[5].Text -like "*$filterBy*") -and ($filterBy -ne "")) {
                $row.SubItems[5].BackColor = [System.Drawing.Color]::$filteredProperty_backcolor
            } else {$row.SubItems[5].ResetStyle()}
        } else {$row.SubItems.Add("---")}

        if ($result.'Office Phone') {
            $row.SubItems.Add($result.'Office Phone')
            if (($row.SubItems[6].Text -like "*$filterBy*") -and ($filterBy -ne "")) {
                $row.SubItems[6].BackColor = [System.Drawing.Color]::$filteredProperty_backcolor
            } else {$row.SubItems[6].ResetStyle()} 
        } else {$row.SubItems.Add("---")}

        
        # $row.SubItems.Add(($result.'Enabled').ToString())
        # if ($row.SubItems[7].Text -eq "False") {
        #     $row.SubItems[7].BackColor = [System.Drawing.Color]::$enabledProperty_backcolor
        # } else {$row.SubItems[7].BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor}
        
        $row.UseItemStyleForSubItems = $false
        $textBoxResults.Items.Add($row) | Out-Null
    }
    
    if ($resultsFiltered) {
        # Update the status label with the number of results
        $statusStriplabel_filterMessage.ForeColor = "darkgreen"
        $statusStriplabel_filterMessage.Text = "$($resultsFiltered.Count) users"
        if ($resultsFiltered.Count -lt 2) {$statusStriplabel_filterMessage.Text = "1 user"}
    } else {
        # If the results array is empty, display a message
        $statusStriplabel_filterMessage.ForeColor = "red"
        $statusStriplabel_filterMessage.Text = "No users"
    }
    if ($filterBy.Length -lt 1) {
        $statusStriplabel_filterMessage.Text = "$($resultsFiltered.Count) users"
        $statusStriplabel_filterMessage.ForeColor = "darkgreen"
    }
    
}
####################

####################
## SEARCH function
####################
function Search-AD {

    ### Adjust form and controls properties ###
    $textBoxResults.Items.Clear()
    # main form properties
    $form.Width = $formWidth_Expanded
    
    #Results box properties
    $textBoxResults.Visible = $true

    #Filter message/label properties
    $textbox_filter.Enabled = $true
    $textbox_filter.Enabled = $true
    $textbox_filter.Focus()
    $textbox_filter.Clear()
    $statusStriplabel_filterMessage.Text = ""
    $statusStriplabel_filterMessage.Visible = $true
    $statusStriplabel_filterMessage.Text = "{0} users" -f $Results.Count
    $statusStriplabel_filter.Enabled = $true
    ####################

    # Get the search criteria from the text boxes
    $name = $textBoxName.Text
    $department = $textBoxDepartment.Text
    $emailaddress = $textBoxEmailAddress.Text
    $displayName = $textBoxDisplayName.Text
    $title = $textBoxTitle.Text
    $Mobile = $textBoxMobile.Text
    $TelephoneNumber = $textBoxTelephoneNumber.Text

        # Build the LDAP filter based on the search criteria
    $filter = ""
    if ($name) {
        $filter += "(name=$name)"
    }
    if ($department) {
        $filter += "(department=$department)"
    }
    if ($emailaddress) {
        $filter += "(mail=$emailaddress)"
    }
    if ($displayName) {
        $filter += "(displayName=$displayName*)"
    }
    if ($title) {
        $filter += "(title=$title)"
    }
        if ($Mobile) {
        $filter += "(mobile=$Mobile)"
    }
        if ($TelephoneNumber) {
        $filter += "(TelephoneNumber=$TelephoneNumber)"            
    }
  
    
    #Get the sort property
    $sortProperty = $comboBoxSort.SelectedItem.ToString()

    # Search Active Directory and display the results
    if ($filter) {
        $results = Get-ADUser -LDAPFilter $filter -Properties name, department, mail, displayName, title, mobile, TelephoneNumber,
            extensionAttribute6, extensionAttribute8, EmployeeID, LastLogonDate, Enabled, LockedOut
    }
    if ($results) {
        #search AD and save results to global variable
        $results = Get-ADUser -LDAPFilter $filter -Properties displayName, name, title, department, mail, mobile, TelephoneNumber, extensionAttribute6,
        extensionAttribute8, EmployeeID, LastLogonDate, Enabled, LockedOut, lockOutTime, pwdLastSet, LastBadPasswordAttempt, city, state, streetaddress, PostalCode, employeeType  |
            Select-Object @{N='Full Name';E={$_.displayName}}, @{N='Username';E={$_.name}}, @{N='Job Title';E={$_.title}}, 
                @{N='Department';E={$_.department}}, @{N='Email';E={$_.mail}}, @{N='Mobile Phone';E={$_.mobile}}, 
                @{N='Office Phone';E={$_.TelephoneNumber}}, @{N='Company';E={$_.extensionAttribute6}}, @{N='Organization';E={$_.extensionAttribute8}},
                @{N='Emp. ID';E={$_.EmployeeID}}, @{N='Last Logon';E={$_.LastLogonDate}}, Enabled, LockedOut, City, State, @{N='Address';E={$_.streetaddress}},
                @{N='Lockout Time';E={$_.lockOutTime}}, @{N='Emp. Type';E={$_.employeeType}}, @{N='Zip';E={$_.PostalCode}}, 
                @{N='Last Bad PW';E={$_.LastBadPasswordAttempt}}, @{N='PW Last Set';E={$_.pwdLastSet}}| 
                    Sort-Object "$sortProperty"  

        ### Adjust form and controls properties ###

        #Search results message
        $buttonSearchMessage.Text = ""

        ### POPULATE LIST VIEW
        
        $textBoxResults.BeginUpdate() 
        $textBoxResults.Items.Clear()
        $itemCount = 0
        foreach ($result in $results) {
            if ($result.'Full Name') {$row = New-Object System.Windows.Forms.ListViewItem($result.'Full Name')} else {$row = New-Object System.Windows.Forms.ListViewItem("---")}
            if ($result.'Username') {$row.SubItems.Add($result.'Username')} else {$row.SubItems.Add("---")}
            if ($result.'Department') {$row.SubItems.Add($result.'Department')} else {$row.SubItems.Add("---")}
            if ($result.'Email') {$row.SubItems.Add($result.'Email')} else {$row.SubItems.Add("---")}
            if ($result.'Job Title') {$row.SubItems.Add($result.'Job Title')} else {$row.SubItems.Add("---")}
            if ($result.'Mobile Phone') {$row.SubItems.Add($result.'Mobile Phone')} else {$row.SubItems.Add("---")}
            if ($result.'Office Phone') {$row.SubItems.Add($result.'Office Phone')} else {$row.SubItems.Add("---")}

            #Add info to Tag property
            $row.Tag = @{
                FullName = $result.'Full Name'
                Username = $result.'Username'
            }
            $row.ToolTipText = $row.Tag.FullName

            $row.UseItemStyleForSubItems = $false
            $textBoxResults.Items.Add($row) | Out-Null

            $itemCount += 1
            $statusStriplabel_filterMessage.Text = "$itemCount users"
            $statusStrip.Refresh()
        }
        $textBoxResults.EndUpdate()


        $script:globalResults = $results

        return $results
                
    } else {
        if (!$filter) {
            # If nothing entered into any property search field
            [System.Windows.Forms.MessageBox]::Show("No search criteria provided.", "AD query Failure", "OK", "Error")
        } else {
            # If text entered but nothing found in AD
            [System.Windows.Forms.MessageBox]::Show("No results for: $filter", 'Active Directory', 'Ok', 'Information')

        }      
    }
}


#endregion FUNCTIONS
########################################

################################################################################
################################################################################
#                          USER DETAILS FORM
################################################################################
################################################################################
function CreateUserDetailsForm {
    # Create the User Details form
    $form_UserDetails = New-Object System.Windows.Forms.Form
    $form_UserDetails.ClientSize = New-Object System.Drawing.Size(600, 375)
    $form_UserDetails.StartPosition = "CenterScreen"
    $form_UserDetails.BackColor = $form.BackColor
    #$form_UserDetails.Add_Shown({OnUserDetailsFormLoad})

    # List Labels
    # Create an array of label titles
    $labelTitles = "Employee Details", "Job Details", "Personal Details", "Account/Authentication Details"

    # Create the labels using a loop
    $label_Y = 0
    for ($i = 0; $i -lt $labelTitles.Length; $i++) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelTitles[$i]
        $label.Font = New-Object System.Drawing.Font('Segoe UI', 10)
        $label.ForeColor = "$textLabel_ForeColor"
        $label.AutoSize = $true
        if ($i -ne 0) {$label_Y += 90}
        $label.Location = New-Object System.Drawing.Point(0, $label_Y)  # set the location of the label
        $form_UserDetails.Controls.Add($label)
    }

    ######################
    #Employee Details List
    ######################
    # Create a text box for displaying the results
    $listView_UserDetails = New-Object System.Windows.Forms.ListView
    $listView_UserDetails.View = "Details"
    $listView_UserDetails.FullRowSelect = $true
    $listView_UserDetails.MultiSelect = $true
    $listView_UserDetails.Location = New-Object System.Drawing.Point(0, 30)
    $listView_UserDetails.Size = New-Object System.Drawing.Size(595, 50)
    $listView_UserDetails.Font = New-Object System.Drawing.Font("Lucidia", 10)
    $listView_UserDetails.BackColor = $textBoxResults.BackColor
    $listView_UserDetails.ForeColor = $textBoxResults.ForeColor
    $listView_UserDetails.Visible = $true
    $listView_UserDetails.GridLines = $true
    $listView_UserDetails.AllowColumnReorder = $true
    $listView_UserDetails.BorderStyle = 'Fixed3D'
    $listView_UserDetails.ShowItemToolTips = $true    

    $listView_UserDetails.Columns.Add("Full Name", 120) | Out-Null
    $listView_UserDetails.Columns.Add("Username", 80) | Out-Null
    $listView_UserDetails.Columns.Add("Email", 155) | Out-Null
    $listView_UserDetails.Columns.Add("Emp. ID", 75) | Out-Null
    $listView_UserDetails.Columns.Add("Emp. Type", 60) | Out-Null

    # Set the width of the last column to automatically fill the available space
    $listView_UserDetails.Columns[$listView_UserDetails.Columns.Count - 1].Width = -2

    # Add the context menu to the listview control
    $listView_UserDetails.ContextMenuStrip = $contextMenu

    # Add event for Context Menu
    # $listView_UserDetails.Add_ItemSelectionChanged({
    #     $textBoxResults.SelectedItems.Clear()
    #     Copy-Property -ListViewRow $listView_UserDetails.SelectedItems[0]
    #     })

    ######################
    #Job Details List
    ######################
    $listView_UserDetails_2 = New-Object System.Windows.Forms.ListView
    $listView_UserDetails_2.View = "Details"
    $listView_UserDetails_2.FullRowSelect = $true
    $listView_UserDetails_2.MultiSelect = $true
    $listView_UserDetails_2.Location = New-Object System.Drawing.Point(0, 115)
    $listView_UserDetails_2.Size = New-Object System.Drawing.Size(595, 50)
    $listView_UserDetails_2.Font = New-Object System.Drawing.Font("Lucidia", 10)
    $listView_UserDetails_2.BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor
    $listView_UserDetails_2.ForeColor = [System.Drawing.Color]::$textBoxResults_DarkOFFforeColor
    $listView_UserDetails_2.Visible = $true
    $listView_UserDetails_2.GridLines = $true
    $listView_UserDetails_2.AllowColumnReorder = $true
    $listView_UserDetails_2.BorderStyle = 'Fixed3D'
    $listView_UserDetails_2.ShowItemToolTips = $true

    $listView_UserDetails_2.Columns.Add("Job Title", 150) | Out-Null
    $listView_UserDetails_2.Columns.Add("Department", 160) | Out-Null
    $listView_UserDetails_2.Columns.Add("Organization", 160) | Out-Null
    $listView_UserDetails_2.Columns.Add("Company", 95) | Out-Null
    # Set the width of the last column to automatically fill the available space
    $listView_UserDetails_2.Columns[$listView_UserDetails_2.Columns.Count - 1].Width = -2

    # Add the context menu to the listview control
    $listView_UserDetails_2.ContextMenuStrip = $contextMenu

    ######################
    #Personal Details List
    ######################
    $listView_UserDetails_3 = New-Object System.Windows.Forms.ListView
    $listView_UserDetails_3.View = "Details"
    $listView_UserDetails_3.FullRowSelect = $true
    $listView_UserDetails_3.MultiSelect = $true
    $listView_UserDetails_3.Location = New-Object System.Drawing.Point(0, 205)
    $listView_UserDetails_3.Size = New-Object System.Drawing.Size(595, 50)
    $listView_UserDetails_3.Font = New-Object System.Drawing.Font("Lucidia", 10)
    $listView_UserDetails_3.BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor
    $listView_UserDetails_3.ForeColor = [System.Drawing.Color]::$textBoxResults_DarkOFFforeColor
    $listView_UserDetails_3.Visible = $true
    $listView_UserDetails_3.GridLines = $true
    $listView_UserDetails_3.AllowColumnReorder = $true
    $listView_UserDetails_3.BorderStyle = 'Fixed3D'
    $listView_UserDetails_3.ShowItemToolTips = $true

    $listView_UserDetails_3.Columns.Add("Address", 250) | Out-Null
    $listView_UserDetails_3.Columns.Add("Mobile #", 120) | Out-Null
    $listView_UserDetails_3.Columns.Add("Office #", 120) | Out-Null
    # Set the width of the last column to automatically fill the available space
    $listView_UserDetails_3.Columns[$listView_UserDetails_3.Columns.Count - 1].Width = -2

    # Add the context menu to the listview control
    $listView_UserDetails_3.ContextMenuStrip = $contextMenu
    # Add event for Context Menu
    # $listView_UserDetails_3.Add_ItemSelectionChanged({
    #     $textBoxResults.SelectedItems.Clear()
    #     Copy-Property -ListViewRow $listView_UserDetails_3.SelectedItems[0]
    #     })

    ######################
    #Account/Authentication List
    ######################
    $listView_UserDetails_4 = New-Object System.Windows.Forms.ListView
    $listView_UserDetails_4.View = "Details"
    $listView_UserDetails_4.FullRowSelect = $true
    $listView_UserDetails_4.MultiSelect = $true
    $listView_UserDetails_4.Location = New-Object System.Drawing.Point(0, 295)
    $listView_UserDetails_4.Size = New-Object System.Drawing.Size(595, 50)
    $listView_UserDetails_4.Font = New-Object System.Drawing.Font("Lucidia", 10)
    $listView_UserDetails_4.BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor
    $listView_UserDetails_4.ForeColor = [System.Drawing.Color]::$textBoxResults_DarkOFFforeColor
    $listView_UserDetails_4.Visible = $true
    $listView_UserDetails_4.GridLines = $true
    $listView_UserDetails_4.AllowColumnReorder = $true
    $listView_UserDetails_4.BorderStyle = 'Fixed3D'
    $listView_UserDetails_4.ShowItemToolTips = $true

    $listView_UserDetails_4.Columns.Add("Enabled", 70) | Out-Null
    $listView_UserDetails_4.Columns.Add("Locked Out", 85) | Out-Null
    $listView_UserDetails_4.Columns.Add("Last Logon", 140) | Out-Null
    $listView_UserDetails_4.Columns.Add("Last Bad PW", 155) | Out-Null
    $listView_UserDetails_4.Columns.Add("PW Last Set", 100) | Out-Null
    # Set the width of the last column to automatically fill the available space
    $listView_UserDetails_4.Columns[$listView_UserDetails_4.Columns.Count - 1].Width = -2

    # Add the context menu to the listview control
    $listView_UserDetails_4.ContextMenuStrip = $contextMenu

    # Add lists  to form
    $form_UserDetails.Controls.Add($listView_UserDetails)
    $form_UserDetails.Controls.Add($listView_UserDetails_2)
    $form_UserDetails.Controls.Add($listView_UserDetails_3)
    $form_UserDetails.Controls.Add($listView_UserDetails_4)
    # $listView_UserDetails.Groups.Add($group1)
    # $listView_UserDetails_2.Groups.Add($group2)

    $selectedUser = $textBoxResults.SelectedItems[0].Tag.Username
    Get-UserDetails -Username $selectedUser

    #OnUserDetailsFormLoad

    foreach ($Control in $form_UserDetails.Controls) {
        #Adjust forecolor label font
        if ($Control.GetType().Name -eq "ListView") {
            $control.BackColor = $textBoxResults.BackColor
            $control.ForeColor = $textBoxResults.ForeColor
        }

    }
    

    $form_UserDetails.Show()| Out-Null
}


######################
######################
# FUNCTIONS
######################
function Get-UserDetails {
    
    param (
        [string]$Username
    )

    $form_UserDetails.Text = "$($textBoxResults.SelectedItems[0].Tag.FullName)"

    $globalResults | Where-Object { $_.Username -eq "$Username" } | ForEach-Object {
        #################
        #Employee Details
        #################
        #Row 1
        if ($_.'Full Name') {$row = New-Object System.Windows.Forms.ListViewItem($_.'Full Name')} else {$row = New-Object System.Windows.Forms.ListViewItem("---")}
        if ($_.'Username') {$row.SubItems.Add($_.'Username')} else {$row.SubItems.Add("---")}
        if ($_.'Email') {$row.SubItems.Add($_.'Email')} else {$row.SubItems.Add("---")}
        if ($_.'Emp. ID') {$row.SubItems.Add($_.'Emp. ID')} else {$row.SubItems.Add("---")}
        if ($_.'Emp. Type') {$row.SubItems.Add($_.'Emp. Type')} else {$row.SubItems.Add("---")}
        # Add Row to Group 1
        # $row.Group = $group1
        
        #Add info to Tag property
        $row.Tag = @{
            Address = $_.'Address'
        }
        # Set tooltips from Tag prop.
        #$row.ToolTipText = $row.Tag.Address

        $row.UseItemStyleForSubItems = $false

        #################
        #Job Details
        #################
        ###Job Title
        if ($_.'Job Title') {$row_2 = New-Object System.Windows.Forms.ListViewItem($_.'Job Title')} else {$row_2 = New-Object System.Windows.Forms.ListViewItem("---")}
        ###Department
        if ($_.'Department') {$row_2.SubItems.Add($_.'Department')} else {$row_2.SubItems.Add("---")}
        ###Organization
        if ($_.'Organization') {$row_2.SubItems.Add($_.'Organization')} else {$row_2.SubItems.Add("---")}
        ###Company
        if ($_.'Company') {$row_2.SubItems.Add($_.'Company')} else {$row_2.SubItems.Add("---")}

        #Add info to Tag property
        $row_2.Tag = @{
            Address = $_.'Address'
        }
        # Set tooltips from Tag prop.
        #$row_2.ToolTipText = $row_2.Tag.Address

        $row_2.UseItemStyleForSubItems = $false

        #################
        #Job Details
        #################
        ### Address
        if ($_.'Address' -or $_.City -or $_.State -or $_.Zip) {
            $Address_Full = "$($_.'Address'), $($_.City), $($_.State) $($_.Zip)"
            $row_3 = New-Object System.Windows.Forms.ListViewItem($Address_Full)
        } else {$row_3 = New-Object System.Windows.Forms.ListViewItem("---")}
        ### Mobile
        if ($_.'Mobile Phone') {$row_3.SubItems.Add($_.'Mobile Phone')} else {$row_3.SubItems.Add("---")}
        ### Office
        if ($_.'Office Phone') {$row_3.SubItems.Add($_.'Office Phone')} else {$row_3.SubItems.Add("---")}

        #Add info to Tag property
        $row_3.Tag = @{
            Address_Full = $Address_Full
        }
        # Set tooltips from Tag prop.
        $row_3.ToolTipText = $row_3.Tag.Address_Full

        $row_3.UseItemStyleForSubItems = $false
        
        #################
        #Account/Authentication Details
        #################
        #Row 2
        
        ###ENABLED
        if ($_.'Enabled') {$row_4 = New-Object System.Windows.Forms.ListViewItem(($_.'Enabled').ToString())} else {
            $row_4 = New-Object System.Windows.Forms.ListViewItem("---")
        }
        if ($row_4.Text -eq "False") {
            $row_4.BackColor = [System.Drawing.Color]::$enabledProperty_backcolor
        } else {$row_4.BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor}
        
        ###LOCKED OUT/LOCKOUT TIME
            # Convertime Unix timestamp of lock out time to be saved in Tag prop.
        if ($_.'Lockout Time') {
            $lockoutTimeStamp = $_.'Lockout Time'
            $dateTimeString = [datetime]::FromFileTime($lockoutTimeStamp)
        } else {$dateTimeString = ""}
        
        $row_4.SubItems.Add(("$($_.LockedOut)   $dateTimeString").ToString())
        if ($_.LockedOut -eq "True") {
            $row_4.SubItems[1].BackColor = [System.Drawing.Color]::$enabledProperty_backcolor
        } else {$row_4.SubItems[1].BackColor = [System.Drawing.Color]::$textBoxResults_DarkOffbackcolor}

        ###LAST LOGON
        if ($_.'Last Logon') {$row_4.SubItems.Add(($_.'Last Logon').ToString())} else {$row_4.SubItems.Add("---")}

        ###LAST BAD PW ATTEMPT
        if ($_.'Last Bad PW') {$row_4.SubItems.Add(($_.'Last Bad PW').ToString())} else {$row_4.SubItems.Add("---")}

        ### PW LAST SET
        if ($_.'PW Last Set') {
            $pwLastSetTimeStamp = $_.'PW Last Set'
            $dateTimeString = [datetime]::FromFileTime($pwLastSetTimeStamp)
        } else {$dateTimeString = "---"}

        $row_4.SubItems.Add(("$dateTimeString").ToString())

        #Add info to Tag property
        $row_4.Tag = @{
            lockOutTime = "$dateTimeString"
        }
        # Set tooltips from Tag prop.
        #$row_4.ToolTipText = $row_4.Tag.lockOutTime

        $row_4.UseItemStyleForSubItems = $false

        #Add both rows to list view
        $listView_UserDetails.Items.Add($row) | Out-Null
        $listView_UserDetails_2.Items.Add($row_2) | Out-Null
        $listView_UserDetails_3.Items.Add($row_3) | Out-Null
        $listView_UserDetails_4.Items.Add($row_4) | Out-Null

    }

    #OnUserDetailsFormLoad
    
}

function OnUserDetailsFormLoad {
    $form_UserDetails.Add_Load({
        $listView_UserDetails.Add_ItemSelectionChanged({
            $textBoxResults.SelectedItems.Clear()
            Copy-Property -ListViewRow $listView_UserDetails.SelectedItems[0]
            })
        $listView_UserDetails_2.Add_ItemSelectionChanged({
            $textBoxResults.SelectedItems.Clear()
            Copy-Property -ListViewRow $listView_UserDetails_2.SelectedItems[0]
            })
        $listView_UserDetails_3.Add_ItemSelectionChanged({
            $textBoxResults.SelectedItems.Clear()
            Copy-Property -ListViewRow $listView_UserDetails_3.SelectedItems[0]
            })
        $listView_UserDetails_4.Add_ItemSelectionChanged({
            $textBoxResults.SelectedItems.Clear()
            Copy-Property -ListViewRow $listView_UserDetails_4.SelectedItems[0]
            })
    })
}
#endregion
####################

# Show the form
$form.ShowDialog() | Out-Null




