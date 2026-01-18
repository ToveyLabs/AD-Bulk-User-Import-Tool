# Required for double-click execution
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables to store CSV data between operations
$script:csvData = $null
$script:csvFilePath = ""
$script:modifiedData = $null

# ===== AD CONNECTION VARIABLES =====
$script:adConnected = $false
$script:adDomain = ""
$script:adCredential = $null

# Main form - optimized for 1366x768 screens
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "Users Import AD Import Tool"
$MainForm.Size = New-Object System.Drawing.Size(900, 700)
$MainForm.StartPosition = "CenterScreen"
$MainForm.MinimumSize = New-Object System.Drawing.Size(900, 700)
$MainForm.MaximumSize = New-Object System.Drawing.Size(1200, 800)
$MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Main layout container (split panel for CSV preview + options)
$SplitContainer = New-Object System.Windows.Forms.SplitContainer
$SplitContainer.Dock = "Fill"
$SplitContainer.SplitterDistance = 450  # Left panel gets more space
$SplitContainer.SplitterWidth = 8
$SplitContainer.Panel1.BackColor = [System.Drawing.Color]::White
$SplitContainer.Panel2.BackColor = [System.Drawing.Color]::FromArgb(248,248,248)
$MainForm.Controls.Add($SplitContainer)


# ===== LEFT PANEL: CSV PREVIEW (450px wide) ===== #

# File selection group
$FileGroup = New-Object System.Windows.Forms.GroupBox
$FileGroup.Text = "1. Select CSV File"
$FileGroup.Location = New-Object System.Drawing.Point(10, 10)
$FileGroup.Size = New-Object System.Drawing.Size(420, 70)
$FileGroup.Anchor = "Left,Right,Top"
$SplitContainer.Panel1.Controls.Add($FileGroup)

$FileTextBox = New-Object System.Windows.Forms.TextBox
$FileTextBox.Location = New-Object System.Drawing.Point(10, 20)
$FileTextBox.Size = New-Object System.Drawing.Size(250, 23)
$FileTextBox.ReadOnly = $true
$FileGroup.Controls.Add($FileTextBox)

$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Text = "Browse..."
$BrowseButton.Location = New-Object System.Drawing.Point(270, 20)
$BrowseButton.Size = New-Object System.Drawing.Size(65, 23)
$FileGroup.Controls.Add($BrowseButton)

$LoadButton = New-Object System.Windows.Forms.Button
$LoadButton.Text = "Load"
$LoadButton.Location = New-Object System.Drawing.Point(340, 20)
$LoadButton.Size = New-Object System.Drawing.Size(65, 23)
$LoadButton.Enabled = $false
$FileGroup.Controls.Add($LoadButton)

$ResetButton = New-Object System.Windows.Forms.Button
$ResetButton.Text = "Clear / Reset"
$ResetButton.Location = New-Object System.Drawing.Point(410, 20) # Adjusted position to fit nicely
$ResetButton.Size = New-Object System.Drawing.Size(85, 23)
# Add this button to the group
$FileGroup.Controls.Add($ResetButton)

# CSV preview group
$PreviewGroup = New-Object System.Windows.Forms.GroupBox
$PreviewGroup.Text = "2. CSV Preview"
$PreviewGroup.Location = New-Object System.Drawing.Point(10, 90)
$PreviewGroup.Size = New-Object System.Drawing.Size(420, 380)
$PreviewGroup.Anchor = "Left,Right,Top,Bottom"
$SplitContainer.Panel1.Controls.Add($PreviewGroup)

$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Location = New-Object System.Drawing.Point(10, 20)
$DataGridView.Size = New-Object System.Drawing.Size(400, 350)
$DataGridView.AllowUserToAddRows = $false
$DataGridView.AllowUserToDeleteRows = $false
$DataGridView.ReadOnly = $true
$DataGridView.AutoSizeColumnsMode = "Fill"
$DataGridView.ColumnHeadersHeightSizeMode = "AutoSize"
$DataGridView.BackgroundColor = [System.Drawing.Color]::White
$DataGridView.Anchor = "Left,Right,Top,Bottom"
$PreviewGroup.Controls.Add($DataGridView)

# Status label
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Point(10, 480)
$StatusLabel.Size = New-Object System.Drawing.Size(420, 40)
$StatusLabel.Text = "Ready. Select a CSV file to begin."
$StatusLabel.Anchor = "Left,Right,Bottom"
$SplitContainer.Panel1.Controls.Add($StatusLabel)

# ===== RIGHT PANEL: CONFIGURATION (342px wide) ===== #

# Create scrollable panel for right side
$RightScrollPanel = New-Object System.Windows.Forms.Panel
$RightScrollPanel.Dock = "Fill"
$RightScrollPanel.AutoScroll = $true
$RightScrollPanel.BackColor = [System.Drawing.Color]::FromArgb(248,248,248)
$SplitContainer.Panel2.Controls.Add($RightScrollPanel)

# Import options group
$OptionsGroup = New-Object System.Windows.Forms.GroupBox
$OptionsGroup.Text = "3. Import Configuration"
$OptionsGroup.Location = New-Object System.Drawing.Point(10, 10)
$OptionsGroup.Size = New-Object System.Drawing.Size(320, 410)  # Increased height for all controls
$OptionsGroup.Anchor = "Left,Right,Top"
$RightScrollPanel.Controls.Add($OptionsGroup)

# ===== OU SELECTION ===== #
$OULabel = New-Object System.Windows.Forms.Label
$OULabel.Text = "Target OU:"
$OULabel.Location = New-Object System.Drawing.Point(15, 25)
$OULabel.Size = New-Object System.Drawing.Size(85, 20)
$OptionsGroup.Controls.Add($OULabel)

$OUTextBox = New-Object System.Windows.Forms.TextBox
$OUTextBox.Location = New-Object System.Drawing.Point(15, 45)
$OUTextBox.Size = New-Object System.Drawing.Size(290, 23)
$OUTextBox.Text = "OU=Students,DC=example,DC=edu"
$OUTextBox.Anchor = "Left,Right,Top"
$OptionsGroup.Controls.Add($OUTextBox)

# ===== USERNAME FORMAT ===== #
$UsernameLabel = New-Object System.Windows.Forms.Label
$UsernameLabel.Text = "Username Format:"
$UsernameLabel.Location = New-Object System.Drawing.Point(15, 75)
$UsernameLabel.Size = New-Object System.Drawing.Size(120, 20)
$OptionsGroup.Controls.Add($UsernameLabel)

$UsernameComboBox = New-Object System.Windows.Forms.ComboBox
$UsernameComboBox.Location = New-Object System.Drawing.Point(15, 95)
$UsernameComboBox.Size = New-Object System.Drawing.Size(290, 23)
$UsernameComboBox.DropDownStyle = "DropDownList"
$UsernameComboBox.Items.AddRange(@("first.last", "finitial.last", "firstl", "student.id"))
$UsernameComboBox.SelectedIndex = 0
$UsernameComboBox.Anchor = "Left,Right,Top"
$OptionsGroup.Controls.Add($UsernameComboBox)

# ===== PASSWORD OPTIONS ===== #
$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Text = "Initial Password:"
$PasswordLabel.Location = New-Object System.Drawing.Point(15, 125)
$PasswordLabel.Size = New-Object System.Drawing.Size(120, 20)
$OptionsGroup.Controls.Add($PasswordLabel)

$PasswordComboBox = New-Object System.Windows.Forms.ComboBox
$PasswordComboBox.Location = New-Object System.Drawing.Point(15, 145)
$PasswordComboBox.Size = New-Object System.Drawing.Size(290, 23)
$PasswordComboBox.DropDownStyle = "DropDownList"
$PasswordComboBox.Items.AddRange(@(
    "Fixed: ChangeMe123!",
    "Student ID (if available)",
    "Random (8 chars)",
    "FirstName + Year",
    "Birthdate (if available)",
    "Custom..."
))
$PasswordComboBox.SelectedIndex = 0
$PasswordComboBox.Anchor = "Left,Right,Top"
$OptionsGroup.Controls.Add($PasswordComboBox)

# Custom password input (hidden by default)
$CustomPasswordTextBox = New-Object System.Windows.Forms.TextBox
$CustomPasswordTextBox.Location = New-Object System.Drawing.Point(15, 175)
$CustomPasswordTextBox.Size = New-Object System.Drawing.Size(290, 23)
$CustomPasswordTextBox.Text = "Enter custom password..."
$CustomPasswordTextBox.ForeColor = [System.Drawing.Color]::Gray
$CustomPasswordTextBox.Visible = $false
$OptionsGroup.Controls.Add($CustomPasswordTextBox)

# Custom password placeholder behavior
$CustomPasswordTextBox.Add_GotFocus({
    if ($CustomPasswordTextBox.Text -eq "Enter custom password...") {
        $CustomPasswordTextBox.Text = ""
        $CustomPasswordTextBox.ForeColor = [System.Drawing.Color]::Black
    }
})

$CustomPasswordTextBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($CustomPasswordTextBox.Text)) {
        $CustomPasswordTextBox.Text = "Enter custom password..."
        $CustomPasswordTextBox.ForeColor = [System.Drawing.Color]::Gray
    }
})

# Password ComboBox event handler
$PasswordComboBox.Add_SelectedIndexChanged({
    $selected = $PasswordComboBox.SelectedItem

    # Show/hide custom password box and adjust layout
    if ($selected -eq "Custom...") {
        $CustomPasswordTextBox.Visible = $true
        $GroupLabel.Location = New-Object System.Drawing.Point(15, 205)
        $AssignGroupsCheck.Location = New-Object System.Drawing.Point(15, 225)
        $SelectedGroupsLabel.Location = New-Object System.Drawing.Point(30, 248)
        $SelectedGroupsTextBox.Location = New-Object System.Drawing.Point(30, 268)
        $BrowseGroupsButton.Location = New-Object System.Drawing.Point(30, 298)
        $ClearGroupsButton.Location = New-Object System.Drawing.Point(155, 298)
        $CheckboxPanel.Location = New-Object System.Drawing.Point(15, 330)
    } else {
        $CustomPasswordTextBox.Visible = $false
        $GroupLabel.Location = New-Object System.Drawing.Point(15, 175)
        $AssignGroupsCheck.Location = New-Object System.Drawing.Point(15, 195)
        $SelectedGroupsLabel.Location = New-Object System.Drawing.Point(30, 218)
        $SelectedGroupsTextBox.Location = New-Object System.Drawing.Point(30, 238)
        $BrowseGroupsButton.Location = New-Object System.Drawing.Point(30, 268)
        $ClearGroupsButton.Location = New-Object System.Drawing.Point(155, 268)
        $CheckboxPanel.Location = New-Object System.Drawing.Point(15, 300)
    }

    # Show warning for weak passwords
    if ($selected -eq "Student ID (if available)" -or 
        $selected -eq "Birthdate (if available)" -or
        $selected -eq "Custom...") {
        $StatusLabel.Text = "Warning: Simple passwords may not meet AD complexity requirements."
        $StatusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
    } else {
        $StatusLabel.Text = "Ready. Select a CSV file to begin."
        $StatusLabel.ForeColor = [System.Drawing.Color]::Black
    }
})

# ===== GROUP MEMBERSHIP ===== #
$GroupLabel = New-Object System.Windows.Forms.Label
$GroupLabel.Text = "Group Membership:"
$GroupLabel.Location = New-Object System.Drawing.Point(15, 175)
$GroupLabel.Size = New-Object System.Drawing.Size(120, 20)
$OptionsGroup.Controls.Add($GroupLabel)

$AssignGroupsCheck = New-Object System.Windows.Forms.CheckBox
$AssignGroupsCheck.Text = "Assign users to AD security groups"
$AssignGroupsCheck.Location = New-Object System.Drawing.Point(15, 195)
$AssignGroupsCheck.Size = New-Object System.Drawing.Size(290, 20)
$AssignGroupsCheck.Checked = $false
$OptionsGroup.Controls.Add($AssignGroupsCheck)

$SelectedGroupsLabel = New-Object System.Windows.Forms.Label
$SelectedGroupsLabel.Text = "Selected Groups:"
$SelectedGroupsLabel.Location = New-Object System.Drawing.Point(30, 218)
$SelectedGroupsLabel.Size = New-Object System.Drawing.Size(100, 20)
$SelectedGroupsLabel.Enabled = $false
$OptionsGroup.Controls.Add($SelectedGroupsLabel)

$SelectedGroupsTextBox = New-Object System.Windows.Forms.TextBox
$SelectedGroupsTextBox.Location = New-Object System.Drawing.Point(30, 238)
$SelectedGroupsTextBox.Size = New-Object System.Drawing.Size(275, 23)
$SelectedGroupsTextBox.ReadOnly = $true
$SelectedGroupsTextBox.Text = "(None selected)"
$SelectedGroupsTextBox.ForeColor = [System.Drawing.Color]::Gray
$SelectedGroupsTextBox.Enabled = $false
$OptionsGroup.Controls.Add($SelectedGroupsTextBox)

$BrowseGroupsButton = New-Object System.Windows.Forms.Button
$BrowseGroupsButton.Text = "Select Groups..."
$BrowseGroupsButton.Location = New-Object System.Drawing.Point(30, 268)
$BrowseGroupsButton.Size = New-Object System.Drawing.Size(120, 25)
$BrowseGroupsButton.Enabled = $false
$OptionsGroup.Controls.Add($BrowseGroupsButton)

$ClearGroupsButton = New-Object System.Windows.Forms.Button
$ClearGroupsButton.Text = "Clear"
$ClearGroupsButton.Location = New-Object System.Drawing.Point(155, 268)
$ClearGroupsButton.Size = New-Object System.Drawing.Size(60, 25)
$ClearGroupsButton.Enabled = $false
$OptionsGroup.Controls.Add($ClearGroupsButton)

# Store selected groups globally
$script:selectedGroups = @()

# Group controls event handlers
$AssignGroupsCheck.Add_CheckedChanged({
    $enabled = $AssignGroupsCheck.Checked
    $SelectedGroupsLabel.Enabled = $enabled
    $SelectedGroupsTextBox.Enabled = $enabled
    $BrowseGroupsButton.Enabled = $enabled -and $script:adConnected
    $ClearGroupsButton.Enabled = $enabled
})

$ClearGroupsButton.Add_Click({
    $script:selectedGroups = @()
    $SelectedGroupsTextBox.Text = "(None selected)"
    $SelectedGroupsTextBox.ForeColor = [System.Drawing.Color]::Gray
})

# ===== CHECKBOXES ===== #
$CheckboxPanel = New-Object System.Windows.Forms.Panel
$CheckboxPanel.Location = New-Object System.Drawing.Point(15, 300)
$CheckboxPanel.Size = New-Object System.Drawing.Size(290, 85)
$CheckboxPanel.Anchor = "Left,Right,Top"
$OptionsGroup.Controls.Add($CheckboxPanel)

$ForcePasswordCheck = New-Object System.Windows.Forms.CheckBox
$ForcePasswordCheck.Text = "Force password change at first login"
$ForcePasswordCheck.Location = New-Object System.Drawing.Point(0, 0)
$ForcePasswordCheck.Size = New-Object System.Drawing.Size(280, 20)
$ForcePasswordCheck.Checked = $true
$CheckboxPanel.Controls.Add($ForcePasswordCheck)

$CreateHomeCheck = New-Object System.Windows.Forms.CheckBox
$CreateHomeCheck.Text = "Create home directory"
$CreateHomeCheck.Location = New-Object System.Drawing.Point(0, 28)
$CreateHomeCheck.Size = New-Object System.Drawing.Size(280, 20)
$CheckboxPanel.Controls.Add($CreateHomeCheck)

$TestModeCheck = New-Object System.Windows.Forms.CheckBox
$TestModeCheck.Text = "Test mode (no changes)"
$TestModeCheck.Location = New-Object System.Drawing.Point(0, 56)
$TestModeCheck.Size = New-Object System.Drawing.Size(280, 20)
$TestModeCheck.Checked = $true
$CheckboxPanel.Controls.Add($TestModeCheck)

# ===== AD CONNECTION GROUP ===== #
$ADConnectionGroup = New-Object System.Windows.Forms.GroupBox
$ADConnectionGroup.Text = "Active Directory Connection"
$ADConnectionGroup.Location = New-Object System.Drawing.Point(10, 430)
$ADConnectionGroup.Size = New-Object System.Drawing.Size(320, 110)
$ADConnectionGroup.Anchor = "Left,Right,Top"
$RightScrollPanel.Controls.Add($ADConnectionGroup)

$ADStatusLabel = New-Object System.Windows.Forms.Label
$ADStatusLabel.Text = "Not connected to AD"
$ADStatusLabel.Location = New-Object System.Drawing.Point(15, 20)
$ADStatusLabel.Size = New-Object System.Drawing.Size(290, 20)
$ADStatusLabel.ForeColor = [System.Drawing.Color]::Red
$ADConnectionGroup.Controls.Add($ADStatusLabel)

$ConnectADButton = New-Object System.Windows.Forms.Button
$ConnectADButton.Text = "Connect to AD"
$ConnectADButton.Location = New-Object System.Drawing.Point(15, 45)
$ConnectADButton.Size = New-Object System.Drawing.Size(90, 25)
$ADConnectionGroup.Controls.Add($ConnectADButton)

$BrowseOUButton = New-Object System.Windows.Forms.Button
$BrowseOUButton.Text = "Browse OU..."
$BrowseOUButton.Location = New-Object System.Drawing.Point(110, 45)
$BrowseOUButton.Size = New-Object System.Drawing.Size(90, 25)
$BrowseOUButton.Enabled = $false
$ADConnectionGroup.Controls.Add($BrowseOUButton)

$DisconnectADButton = New-Object System.Windows.Forms.Button
$DisconnectADButton.Text = "Disconnect"
$DisconnectADButton.Location = New-Object System.Drawing.Point(205, 45)
$DisconnectADButton.Size = New-Object System.Drawing.Size(90, 25)
$DisconnectADButton.Enabled = $false
$ADConnectionGroup.Controls.Add($DisconnectADButton)

$CurrentOULabel = New-Object System.Windows.Forms.Label
$CurrentOULabel.Text = "Current OU: (Not Set)"
$CurrentOULabel.Location = New-Object System.Drawing.Point(15, 80)
$CurrentOULabel.Size = New-Object System.Drawing.Size(290, 20)
$CurrentOULabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$ADConnectionGroup.Controls.Add($CurrentOULabel)

# ===== ACTIONS GROUP ===== #
$OptionsGroup2 = New-Object System.Windows.Forms.GroupBox
$OptionsGroup2.Text = "4. Actions"
$OptionsGroup2.Location = New-Object System.Drawing.Point(10, 550)
$OptionsGroup2.Size = New-Object System.Drawing.Size(320, 150)
$OptionsGroup2.Anchor = "Left,Right,Top"
$RightScrollPanel.Controls.Add($OptionsGroup2)

# Action buttons
$ButtonY = 25
$ButtonSpacing = 35

$AnalyzeButton = New-Object System.Windows.Forms.Button
$AnalyzeButton.Text = "Analyze CSV"
$AnalyzeButton.Location = New-Object System.Drawing.Point(15, $ButtonY)
$AnalyzeButton.Size = New-Object System.Drawing.Size(140, 28)
$AnalyzeButton.Enabled = $false
$OptionsGroup2.Controls.Add($AnalyzeButton)

$PrepareButton = New-Object System.Windows.Forms.Button
$PrepareButton.Text = "Prepare Data"
$PrepareButton.Location = New-Object System.Drawing.Point(165, $ButtonY)
$PrepareButton.Size = New-Object System.Drawing.Size(140, 28)
$PrepareButton.Enabled = $false
$OptionsGroup2.Controls.Add($PrepareButton)

$ButtonY += $ButtonSpacing

$ImportButton = New-Object System.Windows.Forms.Button
$ImportButton.Text = "Import Users"
$ImportButton.Location = New-Object System.Drawing.Point(15, $ButtonY)
$ImportButton.Size = New-Object System.Drawing.Size(140, 28)
$ImportButton.BackColor = [System.Drawing.Color]::LightBlue
$ImportButton.Enabled = $false
$OptionsGroup2.Controls.Add($ImportButton)

$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Text = "Export CSV"
$ExportButton.Location = New-Object System.Drawing.Point(165, $ButtonY)
$ExportButton.Size = New-Object System.Drawing.Size(140, 28)
$ExportButton.Enabled = $false
$OptionsGroup2.Controls.Add($ExportButton)

$ButtonY += $ButtonSpacing

$ColumnsButton = New-Object System.Windows.Forms.Button
$ColumnsButton.Text = "Select Columns"
$ColumnsButton.Location = New-Object System.Drawing.Point(15, $ButtonY)
$ColumnsButton.Size = New-Object System.Drawing.Size(140, 28)
$ColumnsButton.Enabled = $false
$OptionsGroup2.Controls.Add($ColumnsButton)

$ShowAllButton = New-Object System.Windows.Forms.Button
$ShowAllButton.Text = "Show All Columns"
$ShowAllButton.Location = New-Object System.Drawing.Point(165, $ButtonY)
$ShowAllButton.Size = New-Object System.Drawing.Size(140, 28)
$OptionsGroup2.Controls.Add($ShowAllButton)

$ButtonY += $ButtonSpacing

$AutoExportCheck = New-Object System.Windows.Forms.CheckBox
$AutoExportCheck.Text = "Auto-export passwords on completion"
$AutoExportCheck.Location = New-Object System.Drawing.Point(15, $ButtonY)
$AutoExportCheck.Size = New-Object System.Drawing.Size(280, 20)
$AutoExportCheck.Checked = $false
$OptionsGroup2.Controls.Add($AutoExportCheck)

# ===== NOW ADD EVENT HANDLERS AFTER ALL CONTROLS ARE CREATED ===== #

$ConnectADButton.Add_Click({
    Connect-ToActiveDirectory
})

$BrowseOUButton.Add_Click({
    Browse-ADOrganizationalUnits
})

$BrowseGroupsButton.Add_Click({
            Browse-ADGroups
})

$DisconnectADButton.Add_Click({
    $script:adConnected = $false
    $script:adDomain = ""
    $script:adCredential = $null
    Remove-Module ActiveDirectory -ErrorAction SilentlyContinue
    Update-ADConnectionUI
    [System.Windows.Forms.MessageBox]::Show(
        "Disconnected from Active Directory.",
        "Disconnected",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

$ColumnsButton.Add_Click({
    if ($DataGridView.DataSource -ne $null) {
        $dataTable = $DataGridView.DataSource
        Show-ColumnSelector -DataTable $dataTable -DataGridView $DataGridView
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "No data loaded. Please load and prepare data first.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
})

$ShowAllButton.Add_Click({
    if ($DataGridView.DataSource -ne $null) {
        foreach ($column in $DataGridView.Columns) {
            $column.Visible = $true
        }
    }
})

# ===== End Of Event Handlers ===== #

function Update-ADConnectionUI {
    if ($script:adConnected) {
        # Update Status
        $ADStatusLabel.Text = "Connected to: $($script:adDomain)"
        $ADStatusLabel.ForeColor = [System.Drawing.Color]::Green

        # Enable dependent buttons
        $BrowseOUButton.Enabled = $true
        $BrowseGroupsButton.Enabled = $AssignGroupsCheck.Checked
        $ImportButton.Enabled = ($null -ne $script:modifiedData)
        $ConnectADButton.Text = "Reconnect"

        # Update current OU display
        if ($OUTextBox.Text -and $OUTextBox.Text -ne "OU=Students,DC=example,DC=edu") {
            $CurrentOULabel.Text = "Current OU: $(Split-Path $OUTextBox.Text -Leaf)"
            $DisconnectADButton.Enabled = $true
        } else {
            $CurrentOULabel.Text = "Current OU: (Not Set)"
        }
    } else {
        # Update Status
        $ADStatusLabel.Text = "Not connected to AD"
        $ADStatusLabel.ForeColor = [System.Drawing.Color]::Red

        # Disable dependent buttons
        $BrowseOUButton.Enabled = $false
        $BrowseGroupsButton.Enabled = $false
        $ImportButton.Enabled = $false
        $DisconnectADButton.Enabled = $false
        $ConnectADButton.Text = "Connect to AD"
        $CurrentOULabel.Text = "Current OU: (Not Set)"
    }
}

# ===== Function Definitions  ===== #

# ===== AD FUNCTIONS =====
function Connect-ToActiveDirectory {
    # Check if already connected
    if ($script:adConnected) {
        $reconnect = [System.Windows.Forms.MessageBox]::Show(
            "Already connected to Active Directory.`nDomain: $($script:adDomain)`n`nDo you want to reconnect with different credentials?",
            "Already Connected",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($reconnect -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    try {
        # Check if AD module is available
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            $installChoice = [System.Windows.Forms.MessageBox]::Show(
                "Active Directory PowerShell module is not available.`n`n" +
                "This module is required for AD connectivity.`n" +
                "Would you like to install it via RSAT?`n`n" +
                "Note: RSAT is included with Windows 10/11 Pro and Server.",
                "AD Module Required",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($installChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process "https://learn.microsoft.com/en-us/windows-server/remote/remote-server-administration-tools"
            }
            return
        }

        # Import the module
        Import-Module ActiveDirectory -ErrorAction Stop

        # Try to get current domain without credentials first
        try {
            $domain = Get-ADDomain -ErrorAction Stop
            $script:adDomain = $domain.DNSRoot
            $script:adCredential = $null  # Use current user context
            $script:adConnected = $true

            # Update UI
            Update-ADConnectionUI

            [System.Windows.Forms.MessageBox]::Show(
                "Successfully connected to Active Directory.`nDomain: $($script:adDomain)`nUsing current user credentials.",
                "Connected",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            $ADStatusLabel.Text = "Connected to: $($script:adDomain)"
            $ADStatusLabel.ForeColor = [System.Drawing.Color]::Green
            $BrowseOUButton.Enabled = $true
            $ConnectADButton.Text = "Reconnect to AD"

            # Update OU text box with default Students OU (only if not manually set)
            $placeholder = "OU=Students,DC=example,DC=edu"
            if ([string]::IsNullOrEmpty($OUTextBox.Text) -or $OUTextBox.Text -eq $placeholder) {
                $defaultOU = "OU=Students,DC=$($domain.Name.Replace('.',',DC='))"
                $OUTextBox.Text = $defaultOU
                $CurrentOULabel.Text = "Current OU: $defaultOU"
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Successfully connected to Active Directory.`nDomain: $($script:adDomain)`nUsing current user credentials.",
                "Connected",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

        } catch {
            # If current user fails, prompt for credentials
            $credentialForm = New-Object System.Windows.Forms.Form
            $credentialForm.Text = "AD Authentication"
            $credentialForm.Size = New-Object System.Drawing.Size(350, 200)
            $credentialForm.StartPosition = "CenterParent"
            $credentialForm.FormBorderStyle = "FixedDialog"
            $credentialForm.MaximizeBox = $false

            $domainLabel = New-Object System.Windows.Forms.Label
            $domainLabel.Text = "Domain:"
            $domainLabel.Location = New-Object System.Drawing.Point(20, 20)
            $domainLabel.Size = New-Object System.Drawing.Size(100, 20)
            $credentialForm.Controls.Add($domainLabel)

            $domainTextBox = New-Object System.Windows.Forms.TextBox
            $domainTextBox.Location = New-Object System.Drawing.Point(120, 20)
            $domainTextBox.Size = New-Object System.Drawing.Size(200, 20)
            $domainTextBox.Text = $env:USERDNSDOMAIN
            $credentialForm.Controls.Add($domainTextBox)

            $userLabel = New-Object System.Windows.Forms.Label
            $userLabel.Text = "Username:"
            $userLabel.Location = New-Object System.Drawing.Point(20, 50)
            $userLabel.Size = New-Object System.Drawing.Size(100, 20)
            $credentialForm.Controls.Add($userLabel)

            $userTextBox = New-Object System.Windows.Forms.TextBox
            $userTextBox.Location = New-Object System.Drawing.Point(120, 50)
            $userTextBox.Size = New-Object System.Drawing.Size(200, 20)
            $credentialForm.Controls.Add($userTextBox)

            $passwordLabel = New-Object System.Windows.Forms.Label
            $passwordLabel.Text = "Password:"
            $passwordLabel.Location = New-Object System.Drawing.Point(20, 80)
            $passwordLabel.Size = New-Object System.Drawing.Size(100, 20)
            $credentialForm.Controls.Add($passwordLabel)

            $passwordTextBox = New-Object System.Windows.Forms.MaskedTextBox
            $passwordTextBox.Location = New-Object System.Drawing.Point(120, 80)
            $passwordTextBox.Size = New-Object System.Drawing.Size(200, 20)
            $passwordTextBox.PasswordChar = '*'
            $credentialForm.Controls.Add($passwordTextBox)

            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "Connect"
            $okButton.Location = New-Object System.Drawing.Point(120, 120)
            $okButton.Size = New-Object System.Drawing.Size(75, 25)
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $credentialForm.AcceptButton = $okButton
            $credentialForm.Controls.Add($okButton)

            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Text = "Cancel"
            $cancelButton.Location = New-Object System.Drawing.Point(205, 120)
            $cancelButton.Size = New-Object System.Drawing.Size(75, 25)
            $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $credentialForm.CancelButton = $cancelButton
            $credentialForm.Controls.Add($cancelButton)

            if ($credentialForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $domain = $domainTextBox.Text
                $username = $userTextBox.Text
                $password = ConvertTo-String $passwordTextBox.Text -AsPlainText -Force

                $script:adCredential = New-Object System.Management.Automation.PSCredential("$domain$username", $password)

                # Test connection with credentials
                try {
                    $domainInfo = Get-ADDomain -Server $domain -Credential $script:adCredential
                    $script:adDomain = $domainInfo.DNSRoot
                    $script:adConnected = $true

                    $ADStatusLabel.Text = "Connected to: $($script:adDomain)"
                    $ADStatusLabel.ForeColor = [System.Drawing.Color]::Green
                    $BrowseOUButton.Enabled = $true
                    $ConnectADButton.Text = "Reconnect to AD"

                    # Update OU text box (only if not manually set)
                    $placeholder = "OU=Students,DC=example,DC=edu"
                    if ([string]::IsNullOrEmpty($OUTextBox.Text) -or $OUTextBox.Text -eq $placeholder) {
                        $defaultOU = "OU=Students,DC=$($domainInfo.Name.Replace('.',',DC='))"
                        $OUTextBox.Text = $defaultOU
                        $CurrentOULabel.Text = "Current OU: $defaultOU"
                    }

                    [System.Windows.Forms.MessageBox]::Show(
                        "Successfully connected to Active Directory.",
                        "Connected",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )

                } catch {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Failed to connect to AD: $($_.Exception.Message)",
                        "Connection Failed",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                    $script:adConnected = $false
                    $ADStatusLabel.Text = "Connection failed"
                    $ADStatusLabel.ForeColor = [System.Drawing.Color]::Red
                    $BrowseOUButton.Enabled = $false
                }
            }
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error connecting to AD: $($_.Exception.Message)",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# ===== BROWSE AD GROUPS FUNCTION =====
function Browse-ADGroups {
    if (-not $script:adConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Not connected to Active Directory. Please connect first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    try {
        # Create group browser form
        $GroupBrowserForm = New-Object System.Windows.Forms.Form
        $GroupBrowserForm.Text = "Select AD Security Groups"
        $GroupBrowserForm.Size = New-Object System.Drawing.Size(600, 550)
        $GroupBrowserForm.StartPosition = "CenterParent"
        $GroupBrowserForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        # Instruction label
        $InstructionLabel = New-Object System.Windows.Forms.Label
        $InstructionLabel.Text = "Select security groups to assign imported users to:"
        $InstructionLabel.Location = New-Object System.Drawing.Point(10, 10)
        $InstructionLabel.Size = New-Object System.Drawing.Size(565, 20)
        $InstructionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $GroupBrowserForm.Controls.Add($InstructionLabel)

        # Search box
        $SearchLabel = New-Object System.Windows.Forms.Label
        $SearchLabel.Text = "Search:"
        $SearchLabel.Location = New-Object System.Drawing.Point(10, 35)
        $SearchLabel.Size = New-Object System.Drawing.Size(50, 20)
        $GroupBrowserForm.Controls.Add($SearchLabel)

        $SearchTextBox = New-Object System.Windows.Forms.TextBox
        $SearchTextBox.Location = New-Object System.Drawing.Point(65, 35)
        $SearchTextBox.Size = New-Object System.Drawing.Size(430, 23)
        $GroupBrowserForm.Controls.Add($SearchTextBox)

        $SearchButton = New-Object System.Windows.Forms.Button
        $SearchButton.Text = "Search"
        $SearchButton.Location = New-Object System.Drawing.Point(500, 33)
        $SearchButton.Size = New-Object System.Drawing.Size(75, 25)
        $GroupBrowserForm.Controls.Add($SearchButton)

        # CheckedListBox for groups
        $GroupListBox = New-Object System.Windows.Forms.CheckedListBox
        $GroupListBox.Location = New-Object System.Drawing.Point(10, 65)
        $GroupListBox.Size = New-Object System.Drawing.Size(565, 380)
        $GroupListBox.CheckOnClick = $true
        $GroupBrowserForm.Controls.Add($GroupListBox)

        # Status label
        $StatusLabel = New-Object System.Windows.Forms.Label
        $StatusLabel.Text = "Loading groups..."
        $StatusLabel.Location = New-Object System.Drawing.Point(10, 450)
        $StatusLabel.Size = New-Object System.Drawing.Size(565, 20)
        $GroupBrowserForm.Controls.Add($StatusLabel)

        # Buttons
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Text = "OK"
        $OKButton.Location = New-Object System.Drawing.Point(410, 480)
        $OKButton.Size = New-Object System.Drawing.Size(80, 25)
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $GroupBrowserForm.AcceptButton = $OKButton
        $GroupBrowserForm.Controls.Add($OKButton)

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Text = "Cancel"
        $CancelButton.Location = New-Object System.Drawing.Point(495, 480)
        $CancelButton.Size = New-Object System.Drawing.Size(80, 25)
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $GroupBrowserForm.CancelButton = $CancelButton
        $GroupBrowserForm.Controls.Add($CancelButton)

        # Function to load groups
        function Load-ADGroups {
            param([string]$SearchTerm = "*")

            $GroupListBox.Items.Clear()
            $StatusLabel.Text = "Searching..."
            $GroupBrowserForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

            try {
                # Build filter
                $filter = if ($SearchTerm -and $SearchTerm -ne "*") {
                    "Name -like '*$SearchTerm*' -and GroupCategory -eq 'Security'"
                } else {
                    "GroupCategory -eq 'Security'"
                }

                $adParams = @{
                    Filter = $filter
                    Properties = 'Name', 'Description', 'GroupScope'
                    ErrorAction = 'Stop'
                }

                if ($script:adCredential) {
                    $adParams.Credential = $script:adCredential
                    $adParams.Server = $script:adDomain
                }

                # Get security groups
                $groups = Get-ADGroup @adParams | 
                    Sort-Object Name | 
                    Select-Object -First 500  # Limit to 500 for performance

                if ($groups.Count -eq 0) {
                    $StatusLabel.Text = "No groups found matching search."
                    return
                }

                # Add to listbox
                foreach ($group in $groups) {
                    $displayText = $group.Name
                    if ($group.Description) {
                        $displayText += " - $($group.Description)"
                    }
                    $displayText += " [$($group.GroupScope)]"

                    $GroupListBox.Items.Add($displayText) | Out-Null

                    # Check if already selected
                    if ($script:selectedGroups -contains $group.Name) {
                        $GroupListBox.SetItemChecked($GroupListBox.Items.Count - 1, $true)
                    }
                }

                $StatusLabel.Text = "Found $($groups.Count) security groups. Check groups to assign users to them."

            } catch {
                $StatusLabel.Text = "Error loading groups: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show(
                    "Error loading groups: $($_.Exception.Message)",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            } finally {
                $GroupBrowserForm.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }

        # Search button handler
        $SearchButton.Add_Click({
            $searchTerm = $SearchTextBox.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                $searchTerm = "*"
            }
            Load-ADGroups -SearchTerm $searchTerm
        })

        # Enter key in search box
        $SearchTextBox.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                $SearchButton.PerformClick()
                $_.SuppressKeyPress = $true
            }
        })

        # Load groups when form opens
        $GroupBrowserForm.Add_Shown({
            Load-ADGroups
        })

        # Show dialog and process result
        if ($GroupBrowserForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Get checked items
            $script:selectedGroups = @()
            foreach ($item in $GroupListBox.CheckedItems) {
                # Extract group name (Handle hyphens in names correctly)
                $groupName = $item.ToString()
				
				# Split on the description separator " - " (only if it exists)
				if ($groupName -match '^(.*) - .*\[.*\]$|^(.*) - .*') {
					# If format is "Name - Description [Scope]" or "Name - Description"
					$groupName = $matches[1].Trim()
				} elseif ($groupName -match '^(.*) \[.*\]$') {
					# If format is "Name [Scope]"
					$groupName = $matches[1].Trim()
                }
				$script:selectedGroups += $groupName
			}
			
            # Update UI
            if ($script:selectedGroups.Count -eq 0) {
                $SelectedGroupsTextBox.Text = "(None selected)"
                $SelectedGroupsTextBox.ForeColor = [System.Drawing.Color]::Gray
            } else {
                $SelectedGroupsTextBox.Text = "$($script:selectedGroups.Count) group(s) selected"
                $SelectedGroupsTextBox.ForeColor = [System.Drawing.Color]::Black
            }
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error browsing groups: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Browse-ADOrganizationalUnits {
    if (-not $script:adConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Not connected to Active Directory. Please connect first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    try {
        # Create OU browser form
        $OUBrowserForm = New-Object System.Windows.Forms.Form
        $OUBrowserForm.Text = "Browse Active Directory OUs"
        $OUBrowserForm.Size = New-Object System.Drawing.Size(600, 560)
        $OUBrowserForm.StartPosition = "CenterParent"
        $OUBrowserForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        # TreeView for OUs
        $OUTreeView = New-Object System.Windows.Forms.TreeView
        $OUTreeView.Location = New-Object System.Drawing.Point(10, 10)
        $OUTreeView.Size = New-Object System.Drawing.Size(565, 400)
        $OUTreeView.Anchor = "Left,Right,Top,Bottom"
        $OUTreeView.CheckBoxes = $false
        $OUBrowserForm.Controls.Add($OUTreeView)

        # Status label
        $OULabel = New-Object System.Windows.Forms.Label
        $OULabel.Text = "Loading OUs..."
        $OULabel.Location = New-Object System.Drawing.Point(10, 415)
        $OULabel.Size = New-Object System.Drawing.Size(565, 20)
        $OUBrowserForm.Controls.Add($OULabel)

        # Selected OU display
        $SelectedOUTextBox = New-Object System.Windows.Forms.TextBox
        $SelectedOUTextBox.Location = New-Object System.Drawing.Point(10, 440)
        $SelectedOUTextBox.Size = New-Object System.Drawing.Size(565, 20)
        $SelectedOUTextBox.ReadOnly = $true
        $OUBrowserForm.Controls.Add($SelectedOUTextBox)

        # Buttons
        $SelectButton = New-Object System.Windows.Forms.Button
        $SelectButton.Text = "Select"
        $SelectButton.Location = New-Object System.Drawing.Point(410, 465)
        $SelectButton.Size = New-Object System.Drawing.Size(80, 25)
        $SelectButton.Enabled = $false
        $SelectButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $OUBrowserForm.AcceptButton = $SelectButton
        $OUBrowserForm.Controls.Add($SelectButton)

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Text = "Cancel"
        $CancelButton.Location = New-Object System.Drawing.Point(495, 465)
        $CancelButton.Size = New-Object System.Drawing.Size(80, 25)
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $OUBrowserForm.CancelButton = $CancelButton
        $OUBrowserForm.Controls.Add($CancelButton)

        # Define helper function inside Browse-ADOrganizationalUnits scope
        function Populate-OUTree {
            param($ParentNode)

            try {
                $searchBase = if ($null -eq $ParentNode) { 
                    $domain = if ($script:adCredential) {
                        Get-ADDomain -Server $script:adDomain -Credential $script:adCredential
                    } else {
                        Get-ADDomain -Server $script:adDomain
                    }
                    $domain.DistinguishedName
                } else { 
                    $ParentNode.Tag 
                }

                $adParams = @{
                    SearchBase = $searchBase
                    Filter = '*'
                    Server = $script:adDomain
                    ErrorAction = 'SilentlyContinue'
                }

                if ($script:adCredential) {
                    $adParams.Credential = $script:adCredential
                }

                # Get child OUs - FIXED: Only direct children
                $adParams.SearchScope = 'OneLevel'  # ADDED: Prevent deep recursion
                $childOUs = Get-ADOrganizationalUnit @adParams | Sort-Object Name

                foreach ($ou in $childOUs) {
                    $node = New-Object System.Windows.Forms.TreeNode($ou.Name)
                    $node.Tag = $ou.DistinguishedName

                    # FIXED: Check if this OU has children before adding placeholder
                    $hasChildren = $false
                    try {
                        $checkParams = @{
                            SearchBase = $ou.DistinguishedName
                            SearchScope = 'OneLevel'
                            Filter = '*'
                            Server = $script:adDomain
                            ErrorAction = 'SilentlyContinue'
                        }
                        if ($script:adCredential) {
                            $checkParams.Credential = $script:adCredential
                        }
                        $hasChildren = (Get-ADOrganizationalUnit @checkParams | Measure-Object).Count -gt 0
                    } catch {
                        # Assume it might have children if we can't check
                        $hasChildren = $true
                    }

                    if ($hasChildren) {
                        $node.Nodes.Add("Loading...") | Out-Null
                    }

                    if ($null -eq $ParentNode) {
                        $OUTreeView.Nodes.Add($node) | Out-Null
                    } else {
                        $ParentNode.Nodes.Add($node) | Out-Null
                    }
                }

                # Get child containers
                $containerParams = @{
                    SearchBase = $searchBase
                    SearchScope = 'OneLevel'  # ADDED
                    Filter = {ObjectClass -eq "container"}
                    Server = $script:adDomain
                    ErrorAction = 'SilentlyContinue'
                }

                if ($script:adCredential) {
                    $containerParams.Credential = $script:adCredential
                }

                $childContainers = Get-ADObject @containerParams | Sort-Object Name

                foreach ($container in $childContainers) {
                    $node = New-Object System.Windows.Forms.TreeNode("[$($container.Name)]")
                    $node.Tag = $container.DistinguishedName
                    $node.ForeColor = [System.Drawing.Color]::Gray

                    if ($null -eq $ParentNode) {
                        $OUTreeView.Nodes.Add($node) | Out-Null
                    } else {
                        $ParentNode.Nodes.Add($node) | Out-Null
                    }
                }

            } catch {
                Write-Host "Error loading OUs: $_"
                if ($null -eq $ParentNode) {
                    $OULabel.Text = "Error: $($_.Exception.Message)"
                }
            }
        }        

        # TreeView events - FIXED REPLICATION BUG
        $OUTreeView.Add_BeforeExpand({
            param($sender, $e)

            $node = $e.Node
            # Only load if we have the placeholder
            if ($node.Nodes.Count -eq 1 -and $node.Nodes[0].Text -eq "Loading...") {
                $node.Nodes.Clear()
                Populate-OUTree -ParentNode $node

                # If still no children after loading, remove the expand icon
                if ($node.Nodes.Count -eq 0) {
                    # Add a dummy node to prevent the expand happening again
                    $dummyNode = New-Object System.Windows.Forms.TreeNode("")
                    $dummyNode.ForeColor = [System.Drawing.Color]::White
                    $node.Nodes.Add($dummyNode) | Out-Null
                }
            }
        })

        $OUTreeView.Add_AfterSelect({
            param($sender, $e)

            if ($null -ne $OUTreeView.SelectedNode) {
                $SelectedOUTextBox.Text = $OUTreeView.SelectedNode.Tag
                $SelectButton.Enabled = $true
                $OULabel.Text = "Selected: $($OUTreeView.SelectedNode.Text)"  # ADDED: Better feedback
            } else {
                $SelectButton.Enabled = $false
                $OULabel.Text = "Select an OU from the tree"
            }
        })

        # Load root OUs
        $OUBrowserForm.Add_Shown({
            try {
                $OULabel.Text = "Loading domain structure..."
                $OUTreeView.BeginUpdate()

                # FIXED: Use stored credentials
                $domain = if ($script:adCredential) {
                    Get-ADDomain -Server $script:adDomain -Credential $script:adCredential
                } else {
                    Get-ADDomain -Server $script:adDomain
                }

                $rootNode = New-Object System.Windows.Forms.TreeNode($domain.DNSRoot)
                $rootNode.Tag = $domain.DistinguishedName
                $rootNode.Nodes.Add("Loading...") | Out-Null
                $OUTreeView.Nodes.Add($rootNode) | Out-Null
                $OUTreeView.SelectedNode = $rootNode

                # Auto-expand domain
                $rootNode.Expand()

                $OUTreeView.EndUpdate()
                $OULabel.Text = "Select an OU from the tree"

            } catch {
                $OULabel.Text = "Error loading domain: $_"
            }
        })

        # Show dialog
        if ($OUBrowserForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($null -ne $OUTreeView.SelectedNode) {
                $selectedOU = $OUTreeView.SelectedNode.Tag
                $OUTextBox.Text = $selectedOU
                $CurrentOULabel.Text = "Current OU: $(Split-Path $selectedOU -Leaf)"
            }
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error browsing OUs: $($_.Exception.Message)",
            "Browse Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# ===== COLUMN MANAGEMENT =====
function Show-ColumnSelector {
    param(
        [Parameter(Mandatory=$true)]
        [System.Data.DataTable]$DataTable,
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.DataGridView]$DataGridView
    )

    $SelectorForm = New-Object System.Windows.Forms.Form
    $SelectorForm.Text = "Select Columns to Display"
    $SelectorForm.Size = New-Object System.Drawing.Size(450, 500)
    $SelectorForm.StartPosition = "CenterParent"
    $SelectorForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Panel for columns list with checkboxes
    $ColumnsPanel = New-Object System.Windows.Forms.Panel
    $ColumnsPanel.Dock = "Fill"
    $ColumnsPanel.AutoScroll = $true
    $SelectorForm.Controls.Add($ColumnsPanel)

    # Add checkboxes for each column
    $yPos = 10
    $checkBoxes = @{}

    # Common AD columns that student biased
    $usefulColumns = @(
        "FirstName", "LastName", "DisplayName", "SamAccountName", "UserPrincipalName", 
        "EmailAddress", "StudentID", "YearGroup", "FormGroup", "Class"
    )

    # Internal/generated columns (hide by default)
    $internalColumns = @(
        "AD_*", "Generated_*", "AD_InitialPassword", "AD_TargetOU", 
        "AD_ForcePasswordChange", "AD_CreateHomeDir", "AD_TestMode"
    )

    # Categorize columns
    $allColumns = @($DataTable.Columns | ForEach-Object { $_.ColumnName })

    # Group columns by category
    $usefulCols = @()
    $otherCols = @()
    $internalCols = @()

    foreach ($colName in $allColumns) {
        $isInternal = $false
        foreach ($pattern in $internalColumns) {
            if ($colName -like $pattern) {
                $isInternal = $true
                break
            }
        }

        if ($isInternal) {
            $internalCols += $colName
        } elseif ($colName -in $usefulColumns -or $colName -match "name|id|email|group|class|year") {
            $usefulCols += $colName
        } else {
            $otherCols += $colName
        }
    }

    # Sort each category
    $usefulCols = $usefulCols | Sort-Object
    $otherCols = $otherCols | Sort-Object
    $internalCols = $internalCols | Sort-Object

    # Add headers and checkboxes
    function Add-Category {
        param($title, $columns)

        if ($columns.Count -eq 0) { return }

        $label = New-Object System.Windows.Forms.Label
        $label.Text = $title
        $label.Location = New-Object System.Drawing.Point(10, $yPos)
        $label.Size = New-Object System.Drawing.Size(400, 20)
        $label.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $ColumnsPanel.Controls.Add($label)
        $yPos += 25

        foreach ($colName in $columns) {
            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = $colName
            $chk.Location = New-Object System.Drawing.Point(20, $yPos)
            $chk.Size = New-Object System.Drawing.Size(400, 20)
            $chk.Checked = $DataGridView.Columns[$colName].Visible
            $checkBoxes[$colName] = $chk
            $ColumnsPanel.Controls.Add($chk)
            $yPos += 25
        }
        $yPos += 10
    }

    Add-Category "Useful Columns" $usefulCols
    Add-Category "Other Columns" $otherCols
    Add-Category "Internal/Generated Columns (Usually Hidden)" $internalCols

    # Button panel at bottom
    $ButtonPanel = New-Object System.Windows.Forms.Panel
    $ButtonPanel.Dock = "Bottom"
    $ButtonPanel.Height = 40
    $ButtonPanel.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
    $SelectorForm.Controls.Add($ButtonPanel)

    # Action buttons
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "Apply Selection"
    $OKButton.Location = New-Object System.Drawing.Point(150, 8)
    $OKButton.Size = New-Object System.Drawing.Size(100, 25)
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ButtonPanel.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Text = "Cancel"
    $CancelButton.Location = New-Object System.Drawing.Point(260, 8)
    $CancelButton.Size = New-Object System.Drawing.Size(100, 25)
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $ButtonPanel.Controls.Add($CancelButton)

    $ShowAllButton = New-Object System.Windows.Forms.Button
    $ShowAllButton.Text = "Show All"
    $ShowAllButton.Location = New-Object System.Drawing.Point(30, 8)
    $ShowAllButton.Size = New-Object System.Drawing.Size(100, 25)
    $ShowAllButton.Add_Click({
        foreach ($chk in $checkBoxes.Values) {
            $chk.Checked = $true
        }
    })
    $ButtonPanel.Controls.Add($ShowAllButton)

    # Show the form
    $result = $SelectorForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Apply column visibility
        foreach ($colName in $checkBoxes.Keys) {
            $chk = $checkBoxes[$colName]
            if ($DataGridView.Columns.Contains($colName)) {
                $DataGridView.Columns[$colName].Visible = $chk.Checked
            }
        }
        return $true
    }
    return $false
}

# ===== EVENT HANDLERS (Stubs for now) ===== #
$BrowseButton.Add_Click({
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $FileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($FileDialog.ShowDialog() -eq "OK") {
        $FileTextBox.Text = $FileDialog.FileName
        $LoadButton.Enabled = $true

        # Reset states
        $DataGridView.DataSource = $null
        $AnalyzeButton.Enabled = $false
        $PrepareButton.Enabled = $false
        $ImportButton.Enabled = $false
        $ExportButton.Enabled = $false
        $StatusLabel.Text = "File selected. Click 'Load' to preview."
    }
})

$LoadButton.Add_Click({
    $StatusLabel.Text = "Loading file..."
    $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $AnalyzeButton.Enabled = $false
    $PrepareButton.Enabled = $false
    $ImportButton.Enabled = $false
    $ExportButton.Enabled = $false

    try {
        # Check if file exists
        $filePath = $FileTextBox.Text
        if (-not (Test-Path -Path $filePath -PathType Leaf)) {
            [System.Windows.Forms.MessageBox]::Show(
                "File not found: $filePath",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $StatusLabel.Text = "File not found."
            return
        }

        # Clear existing DataGridView
        $DataGridView.DataSource = $null
        $DataGridView.Columns.Clear()

        # Read CSV file
        $StatusLabel.Text = "Reading CSV file..."
        $csvData = Import-Csv -Path $filePath
		

        if ($csvData.Count -eq 0) {
            $StatusLabel.Text = "CSV file is empty or invalid."
            return
        }

        $StatusLabel.Text = "Processing $($csvData.Count) rows..."

        # Create DataTable for DataGridView binding
        $dataTable = New-Object System.Data.DataTable
        $dataTable.TableName = "CSVData"

        # Get column names from first row
        $firstRow = $csvData[0]
        $columnNames = $firstRow.PSObject.Properties.Name

        # Add columns to DataTable
        foreach ($colName in $columnNames) {
            $column = New-Object System.Data.DataColumn($colName, [string])
            $dataTable.Columns.Add($column)
        }

        # Add data rows (limit to 1000 for performance)
        $rowCount = [Math]::Min($csvData.Count, 1000)
        for ($i = 0; $i -lt $rowCount; $i++) {
            $row = $dataTable.NewRow()
            $currentRow = $csvData[$i]

            foreach ($colName in $columnNames) {
                if ($null -ne $currentRow.$colName) {
                    $row[$colName] = $currentRow.$colName.ToString()
                } else {
                    $row[$colName] = [DBNull]::Value
                }
            }
            $dataTable.Rows.Add($row)
        }

        # Bind to DataGridView
        $DataGridView.DataSource = $dataTable

        # Configure DataGridView appearance
        $DataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
        $DataGridView.BackgroundColor = [System.Drawing.Color]::White
        $DataGridView.GridColor = [System.Drawing.Color]::LightGray
        $DataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 248, 248)

        $StatusLabel.Text = "Loaded $($csvData.Count) rows from: $(Split-Path $filePath -Leaf)"
        $AnalyzeButton.Enabled = $true

        # Store the raw CSV data for later use
        $script:csvData = $csvData
        $script:csvFilePath = $filePath
        $script:csvColumnNames = $columnNames

        [System.Windows.Forms.MessageBox]::Show(
            "Successfully loaded $($csvData.Count) rows.`nPreview shows first $rowCount rows.",
            "File Loaded",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

    } catch [System.Management.Automation.PSInvalidCastException] {
        [System.Windows.Forms.MessageBox]::Show(
            "Error: The CSV file format may be invalid.`nPlease ensure it's a proper CSV with headers.",
            "CSV Format Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $StatusLabel.Text = "CSV format error. Check file structure."
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error loading CSV: $($_.Exception.Message)",
            "Load Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $StatusLabel.Text = "Error loading file: $($_.Exception.Message)"
    } finally {
        $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        # Force refresh of DataGridView
        $DataGridView.Refresh()
        $PreviewGroup.Refresh()
    }
})

$ResetButton.Add_Click({
    $script:csvData = $null
    $script:modifiedData = $null
    $script:csvFilePath = ""
    $script:csvColumnNames = $null

    # Clear the grid
    $DataGridView.DataSource = $null
    $DataGridView.Columns.Clear()

    # Clear file path
    $FileTextBox.Text = ""

    # Reset buttons state
    $LoadButton.Enabled = $false
    $AnalyzeButton.Enabled = $false
    $PrepareButton.Enabled = $false
    $ImportButton.Enabled = $false
    $ExportButton.Enabled = $false
    $ColumnsButton.Enabled = $false

    # Reset Status
    $StatusLabel.Text = "Ready. Select a CSV file to begin."
    $StatusLabel.ForeColor = [System.Drawing.Color]::Black
})

$AnalyzeButton.Add_Click({
    if ($null -eq $script:csvData -or $script:csvData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No CSV data loaded. Please load a CSV file first.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $StatusLabel.Text = "Analyzing CSV structure..."
    $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        # Create Analysis Results window
        $AnalysisForm = New-Object System.Windows.Forms.Form
        $AnalysisForm.Text = "CSV Analysis Results"
        $AnalysisForm.Size = New-Object System.Drawing.Size(700, 600)
        $AnalysisForm.StartPosition = "CenterParent"
        $AnalysisForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        # Tab control for different analysis views
        $AnalysisTabs = New-Object System.Windows.Forms.TabControl
        $AnalysisTabs.Dock = "Fill"
        $AnalysisTabs.Padding = New-Object System.Drawing.Point(10, 10)

        # Tab 1: Column Summary
        $SummaryTab = New-Object System.Windows.Forms.TabPage
        $SummaryTab.Text = "Column Summary"
        $AnalysisTabs.TabPages.Add($SummaryTab)

        $SummaryList = New-Object System.Windows.Forms.ListView
        $SummaryList.Dock = "Fill"
        $SummaryList.View = [System.Windows.Forms.View]::Details
        $SummaryList.FullRowSelect = $true
        $SummaryList.Columns.Add("Column Name", 150) | Out-Null
        $SummaryList.Columns.Add("Type", 100) | Out-Null
        $SummaryList.Columns.Add("Non-Empty", 80) | Out-Null
        $SummaryList.Columns.Add("Unique", 80) | Out-Null
        $SummaryList.Columns.Add("Max Length", 80) | Out-Null
        $SummaryList.Columns.Add("Status", 150) | Out-Null
        $SummaryTab.Controls.Add($SummaryList)

        # Tab 2: Required Fields Check
        $RequiredTab = New-Object System.Windows.Forms.TabPage
        $RequiredTab.Text = "Required Fields"
        $AnalysisTabs.TabPages.Add($RequiredTab)

        $RequiredList = New-Object System.Windows.Forms.ListView
        $RequiredList.Dock = "Fill"
        $RequiredList.View = [System.Windows.Forms.View]::Details
        $RequiredList.FullRowSelect = $true
        $RequiredList.Columns.Add("AD Field", 150) | Out-Null
        $RequiredList.Columns.Add("Required", 80) | Out-Null
        $RequiredList.Columns.Add("CSV Column", 150) | Out-Null
        $RequiredList.Columns.Add("Status", 100) | Out-Null
        $RequiredList.Columns.Add("Notes", 200) | Out-Null
        $RequiredTab.Controls.Add($RequiredList)

        # Analyze the data
        $rowCount = $script:csvData.Count
        $columnNames = $script:csvData[0].PSObject.Properties.Name

        # Define required and optional AD fields
        $adFields = @{
            "FirstName" = @{Required=$true; Description="Given Name"}
            "LastName" = @{Required=$true; Description="Surname"}
            "DisplayName" = @{Required=$false; Description="Display Name"}
            "SamAccountName" = @{Required=$false; Description="Username (pre-Windows 2000)"}
            "UserPrincipalName" = @{Required=$false; Description="User Principal Name (email format)"}
            "EmailAddress" = @{Required=$false; Description="Email Address"}
            "Title" = @{Required=$false; Description="Job Title"}
            "Department" = @{Required=$false; Description="Department"}
            "Company" = @{Required=$false; Description="School/Company"}
            "Office" = @{Required=$false; Description="Office/Room"}
            "TelephoneNumber" = @{Required=$false; Description="Phone Number"}
            "StreetAddress" = @{Required=$false; Description="Street Address"}
            "City" = @{Required=$false; Description="City"}
            "PostalCode" = @{Required=$false; Description="Postal Code"}
            "Country" = @{Required=$false; Description="Country"}
        }

        # Initialize statistics
        $columnStats = @{ }
        $warnings = @()
        $suggestions = @()
        $missingRequired = @()

        # Analyze each column
        foreach ($col in $columnNames) {
            $values = @()
            foreach ($row in $script:csvData) {
                $values += $row.$col
            }

            $nonEmpty = @($values | Where-Object { $_ -and $_.ToString().Trim() -ne "" })
            $unique = @($nonEmpty | Select-Object -Unique)

            # Calculate max length safely
            $maxLength = 0
            foreach ($value in $nonEmpty) {
                if ($value -ne $null) {
                    $len = $value.ToString().Length
                    if ($len -gt $maxLength) { $maxLength = $len }
                }
            }

            # Try to infer data type
            $sample = $nonEmpty | Select-Object -First 5
            $dataType = "String"
            foreach ($s in $sample) {
                if ($s -match '^\d+$') { $dataType = "Integer"; break }
                if ($s -match '^\d{1,2}/\d{1,2}/\d{4}') { $dataType = "Date"; break }
                if ($s -match '@') { $dataType = "Email"; break }
            }

            $columnStats[$col] = @{
                NonEmpty = $nonEmpty.Count
                Empty = $rowCount - $nonEmpty.Count
                Unique = $unique.Count
                MaxLength = $maxLength
                DataType = $dataType
                Sample = $sample
            }

            # Check for common issues
            if ($nonEmpty.Count -eq 0) {
                $warnings += "Column '$col' is completely empty"
            }

            if ($dataType -eq "Email" -and $unique.Count -lt [math]::Round($rowCount * 0.9)) {
                $suggestions += "Column '$col' appears to contain emails but may have duplicates"
            }

            # Add to summary list
            $status = "OK"
            if ($nonEmpty.Count -eq 0) { $status = "EMPTY" }
            elseif ($nonEmpty.Count -lt $rowCount) { $status = "HAS EMPTY" }

            $item = New-Object System.Windows.Forms.ListViewItem($col)
            $item.SubItems.Add($dataType)
            $item.SubItems.Add("$($nonEmpty.Count)/$rowCount")
            $item.SubItems.Add($unique.Count)
            if ($maxLength -gt 0) {
                $item.SubItems.Add($maxLength.ToString())
            } else {
                $item.SubItems.Add("0")
            }
            $item.SubItems.Add($status)

            if ($nonEmpty.Count -eq 0) { 
                $item.BackColor = [System.Drawing.Color]::LightPink 
            } elseif ($nonEmpty.Count -lt $rowCount) { 
                $item.BackColor = [System.Drawing.Color]::LightYellow 
            }

            $SummaryList.Items.Add($item) | Out-Null
        }

        # Check for required AD fields
        foreach ($field in $adFields.Keys) {
            $isRequired = $adFields[$field].Required
            $foundColumn = $null

            # Look for exact or partial matches
            foreach ($col in $columnNames) {
                if ($col -eq $field) {
                    $foundColumn = $col
                    break
                }
                if ($col -like "*$field*") {
                    $foundColumn = $col
                    break
                }
            }

            # Special case lookups for common variations
            if (-not $foundColumn) {
                switch ($field) {
                    "FirstName" { 
                        $foundColumn = $columnNames | Where-Object { $_ -match "First|Given|Forename" } | Select-Object -First 1 
                    }
                    "LastName" { 
                        $foundColumn = $columnNames | Where-Object { $_ -match "Last|Surname|Family" } | Select-Object -First 1 
                    }
                    "EmailAddress" { 
                        $foundColumn = $columnNames | Where-Object { $_ -match "Email|Mail" } | Select-Object -First 1 
                    }
                    "SamAccountName" { 
                        $foundColumn = $columnNames | Where-Object { $_ -match "User|Login|Account" } | Select-Object -First 1 
                    }
                    "UserPrincipalName" { 
                        $foundColumn = $columnNames | Where-Object { $_ -match "UPN|Principal" } | Select-Object -First 1 
                    }
                }
            }

            $status = if ($foundColumn) { "PRESENT" } else { "MISSING" }
            $notes = ""

            if (-not $foundColumn -and $isRequired) {
                $missingRequired += $field
                $notes = "Will be auto-generated"
				# Only say we can generate it if it's actually possible
                if ($field -in @("SamAccountName", "UserPrincipalName", "EmailAddress", "DisplayName")) {
                    $notes = "Will be auto-generated"
                } else {
                    $notes = "CRITICAL: Cannot create user without this data."
                }
            }

            # Suggest likely matches for missing fields
            if (-not $foundColumn) {
                $suggestedMatches = @()
                switch ($field) {
                    "FirstName" { 
                        $suggestedMatches = $columnNames | Where-Object { $_ -match "First|Given|Forename" } 
                    }
                    "LastName" { 
                        $suggestedMatches = $columnNames | Where-Object { $_ -match "Last|Surname|Family" } 
                    }
                    "EmailAddress" { 
                        $suggestedMatches = $columnNames | Where-Object { $_ -match "Email|Mail" } 
                    }
                    "SamAccountName" { 
                        $suggestedMatches = $columnNames | Where-Object { $_ -match "User|Login|Account" } 
                    }
                    "UserPrincipalName" { 
                        $suggestedMatches = $columnNames | Where-Object { $_ -match "UPN|Principal" } 
                    }
                }

                if ($suggestedMatches.Count -gt 0) {
                    $notes = "Consider mapping to: $($suggestedMatches[0])"
                }
            }

            $item = New-Object System.Windows.Forms.ListViewItem($field)
            if ($isRequired) {
                $item.SubItems.Add("Yes")
            } else {
                $item.SubItems.Add("No")
            }

            if ($foundColumn) {
                $item.SubItems.Add($foundColumn)
            } else {
                $item.SubItems.Add("---")
            }

            $item.SubItems.Add($status)
            $item.SubItems.Add($notes)

            if (-not $foundColumn -and $isRequired) { 
                $item.BackColor = [System.Drawing.Color]::LightPink 
            } elseif ($foundColumn -and $isRequired) {
                $colData = $columnStats[$foundColumn]
                if ($colData.NonEmpty -lt $rowCount) {
                    $item.BackColor = [System.Drawing.Color]::LightYellow
                } else {
                    $item.BackColor = [System.Drawing.Color]::LightGreen
                }
            }

            $RequiredList.Items.Add($item) | Out-Null
        }

        # Create summary panel at bottom
        $SummaryPanel = New-Object System.Windows.Forms.Panel
        $SummaryPanel.Dock = "Bottom"
        $SummaryPanel.Height = 100
        $SummaryPanel.BorderStyle = "FixedSingle"
        $AnalysisForm.Controls.Add($SummaryPanel)

        $SummaryText = New-Object System.Windows.Forms.TextBox
        $SummaryText.Multiline = $true
        $SummaryText.Dock = "Fill"
        $SummaryText.ReadOnly = $true
        $SummaryText.Font = New-Object System.Drawing.Font("Consolas", 9)
        $SummaryText.ScrollBars = "Vertical"

        # Build summary text
        $summary = @()
        $summary += "=== CSV ANALYSIS SUMMARY ==="
        $summary += "Total Rows: $rowCount"
        $summary += "Total Columns: $($columnNames.Count)"
        $summary += ""

        if ($missingRequired.Count -gt 0) {
            $summary += "WARNING: MISSING REQUIRED FIELDS:"
            $summary += "The following required AD fields are missing:"
            foreach ($field in $missingRequired) {
                $summary += "  * $field"
            }
            $summary += ""
            $summary += "ACTION: These will be auto-generated during preparation."
        } else {
            $summary += "OK: All required AD fields are present in CSV."
        }

        if ($warnings.Count -gt 0) {
            $summary += ""
            $summary += "WARNINGS:"
            foreach ($warning in $warnings) {
                $summary += "  * $warning"
            }
        }

        if ($suggestions.Count -gt 0) {
            $summary += ""
            $summary += "SUGGESTIONS:"
            foreach ($suggestion in $suggestions) {
                $summary += "  * $suggestion"
            }
        }

        # Check username uniqueness
        if ($columnNames -contains "SamAccountName" -or $columnNames -contains "UserName") {
            $userCol = if ($columnNames -contains "SamAccountName") { "SamAccountName" } else { "UserName" }
            $usernames = @()
            foreach ($row in $script:csvData) {
                $usernames += $row.$userCol
            }
            $usernames = $usernames | Where-Object { $_ -and $_.ToString().Trim() -ne "" }
            $uniqueUsernames = $usernames | Select-Object -Unique
            if ($usernames.Count -ne $uniqueUsernames.Count) {
                $summary += ""
                $summary += "WARNING: DUPLICATE USERNAMES FOUND:"
                $summary += "$($usernames.Count - $uniqueUsernames.Count) duplicate usernames detected."
            }
        }

        # Check email format if present
        $emailColumns = $columnNames | Where-Object { $_ -match "Email|Mail" }
        foreach ($emailCol in $emailColumns) {
            $emails = @()
            foreach ($row in $script:csvData) {
                $emails += $row.$emailCol
            }
            $emails = $emails | Where-Object { $_ -and $_.ToString().Trim() -ne "" }
            $invalidEmails = $emails | Where-Object { $_ -notmatch '^[^@]+@[^@]+\.[^@]+$' }
            if ($invalidEmails.Count -gt 0) {
                $summary += ""
                $summary += "WARNING: INVALID EMAIL FORMATS:"
                $summary += "Column '$emailCol' has $($invalidEmails.Count) invalid email addresses."
            }
        }

        $SummaryText.Text = $summary -join "`r`n"
        $SummaryPanel.Controls.Add($SummaryText)

        # Add tabs to form
        $AnalysisForm.Controls.Add($AnalysisTabs)

        # Action buttons
        $CloseButton = New-Object System.Windows.Forms.Button
        $CloseButton.Text = "Close"
        $CloseButton.Size = New-Object System.Drawing.Size(100, 30)
        $CloseButton.Location = New-Object System.Drawing.Point(300, 565)
        $CloseButton.Anchor = "Bottom"
        $CloseButton.Add_Click({ $AnalysisForm.Close() })
        $AnalysisForm.Controls.Add($CloseButton)

         # Enable Prepare button if we have data loaded
        # (Don't check headers here - user can fix them in Prepare step via Smart Fallback)
        if ($rowCount -gt 0) {
            $PrepareButton.Enabled = $true
        } else {
            $PrepareButton.Enabled = $false
        }

        $StatusLabel.Text = "Analysis complete. $rowCount rows, $($columnNames.Count) columns."

        # Show analysis form
        $AnalysisForm.Add_Shown({ $AnalysisForm.Activate() })
        [void]$AnalysisForm.ShowDialog()

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error during analysis: $($_.Exception.Message)`n`nLine: $($_.InvocationInfo.ScriptLineNumber)`nScript: $($_.InvocationInfo.ScriptName)",
            "Analysis Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $StatusLabel.Text = "Analysis failed: $($_.Exception.Message)"
    } finally {
        $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# ===== PASSWORD GENERATION FUNCTION =====
function Generate-Password {
    param(
        [Parameter(Mandatory=$true)]
        $Row,
        [Parameter(Mandatory=$true)]
        [String]$PasswordOption,
        $FirstNameColumn,
        $StudentIDColumn,
        $YearGroupColumn,
        $BirthdateColumn,
        [String] $CustomPassword  # NEW PARAMETER
    )

    try {
        switch ($PasswordOption) {
            "Fixed: ChangeMe123!" {
                return "ChangeMe123!"
            }

            "Custom..." {
                # Use custom password if provided, otherwise fallback
                if ($CustomPassword -and $CustomPassword.Trim() -ne "") {
                    return $CustomPassword.Trim()
                } else {
                    return "ChangeMe123!"
                }
            }

            "Student ID (if available)" {
                if ($StudentIDColumn -and $Row.$StudentIDColumn) {
                    $id = $Row.$StudentIDColumn.ToString().Trim()
                    if ($id.Length -ge 6) {
                        return "$id!"
                    } else {
                        return "Student$id!"
                    }
                }
                return "ChangeMe123!"
            }

            "Random (8 chars)" {
                # Generate cryptographically random password
                $chars = "abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789"
                $special = "!@#$%"

                # Build password: 2 uppercase, 4 lowercase, 1 number, 1 special
                $password = ""

                # 2 uppercase
                for ($i = 0; $i -lt 2; $i++) {
                    $password += $chars[(Get-Random -Maximum 26) + 26]
                }

                # 4 lowercase
                for ($i = 0; $i -lt 4; $i++) {
                    $password += $chars[(Get-Random -Maximum 26)]
                }

                # 1 number
                $password += $chars[(Get-Random -Maximum 8) + 52]

                # 1 special character
                $password += $special[(Get-Random -Maximum $special.Length)]

                # Shuffle the password
                $passwordArray = $password.ToCharArray()
                $random = New-Object System.Random
                for ($i = $passwordArray.Length - 1; $i -gt 0; $i--) {
                    $j = $random.Next(0, $i + 1)
                    $temp = $passwordArray[$i]
                    $passwordArray[$i] = $passwordArray[$j]
                    $passwordArray[$j] = $temp
                }

                return -join $passwordArray
            }

            "FirstName + Year" {
                $firstName = if ($FirstNameColumn -and $Row.$FirstNameColumn) {
                    $Row.$FirstNameColumn.ToString().Trim()
                } else {
                    "Student"
                }

                $year = if ($YearGroupColumn -and $Row.$YearGroupColumn) {
                    $Row.$YearGroupColumn.ToString().Trim()
                } else {
                    (Get-Date).Year.ToString()
                }

                # Capitalize first letter of first name
                $firstName = (Get-Culture).TextInfo.ToTitleCase($firstName.ToLower())

                return "$firstName$year!"
            }

            "Birthdate (if available)" {
                if ($BirthdateColumn -and $Row.$BirthdateColumn) {
                    $birthdate = $Row.$BirthdateColumn.ToString().Trim()

                    # Try to parse different date formats
                    $date = $null
                    $formats = @(
                        "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd",
                        "dd-MM-yyyy", "MM-dd-yyyy", "d/M/yyyy"
                    )

                    foreach ($format in $formats) {
                        try {
                            $date = [DateTime]::ParseExact($birthdate, $format, $null)
                            break
                        } catch {
                            continue
                        }
                    }

                    if ($date) {
                        # Format as DDMMYYYY (8 chars, meets length requirements)
                        return $date.ToString("ddMMyyyy")
                    }
                }
                # Fallback
                return "ChangeMe123!"
            }

            default {
                return "ChangeMe123!"
            }
        }
    } catch {
        Write-Host "Error generating password: $_"
        return "ChangeMe123!"
    }
}

# ===== GROUP NAME GENERATION FUNCTION =====
function Generate-GroupName {
    param(
        [Parameter(Mandatory=$true)]
        $Row,
        [Parameter(Mandatory=$true)]
        [string]$Pattern,
        $YearGroupColumn,
        $FormGroupColumn
    )

    try {
        $groupName = $Pattern

        # Replace placeholders
        if ($YearGroupColumn -and $Row.$YearGroupColumn) {
            $year = $Row.$YearGroupColumn.ToString().Trim()
            $groupName = $groupName -replace '\[YearGroup\]', $year
        } else {
            $groupName = $groupName -replace '\[YearGroup\]', 'Unknown'
        }

        if ($FormGroupColumn -and $Row.$FormGroupColumn) {
            $form = $Row.$FormGroupColumn.ToString().Trim()
            $groupName = $groupName -replace '\[FormGroup\]', $form
        } else {
            $groupName = $groupName -replace '\[FormGroup\]', 'Unknown'
        }

        # Clean up invalid characters for AD group names
        $groupName = $groupName -replace '[\\/:*?"<>|]', '_'

        # Remove any remaining placeholders
        $groupName = $groupName -replace '\[.*?\]', ''

        return $groupName.Trim()

    } catch {
        Write-Host "Error generating group name: $_"
        return $null
    }
}

# ===== CREATE OR GET AD GROUP FUNCTION =====
function Get-OrCreateADGroup {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GroupName,
        [Parameter(Mandatory=$true)]
        [string]$TargetOU,
        [Parameter(Mandatory=$false)]
        [bool]$TestMode = $false
    )

    try {
        if ([string]::IsNullOrWhiteSpace($GroupName) -or $GroupName -eq "Unknown") {
            return $null
        }

        # Check if group already exists
        $adParams = @{
            Filter = "Name -eq '$GroupName'"
            ErrorAction = 'SilentlyContinue'
        }

        if ($script:adCredential) {
            $adParams.Credential = $script:adCredential
            $adParams.Server = $script:adDomain
        }

        $existingGroup = Get-ADGroup @adParams

        if ($existingGroup) {
            return $existingGroup
        }

        # Create new group if it doesn't exist
        if (-not $TestMode) {
            $createParams = @{
                Name = $GroupName
                GroupScope = 'Global'
                GroupCategory = 'Security'
                Path = $TargetOU
                Description = "Auto-created student group"
                PassThru = $true
            }

            if ($script:adCredential) {
                $createParams.Credential = $script:adCredential
                $createParams.Server = $script:adDomain
            }

            $newGroup = New-ADGroup @createParams
            Write-Host "Created new group: $GroupName"
            return $newGroup
        } else {
            Write-Host "[TEST MODE] Would create group: $GroupName"
            return [PSCustomObject]@{ Name = $GroupName; TestMode = $true }
        }

    } catch {
        Write-Host "Error creating/getting group '$GroupName': $_"
        return $null
    }
}

$PrepareButton.Add_Click({
    if ($null -eq $script:csvData -or $script:csvData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No CSV data loaded. Please load and analyze a CSV file first.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $StatusLabel.Text = "Preparing data for import..."
    $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    # Validate custom password if selected
    if ($PasswordComboBox.SelectedItem -eq "Custom...") {
        $customPwdText = $CustomPasswordTextBox.Text.Trim()

        # Check if placeholder text is still there or empty
        if ([string]::IsNullOrWhiteSpace($customPwdText) -or 
            $customPwdText -eq "Enter custom password...") {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter a custom password or select a different password option.",
                "Custom Password Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }

        # Warn about simple passwords
        if ($customPwdText.Length -lt 3) {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Your custom password '$customPwdText' is very short.`n`nAre you sure you want to use this?",
                "Weak Password Warning",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
                return
            }
        }
    }

    # Get custom password, filtering out placeholder text
    $customPwd = if ($passwordOption -eq "Custom...") { 
        $pwd = $CustomPasswordTextBox.Text.Trim()
        if ($pwd -eq "Enter custom password...") { $null } else { $pwd }
    } else { 
        $null 
    }


    try {
        # Get configuration from UI
        $usernameFormat = $UsernameComboBox.SelectedItem
        $targetOU = $OUTextBox.Text.Trim()
        $forcePasswordChange = $ForcePasswordCheck.Checked
        $createHomeDir = $CreateHomeCheck.Checked
        $testMode = $TestModeCheck.Checked

        # Validate target OU format
        if (-not $targetOU -or -not ($targetOU -match "^OU=|^CN=")) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter a valid Active Directory OU path (e.g., OU=Students,DC=school,DC=local)",
                "Invalid OU",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }

        # Ask for confirmation
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Prepare data with these settings?`n`n" +
            "Username format: $usernameFormat`n" +
            "Target OU: $targetOU`n" +
            "Test mode: $(if ($testMode) { 'Yes (no changes will be made)' } else { 'No (users will be created)' })`n`n" +
            "Do you want to continue?",
            "Confirm Preparation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $StatusLabel.Text = "Preparation cancelled."
            $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }


        # Create a copy of the data to modify
        $modifiedData = @()

        # Get column names from CSV
        $columnNames = @($script:csvData[0].PSObject.Properties.Name)

        # Determine which columns we have
        $hasFirstName = $columnNames | Where-Object { $_ -match "FirstName|First Name|GivenName|Given" } | Select-Object -First 1
        $hasLastName = $columnNames | Where-Object { $_ -match "LastName|Last Name|Surname|Family" } | Select-Object -First 1

        # ===== SMART FALLBACK: Handle Unknown Headers =====
        # This triggers if standard headers (FirstName, LastName) are NOT found
        if (-not $hasFirstName -or -not $hasLastName) {
            if ($columnNames.Count -ge 2) {
                # Ask user to map the first two columns
                $col1Name = $columnNames[0]
                $col2Name = $columnNames[1]

                $guessMsg = "Standard headers (FirstName/LastName) not found.`n`n" +
                            "The following columns are available:`n`n" +
                            "Column 1: $col1Name`n" +
                            "Column 2: $col2Name`n`n" +
                            "`nDo you want to map:`n" +
                            "Column 1 -> First Name`n" +
                            "Column 2 -> Last Name"

                $guessConfirm = [System.Windows.Forms.MessageBox]::Show(
                    $guessMsg,
                    "Confirm Column Mapping",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )

                if ($guessConfirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                    # Apply the mapping
                    if (-not $hasFirstName) { $hasFirstName = $col1Name }
                    if (-not $hasLastName) { $hasLastName = $col2Name }
                    $StatusLabel.Text = "Using fallback mapping: $col1Name -> First Name"
                } else {
                    $StatusLabel.Text = "Preparation cancelled: Column mapping rejected."
                    $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
                    return
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Cannot proceed. Standard name columns not found and file has fewer than 2 columns.",
                    "Data Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                $StatusLabel.Text = "Preparation cancelled."
                $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
                return
            }
        }
        # ===== END SMART FALLBACK =====

    # ===== VALIDATION: Check for Valid Users =====
    # Instead of counting 'bad' rows, let's count 'good' rows.
    # Only warn if we have ZERO valid users.
    $validUserCount = 0

    foreach ($row in $script:csvData) {
        # Get values from columns we identified
        $fnCheck = if ($hasFirstName) { $row.$hasFirstName } else { "" }
        $lnCheck = if ($hasLastName) { $row.$hasLastName } else { "" }

        # Check if we have valid names (Trim removes accidental spaces)
        if (-not [string]::IsNullOrWhiteSpace($fnCheck.Trim()) -and -not [string]::IsNullOrWhiteSpace($lnCheck.Trim())) {
            $validUserCount++
        }
    }

    # Only warn if we have absolutely NO valid users to import
    if ($validUserCount -eq 0) {
        $warnMsg = "Validation Failed: No valid users found in mapped columns ($hasFirstName / $hasLastName).`n`n" +
                     "Please check your CSV data to ensure these columns contain names."

        [System.Windows.Forms.MessageBox]::Show(
            $warnMsg,
            "No Data Found",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Exclamation
        )

        $StatusLabel.Text = "Preparation cancelled: No valid users found."
        $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        return
    }
    # ===== END VALIDATION =====

        $hasStudentID = $columnNames | Where-Object { $_ -match "StudentID|Student ID|ID|StudentNumber" } | Select-Object -First 1
        $hasEmail = $columnNames | Where-Object { $_ -match "Email|Mail" } | Select-Object -First 1
        $hasYearGroup = $columnNames | Where-Object { $_ -match "YearGroup|Year|Class" } | Select-Object -First 1
        $hasFormGroup = $columnNames | Where-Object { $_ -match "FormGroup|Form|TutorGroup" } | Select-Object -First 1
        $hasBirthdate = $columnNames | Where-Object { $_ -match "Birth|DOB|DateOfBirth" } | Select-Object -First 1

        #if (-not $hasFirstName -or -not $hasLastName) {
        #   [System.Windows.Forms.MessageBox]::Show(
        #        "CSV must contain First Name and Last Name columns to generate usernames.",
        #        "Missing Required Columns",
        #        [System.Windows.Forms.MessageBoxButtons]::OK,
        #        [System.Windows.Forms.MessageBoxIcon]::Error
        #    )
        #    $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
        #    return
        #}

        # Prepare progress dialog
        $ProgressForm = New-Object System.Windows.Forms.Form
        $ProgressForm.Text = "Preparing Data"
        $ProgressForm.Size = New-Object System.Drawing.Size(400, 150)
        $ProgressForm.StartPosition = "CenterParent"
        $ProgressForm.FormBorderStyle = "FixedDialog"
        $ProgressForm.ControlBox = $false

        $ProgressLabel = New-Object System.Windows.Forms.Label
        $ProgressLabel.Text = "Processing records..."
        $ProgressLabel.Location = New-Object System.Drawing.Point(20, 20)
        $ProgressLabel.Size = New-Object System.Drawing.Size(360, 20)
        $ProgressForm.Controls.Add($ProgressLabel)

        $ProgressBar = New-Object System.Windows.Forms.ProgressBar
        $ProgressBar.Location = New-Object System.Drawing.Point(20, 50)
        $ProgressBar.Size = New-Object System.Drawing.Size(350, 20)
        $ProgressBar.Minimum = 0
        $ProgressBar.Maximum = $script:csvData.Count
        $ProgressForm.Controls.Add($ProgressBar)

        # Show progress form
        $ProgressForm.Show()
        $ProgressForm.Refresh()

        # Track generated usernames to avoid duplicates
        $generatedUsernames = @{}

        # Process each row
        for ($i = 0; $i -lt $script:csvData.Count; $i++) {
            $row = $script:csvData[$i]

            # Update progress
            $ProgressBar.Value = $i + 1
            $ProgressLabel.Text = "Processing record $($i + 1) of $($script:csvData.Count)..."
            $ProgressForm.Refresh()

            # Create a new object for modified data
            $newRow = New-Object PSObject

            # Copy existing columns
            $row.PSObject.Properties | ForEach-Object {
                $newRow | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value
            }

            # Get base values
            $firstName = (Get-Culture).TextInfo.ToTitleCase($row.$hasFirstName.Trim())
            $lastName = (Get-Culture).TextInfo.ToTitleCase($row.$hasLastName.Trim())
            $studentID = if ($hasStudentID) { $row.$hasStudentID.Trim() } else { "" }
            $yearGroup = if ($hasYearGroup) { $row.$hasYearGroup.Trim() } else { "" }
            $formGroup = if ($hasFormGroup) { $row.$hasFormGroup.Trim() } else { "" }

            # Generate username based on selected format
            $baseUsername = ""
            switch ($usernameFormat) {
                "first.last" {
                    $baseUsername = "$($firstName.ToLower()).$($lastName.ToLower())"
                }
                "finitial.last" {
                    $firstInitial = $firstName.Substring(0,1).ToLower()
                    $baseUsername = "$($firstInitial).$($lastName.ToLower())"
                }
                "firstl" {
                    $firstInitial = $firstName.Substring(0,1).ToLower()
                    $baseUsername = "$($firstInitial)$($lastName.ToLower())"
                }
                "student.id" {
                    if ($studentID) {
                        $baseUsername = $studentID.ToLower()
                    } else {
                        $baseUsername = "$($firstName.ToLower()).$($lastName.ToLower())"
                    }
                }
            }

            # Clean username (remove spaces, special characters)
            $baseUsername = $baseUsername -replace '[^a-zA-Z0-9._-]', ''

            # Handle duplicates
            $username = $baseUsername
            $counter = 1
            while ($generatedUsernames.ContainsKey($username)) {
                $username = "$baseUsername$counter"
                $counter++
            }
            $generatedUsernames[$username] = $true

            # Add generated fields
            $newRow | Add-Member -MemberType NoteProperty -Name "Username" -Value $username -Force
            $newRow | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "$firstName $lastName" -Force

            # Generate email if not present
            if (-not $hasEmail) {
                $domain = if ($targetOU -match "DC=([^,]+)") { $matches[1] } else { "school.local" }
                $email = "$username@$domain"
                $newRow | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $email -Force
            }

            # Add student-specific fields if they exist
            if ($studentID) {
                $newRow | Add-Member -MemberType NoteProperty -Name "StudentID" -Value $studentID -Force
            }
            if ($yearGroup) {
                $newRow | Add-Member -MemberType NoteProperty -Name "YearGroup" -Value $yearGroup -Force
            }
            if ($formGroup) {
                $newRow | Add-Member -MemberType NoteProperty -Name "FormGroup" -Value $formGroup -Force
            }

            # Add AD import fields (will be hidden by default)
            $newRow | Add-Member -MemberType NoteProperty -Name "AD_Username" -Value $username -Force
            $newRow | Add-Member -MemberType NoteProperty -Name "AD_FirstName" -Value $firstName -Force
            $newRow | Add-Member -MemberType NoteProperty -Name "AD_LastName" -Value $lastName -Force
            $newRow | Add-Member -MemberType NoteProperty -Name "AD_TargetOU" -Value $targetOU -Force
            # Generate password based on selected option
            $passwordOption = $PasswordComboBox.SelectedItem
            $customPwd = if ($passwordOption -eq "Custom...") { $CustomPasswordTextBox.Text } else { $null }

            $generatedPassword = Generate-Password -Row $row `
                -PasswordOption $passwordOption `
                -FirstNameColumn $hasFirstName `
                -StudentIDColumn $hasStudentID `
                -YearGroupColumn $hasYearGroup `
                -BirthdateColumn $hasBirthdate `
                -CustomPassword $customPwd  # NEW PARAMETER

            $newRow | Add-Member -MemberType NoteProperty -Name "AD_InitialPassword" -Value $generatedPassword -Force
            $newRow | Add-Member -MemberType NoteProperty -Name "AD_ForcePasswordChange" -Value $forcePasswordChange -Force
            $newRow | Add-Member -MemberType NoteProperty -Name "AD_CreateHomeDir" -Value $createHomeDir -Force
            $newRow | Add-Member -MemberType NoteProperty -Name "AD_TestMode" -Value $testMode -Force

            # ADD GROUP ASSIGNMENT HERE (inside the loop)
            if ($AssignGroupsCheck.Checked -and $script:selectedGroups.Count -gt 0) {
                $groupsList = $script:selectedGroups -join ';'
                $newRow | Add-Member -MemberType NoteProperty -Name "AD_AssignToGroups" -Value $groupsList -Force
            }

            $modifiedData += $newRow
        }   

        # Close progress form
        $ProgressForm.Close()

        # Update DataGridView with modified data
        $DataGridView.DataSource = $null
        $DataGridView.Columns.Clear()

        # Create DataTable for modified data
        $dataTable = New-Object System.Data.DataTable
        $firstRow = $modifiedData[0]
        $firstRow.PSObject.Properties.Name | ForEach-Object {
            $column = New-Object System.Data.DataColumn($_, [string])
            $dataTable.Columns.Add($column)
        }

        foreach ($row in $modifiedData) {
            $dataRow = $dataTable.NewRow()
            $row.PSObject.Properties | ForEach-Object {
                $dataRow[$_.Name] = if ($null -ne $_.Value) { $_.Value.ToString() } else { "" }
            }
            $dataTable.Rows.Add($dataRow)
        }

        $DataGridView.DataSource = $dataTable
        $DataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells

        # Hide internal columns by default
        foreach ($column in $DataGridView.Columns) {
            if ($column.Name -like "AD_*" -or $column.Name -eq "AD_InitialPassword") {
                $column.Visible = $false
            }
        }

        # Ask user if they want to select visible columns
        $selectColumns = [System.Windows.Forms.MessageBox]::Show(
            "Data preparation complete!`n`n" +
            "Records processed: $($modifiedData.Count)`n" +
            "Username format: $usernameFormat`n`n" +
            "Some internal columns have been hidden.`n" +
            "Would you like to customize which columns are visible?",
            "Column Selection",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($selectColumns -eq [System.Windows.Forms.DialogResult]::Yes) {
            Show-ColumnSelector -DataTable $dataTable -DataGridView $DataGridView
        }

        # Store modified data globally
        $script:modifiedData = $modifiedData

        # Enable Import and Export buttons
        $ImportButton.Enabled = $true
        $ExportButton.Enabled = $true
        $ColumnsButton.Enabled = $true

        $StatusLabel.Text = "Preparation complete. Ready for import."

    } catch {
        # Safely close form (only if it was created successfully)
        if ($ProgressForm -ne $null) {
            $ProgressForm.Close()
        }
        [System.Windows.Forms.MessageBox]::Show(
            "Error during preparation: $($_.Exception.Message)",
            "Preparation Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $StatusLabel.Text = "Preparation failed."
        $ImportButton.Enabled = $false
        $ExportButton.Enabled = $false
    } finally {
        $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$ImportButton.Add_Click({
    if ($null -eq $script:modifiedData -or $script:modifiedData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No data prepared for import. Please load, analyze, and prepare data first.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Check if we're in test mode
    $testMode = $TestModeCheck.Checked
    $actionWord = if ($testMode) { "simulate" } else { "create" }

    # Confirmation dialog
    $confirmMessage = "You are about to $actionWord Active Directory users.`n`n"
    $confirmMessage += "Target OU: $($OUTextBox.Text)`n"
    $confirmMessage += "Number of users: $($script:modifiedData.Count)`n"
    $confirmMessage += "Test Mode: $(if ($testMode) { 'ON (no changes will be made)' } else { 'OFF (users WILL be created)' })`n"
    $confirmMessage += "Force Password Change: $($ForcePasswordCheck.Checked)`n"
    $confirmMessage += "Create Home Directory: $($CreateHomeCheck.Checked)`n`n"

    if (-not $testMode) {
        $confirmMessage += "WARNING: This will create real Active Directory users!`n`n"
    }

    $confirmMessage += "Do you want to continue?"

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage,
        "Confirm User Import",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        $(if ($testMode) { [System.Windows.Forms.MessageBoxIcon]::Question } else { [System.Windows.Forms.MessageBoxIcon]::Warning })
    )

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
        $StatusLabel.Text = "Import cancelled."
        return
    }

    $StatusLabel.Text = "Importing users to Active Directory..."
    $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    # Create import results window
    $ImportForm = New-Object System.Windows.Forms.Form
    $ImportForm.Text = "AD Import Progress"
    $ImportForm.Size = New-Object System.Drawing.Size(600, 500)
    $ImportForm.StartPosition = "CenterParent"
    $ImportForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Progress bar
    $ImportProgressBar = New-Object System.Windows.Forms.ProgressBar
    $ImportProgressBar.Location = New-Object System.Drawing.Point(20, 20)
    $ImportProgressBar.Size = New-Object System.Drawing.Size(550, 30)
    $ImportProgressBar.Minimum = 0
    $ImportProgressBar.Maximum = $script:modifiedData.Count
    $ImportForm.Controls.Add($ImportProgressBar)

    # Status label
    $ImportStatusLabel = New-Object System.Windows.Forms.Label
    $ImportStatusLabel.Text = "Starting import..."
    $ImportStatusLabel.Location = New-Object System.Drawing.Point(20, 60)
    $ImportStatusLabel.Size = New-Object System.Drawing.Size(550, 25)
    $ImportForm.Controls.Add($ImportStatusLabel)

    # Results list view
    $ImportResultsList = New-Object System.Windows.Forms.ListView
    $ImportResultsList.Location = New-Object System.Drawing.Point(20, 90)
    $ImportResultsList.Size = New-Object System.Drawing.Size(550, 320)
    $ImportResultsList.View = [System.Windows.Forms.View]::Details
    $ImportResultsList.FullRowSelect = $true
    $ImportResultsList.Columns.Add("Username", 100) | Out-Null
    $ImportResultsList.Columns.Add("Name", 150) | Out-Null
    $ImportResultsList.Columns.Add("Status", 150) | Out-Null
    $ImportResultsList.Columns.Add("Message", 250) | Out-Null
    $ImportForm.Controls.Add($ImportResultsList)

    # Action buttons (initially hidden)
    $ImportCloseButton = New-Object System.Windows.Forms.Button
    $ImportCloseButton.Text = "Close"
    $ImportCloseButton.Location = New-Object System.Drawing.Point(250, 420)
    $ImportCloseButton.Size = New-Object System.Drawing.Size(100, 30)
    $ImportCloseButton.Enabled = $false
    $ImportCloseButton.Add_Click({ $ImportForm.Close() })
    $ImportForm.Controls.Add($ImportCloseButton)

    $ImportSaveLogButton = New-Object System.Windows.Forms.Button
    $ImportSaveLogButton.Text = "Save Log"
    $ImportSaveLogButton.Location = New-Object System.Drawing.Point(360, 420)
    $ImportSaveLogButton.Size = New-Object System.Drawing.Size(100, 30)
    $ImportSaveLogButton.Enabled = $false
    $ImportForm.Controls.Add($ImportSaveLogButton)

    # Auto-export passwords if requested
    if ($AutoExportCheck.Checked) {
        $date = Get-Date -Format "yyyyMMdd_HHmm"
        $autoPath = "$env:USERPROFILE\Desktop\AD_Passwords_$date.csv"

        # Create export list with just Username, Name, and Password
        $autoExportList = @()
        foreach ($row in $script:modifiedData) {
            $obj = [PSCustomObject]@{
                Username = $row.AD_Username
                Name     = $row.DisplayName
                Password = $row.AD_InitialPassword
            }
            $autoExportList += $obj
        }

        # Save to Desktop
        $autoExportList | Export-Csv -Path $autoPath -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show(
            "Passwords auto-exported to:`n`n$autoPath",
            "Auto-Export Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }

    # Create a timer for UI updates
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 500

    # Store results across timer ticks
    $script:importResults = @()
    $script:currentUserIndex = 0
    $script:importJobRunning = $true

    # Function to process a single user
    function Process-UserImport {
        param($user, $testMode, [String] $forcePasswordChange, $createHomeDir, $targetOU)

        try {
            # Extract user data
            $username = $user.AD_Username
            $firstName = $user.AD_FirstName
            $lastName = $user.AD_LastName
            $displayName = if ($user.DisplayName) { $user.DisplayName } else { "$firstName $lastName" }
            $password = $user.AD_InitialPassword

            # Check if user already exists
            $existingUser = $null
            try {
                $existingUser = Get-ADUser -Filter "SamAccountName -eq '$username'" -ErrorAction SilentlyContinue
            } catch {
                # User doesn't exist - that's fine
            }

            if ($existingUser) {
                return [PSCustomObject]@{
                    Username = $username
                    Name = $displayName
                    Status = "SKIPPED"
                    Message = "User already exists in AD"
                }
            }

            if (-not $testMode) {
                # Create the user in Active Directory
                # Explicitly cast types to ensure Strings aren't passed where Booleans are expected
                $userParams = @{
                    SamAccountName = $username
                    Name = $displayName
                    GivenName = $firstName
                    Surname = $lastName
                    DisplayName = $displayName
                    UserPrincipalName = "$username@$((Get-ADDomain).DNSRoot)"
                    AccountPassword = (ConvertTo-SecureString $password -AsPlainText -Force)
                    Enabled = [bool]$true
                    ChangePasswordAtLogon = [bool]$forcePasswordChange
                    Path = $targetOU
                }

                # Add optional fields if they exist in the data
                # Ensure we only pass data if it is not null/empty
                # Check for Email
                if ($user.EmailAddress) { $userParams.EmailAddress = $user.EmailAddress }

                # Check for StudentID (or variations)
                if ($user.EmployeeID) { $userParams.EmployeeID = $user.EmployeeID }

                # Check for Title (Support variations: "Job Title", "Position", "Role")
                $titleCols = $columnNames | Where-Object { $_ -match "Title|Job Title|Position|Role" }
                if ($titleCols) { 
                    $userParams.Title = $user.($titleCols | Select-Object -First 1) 
                }

                # Check for Office (Support variations: "Room", "Location")
                $officeCols = $columnNames | Where-Object { $_ -match "Office|Room|Location" }
                if ($officeCols) { 
                    $userParams.Office = $user.($officeCols | Select-Object -First 1) 
                }

                if ($user.Department) { $userParams.Department = $user.Department }
                if ($user.Company) { $userParams.Company = $user.Company }
                if ($user.Office) { $userParams.Office = $user.Office }
                if ($user.TelephoneNumber) { $userParams.OfficePhone = $user.TelephoneNumber }
                if ($user.StreetAddress) { $userParams.StreetAddress = $user.StreetAddress }
                if ($user.City) { $userParams.City = $user.City }
                if ($user.PostalCode) { $userParams.PostalCode = $user.PostalCode }
                if ($user.Country) { $userParams.Country = $user.Country }
                if ($user.YearGroup) { $userParams.Description = "Year Group: $($user.YearGroup)" }
                

				# Create the user in Active Directory
                $newUser = New-ADUser @userParams -PassThru
              
                # Add to groups if specified
                if ($user.AD_AssignToGroups) {
                    $groupNames = $user.AD_AssignToGroups -split ';'
                    $groupsAdded = @()
                    $groupsFailed = @()

                    foreach ($groupName in $groupNames) {
                        try {
                            # Get the group
                            $adParams = @{
                                Filter = "Name -eq '$groupName'"
                                ErrorAction = 'Stop'
                            }

                            if ($script:adCredential) {
                                $adParams.Credential = $script:adCredential
                                $adParams.Server = $script:adDomain
                            }

                            $group = Get-ADGroup @adParams

                            if ($group) {
                                # Add user to group
                                $addParams = @{
                                    Identity = $group
                                    Members = $newUser
                                    ErrorAction = 'Stop'
                                }

                                if ($script:adCredential) {
                                    $addParams.Credential = $script:adCredential
                                    $addParams.Server = $script:adDomain
                                }

                                Add-ADGroupMember @addParams
                                $groupsAdded += $groupName
                                Write-Host "Added $username to group: $groupName"
                            }
                        } catch {
                            $groupsFailed += $groupName
                            Write-Host "Warning: Could not add $username to group $groupName : $_"
                        }
                    }

                    # Return success with group info
                    $groupMessage = ""
                    if ($groupsAdded.Count -gt 0) {
                        $groupMessage += " Added to: $($groupsAdded -join ', ')"
                    }
                    if ($groupsFailed.Count -gt 0) {
                        $groupMessage += " Failed: $($groupsFailed -join ', ')"
                    }

                    return [PSCustomObject]@{
                        Username = $username
                        Name = $displayName
                        Status = "CREATED"
                        Message = "User created in $targetOU.$groupMessage"
                    }
                }

                # Create home directory if requested
                if ($createHomeDir) {
                    $homeDrive = "H:"
                    $homeDirectory = "\\fileserver\users$username"
                    Set-ADUser $newUser -HomeDrive $homeDrive -HomeDirectory $homeDirectory
                }

                return [PSCustomObject]@{
                    Username = $username
                    Name = $displayName
                    Status = "CREATED"
                    Message = "User created successfully in $targetOU"
                }
            } else {

                    # Test mode - just simulate
                    $groupInfo = ""
                    if ($user.AD_AssignToGroups) {
                        $groups = ($user.AD_AssignToGroups -split ';') -join ', '
                        $groupInfo = " (Would add to: $groups)"
                    }

                    return [PSCustomObject]@{
                        Username = $username
                        Name = $displayName
                        Status = "TEST"
                        Message = "Would create in $targetOU$groupInfo"
                    }
                }

        } catch {
            return [PSCustomObject]@{
                Username = $username
                Name = $displayName
                Status = "ERROR"
                Message = $_.Exception.Message
            }
        }
    }

    # Timer tick event
    $timer.Add_Tick({
        if ($script:currentUserIndex -lt $script:modifiedData.Count) {
            # Process next user
            $user = $script:modifiedData[$script:currentUserIndex]
            $result = Process-UserImport -user $user -testMode $testMode -forcePasswordChange $ForcePasswordCheck.Checked -createHomeDir $CreateHomeCheck.Checked -targetOU $OUTextBox.Text

            # Store result
            $script:importResults += $result

            # Update UI
            $item = New-Object System.Windows.Forms.ListViewItem($result.Username)
            $item.SubItems.Add($result.Name)
            $item.SubItems.Add($result.Status)
            $item.SubItems.Add($result.Message)

            # Color code based on status
            switch ($result.Status) {
                "CREATED" { $item.BackColor = [System.Drawing.Color]::LightGreen }
                "TEST" { $item.BackColor = [System.Drawing.Color]::LightBlue }
                "SKIPPED" { $item.BackColor = [System.Drawing.Color]::LightYellow }
                "ERROR" { $item.BackColor = [System.Drawing.Color]::LightPink }
            }

            $ImportResultsList.Items.Add($item) | Out-Null

            # Auto-scroll to bottom
            if ($ImportResultsList.Items.Count -gt 0) {
                $ImportResultsList.Items[$ImportResultsList.Items.Count - 1].EnsureVisible()
            }

            # Update progress
            $script:currentUserIndex++
            $ImportProgressBar.Value = $script:currentUserIndex
            $ImportStatusLabel.Text = "Processing user $($script:currentUserIndex) of $($script:modifiedData.Count)..."

        } else {
            # All users processed
            $timer.Stop()
            $script:importJobRunning = $false

            # Show summary
            $successCount = ($script:importResults | Where-Object { $_.Status -eq "CREATED" }).Count
            $testCount = ($script:importResults | Where-Object { $_.Status -eq "TEST" }).Count
            $skipCount = ($script:importResults | Where-Object { $_.Status -eq "SKIPPED" }).Count
            $errorCount = ($script:importResults | Where-Object { $_.Status -eq "ERROR" }).Count

            $summary = "Import complete! "
            if ($testMode) {
                $summary += "Test mode: $testCount users would be created."
            } else {
                $summary += "$successCount created, $skipCount skipped, $errorCount errors."
            }

            $ImportStatusLabel.Text = $summary

            # Enable action buttons
            $ImportCloseButton.Enabled = $true
            $ImportSaveLogButton.Enabled = $true
        }
    })

    # Save log button handler
    $ImportSaveLogButton.Add_Click({
        $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveDialog.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $SaveDialog.FileName = "AD_Import_Log_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".log"
        $SaveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

        if ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $logContent = "Active Directory Import Log`n"
            $logContent += "Generated: $(Get-Date)`n"
            $logContent += "Test Mode: $testMode`n"
            $logContent += "Target OU: $($OUTextBox.Text)`n"
            $logContent += "=" * 50 + "`n`n"

            foreach ($result in $script:importResults) {
                $logContent += "$($result.Username) | $($result.Name) | $($result.Status) | $($result.Message)`n"
            }

            $logContent | Out-File -FilePath $SaveDialog.FileName -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show(
                "Log saved to:`n$($SaveDialog.FileName)",
                "Log Saved",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })

    # Show the form and start the timer
    $ImportForm.Add_Shown({
        $ImportForm.Activate()
        $timer.Start()
    })

    # Show the import form
    [void]$ImportForm.ShowDialog()

    # Clean up timer
    $timer.Stop()
    $timer.Dispose()

    # Reset cursor and status
    $StatusLabel.Text = "Import process completed."
    $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
})

$ExportButton.Add_Click({
    if ($null -eq $script:modifiedData -or $script:modifiedData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No prepared data to export. Please load and prepare data first.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $StatusLabel.Text = "Exporting data..."
    $MainForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        # Ask user where to save the file
        $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $SaveDialog.FileName = "Students_Prepared_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".csv"
        $SaveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        $SaveDialog.OverwritePrompt = $true

        if ($SaveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            $StatusLabel.Text = "Export cancelled."
            $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
            return
        }

        $exportPath = $SaveDialog.FileName

        # Ask which columns to export
        $ExportForm = New-Object System.Windows.Forms.Form
        $ExportForm.Text = "Select Columns to Export"
        $ExportForm.Size = New-Object System.Drawing.Size(400, 500)
        $ExportForm.StartPosition = "CenterParent"
        $ExportForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        # Panel for columns list with checkboxes
        $ExportPanel = New-Object System.Windows.Forms.Panel
        $ExportPanel.Dock = "Fill"
        $ExportPanel.AutoScroll = $true
        $ExportForm.Controls.Add($ExportPanel)

        # Get all column names from modified data
        $allColumns = @()
        if ($script:modifiedData.Count -gt 0) {
            $allColumns = $script:modifiedData[0].PSObject.Properties.Name
        }

        # Create checkboxes for each column
        $yPos = 10
        $checkBoxes = @{ }
        $selectedColumns = @()

        # Presets for different export types
        $presetPanel = New-Object System.Windows.Forms.Panel
        $presetPanel.Location = New-Object System.Drawing.Point(10, $yPos)
        $presetPanel.Size = New-Object System.Drawing.Size(360, 80)

        $presetLabel = New-Object System.Windows.Forms.Label
        $presetLabel.Text = "Export Presets:"
        $presetLabel.Location = New-Object System.Drawing.Point(0, 0)
        $presetLabel.Size = New-Object System.Drawing.Size(100, 20)
        $presetPanel.Controls.Add($presetLabel)

        $FullExportButton = New-Object System.Windows.Forms.Button
        $FullExportButton.Text = "All Columns"
        $FullExportButton.Location = New-Object System.Drawing.Point(0, 25)
        $FullExportButton.Size = New-Object System.Drawing.Size(100, 25)
        $FullExportButton.Add_Click({
            foreach ($chk in $checkBoxes.Values) {
                $chk.Checked = $true
            }
        })
        $presetPanel.Controls.Add($FullExportButton)

        $ADExportButton = New-Object System.Windows.Forms.Button
        $ADExportButton.Text = "AD Import Only"
        $ADExportButton.Location = New-Object System.Drawing.Point(110, 25)
        $ADExportButton.Size = New-Object System.Drawing.Size(120, 25)
        $ADExportButton.Add_Click({
            foreach ($key in $checkBoxes.Keys) {
                $checkBoxes[$key].Checked = ($key -like "AD_*" -or 
                                           $key -in @("FirstName", "LastName", "Username", "DisplayName", 
                                                     "EmailAddress", "StudentID", "YearGroup", "FormGroup"))
            }
        })
        $presetPanel.Controls.Add($ADExportButton)

        $StudentExportButton = New-Object System.Windows.Forms.Button
        $StudentExportButton.Text = "Student Data Only"
        $StudentExportButton.Location = New-Object System.Drawing.Point(240, 25)
        $StudentExportButton.Size = New-Object System.Drawing.Size(120, 25)
        $StudentExportButton.Add_Click({
            foreach ($key in $checkBoxes.Keys) {
                $checkBoxes[$key].Checked = ($key -notlike "AD_*" -and $key -ne "AD_InitialPassword")
            }
        })
        $presetPanel.Controls.Add($StudentExportButton)

        $ExportPanel.Controls.Add($presetPanel)
        $yPos += 100

        $selectAllLabel = New-Object System.Windows.Forms.Label
        $selectAllLabel.Text = "Select Columns to Export:"
        $selectAllLabel.Location = New-Object System.Drawing.Point(10, $yPos)
        $selectAllLabel.Size = New-Object System.Drawing.Size(200, 20)
        $selectAllLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $ExportPanel.Controls.Add($selectAllLabel)
        $yPos += 25

        # Add Select All/None buttons
        $SelectAllButton = New-Object System.Windows.Forms.Button
        $SelectAllButton.Text = "Select All"
        $SelectAllButton.Location = New-Object System.Drawing.Point(10, $yPos)
        $SelectAllButton.Size = New-Object System.Drawing.Size(80, 25)
        $SelectAllButton.Add_Click({
            foreach ($chk in $checkBoxes.Values) {
                $chk.Checked = $true
            }
        })
        $ExportPanel.Controls.Add($SelectAllButton)

        $SelectNoneButton = New-Object System.Windows.Forms.Button
        $SelectNoneButton.Text = "Select None"
        $SelectNoneButton.Location = New-Object System.Drawing.Point(100, $yPos)
        $SelectNoneButton.Size = New-Object System.Drawing.Size(80, 25)
        $SelectNoneButton.Add_Click({
            foreach ($chk in $checkBoxes.Values) {
                $chk.Checked = $false
            }
        })
        $ExportPanel.Controls.Add($SelectNoneButton)

        $yPos += 35

        # Add checkboxes for each column
        $sortedColumns = $allColumns | Sort-Object
        foreach ($col in $sortedColumns) {
            # Skip internal columns by default
            $isChecked = $true
            if ($col -like "AD_*" -or $col -eq "AD_InitialPassword") {
                $isChecked = $false
            }

            $chk = New-Object System.Windows.Forms.CheckBox
            $chk.Text = $col
            $chk.Location = New-Object System.Drawing.Point(20, $yPos)
            $chk.Size = New-Object System.Drawing.Size(350, 20)
            $chk.Checked = $isChecked

            # Color code internal columns
            if ($col -like "AD_*") {
                $chk.ForeColor = [System.Drawing.Color]::DarkRed
                $chk.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
            } elseif ($col -like "*Generated*" -or $col -eq "AD_InitialPassword") {
                $chk.ForeColor = [System.Drawing.Color]::DarkBlue
                $chk.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
            }

            $checkBoxes[$col] = $chk
            $ExportPanel.Controls.Add($chk)
            $yPos += 25
        }

        # Button panel at bottom
        $ButtonPanel = New-Object System.Windows.Forms.Panel
        $ButtonPanel.Dock = "Bottom"
        $ButtonPanel.Height = 50
        $ButtonPanel.BackColor = [System.Drawing.Color]::FromArgb(240,240,240)
        $ExportForm.Controls.Add($ButtonPanel)

        $ExportOKButton = New-Object System.Windows.Forms.Button
        $ExportOKButton.Text = "Export"
        $ExportOKButton.Location = New-Object System.Drawing.Point(120, 10)
        $ExportOKButton.Size = New-Object System.Drawing.Size(75, 30)
        $ExportOKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $ButtonPanel.Controls.Add($ExportOKButton)

        $ExportCancelButton = New-Object System.Windows.Forms.Button
        $ExportCancelButton.Text = "Cancel"
        $ExportCancelButton.Location = New-Object System.Drawing.Point(205, 10)
        $ExportCancelButton.Size = New-Object System.Drawing.Size(75, 30)
        $ExportCancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $ButtonPanel.Controls.Add($ExportCancelButton)

        # Show the form
        $result = $ExportForm.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # Get selected columns
            $selectedColumns = @()
            foreach ($key in $checkBoxes.Keys) {
                if ($checkBoxes[$key].Checked) {
                    $selectedColumns += $key
                }
            }

            if ($selectedColumns.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "No columns selected for export.",
                    "Export Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                $StatusLabel.Text = "Export cancelled - no columns selected."
                return
            }

            # Prepare data for export (only selected columns)
            $exportData = @()
            foreach ($row in $script:modifiedData) {
                $exportRow = New-Object PSObject
                foreach ($col in $selectedColumns) {
                    $exportRow | Add-Member -MemberType NoteProperty -Name $col -Value $row.$col
                }
                $exportData += $exportRow
            }

            # Export to CSV
            $exportData | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8

            # Show success message
            $exportSummary = @"
Export completed successfully!

File: $(Split-Path $exportPath -Leaf)
Location: $(Split-Path $exportPath -Parent)
Rows: $($exportData.Count)
Columns: $($selectedColumns.Count)

Selected columns:
$($selectedColumns -join ", ")
"@

            [System.Windows.Forms.MessageBox]::Show(
                $exportSummary,
                "Export Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            $StatusLabel.Text = "Exported $($exportData.Count) rows to: $(Split-Path $exportPath -Leaf)"

            # Ask if user wants to open the exported file
            $openFile = [System.Windows.Forms.MessageBox]::Show(
                "Would you like to open the exported CSV file?",
                "Open File",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($openFile -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process $exportPath
            }
        } else {
            $StatusLabel.Text = "Export cancelled."
        }

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error during export: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $StatusLabel.Text = "Export failed: $($_.Exception.Message)"
    } finally {
        $MainForm.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})
# Initialize UI state
Update-ADConnectionUI

# Disable Import button until data is prepared AND connected to AD
$ImportButton.Enabled = $false

# ===== SHOW THE FORM ===== #
[void]$MainForm.ShowDialog()
