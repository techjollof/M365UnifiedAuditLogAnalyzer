# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

################# Parameters #####################
# Define standard button size
$buttonWidth = 100
$buttonHeight = 25
$uIMargin = 5

#################### Window Configuration #################################
# Create Window
$window = New-Object System.Windows.Window
$window.Title = "Unified Audit Log Viewer"
$window.Width = 1400
$window.Height = 800
$window.MinWidth = 800
$window.MinHeight = 500
$window.WindowStartupLocation = "CenterScreen"

# Styling for a modern UI feel
$window.FontFamily = "Segoe UI"
$window.FontSize = 13   # Slightly larger for readability
$window.FontWeight = "Normal"

# Enable resizing while maintaining structure
$window.ResizeMode = "CanResize"

# Optional: Add a subtle border around the window
$window.BorderThickness = 1
$window.BorderBrush = [System.Windows.Media.Brushes]::Gray

# Optional: Add a drop shadow effect for a modern UI
$shadowEffect = New-Object System.Windows.Media.Effects.DropShadowEffect
$shadowEffect.BlurRadius = 10
$shadowEffect.ShadowDepth = 5
$shadowEffect.Opacity = 0.4
$window.Effect = $shadowEffect

# Create Grid
$grid = New-Object System.Windows.Controls.Grid
$grid.Margin = "10"  # Add margin around the grid
$window.Content = $grid  # Set the grid as the window's content

# Define columns for the main Grid (percentage-based)
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "3*" }))  # Column 0: TreeView (25%)
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "5*" }))  # Column 1: Preview Pane (50%)
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "3*" }))  # Column 2: Detailed Info Pane (25%)

# Define rows for the main Grid
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 0: Header
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))      # Row 1: Main Content (stretches to fill remaining space)
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 2: Connect Online
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 3: Footer

# Create a StackPanel to hold all controls vertically
$auditSearchPaneStackPanel = New-Object System.Windows.Controls.WrapPanel
$auditSearchPaneStackPanel.Orientation = "horizontal"
$auditSearchPaneStackPanel.Margin = $uIMargin
$auditSearchPaneStackPanel.VerticalAlignment = "Stretch"
$auditSearchPaneStackPanel.HorizontalAlignment = "Stretch"


# Add the StackPanel to the main Grid
$grid.Children.Add($auditSearchPaneStackPanel)
[System.Windows.Controls.Grid]::SetRow($auditSearchPaneStackPanel, 2)
[System.Windows.Controls.Grid]::SetColumn($auditSearchPaneStackPanel, 0)
[System.Windows.Controls.Grid]::SetColumnSpan($auditSearchPaneStackPanel, 3)  # Span across all 3 columns

# Function to add a label and input control with horizontal grouping
function Add-InputControl {
    param (
        [string]$LabelText,
        [string]$ControlType,
        [int]$Row,
        [int]$objMargin = 5
    )

    # Create a StackPanel for horizontal grouping
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Orientation = "Horizontal"
    $stackPanel.Margin = $objMargin

    # Add Label
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $LabelText
    $label.Margin = "0,0,10,0"  # Add right margin to separate label and control
    $stackPanel.Children.Add($label)

    # Add Input Control
    switch ($ControlType) {
        "DatePicker" {
            $control = New-Object System.Windows.Controls.DatePicker
            $control.Margin = $objMargin
        }
        "TextBox" {
            $control = New-Object System.Windows.Controls.TextBox
            $control.Margin = $objMargin
            $control.Width = 300
            $control.Height = 25
        }
        "CheckBox" {
            $control = New-Object System.Windows.Controls.CheckBox
            $control.Margin = $objMargin
        }
    }

    $stackPanel.Children.Add($control)

    # Add the StackPanel to the main StackPanel
    $auditSearchPaneStackPanel.Children.Add($stackPanel)

    return $control
}

# Add input controls for each parameter
# Start Date
$startDatePicker = Add-InputControl -LabelText "Start Date:" -ControlType "DatePicker"
# End Date
$endDatePicker = Add-InputControl -LabelText "End Date:" -ControlType "DatePicker"
# Free Text
$freeTextTextBox = Add-InputControl -LabelText "Free Text:" -ControlType "TextBox"
# High Completeness
$highCompletenessCheckBox = Add-InputControl -LabelText "High Completeness:" -ControlType "CheckBox"
# IP Addresses
$ipAddressesTextBox = Add-InputControl -LabelText "IP Addresses:" -ControlType "TextBox"
# Longer Retention Enabled
$longerRetentionEnabledTextBox = Add-InputControl -LabelText "Longer Retention Enabled:" -ControlType "TextBox"
# Object IDs
$objectIdsTextBox = Add-InputControl -LabelText "Object IDs:" -ControlType "TextBox"
# Operations
$operationsTextBox = Add-InputControl -LabelText "Operations:" -ControlType "TextBox"
# Record Type
$recordTypeTextBox = Add-InputControl -LabelText "Record Type:" -ControlType "TextBox"
# Result Size
$resultSizeTextBox = Add-InputControl -LabelText "Result Size:" -ControlType "TextBox"
# Session Command
$sessionCommandTextBox = Add-InputControl -LabelText "Session Command:" -ControlType "TextBox"
# Session ID
$sessionIdTextBox = Add-InputControl -LabelText "Session ID:" -ControlType "TextBox"
# Site IDs
$siteIdsTextBox = Add-InputControl -LabelText "Site IDs:" -ControlType "TextBox"
# User IDs
$userIdsTextBox = Add-InputControl -LabelText "User IDs:" -ControlType "TextBox"
# Formatted Output
$formattedCheckBox = Add-InputControl -LabelText "Formatted" -ControlType "CheckBox"



# Search Button
$searchButton = New-Object System.Windows.Controls.Button
$searchButton.Content = "Online Audit Query"
$searchButton.Margin = "5"
$searchButton.Width = 100
$searchButton.Height = 30
$searchButton.Add_Click({
    try {
        # Build the parameters for the function
        $params = @{
            StartDate = $startDatePicker.SelectedDate
            EndDate   = $endDatePicker.SelectedDate
        }

        if ($freeTextTextBox.Text) {
            $params['FreeText'] = $freeTextTextBox.Text
        }
        if ($highCompletenessCheckBox.IsChecked) {
            $params['HighCompleteness'] = $true
        }
        if ($ipAddressesTextBox.Text) {
            $params['IPAddresses'] = $ipAddressesTextBox.Text -split ','
        }
        if ($longerRetentionEnabledTextBox.Text) {
            $params['LongerRetentionEnabled'] = $longerRetentionEnabledTextBox.Text
        }
        if ($objectIdsTextBox.Text) {
            $params['ObjectIds'] = $objectIdsTextBox.Text -split ','
        }
        if ($operationsTextBox.Text) {
            $params['Operations'] = $operationsTextBox.Text -split ','
        }
        if ($recordTypeTextBox.Text) {
            $params['RecordType'] = $recordTypeTextBox.Text
        }
        if ($resultSizeTextBox.Text) {
            $params['ResultSize'] = [int]$resultSizeTextBox.Text
        }
        if ($sessionCommandTextBox.Text) {
            $params['SessionCommand'] = $sessionCommandTextBox.Text
        }
        if ($sessionIdTextBox.Text) {
            $params['SessionId'] = $sessionIdTextBox.Text
        }
        if ($siteIdsTextBox.Text) {
            $params['SiteIds'] = $siteIdsTextBox.Text -split ','
        }
        if ($userIdsTextBox.Text) {
            $params['UserIds'] = $userIdsTextBox.Text -split ','
        }
        if ($formattedCheckBox.IsChecked) {
            $params['Formatted'] = $true
        }
        
        Get-command Search-UnifiedAuditLog
        # Call the Search-UnifiedAuditLog function (ensure this function exists in your environment)
        $results = Search-UnifiedAuditLog @params

        # Display the results
        if ($results) {
            $resultsTextBox.Text = $results | Out-String
        } else {
            $resultsTextBox.Text = "No results found."
        }
    } catch {
        $resultsTextBox.Text = "Error: $_"
    }
})
$auditSearchPaneStackPanel.Children.Add($searchButton)



# Results Display
$resultsLabel = New-Object System.Windows.Controls.Label
$resultsLabel.Content = "Results:"
$resultsLabel.Margin = "5"
$auditSearchPaneStackPanel.Children.Add($resultsLabel)

$resultsTextBox = New-Object System.Windows.Controls.TextBox
$resultsTextBox.Margin = "5"
$resultsTextBox.IsReadOnly = $true
$resultsTextBox.Width = "500"
$resultsTextBox.Width = "300"
$resultsTextBox.VerticalScrollBarVisibility = "Visible"
$resultsTextBox.TextWrapping = "Wrap"
$auditSearchPaneStackPanel.Children.Add($resultsTextBox)

# Show the window
$window.ShowDialog()