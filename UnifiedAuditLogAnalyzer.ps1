param (
    [object]$InputData   # Can be a CSV file path (string) OR in-memory data (Hashtable/Array)
)



# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


################### All functions ###################################

# Function to check if input is a valid file
function Test-ValidFile {
    param ([string]$FilePath)

    if (-not (Test-Path $FilePath)) {
        [System.Windows.MessageBox]::Show("Error: File not found at $FilePath") | Out-Null
        return $false
    }

    if ((Get-Item $FilePath).Length -eq 0) {
        [System.Windows.MessageBox]::Show("Error: File is empty.") | Out-Null
        return $false
    }

    return $true
}

# Function to check if input is valid in-memory data
function Test-ValidData {
    param ([object]$Data)

    if (($Data | Measure-Object).Count -eq 0) {
        [System.Windows.MessageBox]::Show("Error: Provided data is empty or invalid.") | Out-Null
        return $false
    }

    return $true
}

# Function to load data from a file
function Import-DataFromFile {
    param ([string]$FilePath)

    if ($FilePath -match '\.csv$') {
        return Import-Csv -Path $FilePath
    }
    elseif ($FilePath -match '\.json$') {
        return Get-Content -Path $FilePath | ConvertFrom-Json
    }
    else {
        [System.Windows.MessageBox]::Show("Error: Unsupported file format. Only CSV and JSON files are supported.") | Out-Null
        return $null
    }
}

# the status bar
function Update-StatusBar {
    param (
        $Message,
        [string]$TextColor = "Black"
    )
    $brush = [System.Windows.Media.Brushes]::$TextColor
    if (-not $statusBarText) {
        Update-StatusBar "statusBarText is null. Please ensure it is initialized."
        return
    }
    $statusBarText.Dispatcher.Invoke([action] {
            $statusBarText.Text = $Message
            $statusBarText.Foreground = $brush
        }, "Normal")
}


# Check the visibility of the UIElement
function Set-UIElementVisibility {
    param (
        [Parameter(Mandatory = $true)]
        [System.Windows.UIElement]$UIElement, # Accept any UIElement (Button, Grid, etc.)

        [Parameter(Mandatory = $true)]
        [ValidateSet("Hide", "Show")]
        $Action  
    )
    if ($Action -eq 'Hide') {
        $UIElement.Visibility = 'Collapsed'
    }
    elseif ($Action -eq 'Show') {
        $UIElement.Visibility = 'Visible'
    }
    else {
        Update-StatusBar "Invalid action. Use 'Hide' or 'Show'."
    }
}

##  Export filtered data
Function Export-FilteredAuditData {

    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$FilteredAuditData,
    
        [string]$Prefix = ""
    )

    $ProcessedLogs = @()

    function Expand-UnifiedAuditData {

        param(
            [Parameter(Mandatory = $true)]
            [PSCustomObject]$Data,
        
            [string]$Prefix = ""
        )
    
        $Expanded = @()
        foreach ($Property in $Data.PSObject.Properties) {
            $Name = if ($Prefix) { "$Prefix.$($Property.Name)" } else { $Property.Name }
        
            if ($Property.Value -is [PSCustomObject]) {
                $Expanded += Expand-UnifiedAuditData -Data $Property.Value -Prefix $Name
            }
            elseif ($Property.Value -is [System.Array]) {
                for ($i = 0; $i -lt $Property.Value.Count; $i++) {
                    $Expanded += Expand-UnifiedAuditData -Data $Property.Value[$i] -Prefix "$Name[$i]"
                }
            }
            else {
                $Expanded += [PSCustomObject]@{
                    Name  = $Name
                    Value = $Property.Value
                }
            }
        }
        return $Expanded
    }

    
    foreach ($Log in $FilteredAuditData) {
        try {
            $AuditData = $Log.AuditData | ConvertFrom-Json -ErrorAction Stop
            $ExpandedData = Expand-UnifiedAuditData -Data $AuditData
    
            $FlattenedLog = [PSCustomObject]@{
                RecordType   = $Log.RecordType
                CreationDate = $Log.CreationDate
                UserIds      = $Log.UserIds
                Operations   = $Log.Operations
            }
    
            foreach ($Item in $ExpandedData) {
                Add-Member -InputObject $FlattenedLog -MemberType NoteProperty -Name $Item.Name -Value $Item.Value -Force
            }
    
            $ProcessedLogs += $FlattenedLog
        }
        catch {
            Write-Warning "Failed to process log entry: $_"
        }
    }

    return  $ProcessedLogs
    
}

# Function to Update the Preview Pane
function Format-PropertyValue {
    param (
        [object]$Value,
        [int]$IndentLevel = 0
    )

    $indent = "  " * $IndentLevel  # Create indentation based on the nesting level

    if ($null -eq $Value) {
        return "${indent}N/A"
    }
    elseif ($Value -is [array]) {
        $result = @()
        foreach ($item in $Value) {
            $result += Format-PropertyValue -Value $item -IndentLevel ($IndentLevel + 1)
        }
        return $result -join "`n"
    }
    elseif ($Value -is [System.Management.Automation.PSCustomObject] -or $Value -is [System.Collections.IDictionary]) {
        $result = @()
        foreach ($key in $Value.PSObject.Properties.Name) {
            $subValue = Format-PropertyValue -Value $Value.$key -IndentLevel ($IndentLevel + 1)
            
            # If the value is an array or nested object, format it with proper indentation
            if ($Value.$key -is [array] -or $Value.$key -is [System.Management.Automation.PSCustomObject]) {
                $result += "${indent}- ${key}:`n$subValue"
            }
            else {
                $result += "${indent}- ${key}: $subValue"
            }
        }
        return $result -join "`n"
    }
    elseif ($Value -is [string]) {
        # Attempt to parse JSON strings
        try {
            $parsedValue = $Value | ConvertFrom-Json -ErrorAction Stop
            return Format-PropertyValue -Value $parsedValue -IndentLevel $IndentLevel
        }
        catch {
            # If parsing fails, treat it as a regular string
            return "${indent}$($Value)"
        }
    }
    else {
        return "${indent}$($Value.ToString())"
    }
}

function Update-PreviewPane {
    param (
        [object]$SelectedItem
    )

    if ($null -eq $SelectedItem) {
        $previewPane.Text = "No log entry selected."
        return
    }

    # Extract all properties dynamically
    $previewText = @()

    $previewText += "===== Log Entry =====`n"
    
    foreach ($key in $SelectedItem.PSObject.Properties.Name) {
        $value = $SelectedItem.$key
        $formattedValue = Format-PropertyValue -Value $value
        $previewText += "$($key):   $($formattedValue)`n"
    }

    $previewText += "`n===== Log End =====`n"

    # Update the Preview Pane
    $previewPane.Text = $previewText -join ""
}

# Function to recursively add JSON data to the TreeView
function Add-TreeNode {
    param (
        [System.Windows.Controls.TreeViewItem]$parentNode,
        [string]$key,
        $value
    )

    # Create a new TreeViewItem for the key-value pair
    $node = New-Object System.Windows.Controls.TreeViewItem

    # Display the key and value in the header
    if ($null -eq $value) {
        $node.Header = "$($key)"
    }
    else {
        $node.Header = "$($key)"
    }

    $node.Tag = $value

    # Add a tooltip to display the full value
    $node.ToolTip = if ($null -eq $value) { "Null or empty data" } else { $value.ToString() }

    if ($value -is [System.Collections.IDictionary]) {
        foreach ($subKey in $value.Keys) {
            Add-TreeNode -parentNode $node -key $subKey -value $value[$subKey]
        }
    }
    elseif ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) {
        # If the value is a collection (but not a string), add each item
        $index = 0
        foreach ($item in $value) {
            Add-TreeNode -parentNode $node -key "Item $index" -value $item
            $index++
        }
    }
    else {
        # For simple values, add a click event handler to display detailed information
        $node.AddHandler(
            [System.Windows.Controls.TreeViewItem]::MouseLeftButtonUpEvent,
            [System.Windows.RoutedEventHandler] {
                param ($eventSender, $e)
                if ($null -eq $eventSender.Tag) {
                    $detailedInfoTextBox.Text = "Null or empty data"
                }
                else {
                    $detailedInfoTextBox.Text = $eventSender.Tag | ConvertTo-Json -Depth 10
                }
            }
        )

        # Highlight search results if a search term is provided
        if (-not [string]::IsNullOrEmpty($searchBox.Text)) {
            $searchText = $searchBox.Text
            $text = if ($null -eq $value) { "" } else { $value.ToString() }

            if ($regexCheckBox.IsChecked) {
                try {
                    $matchesText = [regex]::Matches($text, $searchText)
                    $lastIndex = 0
                    $textBlock = New-Object System.Windows.Controls.TextBlock

                    foreach ($match in $matchesText) {
                        # Add text before the match
                        if ($match.Index -gt $lastIndex) {
                            $textBlock.Inlines.Add($text.Substring($lastIndex, $match.Index - $lastIndex))
                        }

                        # Add the match with highlighting
                        $run = New-Object System.Windows.Documents.Run $match.Value
                        $run.Background = [System.Windows.Media.Brushes]::Yellow
                        $run.FontWeight = "Bold"
                        $textBlock.Inlines.Add($run)

                        # Update the last index
                        $lastIndex = $match.Index + $match.Length
                    }

                    # Add the remaining text after the last match
                    if ($lastIndex -lt $text.Length) {
                        $textBlock.Inlines.Add($text.Substring($lastIndex))
                    }

                    # Set the TextBlock as the header of the TreeViewItem
                    $node.Header = $textBlock
                }
                catch {
                    # If regex is invalid, display the text without highlighting
                    $node.Header = "$($key)"
                }
            }
            else {
                # Simple text search (case-insensitive)
                $index = $text.IndexOf($searchText, [System.StringComparison]::OrdinalIgnoreCase)
                if ($index -ge 0) {
                    $textBlock = New-Object System.Windows.Controls.TextBlock

                    # Add text before the match
                    $textBlock.Inlines.Add($text.Substring(0, $index))

                    # Add the match with highlighting
                    $run = New-Object System.Windows.Documents.Run $searchText
                    $run.Background = [System.Windows.Media.Brushes]::Yellow
                    $run.FontWeight = "Bold"
                    $textBlock.Inlines.Add($run)

                    # Add the remaining text after the match
                    $textBlock.Inlines.Add($text.Substring($index + $searchText.Length))

                    # Set the TextBlock as the header of the TreeViewItem
                    $node.Header = $textBlock
                }
            }
        }
    }

    # Add the node to the parent node
    $parentNode.Items.Add($node)
}

# Function to Load Audit Log Data with Progress
function Import-AuditLogData {
    param ([object]$DataInput)

    $progressBar.Value = 0
    Update-StatusBar -Message "Loading data..."

    # Import or assign data dynamically
    if ($DataInput -is [string] -and (Test-Path $DataInput)) {
        $ParsedDataInput = Import-Csv -Path $DataInput
    }
    elseif ($DataInput -is [System.Collections.IEnumerable]) {
        $ParsedDataInput = $DataInput
    }
    else {
        Write-Host "Invalid input! Provide a valid CSV file path or in-memory data."
        exit
    }

    $logDataArray = @()
    $totalCount = $ParsedDataInput.Count
    $currentCount = 0

    # Process each record dynamically
    $ParsedDataInput | ForEach-Object {
        $currentCount++
        $progressBar.Value = ($currentCount / $totalCount) * 100
        Update-StatusBar -Message "Loading item $currentCount of $totalCount..."

        # Iterate over all columns dynamically
        $_.PSObject.Properties | ForEach-Object {
            $columnName = $_.Name
            $columnValue = $_.Value

            # Try parsing any JSON-like column dynamically
            if ($columnValue -is [string] -and $columnValue.Trim().StartsWith("{")) {
                try {
                    $_.Value = ConvertFrom-Json $columnValue -ErrorAction Stop
                }
                catch {
                    Write-Warning "Failed to parse JSON for column: $columnName"
                }
            }
        }

        $logDataArray += $_
    }

    $progressBar.Value = 100
    Update-StatusBar -Message "Data loaded successfully!"

    return $logDataArray
}

# Function to update the selected RecordTypes
function Update-SelectedRecordTypes {
    $selectedRecordTypes = $recordTypeCheckBoxPanel.Children | Where-Object { $_.IsChecked -eq $true } | ForEach-Object { $_.Content }
    $recordTypeFilter.Text = $selectedRecordTypes -join ", "
    Update-TreeView
}

function Update-SelectedOperations {
    $selectedOperations = $operationsCheckBoxPanel.Children | Where-Object { $_.IsChecked -eq $true } | ForEach-Object { $_.Content }
    $operationsFilter.Text = $selectedOperations -join ", "
    Update-TreeView
}

# Treeview filtering
function Update-Filters {
    # Clear existing items
    $recordTypeFilter.Items.Clear()
    $operationsFilter.Items.Clear()

    # Clear existing CheckBoxes
    $recordTypeCheckBoxPanel.Children.Clear()
    $operationsCheckBoxPanel.Children.Clear()

    # Add "All" option to RecordType dropdown
    $script:allRecordTypeCheckBox = New-Object System.Windows.Controls.CheckBox
    $script:allRecordTypeCheckBox.Content = "All"
    $script:allRecordTypeCheckBox.Margin = "2"
    $script:allRecordTypeCheckBox.IsChecked = $true  # Default to checked

    $script:allRecordTypeCheckBox.Add_Checked({
            if (-not $script:isUpdatingCheckBoxes) {
                $script:isUpdatingCheckBoxes = $true
                foreach ($checkBox in $recordTypeCheckBoxPanel.Children) {
                    if ($checkBox -ne $script:allRecordTypeCheckBox) {
                        $checkBox.IsChecked = $false
                    }
                }
                Update-SelectedRecordTypes
                $script:isUpdatingCheckBoxes = $false
            }
        })

    $script:allRecordTypeCheckBox.Add_Unchecked({
            if (-not $script:isUpdatingCheckBoxes) {
                $script:isUpdatingCheckBoxes = $true
                Update-SelectedRecordTypes
                $script:isUpdatingCheckBoxes = $false
            }
        })

    # Add the "All" CheckBox for RecordType to the StackPanel
    $recordTypeCheckBoxPanel.Children.Add($script:allRecordTypeCheckBox) | Out-File

    # Get unique RecordType values
    $uniqueRecordTypes = $logDataArray | ForEach-Object { $_.RecordType } | Sort-Object -Unique

    # Add unique RecordType values with CheckBoxes
    foreach ($recordType in $uniqueRecordTypes) {
        $checkBox = New-Object System.Windows.Controls.CheckBox
        $checkBox.Content = $recordType
        $checkBox.Margin = "2"

        $checkBox.Add_Checked({
                if (-not $script:isUpdatingCheckBoxes) {
                    $script:isUpdatingCheckBoxes = $true
                    $script:allRecordTypeCheckBox.IsChecked = $false  # Ensure "All" is unchecked
                    Update-SelectedRecordTypes
                    $script:isUpdatingCheckBoxes = $false
                }
            })

        $checkBox.Add_Unchecked({
                if (-not $script:isUpdatingCheckBoxes) {
                    $script:isUpdatingCheckBoxes = $true
                    if (-not ($recordTypeCheckBoxPanel.Children | Where-Object { $_.IsChecked -eq $true -and $_.Content -ne "All" })) {
                        $script:allRecordTypeCheckBox.IsChecked = $true  # Recheck "All" if no others are checked
                    }
                    Update-SelectedRecordTypes
                    $script:isUpdatingCheckBoxes = $false
                }
            })

        $recordTypeCheckBoxPanel.Children.Add($checkBox) | out-null
    }

    $recordTypeFilter.Items.Add($recordTypeCheckBoxPanel)

    # Add "All" option to Operations dropdown
    $script:allOperationsCheckBox = New-Object System.Windows.Controls.CheckBox
    $script:allOperationsCheckBox.Content = "All"
    $script:allOperationsCheckBox.Margin = "2"
    $script:allOperationsCheckBox.IsChecked = $true  # Default to checked

    $script:allOperationsCheckBox.Add_Checked({
            if (-not $script:isUpdatingCheckBoxes) {
                $script:isUpdatingCheckBoxes = $true
                foreach ($checkBox in $operationsCheckBoxPanel.Children) {
                    if ($checkBox -ne $script:allOperationsCheckBox) {
                        $checkBox.IsChecked = $false
                    }
                }
                Update-SelectedOperations
                $script:isUpdatingCheckBoxes = $false
            }
        })

    $script:allOperationsCheckBox.Add_Unchecked({
            if (-not $script:isUpdatingCheckBoxes) {
                $script:isUpdatingCheckBoxes = $true
                Update-SelectedOperations
                $script:isUpdatingCheckBoxes = $false
            }
        })

    # Add the "All" CheckBox for Operations to the StackPanel
    $operationsCheckBoxPanel.Children.Add($script:allOperationsCheckBox) | Out-Null

    # Get unique Operations values
    $uniqueOperations = $logDataArray | ForEach-Object { $_.Operations } | Sort-Object -Unique

    # Add unique Operations values with CheckBoxes
    foreach ($operation in $uniqueOperations) {
        $checkBox = New-Object System.Windows.Controls.CheckBox
        $checkBox.Content = $operation
        $checkBox.Margin = "2"

        $checkBox.Add_Checked({
                if (-not $script:isUpdatingCheckBoxes) {
                    $script:isUpdatingCheckBoxes = $true
                    $script:allOperationsCheckBox.IsChecked = $false  # Ensure "All" is unchecked
                    Update-SelectedOperations
                    $script:isUpdatingCheckBoxes = $false
                }
            })

        $checkBox.Add_Unchecked({
                if (-not $script:isUpdatingCheckBoxes) {
                    $script:isUpdatingCheckBoxes = $true
                    if (-not ($operationsCheckBoxPanel.Children | Where-Object { $_.IsChecked -eq $true -and $_.Content -ne "All" })) {
                        $script:allOperationsCheckBox.IsChecked = $true  # Recheck "All" if no others are checked
                    }
                    Update-SelectedOperations
                    $script:isUpdatingCheckBoxes = $false
                }
            })

        $operationsCheckBoxPanel.Children.Add($checkBox) | out-null
    }

    $operationsFilter.Items.Add($operationsCheckBoxPanel)
}

# Function to Filter TreeView
function Update-TreeView {
    # Get selected RecordTypes, excluding "All" if it's checked
    $selectedRecordTypes = $recordTypeCheckBoxPanel.Children | Where-Object { $_.IsChecked -eq $true -and $_.Content -ne "All" } | ForEach-Object { $_.Content }
    $isAllRecordTypesChecked = ($recordTypeCheckBoxPanel.Children | Where-Object { $_.Content -eq "All" }).IsChecked

    # Get selected Operations, excluding "All" if it's checked
    $selectedOperations = $operationsCheckBoxPanel.Children | Where-Object { $_.IsChecked -eq $true -and $_.Content -ne "All" } | ForEach-Object { $_.Content }
    $isAllOperationsChecked = ($operationsCheckBoxPanel.Children | Where-Object { $_.Content -eq "All" }).IsChecked

    # Combine date and time for start and end filters
    
    $startDate = $startDatePicker.SelectedDate
    $startTime = $startTimeComboBox.Text
    $endDate = $endDatePicker.SelectedDate
    $endTime = $endTimeComboBox.Text

    $script:exportFilteredLogData = @()

    # Parse start and end DateTime
    $startDateTime = $null
    $endDateTime = $null

    if ($null -ne $startDate -and -not [string]::IsNullOrEmpty($startTime)) {
        try {
            $startDateTime = [datetime]::ParseExact("$($startDate.ToString('MM/dd/yyyy')) $startTime", "MM/dd/yyyy HH:mm:ss", $null)
        }
        catch {
            Write-Warning "Invalid start date or time format."
        }
    }

    if ($null -ne $endDate -and -not [string]::IsNullOrEmpty($endTime)) {
        try {
            $endDateTime = [datetime]::ParseExact("$($endDate.ToString('MM/dd/yyyy')) $endTime", "MM/dd/yyyy HH:mm:ss", $null)
        }
        catch {
            Write-Warning "Invalid end date or time format."
        }
    }

    # Clear the TreeView before populating it
    $treeView.Items.Clear()

    # Iterate through log data and apply filters
    foreach ($logData in $logDataArray) {
        $logDate = $null

        # Parse the log entry's CreationDate
        if (-not [string]::IsNullOrEmpty($logData.CreationDate)) {
            $dateFormats = @("M/d/yyyy h:mm:ss tt", "MM/dd/yyyy HH:mm:ss", "yyyy-MM-ddTHH:mm:ss", "yyyy/MM/dd HH:mm:ss")

            $parsedSuccessfully = $false
            foreach ($format in $dateFormats) {
                try {
                    $logDate = [datetime]::ParseExact($logData.CreationDate, $format, $null)
                    $parsedSuccessfully = $true
                    break  # Exit loop once parsing succeeds
                }
                catch {
                    continue  # Try the next format
                }
            }

            if (-not $parsedSuccessfully) {
                Write-Warning "Failed to parse CreationDate for entry: $($logData.CreationDate)"
                continue  # Skip this entry if parsing fails
            }
        }

        # Apply filters
        $matchesRecordType = $isAllRecordTypesChecked -or $logData.RecordType -in $selectedRecordTypes
        $matchesOperation = $isAllOperationsChecked -or $logData.Operations -in $selectedOperations
        $matchesSearch = ($logData.AuditData -match $searchBox.Text -or $logData.ResultIndex -match $searchBox.Text)
        $matchesDateRange = ($null -eq $startDateTime -or $logDate -ge $startDateTime) -and ($null -eq $endDateTime -or $logDate -le $endDateTime)

        # If all filters match, add the log entry to the TreeView
        if ($matchesRecordType -and $matchesOperation -and $matchesSearch -and $matchesDateRange) {

            $script:exportFilteredLogData += $logData

            $entryNode = New-Object System.Windows.Controls.TreeViewItem
            $entryNode.Header = "$($logData.RecordType) - $($logData.Operations)"
            $entryNode.Tag = $logData  # Store the log entry data in the Tag property
            $treeView.Items.Add($entryNode)

            foreach ($key in $logData.PSObject.Properties.Name) {
                $value = $logData.$key
            
                # Check if the value is a JSON string and try converting it
                if ($value -is [string]) {
                    try {
                        $parsedValue = ConvertFrom-Json $value -ErrorAction Stop
                        $value = $parsedValue
                    }
                    catch {
                        # continue
                    }
                }
            
                # If the parsed value is an object, iterate over its properties
                if ($value -is [PSCustomObject]) {
                    $node = New-Object System.Windows.Controls.TreeViewItem
                    $node.Header = $key
                    $entryNode.Items.Add($node)
            
                    foreach ($subKey in $value.PSObject.Properties.Name) {
                        $subValue = $value.$subKey
            
                        # Special handling for collections (e.g., "Parameters")
                        if ($subValue -is [System.Collections.IEnumerable] -and $subValue -isnot [string]) {
                            $subNode = New-Object System.Windows.Controls.TreeViewItem
                            $subNode.Header = $subKey
                            $node.Items.Add($subNode)
            
                            foreach ($item in $subValue) {
                                try {
                                    $itemValue = ConvertFrom-Json $item.Value -ErrorAction Stop
                                }
                                catch {
                                    $itemValue = $item.Value
                                }
                                Add-TreeNode -parentNode $subNode -key $item.Name -value $itemValue
                            }
                        }
                        else {
                            Add-TreeNode -parentNode $node -key $subKey -value $subValue
                        }
                    }
                }
                else {
                    Add-TreeNode -parentNode $entryNode -key $key -value $value
                }
            }
        }
    }
}

# Function to convert System.Drawing.Color to WPF Brush
function ConvertTo-Brush {
    param ($drawingColor)
    return New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromArgb($drawingColor.A, $drawingColor.R, $drawingColor.G, $drawingColor.B))
}


function Update-UITheme {
    param ($theme)

    # Apply theme to main window
    $window.Background = $theme.Background

    # Recursively apply theme to all controls in the window
    Update-UIThemeToControl -control $window.Content -theme $theme
}

function Update-UIThemeToControl {
    param ($control, $theme)

    # Apply theme to the current control
    if ($control -is [System.Windows.Controls.Control]) {
        # Set background for non-label controls
        if ($control -isnot [System.Windows.Controls.Label]) {
            $control.Background = $theme.Background
        }

        # Set foreground for text-based controls
        if ($control -is [System.Windows.Controls.TextBox] -or
            $control -is [System.Windows.Controls.RichTextBox] -or
            $control -is [System.Windows.Controls.TextBlock] -or
            $control -is [System.Windows.Controls.Label] -or
            $control -is [System.Windows.Controls.Button] -or
            $control -is [System.Windows.Controls.TreeViewItem]) {
            $control.Foreground = $theme.Foreground
        }

        # Set background for buttons
        if ($control -is [System.Windows.Controls.Button]) {
            $control.Background = $theme.ButtonBackground
        }

        # Set background and foreground for TreeView and TreeViewItem
        if ($control -is [System.Windows.Controls.TreeView]) {
            $control.Background = $theme.Background
            $control.Foreground = $theme.Foreground
        }
        if ($control -is [System.Windows.Controls.TreeViewItem]) {
            $control.Background = $theme.Background
            $control.Foreground = $theme.Foreground
        }
    }

    # If the control is a container, recursively apply the theme to its children
    if ($control -is [System.Windows.Controls.Panel]) {
        foreach ($child in $control.Children) {
            Update-UIThemeToControl -control $child -theme $theme
        }
    }
    elseif ($control -is [System.Windows.Controls.ContentControl]) {
        if ($control.Content -is [System.Windows.Controls.Control]) {
            Update-UIThemeToControl -control $control.Content -theme $theme
        }
    }
    elseif ($control -is [System.Windows.Controls.ItemsControl]) {
        foreach ($item in $control.Items) {
            if ($item -is [System.Windows.Controls.Control]) {
                Update-UIThemeToControl -control $item -theme $theme
            }
        }
    }

    # Handle TreeView and TreeViewItem specifically
    if ($control -is [System.Windows.Controls.TreeView]) {
        foreach ($item in $control.Items) {
            if ($item -is [System.Windows.Controls.TreeViewItem]) {
                Update-UIThemeToControl -control $item -theme $theme
            }
        }
    }
    if ($control -is [System.Windows.Controls.TreeViewItem]) {
        foreach ($item in $control.Items) {
            if ($item -is [System.Windows.Controls.TreeViewItem]) {
                Update-UIThemeToControl -control $item -theme $theme
            }
        }
    }
}


# Add Expand/Collapse Buttons
# Function to set expansion state recursively
function Set-ExpansionState($items, $state) {
    foreach ($item in $items) {
        $item.IsExpanded = $state
        if ($item.Items.Count -gt 0) {
            Set-ExpansionState $item.Items $state
        }
    }
}



# Function to check if the current WPF theme is dark or light
function Get-WpfCurrentTheme {
    # Get the current application's resource dictionary
    $appResources = [System.Windows.Application]::Current.Resources

    # Check if a dark theme resource is applied (this is a common approach, customize it as needed)
    if ($appResources.MergedDictionaries -match "Dark") {
        return "Dark"
    }
    elseif ($appResources.MergedDictionaries -match "Light") {
        return "Light"
    }
    return "Unknown"
}

# Function to generate a random color based on the current theme (light/dark)
function Get-RandomColorBasedOnTheme {
    param(
        [ValidateSet("Dark", "Light")]
        [string]$currentTheme
    )

    # Random color generator logic
    $random = New-Object System.Random
    $r = $random.Next(0, 256)
    $g = $random.Next(0, 256)
    $b = $random.Next(0, 256)

    if ($currentTheme -eq "Dark") {
        # Dark theme: Keep colors darker
        $r = [math]::Min($r, 100)
        $g = [math]::Min($g, 100)
        $b = [math]::Min($b, 100)
    }
    elseif ($currentTheme -eq "Light") {
        # Light theme: Keep colors lighter
        $r = [math]::Max($r, 150)
        $g = [math]::Max($g, 150)
        $b = [math]::Max($b, 150)
    }

    # Return the SolidColorBrush with the random color
    return New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb($r, $g, $b))
}


################## Parameters #####################
$buttonWidth = 105
$buttonHeight = 25
$uIMargin = 5
# Define themes with both Background & Text color
$themes = @(
    @{Background = [System.Windows.Media.Brushes]::White; Foreground = [System.Windows.Media.Brushes]::Black},    # Light Mode
    @{Background = [System.Windows.Media.Brushes]::Black; Foreground = [System.Windows.Media.Brushes]::White},    # Dark Mode
    @{Background = [System.Windows.Media.Brushes]::Gray; Foreground = [System.Windows.Media.Brushes]::Yellow},    # High Contrast
    @{Background = [System.Windows.Media.Brushes]::Navy; Foreground = [System.Windows.Media.Brushes]::LightGray},  # Custom Navy Theme
    @{Background = [System.Windows.Media.Brushes]::Beige; Foreground = [System.Windows.Media.Brushes]::DarkSlateGray}, # Solarized Light
    @{Background = [System.Windows.Media.Brushes]::DarkSlateBlue; Foreground = [System.Windows.Media.Brushes]::LightSlateGray}, # Solarized Dark
    @{Background = [System.Windows.Media.Brushes]::Black; Foreground = [System.Windows.Media.Brushes]::GhostWhite},  # Monokai
    @{Background = [System.Windows.Media.Brushes]::DarkViolet; Foreground = [System.Windows.Media.Brushes]::WhiteSmoke},  # Dracula
    @{Background = [System.Windows.Media.Brushes]::Black; Foreground = [System.Windows.Media.Brushes]::DarkOrange},  # One Dark
    @{Background = [System.Windows.Media.Brushes]::AntiqueWhite; Foreground = [System.Windows.Media.Brushes]::DarkOliveGreen} # Gruvbox Light
)


# Track current theme index
$script:currentThemeIndex = 0

$global:logDataArray = @()


############################ Window Configuration and  Drop event handler #################

$window = New-Object System.Windows.Window
$window.Title = "MAuditMaster 365 - Microsoft 365 Unified Audit Log Viewer"
$window.Width = 1400
$window.Height = 900
$window.MinWidth = 800
$window.MinHeight = 500
$window.WindowStartupLocation = "CenterScreen"
$window.AllowDrop = $true
$window.FontFamily = "Segoe UI"
$window.FontSize = 13   # Slightly larger for readability
$window.FontWeight = "Normal"
$window.ResizeMode = "CanResize"
$window.BorderThickness = 1
$window.BorderBrush = [System.Windows.Media.Brushes]::Gray

$window.Add_DragEnter({
        param ($eventSender, $e)

        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            $e.Effects = [System.Windows.DragDropEffects]::Copy
        }
        else {
            $e.Effects = [System.Windows.DragDropEffects]::None
        }
    })

$window.Add_Drop({
        param ($eventSender, $e)

        $filePaths = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)

        if ($filePaths.Count -gt 0) {
            $filePath = $filePaths[0]

            if (Test-ValidFile -FilePath $filePath) {
                $global:logDataArray = Import-DataFromFile -FilePath $filePath

                if ($null -ne $global:logDataArray) {
                    Update-Filters
                    Update-TreeView
                    Update-StatusBar -Message "File loaded successfully: $filePath"
                }
            }
        }
    })


# Styling for a modern UI feel
$shadowEffect = New-Object System.Windows.Media.Effects.DropShadowEffect
$shadowEffect.BlurRadius = 10
$shadowEffect.ShadowDepth = 5
$shadowEffect.Opacity = 0.4


######################### Create Grid ###########################
$grid = New-Object System.Windows.Controls.Grid
$grid.Margin = "10"  
$window.Content = $grid 

# Define columns for the main Grid (percentage-based)
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "3*" }))  # Column 0: TreeView (25%)
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "5*" }))  # Column 1: Preview Pane (50%)
$grid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "3*" }))  # Column 2: Detailed Info Pane (25%)

# Define rows for the main Grid
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 0: Header
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))      # Row 1: Main Content (stretches to fill remaining space)
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 2: Connect Online
$grid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 2: Footer


################# Status bar and progress bar ###################

# Create a Border to wrap the Grid and add a border effect
$borderContainer = New-Object System.Windows.Controls.Border
$borderContainer.Margin = $uIMargin
$borderContainer.BorderBrush = [System.Windows.Media.Brushes]::Gray
$borderContainer.BorderThickness = "1"

# Create a Grid for the status bar and progress bar
$statusBarGrid = New-Object System.Windows.Controls.Grid
$statusBarGrid.VerticalAlignment = "Bottom"
$statusBarGrid.HorizontalAlignment = "Stretch"
$statusBarGrid.Background = [System.Windows.Media.Brushes]::LightGray

# Define columns for the status bar grid
$statusBarGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))  # Column 0: Status text
$statusBarGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))    # Column 1: Spacer
$statusBarGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))  # Column 2: Progress bar

# Add the status bar grid to the border container
$borderContainer.Child = $statusBarGrid

# Add the border container to the main grid
$grid.Children.Add($borderContainer) | out-null
[System.Windows.Controls.Grid]::SetRow($borderContainer, 3)
[System.Windows.Controls.Grid]::SetColumnSpan($borderContainer, 3)  # Span across all columns

# Create the status text block with improved design
$statusBarText = New-Object System.Windows.Controls.TextBox
$statusBarText.Margin = $uIMargin
$statusBarText.Text = "Unified log analyzer application started successfully."
$statusBarText.VerticalAlignment = "Center"
$statusBarText.HorizontalAlignment = "Left"
$statusBarText.FontSize = "14"
$statusBarText.FontWeight = "Bold"
$statusBarText.Padding = "5,0,0,0"
$statusBarText.IsReadOnly = $true
$statusBarText.Background = [System.Windows.Media.Brushes]::Transparent  # Remove background color
$statusBarText.BorderBrush = [System.Windows.Media.Brushes]::Transparent  # Remove outline color
$statusBarText.BorderThickness = 0  # Remove outline thickness


$statusBarGrid.Children.Add($statusBarText) | out-null
[System.Windows.Controls.Grid]::SetRow($statusBarText, 2)
[System.Windows.Controls.Grid]::SetColumn($statusBarText, 0)

# Create a ProgressBar for loading/processing status with better styling
$progressBar = New-Object System.Windows.Controls.ProgressBar
$progressBar.Height = $buttonHeight
$progressBar.Margin = $uIMargin
$progressBar.VerticalAlignment = "Center"
$progressBar.HorizontalAlignment = "Stretch"
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Visibility = "Visible"  # Visible by default
$progressBar.Foreground = [System.Windows.Media.Brushes]::DarkBlue
$progressBar.Background = [System.Windows.Media.Brushes]::LightSteelBlue
$progressBar.BorderBrush = [System.Windows.Media.Brushes]::DarkSlateGray
$progressBar.BorderThickness = "1"
$progressBar.Height = 25

# Add the ProgressBar to the last column of the status bar grid
$statusBarGrid.Children.Add($progressBar) | out-null
[System.Windows.Controls.Grid]::SetColumn($progressBar, 2)



############ Tree View Configuration #####################

# Create the TreeView Grid for the section
$treeViewGrid = New-Object System.Windows.Controls.Grid
$treeViewGrid.Margin = $uIMargin
$treeViewGrid.VerticalAlignment = "Stretch"
$treeViewGrid.HorizontalAlignment = "Stretch"

# Define rows for the Grid
$treeViewGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) # Row 0: Label and CheckBox
$treeViewGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))    # Row 1: TreeView (fills remaining space)

# Add the TreeView Grid to the main Grid
$grid.Children.Add($treeViewGrid) | out-null
[System.Windows.Controls.Grid]::SetRow($treeViewGrid, 1)
[System.Windows.Controls.Grid]::SetColumn($treeViewGrid, 0)

# Create the Label and Button Grid
$treeViewLabelGrid = New-Object System.Windows.Controls.Grid
$treeViewLabelGrid.HorizontalAlignment = "Stretch"
$treeViewLabelGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))  # Column 0: Label
$treeViewLabelGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))     # Column 1: Spacer
$treeViewLabelGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))  # Column 2: CheckBox Panel

# Add Label to the TreeView Label Grid
$treeViewLabel = New-Object System.Windows.Controls.Label
$treeViewLabel.Content = "Audit Data:"
$treeViewLabel.FontSize = 18
$treeViewLabel.FontWeight = "Bold"
$treeViewLabelGrid.Children.Add($treeViewLabel) | out-null
[System.Windows.Controls.Grid]::SetColumn($treeViewLabel, 0)


# Create a StackPanel for Checkboxes
$OnlineDataCheckBoxPanel = New-Object System.Windows.Controls.StackPanel
$OnlineDataCheckBoxPanel.Orientation = "Horizontal"
$OnlineDataCheckBoxPanel.Margin = 3
$OnlineDataCheckBoxPanel.HorizontalAlignment = "Right"
$OnlineDataCheckBoxPanel.Background = "black"
$OnlineDataCheckBoxPanel.HorizontalAlignment = "Right"  
$OnlineDataCheckBoxPanel.VerticalAlignment = "Center" 
$OnlineDataCheckBoxPanel.ToolTip = "Open Unified Audit Panel, if inactive, Please, click on 'Connect EXO'."


$dropShadowEffect = New-Object System.Windows.Media.Effects.DropShadowEffect
$dropShadowEffect.Color = [System.Windows.Media.Color]::FromArgb(255, 0, 0, 0)
$dropShadowEffect.Direction = 330
$dropShadowEffect.ShadowDepth = 2
$dropShadowEffect.BlurRadius = 2
$OnlineDataCheckBoxPanel.Effect = $dropShadowEffect


# Add CheckBox to the CheckBox Panel
$OnlineDataCheckBox = New-Object System.Windows.Controls.CheckBox
$OnlineDataCheckBox.ToolTip = "Show Online Pane"
$OnlineDataCheckBox.Margin = $uIMargin
$OnlineDataCheckBox.VerticalAlignment = "Center"
$OnlineDataCheckBoxPanel.Children.Add($OnlineDataCheckBox) | out-null
$OnlineDataCheckBox.Add_Checked({
        # Show the UI element and update the status bar
        if (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue) {
            Set-UIElementVisibility -UIElement $auditSearchContainer -Action Show
            $connectEXOButton.Content = "Disconnect EXO"
            Update-StatusBar -Message "Open Unified Audit Panel opened and Exchange Online is already connected."
        }
        else {
            Update-StatusBar -Message "The Open Unified Audit Panel is automatically close because Exchange Online PowerShell is not connected. Click on Connect EXO button."
            $OnlineDataCheckBox.IsChecked = $false
            $OnlineDataCheckBox.IsEnabled = $false  # Disable interaction until Exchange Online is connected
            $OnlineDataCheckBox.ToolTip = "Please click on Connect EXO button to connect"
        
            Set-UIElementVisibility -UIElement $auditSearchContainer -Action Hide
        }
    })
$OnlineDataCheckBox.Add_Unchecked({
        Set-UIElementVisibility -UIElement $auditSearchContainer -Action Hide
        Update-StatusBar -Message "The Unified Audit Panel is temporarily minimized. You can reopen it anytime. Exchange Online is already connected."
        if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
            Update-StatusBar -Message "Exchange Online is not connected. Please click on Connect EXO button to connect."
        }
    })



# Add Label to the CheckBox Panel
$OnlineDataCheckBoxPanelLabel = New-Object System.Windows.Controls.Label
$OnlineDataCheckBoxPanelLabel.Content = "Online Data"
$OnlineDataCheckBoxPanelLabel.MinWidth = 100  # Set a minimum width
$OnlineDataCheckBoxPanelLabel.HorizontalAlignment = "Stretch"
$OnlineDataCheckBoxPanelLabel.Foreground = "white"
$OnlineDataCheckBoxPanelLabel.Margin = "0,0,0,2"
$OnlineDataCheckBoxPanelLabel.FontWeight = "Bold"
$OnlineDataCheckBoxPanelLabel.VerticalAlignment = "Center"
$OnlineDataCheckBoxPanelLabel.FontSize = 14
$OnlineDataCheckBoxPanelLabel.VerticalContentAlignment = "center"
$OnlineDataCheckBoxPanelLabel.HorizontalContentAlignment = "Left"

$treeViewLabelGrid.Children.Add($OnlineDataCheckBoxPanel) | out-null
$OnlineDataCheckBoxPanel.Children.Add($OnlineDataCheckBoxPanelLabel) | out-null
[System.Windows.Controls.Grid]::SetColumn($OnlineDataCheckBoxPanel, 1)

# Add the TreeView Label Grid to the TreeView Grid
$treeViewGrid.Children.Add($treeViewLabelGrid) | out-null
[System.Windows.Controls.Grid]::SetRow($treeViewLabelGrid, 0)

# Create the TreeView
$treeView = New-Object System.Windows.Controls.TreeView
$treeView.ToolTip = "Browse the audit log data hierarchically."
$treeView.VerticalAlignment = "Stretch"  # Ensure TreeView stretches vertically
$treeView.HorizontalAlignment = "Stretch"
$treeView.Background = "LightYellow"
$treeView.Padding = "10"
$treeView.Margin = $uIMargin
$treeView.FontFamily = "Segoe UI"
$treeView.FontSize = "12"
$treeView.BorderBrush = "DarkGreen"
$treeView.BorderThickness = "1"


# Enable virtualization for TreeView
$treeView.ItemsPanel = New-Object System.Windows.Controls.ItemsPanelTemplate -ArgumentList @(
    [System.Windows.FrameworkElementFactory]::new([System.Windows.Controls.VirtualizingStackPanel])
    $treeView.SetValue([System.Windows.Controls.VirtualizingStackPanel]::IsVirtualizingProperty, $true)
    $treeView.SetValue([System.Windows.Controls.VirtualizingStackPanel]::VirtualizationModeProperty, [System.Windows.Controls.VirtualizationMode]::Recycling))

# Add the TreeView to Row 1 of the TreeView Grid
$treeViewGrid.Children.Add($treeView) | out-null
[System.Windows.Controls.Grid]::SetRow($treeView, 1)

# Event Handler for TreeView Selection Changed
$treeView.Add_SelectedItemChanged({
        $selectedItem = $treeView.SelectedItem

        if ($selectedItem -and $selectedItem.Tag) {
            Update-PreviewPane -SelectedItem $selectedItem.Tag
        }
        else {
            $previewPane.Text = "No log entry selected."
        }
    })

$treeView.Add_MouseDoubleClick({
        param ($eventSender, $e)

        if ($treeView.Items.Count -eq 0) {
            $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
            $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|JSON Files (*.json)|*.json"
            $openFileDialog.Title = "Select a CSV or JSON file to load"

            if ($openFileDialog.ShowDialog() -eq $true) {
                $filePath = $openFileDialog.FileName

                if (Test-ValidFile -FilePath $filePath) {
                    $global:logDataArray = Import-DataFromFile -FilePath $filePath

                    if ($null -ne $global:logDataArray) {
                        Update-Filters
                        Update-TreeView
                        Update-StatusBar -Message "File loaded successfully: $filePath"
                    }
                }
            }
        }
    })


############ Preview Pane Configuration ##################

# Create the Preview Pane Grid
$previewPaneGrid = New-Object System.Windows.Controls.Grid
$previewPaneGrid.Margin = $uIMargin
$previewPaneGrid.VerticalAlignment = "Stretch"
$previewPaneGrid.HorizontalAlignment = "Stretch"

# Define rows for the Grid
$previewPaneGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) # Row 0: Label
$previewPaneGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))    # Row 1: TextBox (fills remaining space)

# Add the Preview Pane Grid to the main Grid
$grid.Children.Add($previewPaneGrid) | out-null
[System.Windows.Controls.Grid]::SetRow($previewPaneGrid, 1)
[System.Windows.Controls.Grid]::SetColumn($previewPaneGrid, 1)

# Create a Grid for the Label and Copy Button
$labelButtonGrid = New-Object System.Windows.Controls.Grid
$labelButtonGrid.HorizontalAlignment = "Stretch"  # Ensure it stretches horizontally

# Define columns for the Label and Button Grid
$labelButtonGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" })) # Column 0: Label (left-aligned)
$labelButtonGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))    # Column 1: Spacer (fills remaining space)
$labelButtonGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" })) # Column 2: Button (right-aligned)

# Create the Label
$previewPaneLabel = New-Object System.Windows.Controls.Label
$previewPaneLabel.Content = "Audit Preview:"
$previewPaneLabel.HorizontalAlignment = "Left"
$previewPaneLabel.FontSize = 18
$previewPaneLabel.FontWeight = "Bold"
$previewPaneLabel.Margin = "0,0,10,0"  # Add some margin between the Label and the Copy Button

# Add the Label to Column 0 of the Label and Button Grid
$labelButtonGrid.Children.Add($previewPaneLabel) | out-null
[System.Windows.Controls.Grid]::SetColumn($previewPaneLabel, 0)

# Add a Copy Button next to the Preview Pane
$copyButton = New-Object System.Windows.Controls.Button
$copyButton.Content = "Copy"
$copyButton.FontWeight = "Bold"
$copyButton.Width = $buttonWidth
$copyButton.Height = $buttonHeight
$copyButton.Margin = $uIMargin
$copyButton.HorizontalAlignment = "Right"
$copyButton.ToolTip = "Copy the preview text to the clipboard"
$copyButton.Add_Click({

        $previewPane.SelectAll()
        $previewPane.Copy()
        $previewPane.SelectionLength = 0
        Update-StatusBar -Message "Text copied to clipboard."
    })

# Create the Preview Pane (TextBox)
$previewPane = New-Object System.Windows.Controls.TextBox
$previewPane.IsReadOnly = $true
$previewPane.VerticalScrollBarVisibility = "Auto"
$previewPane.HorizontalScrollBarVisibility = "Auto"
$previewPane.ToolTip = "Preview of the selected log entry."
$previewPane.VerticalAlignment = "Stretch" 
$previewPane.HorizontalAlignment = "Stretch" 
$previewPane.Padding = "10"
$previewPane.Margin = $uIMargin

$previewPane.BorderBrush = [System.Windows.Media.Brushes]::Gray
$previewPane.BorderThickness = 1
$previewPane.Background = [System.Windows.Media.Brushes]::White  # Subtle background color for contrast

# Add the Copy Button to Column 2 of the Label and Button Grid
$labelButtonGrid.Children.Add($copyButton) | out-null
[System.Windows.Controls.Grid]::SetColumn($copyButton, 2)
$previewPaneGrid.Children.Add($labelButtonGrid) | out-null
[System.Windows.Controls.Grid]::SetRow($labelButtonGrid, 0)
$previewPaneGrid.Children.Add($previewPane) | out-null
[System.Windows.Controls.Grid]::SetRow($previewPane, 1)



############ Detailed Info Pane Configuration ##################

# Create the Detailed Info Pane Grid
$detailedInfoPaneGrid = New-Object System.Windows.Controls.Grid
$detailedInfoPaneGrid.Margin = $uIMargin
$detailedInfoPaneGrid.VerticalAlignment = "Stretch"
$detailedInfoPaneGrid.HorizontalAlignment = "Stretch"

# Define rows for the Grid
$detailedInfoPaneGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" })) # Row 0: Label
$detailedInfoPaneGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "*" }))    # Row 1: TextBox (fills remaining space)

# Add the Detailed Info Pane Grid to the main Grid
$grid.Children.Add($detailedInfoPaneGrid) | out-null
[System.Windows.Controls.Grid]::SetRow($detailedInfoPaneGrid, 1)
[System.Windows.Controls.Grid]::SetColumn($detailedInfoPaneGrid, 2)

# Create the Label
$detailedInfoPaneLabel = New-Object System.Windows.Controls.Label
$detailedInfoPaneLabel.Content = "Audit Detail:"
$detailedInfoPaneLabel.FontSize = 18
$detailedInfoPaneLabel.FontWeight = "Bold"

# Add the Label to Row 0 of the Detailed Info Pane Grid
$detailedInfoPaneGrid.Children.Add($detailedInfoPaneLabel) | out-null
[System.Windows.Controls.Grid]::SetRow($detailedInfoPaneLabel, 0)

# Create the Detailed Info Pane (TextBox)
$detailedInfoTextBox = New-Object System.Windows.Controls.TextBox
$detailedInfoTextBox.IsReadOnly = $true
$detailedInfoTextBox.VerticalAlignment = "Stretch" 
$detailedInfoTextBox.HorizontalAlignment = "Stretch" 
$detailedInfoTextBox.VerticalScrollBarVisibility = "Auto"
$detailedInfoTextBox.HorizontalScrollBarVisibility = "Auto"
$detailedInfoTextBox.Padding = "10"
$detailedInfoTextBox.Margin = $uIMargin
$detailedInfoTextBox.ToolTip = "View detailed information about the selected item."
$detailedInfoTextBox.TextWrapping = "Wrap"
$detailedInfoTextBox.BorderBrush = [System.Windows.Media.Brushes]::Gray
$detailedInfoTextBox.BorderThickness = 1
$detailedInfoTextBox.Background = [System.Windows.Media.Brushes]::WhiteSmoke

# Add the TextBox to Row 1 of the Detailed Info Pane Grid
$detailedInfoPaneGrid.Children.Add($detailedInfoTextBox) | out-null
[System.Windows.Controls.Grid]::SetRow($detailedInfoTextBox, 1)

################### Filter by Search ###################

# Add a Grid to hold the search components
$searchPanel = New-Object System.Windows.Controls.Grid
$searchPanel.Margin = "10"
$searchPanel.VerticalAlignment = "Center"
$searchPanel.HorizontalAlignment = "Stretch"  # Stretch to fill available space

# Define columns for the search panel
$searchPanel.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "*" }))  # Column 0: Search Box (fills remaining space)
$searchPanel.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = "Auto" }))  # Column 1: Filter Button (fits content)

# Add Search Box
$searchBox = New-Object System.Windows.Controls.TextBox
$searchBox.Width = [double]::NaN  # Auto-size width to fill available space
$searchBox.Height = 30
$searchBox.Margin = $uIMargin
$searchBox.VerticalAlignment = "Center"
$searchBox.HorizontalAlignment = "Stretch"  
$searchBox.TextAlignment = "Left"
$searchBox.FontSize = "14"
$searchBox.ToolTip = "Enter a keyword to filter log entries."
$searchBox.Background = "White"
$searchBox.BorderBrush = "Gray"
$searchBox.BorderThickness = "1"
$searchBox.Padding = "5"

# Add Filter Button
$keywordButton = New-Object System.Windows.Controls.Button
$keywordButton.Content = "Keyword"
$keywordButton.Height = 30
$keywordButton.Margin = $uIMargin
$keywordButton.VerticalAlignment = "Center"
$keywordButton.HorizontalAlignment = "Left"
$keywordButton.FontSize = "14"
$keywordButton.Padding = "10,5,10,5" 
$keywordButton.BorderThickness = "1"
$keywordButton.ToolTip = "Filter log entries based on the search term/keyword."
$keywordButton.Add_Click({
        Update-TreeView  
    })

# Add the search components to the grid
$searchPanel.Children.Add($searchBox) | out-null
[System.Windows.Controls.Grid]::SetColumn($searchBox, 0)

$searchPanel.Children.Add($keywordButton) | out-null
[System.Windows.Controls.Grid]::SetColumn($keywordButton, 1)

# Add the search panel to the grid or parent container
$grid.Children.Add($searchPanel) | out-null
[System.Windows.Controls.Grid]::SetRow($searchPanel, 0)  #
[System.Windows.Controls.Grid]::SetColumn($searchPanel, 0)  

#################### Filter by RecordType, Operation, Date and Time #####################

# Define StackPanels for CheckBoxes globally
$recordTypeCheckBoxPanel = New-Object System.Windows.Controls.StackPanel
$operationsCheckBoxPanel = New-Object System.Windows.Controls.StackPanel

# Define these variables at script scope level (outside any function)
$script:allRecordTypeCheckBox = $null
$script:allOperationsCheckBox = $null
$script:isUpdatingCheckBoxes = $false


$filterPanel = New-Object System.Windows.Controls.WrapPanel  # Use StackPanel for vertical layout
$filterPanel.Orientation = "Horizontal"
$filterPanel.HorizontalAlignment = "Stretch"
$filterPanel.VerticalAlignment = "Center"
$filterPanel.Margin = "10"
$grid.Children.Add($filterPanel) | out-null
[System.Windows.Controls.Grid]::SetRow($filterPanel, 0)
[System.Windows.Controls.Grid]::SetColumn($filterPanel, 1)

# Add Parent Label "Filters"
$filtersLabel = New-Object System.Windows.Controls.Label
$filtersLabel.Content = "Filters:"
$filtersLabel.Margin = $uIMargin
$filtersLabel.VerticalAlignment = "Center"
$filtersLabel.FontSize = "16"
$filtersLabel.FontWeight = "Bold"
$filterPanel.Children.Add($filtersLabel) | out-null

# Group RecordType Label and Dropdown
$recordTypeGroup = New-Object System.Windows.Controls.StackPanel
$recordTypeGroup.Orientation = "Horizontal"
$recordTypeGroup.Margin = $uIMargin
$recordTypeGroup.VerticalAlignment = "Center"

# Add Label for RecordType Filter
$recordTypeLabel = New-Object System.Windows.Controls.Label
$recordTypeLabel.Content = "RecordType:"
$recordTypeLabel.Margin = $uIMargin
$recordTypeLabel.VerticalAlignment = "Center"
$recordTypeLabel.FontSize = "14"

# Add RecordType Filter Dropdown
$recordTypeFilter = New-Object System.Windows.Controls.ComboBox
$recordTypeFilter.Width = 150
$recordTypeFilter.Height = $buttonHeight
$recordTypeFilter.Margin = $uIMargin
$recordTypeFilter.ToolTip = "Filter by RecordType"
$recordTypeFilter.FontSize = "14"
$recordTypeFilter.IsEditable = $true  
$recordTypeFilter.IsReadOnly = $true  
$recordTypeFilter.StaysOpenOnEdit = $true  
$recordTypeFilter.Add_SelectionChanged({
        Update-TreeView
    })

# Add Label and Dropdown to the RecordType Group
$recordTypeGroup.Children.Add($recordTypeLabel) | out-null
$recordTypeGroup.Children.Add($recordTypeFilter) | out-null

# Group Operations Label and Dropdown
$operationsGroup = New-Object System.Windows.Controls.StackPanel
$operationsGroup.Orientation = "Horizontal"
$operationsGroup.Margin = $uIMargin
$operationsGroup.VerticalAlignment = "Center"

# Add Label for Operations Filter
$operationsLabel = New-Object System.Windows.Controls.Label
$operationsLabel.Content = "Operation:"
$operationsLabel.Margin = $uIMargin
$operationsLabel.VerticalAlignment = "Center"
$operationsLabel.FontSize = "14"

# Add Operations Filter Dropdown
$operationsFilter = New-Object System.Windows.Controls.ComboBox
$operationsFilter.Width = 150
$operationsFilter.Height = 25
$operationsFilter.Margin = $uIMargin
$operationsFilter.ToolTip = "Filter by Operations"
$operationsFilter.FontSize = "14"
$operationsFilter.IsEditable = $true  
$operationsFilter.IsReadOnly = $true  
$operationsFilter.StaysOpenOnEdit = $true 
$operationsFilter.Add_SelectionChanged({
        Update-TreeView
    })

# Add Label and Dropdown to the Operations Group
$operationsGroup.Children.Add($operationsLabel) | out-null
$operationsGroup.Children.Add($operationsFilter) | out-null

# Group Date Range Label, Start Date, and End Date
$dateRangeGroup = New-Object System.Windows.Controls.StackPanel
$dateRangeGroup.Orientation = "Horizontal"
$dateRangeGroup.Margin = $uIMargin
$dateRangeGroup.VerticalAlignment = "Center"

# Add Label for Date Range Filter
$dateRangeLabel = New-Object System.Windows.Controls.Label
$dateRangeLabel.Content = "Date Range:"
$dateRangeLabel.Margin = $uIMargin
$dateRangeLabel.VerticalAlignment = "Center"
$dateRangeLabel.FontSize = "14"

# Add Start Date Picker
$startDatePicker = New-Object System.Windows.Controls.DatePicker
$startDatePicker.Width = 110
$startDatePicker.Margin = $uIMargin
$startDatePicker.ToolTip = "Start Date"
$startDatePicker.FontSize = "14"
$startDatePicker.Add_SelectedDateChanged({
        Update-TreeView
    })

# Add End Date Picker
$endDatePicker = New-Object System.Windows.Controls.DatePicker
$endDatePicker.Width = 110
$endDatePicker.Margin = $uIMargin
$endDatePicker.ToolTip = "End Date"
$endDatePicker.FontSize = "14"
$endDatePicker.Add_SelectedDateChanged({
        Update-TreeView
    })

# Add Date Range components to the Date Range Group
$dateRangeGroup.Children.Add($dateRangeLabel) | out-null
$dateRangeGroup.Children.Add($startDatePicker) | out-null
$dateRangeGroup.Children.Add($endDatePicker) | out-null

# Group Time Label, Start Time, and End Time
$timeGroup = New-Object System.Windows.Controls.StackPanel
$timeGroup.Orientation = "Horizontal"
$timeGroup.Margin = $uIMargin
$timeGroup.VerticalAlignment = "Center"

# Add Label for Time Filter
$timeLabel = New-Object System.Windows.Controls.Label
$timeLabel.Content = "Time:"
$timeLabel.Margin = $uIMargin
$timeLabel.VerticalAlignment = "Center"
$timeLabel.FontSize = "14"

# Add Start Time ComboBox
$startTimeComboBox = New-Object System.Windows.Controls.ComboBox
$startTimeComboBox.Width = 80
$startTimeComboBox.Margin = $uIMargin
$startTimeComboBox.ToolTip = "Select start time"
$startTimeComboBox.IsEditable = $true  
$startTimeComboBox.Text = "00:00:00"  
$startTimeComboBox.FontSize = "14"
$startTimeComboBox.AddHandler([System.Windows.Controls.TextBox]::TextChangedEvent, [System.Windows.RoutedEventHandler] {
        Update-TreeView
    })

# Add End Time ComboBox
$endTimeComboBox = New-Object System.Windows.Controls.ComboBox
$endTimeComboBox.Width = 80
$endTimeComboBox.Margin = $uIMargin
$endTimeComboBox.ToolTip = "Select end time"
$endTimeComboBox.IsEditable = $true    
$endTimeComboBox.Text = "23:59:59"     
$endTimeComboBox.FontSize = "14"
$endTimeComboBox.AddHandler([System.Windows.Controls.TextBox]::TextChangedEvent, [System.Windows.RoutedEventHandler] {
        Update-TreeView
    })

# Populate ComboBoxes with time values
$timeValues = @()
for ($hour = 0; $hour -lt 24; $hour++) {
    for ($minute = 0; $minute -lt 60; $minute++) {
        $timeValues += "{0:D2}:{1:D2}:00" -f $hour, $minute
    }
}

# Add time values to ComboBoxes
$timeValues | ForEach-Object {
    $startTimeComboBox.Items.Add($_) | Out-Null
    $endTimeComboBox.Items.Add($_) | Out-Null
}

# Add Time components to the Time Group
$timeGroup.Children.Add($timeLabel) | out-null
$timeGroup.Children.Add($startTimeComboBox) | out-null
$timeGroup.Children.Add($endTimeComboBox) | out-null
$filterPanel.Children.Add($recordTypeGroup) | out-null
$filterPanel.Children.Add($operationsGroup) | out-null
$filterPanel.Children.Add($dateRangeGroup) | out-null
$filterPanel.Children.Add($timeGroup) | out-null

################ Other feature ####################

# Add a StackPanel to Row 1, Column 1 for buttons
$buttonPanel = New-Object System.Windows.Controls.WrapPanel
$buttonPanel.Orientation = "Horizontal"
$buttonPanel.HorizontalAlignment = "Stretch"
$buttonPanel.VerticalAlignment = "Center"
$buttonPanel.Margin = "10"

$grid.Children.Add($buttonPanel) | out-null
[System.Windows.Controls.Grid]::SetRow($buttonPanel, 0)
[System.Windows.Controls.Grid]::SetColumn($buttonPanel, 2)

# Add Export to JSON Button

$exportJsonButton = New-Object System.Windows.Controls.Button
$exportJsonButton.Content = "Export to JSON"
$exportJsonButton.Width = $buttonWidth
$exportJsonButton.Height = $buttonHeight
$exportJsonButton.Margin = $uIMargin
$exportJsonButton.ToolTip = "Export the displayed data to a JSON file."

$exportJsonButton.Add_Click({
        try {
            if (-not $script:exportFilteredLogData) {
                [System.Windows.MessageBox]::Show("No data available for export.", "Error", "OK", "Error") | Out-Null
                return
            }

            $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveFileDialog.Filter = "JSON Files (*.json)|*.json"
            $saveFileDialog.Title = "Save JSON File"
            $saveFileDialog.DefaultExt = "json"

            if ($saveFileDialog.ShowDialog() -eq $true) {
                # Convert each log entry's AuditData into structured form only if needed
                $structuredData = $script:exportFilteredLogData | ForEach-Object {
                    $logEntry = $_ | Select-Object * -ExcludeProperty AuditData
                
                    # Auto-detect if AuditData is a string (JSON) or already an object
                    if ($_.AuditData -is [string]) {
                        try {
                            $parsedAuditData = $_.AuditData | ConvertFrom-Json -ErrorAction Stop
                        }
                        catch {
                            $parsedAuditData = $_.AuditData  # Keep as string if JSON conversion fails
                        }
                    }
                    else {
                        $parsedAuditData = $_.AuditData  # Already an object, no conversion needed
                    }

                    $logEntry | Add-Member -MemberType NoteProperty -Name "AuditData" -Value $parsedAuditData -Force
                    $logEntry
                }

                $structuredData | ConvertTo-Json -Depth 10 | Out-File -FilePath $saveFileDialog.FileName -Encoding utf8

                [System.Windows.MessageBox]::Show("Data exported to JSON successfully!", "Success", "OK", "Information") | Out-Null
                Update-StatusBar -Message "Data exported to JSON successfully!  Dir: ($saveFileDialog.FileName)"
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("An error occurred while exporting: $_", "Export Error", "OK", "Error") | Out-Null
        }
    })


#Add Export to CSV Button
$exportCsvButton = New-Object System.Windows.Controls.Button
$exportCsvButton.Content = "Export to CSV"
$exportCsvButton.Width = $buttonWidth
$exportCsvButton.Height = $buttonHeight
$exportCsvButton.Margin = $uIMargin
$exportCsvButton.ToolTip = "Export the displayed data to a CSV file."
$exportCsvButton.Add_Click({
        $saveFileDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"

        if ($saveFileDialog.ShowDialog() -eq $true) {
            $filteredData = Export-FilteredAuditData -FilteredAuditData $script:exportFilteredLogData
            $filteredData | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            [System.Windows.MessageBox]::Show("Data exported to CSV successfully!") | Out-Null
            Update-StatusBar -Message "Data exported to CSV successfully! Dir: ($saveFileDialog.FileName)"
        }
    })

# Add Refresh Button
$refreshButton = New-Object System.Windows.Controls.Button
$refreshButton.Content = "Refresh"
$refreshButton.Width = $buttonWidth
$refreshButton.Height = $buttonHeight
$refreshButton.Margin = $uIMargin
$refreshButton.ToolTip = "Refresh reloads the data sets"
$refreshButton.Add_Click({
        # Prevent triggering events while updating
        $isUpdatingCheckBoxes = $true

        # Clear all filters
        $searchBox.Text = ""
        $recordTypeFilter.SelectedIndex = -1
        $operationsFilter.SelectedIndex = -1
        $startDatePicker.SelectedDate = $null
        $endDatePicker.SelectedDate = $null
        $startTimeComboBox.Text = "00:00:00"
        $endTimeComboBox.Text = "23:59:59"

        $treeView.Items.Clear()

        # Reload the data if a file was previously loaded via drag-and-drop or double-click
        if ($global:logDataArray) {
            $global:logDataArray = $global:logDataArray  # Reuse existing data
        }
        else {
            # If no data is loaded, prompt the user to select a file
            $openFileDialog = New-Object Microsoft.Win32.OpenFileDialog
            $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|JSON Files (*.json)|*.json"
            $openFileDialog.Title = "Select a CSV or JSON file to load"

            if ($openFileDialog.ShowDialog() -eq $true) {
                $filePath = $openFileDialog.FileName

                if (Test-ValidFile -FilePath $filePath) {
                    $global:logDataArray = Import-DataFromFile -FilePath $filePath
                }
            }
        }

        if ($null -ne $global:logDataArray) {
            Update-Filters  # Ensure filters are populated correctly

            # Ensure "All" is checked for RecordType filter
            foreach ($checkBox in $recordTypeCheckBoxPanel.Children) {
                if ($checkBox.Content -eq "All") {
                    $checkBox.IsChecked = $true
                }
                else {
                    $checkBox.IsChecked = $false
                }
            }
            Update-SelectedRecordTypes  # Ensure UI updates correctly

            # Ensure "All" is checked for Operations filter
            foreach ($checkBox in $operationsCheckBoxPanel.Children) {
                if ($checkBox.Content -eq "All") {
                    $checkBox.IsChecked = $true
                }
                else {
                    $checkBox.IsChecked = $false
                }
            }
            Update-SelectedOperations  # Ensure UI updates correctly

            Update-TreeView
            Update-StatusBar -Message "Data refreshed successfully! 'All' selected for RecordType & Operations."
        }
        else {
            Update-StatusBar -Message "No data loaded."
        }

        # Re-enable updates
        $isUpdatingCheckBoxes = $false
    })


# Add Theme Toggle Button
$themeButton = New-Object System.Windows.Controls.Button
$themeButton.Content = "Toggle Theme"
$themeButton.Width = $buttonWidth
$themeButton.Height = $buttonHeight
$themeButton.ToolTip = "Switch between different UI themes"
$themeButton.Margin = $uIMargin
$themeButton.Add_Click({
        # Cycle through themes
        $script:currentThemeIndex = ($script:currentThemeIndex + 1) % $themes.Count
        Update-UITheme -theme $themes[$script:currentThemeIndex]

        # Update-StatusBar $script:currentThemeIndex

        $OnlineDataCheckBoxPanel.Background = $themes[$script:currentThemeIndex].foreground 
        $OnlineDataCheckBoxPanelLabel.Foreground = $themes[$script:currentThemeIndex].background
    
    })

# Add Custom Color Picker Button
$customThemeButton = New-Object System.Windows.Controls.Button
$customThemeButton.Content = "Custom Theme"
$customThemeButton.Width = $buttonWidth
$customThemeButton.Height = $buttonHeight
$customThemeButton.ToolTip = "Pick custom colors for UI"
$customThemeButton.Margin = $uIMargin
$customThemeButton.Add_Click({
        # Open Color Picker for Background
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $customBackground = ConvertTo-Brush -drawingColor $colorDialog.Color
        }

        # Open Color Picker for Foreground (Text)
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $customForeground = ConvertTo-Brush -drawingColor $colorDialog.Color
        }

        # Apply custom colors
        if ($customBackground -and $customForeground) {
            Update-UITheme -theme @{Background = $customBackground; Foreground = $customForeground }

            $OnlineDataCheckBoxPanel.Background = $customForeground 
            $OnlineDataCheckBoxPanelLabel.Foreground = $customBackground
        }

    })

# Toggle Expand/Collapse Button
$toggleButtonExpandCollapse = New-Object System.Windows.Controls.Button
$toggleButtonExpandCollapse.Content = "Expand All"
$toggleButtonExpandCollapse.Width = $buttonWidth
$toggleButtonExpandCollapse.Height = $buttonHeight
$toggleButtonExpandCollapse.Margin = $uIMargin
$toggleButtonExpandCollapse.ToolTip = "Expand or collapse all logs"
$toggleButtonExpandCollapse.Add_Click({
        # Determine if we need to expand or collapse
        if ($treeView.Items) {
            $expand = -not ($treeView.Items | Where-Object { $_.IsExpanded } | Measure-Object).Count
            Set-ExpansionState $treeView.Items $expand
            $toggleButtonExpandCollapse.Content = if ($expand) { "Collapse All" } else { "Expand All" }
        }
    })




$resetFiltersButton = New-Object System.Windows.Controls.Button
$resetFiltersButton.Content = "Reset Filters"
$resetFiltersButton.ToolTip = "Reset Filters all selected filter and reload the data"
$resetFiltersButton.Width = $buttonWidth
$resetFiltersButton.Height = $buttonHeight
$resetFiltersButton.Margin = $uIMargin
$resetFiltersButton.Add_Click({
        # Prevent triggering events while updating
        $isUpdatingCheckBoxes = $true

        # Clear search box
        $searchBox.Text = ""

        # Reset RecordType (RecipientType) filter
        foreach ($checkBox in $recordTypeCheckBoxPanel.Children) {
            if ($checkBox.Content -eq "All") {
                $checkBox.IsChecked = $true 
            }
            else {
                $checkBox.IsChecked = $false  
            }
        }
        Update-SelectedRecordTypes  

        # Reset Operations filter
        foreach ($checkBox in $operationsCheckBoxPanel.Children) {
            if ($checkBox.Content -eq "All") {
                $checkBox.IsChecked = $true  
            }
            else {
                $checkBox.IsChecked = $false 
            }
        }
        Update-SelectedOperations  
        
        # Reset date and time filters
        $startDatePicker.SelectedDate = $null
        $endDatePicker.SelectedDate = $null
        $startTimeComboBox.Text = "00:00:00"
        $endTimeComboBox.Text = "23:59:59"

        # Update the TreeView
        Update-TreeView

        # Re-enable event handlers
        $isUpdatingCheckBoxes = $false

        # Update the status bar
        Update-StatusBar -Message "Filters cleared. 'All' is selected for both RecordType and Operations."
    })



# Add Clear Audit Data Button (NEW: Added for clearing all data)
$clearLoadedData = New-Object System.Windows.Controls.Button
$clearLoadedData.Content = "Clear Audit Data"
$clearLoadedData.Width = $buttonWidth
$clearLoadedData.Height = $buttonHeight
$clearLoadedData.Margin = $uIMargin
$clearLoadedData.ToolTip = "Clear all loaded data and reset the UI"
$clearLoadedData.Add_Click({
        # Prevent triggering events while updating
        $isUpdatingCheckBoxes = $true

        # Clear search box
        $searchBox.Text = ""

        $recordTypeCheckBoxPanel.Children.Clear()
        Update-SelectedRecordTypes  

        $operationsCheckBoxPanel.Children.Clear()
        Update-SelectedOperations 

        $startDatePicker.SelectedDate = $null
        $endDatePicker.SelectedDate = $null
        $startTimeComboBox.Text = "00:00:00"
        $endTimeComboBox.Text = "23:59:59"

        $treeView.Items.Clear()

        $global:logDataArray = @()  # Empty array instead of $null for stability

        # Update the status bar
        Update-StatusBar -Message "All data cleared. No RecordType or Operations entries remain."

        # Re-enable updates
        $isUpdatingCheckBoxes = $false
    })



   
################### Retrieve content online connection #######################

# Connect/Disconnect EXO Button
$connectEXOButton = New-Object System.Windows.Controls.Button
$connectEXOButton.Content = "Connect EXO"
$connectEXOButton.Width = $buttonWidth
$connectEXOButton.Height = $buttonHeight
$connectEXOButton.Margin = $uIMargin
$connectEXOButton.FontWeight = "Bold"
$connectEXOButton.BorderBrush = "Green"
$connectEXOButton.BorderThickness = 2
$connectEXOButton.ToolTip = "Connect and get logs from online"
$connectEXOButton.Add_Click({
        if ($connectEXOButton.Content -eq "Connect EXO") {
            # Check if Exchange Online is already connected
            if (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue) {
                Update-StatusBar -Message "Exchange Online is already connected, good to go!" -TextColor "DarkGreen"
                $connectEXOButton.Content = "Disconnect EXO"
                return
            }
            else {
                if (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
                    try {
                        
                        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                        $OnlineDataCheckBox.IsEnabled = $true  # Enable the checkbox
                        $OnlineDataCheckBox.IsChecked = $true
                        # Verify connection
                        if (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue) {
                            Update-StatusBar -Message "Successfully connected to Exchange Online!" -TextColor "Green"
                            $connectEXOButton.Content = "Disconnect EXO"
                        }
                        else {
                            Update-StatusBar -Message "Connection established, but some commands are missing. Try reconnecting." -TextColor "Orange"
                        }
                    }
                    catch {
                        Update-StatusBar -Message "Failed to connect: $_" -TextColor "Red"
                    }
                }
                else {
                    Update-StatusBar -Message "Exchange Online module is missing. Install it using, use -Scope CurrentUser if you are not admin
                1️⃣ Open PowerShell as Administrator
                2️⃣ Run: Set-ExecutionPolicy RemoteSigned # -Scope CurrentUser
                3️⃣ Run: Install-Module ExchangeOnlineManagement # -Scope CurrentUser
                4️⃣ Retry connecting!" -TextColor "Red"
                }
            }
        }
        else {
            # Disconnect Exchange Online
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                Start-Sleep -Seconds 2
                Update-StatusBar -Message "Disconnected from Exchange Online!" -TextColor "Red"
                $connectEXOButton.Content = "Connect EXO"
            }
            catch {
                Update-StatusBar -Message "Failed to disconnect: $_" -TextColor "Red"
            }
        }
    })




###################### Add buttons #######################

# Add buttons to the button panel
$buttonPanel.Children.Add($exportJsonButton) | out-null
$buttonPanel.Children.Add($exportCsvButton) | out-null
$buttonPanel.Children.Add($refreshButton) | out-null
$buttonPanel.Children.Add($toggleButtonExpandCollapse) | out-null
$buttonPanel.Children.Add($resetFiltersButton) | out-null
$buttonPanel.Children.Add($clearLoadedData) | out-null
$buttonPanel.Children.Add($connectEXOButton) | out-null
$buttonPanel.Children.Add($themeButton) | out-null
$buttonPanel.Children.Add($customThemeButton) | out-null



##################   Creating Online audit search pannel ###################

# Create a parent container for the search form
$auditSearchContainer = New-Object System.Windows.Controls.Grid
$auditSearchContainer.Margin = "5"
$auditSearchContainer.Visibility = "Collapse"
$grid.Children.Add($auditSearchContainer) | out-null
[System.Windows.Controls.Grid]::SetRow($auditSearchContainer, 2)
[System.Windows.Controls.Grid]::SetColumn($auditSearchContainer, 0)
[System.Windows.Controls.Grid]::SetColumnSpan($auditSearchContainer, 3)

# Define rows for the search form container
$auditSearchContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 0: Form fields
$auditSearchContainer.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = "Auto" }))  # Row 1: Search button

# Create a WrapPanel to organize the form fields
$auditFieldsPanel = New-Object System.Windows.Controls.WrapPanel
$auditFieldsPanel.Orientation = "Horizontal"
$auditFieldsPanel.HorizontalAlignment = "Left"
$auditFieldsPanel.Margin = "0,0,0,10"
$auditSearchContainer.Children.Add($auditFieldsPanel) | out-null
[System.Windows.Controls.Grid]::SetRow($auditFieldsPanel, 1)

# Function to add a label and input control to the WrapPanel
function Add-InputControl {
    param (
        [string]$LabelText,
        [string]$ControlType,
        [string[]]$DropDownItems,
        [string[]]$toolTip = "",
        [int]$Width = 150
    )

    # Create a StackPanel for vertical grouping
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Orientation = "Vertical"
    $stackPanel.Margin = "10,5,10,5"
    $stackPanel.Width = $Width + 20  # Add extra space for margins

    # Add Label
    $label = New-Object System.Windows.Controls.Label
    $label.Content = $LabelText
    $label.Margin = "0,0,0,2"  # Small margin below label
    $stackPanel.Children.Add($label) | out-null

    # Add Input Control
    switch ($ControlType) {
        "DatePicker" {
            $control = New-Object System.Windows.Controls.DatePicker
            $control.ToolTip = [string]$toolTip
            $control.Width = $Width
        }
        "TextBox" {
            $control = New-Object System.Windows.Controls.TextBox
            $control.Width = $Width
            $control.ToolTip = [string]$toolTip
            $control.Height = 24
        }
        "ComboBox" {
            $control = New-Object System.Windows.Controls.ComboBox
            $control.Width = $Width
            $control.ToolTip = $toolTip
            $control.Height = 24

            # Add items to ComboBox
            foreach ($item in $DropDownItems) {
                $comboBoxItem = New-Object System.Windows.Controls.ComboBoxItem
                $comboBoxItem.Content = $item
                if ($toolTip) {
                    $comboBoxItem.ToolTip = $toolTip[$DropDownItems.IndexOf($item)]
                }
                $control.Items.Add($comboBoxItem)
            }
        }
        "CheckBox" {
            $control = New-Object System.Windows.Controls.CheckBox
            $control.ToolTip = [string]$toolTip
            $stackPanel.Orientation = "Horizontal" 
            $stackPanel.Margin = "10,30,10,5"
            $label.VerticalAlignment = "Center" 
            $control.VerticalAlignment = "Center" 
            # No width needed for checkbox
        }
    }

    $stackPanel.Children.Add($control) | out-null
    $auditFieldsPanel.Children.Add($stackPanel) | out-null

    return $control
}


# Add input controls for each parameter
$recordTypeTextBox = Add-InputControl -LabelText "Record Type:" -ControlType "TextBox"
$operationsTextBox = Add-InputControl -LabelText "Operations - ('a','b','c'):" -ControlType "TextBox" -toolTip "The Operations parameter filters the log entries by operation. `nThe available values for this parameter depend on the RecordType value"
$freeTextTextBox = Add-InputControl -LabelText "Free Text:" -ControlType "TextBox" -toolTip "The FreeText parameter filters the log entries by the specified text string. `nIf the value contains spaces, enclose the value in quotation marks (`")"
$ipAddressesTextBox = Add-InputControl -LabelText "IP Addresses - ('a','b','c'):" -ControlType "TextBox" 
$objectIdsTextBox = Add-InputControl -LabelText "Object IDs - ('a','b','c'):" -ControlType "TextBox" 
$resultSizeTextBox = Add-InputControl -LabelText "Result Size:" -ControlType "TextBox" -toolTip "The ResultSize parameter specifies the maximum number of results to return.`nThe default value is 100, maximum is 5,000."
$longerRetentionEnabledTextBox = Add-InputControl -LabelText "Longer Retention Enabled:" -ControlType "TextBox"
$sessionCommandComboBox = Add-InputControl -LabelText "Session Command:" -ControlType "ComboBox" -DropDownItems "ReturnNextPreviewPage", "ReturnLargeSet" `
    -toolTip "Returns sorted data by date with a maximum of 5,000 results.", "Returns unsorted data with a maximum of 50,000 results, optimized for faster search."
$sessionIdTextBox = Add-InputControl -LabelText "Session ID:" -ControlType "TextBox"
$siteIdsTextBox = Add-InputControl -LabelText "Site IDs - ('a','b','c'):" -ControlType "TextBox" -toolTip "The SiteIds parameter filters the log entries by the SharePoint SiteId (GUID). `nYou can enter multiple values separated by commas: Value1, Value2,...ValueN."
$userIdsTextBox = Add-InputControl -LabelText "User IDs - ('a','b','c'):" -ControlType "TextBox" -toolTip "The UserIds parameter filters the log entries by the account (UserPrincipalName)"
$highCompletenessCheckBox = Add-InputControl -LabelText "High Completeness:" -ControlType "CheckBox" -toolTip "The HighCompleteness switch specifies completeness instead performance in the results, `nreturns more complete search results but might take significantly longer to run"
$formattedCheckBox = Add-InputControl -LabelText "Formatted:" -ControlType "CheckBox" -toolTip "The Formatted switch causes attributes that are normally returned as integers `n(for example, RecordType and Operation) to be formatted as descriptive strings."

# Search Button - centered below the form fields
$searchButtonContainer = New-Object System.Windows.Controls.StackPanel
$searchButtonContainer.Orientation = "Horizontal"
$searchButtonContainer.HorizontalAlignment = "left"
$searchButtonContainer.Margin = "0,10,0,10"
$auditSearchContainer.Children.Add($searchButtonContainer) | out-null
[System.Windows.Controls.Grid]::SetRow($searchButtonContainer, 0)


$auditSearchButton = New-Object System.Windows.Controls.Button
$auditSearchButton.Content = "M365 Audit Query - Online"
$auditSearchButton.Width = 300
$auditSearchButton.Padding = "10,5,10,5"
$auditSearchButton.HorizontalAlignment = "Left";
$auditSearchButton.Height = 30
$auditSearchButton.FontWeight = "Bold"
$auditSearchButton.BorderBrush = "Green"
$auditSearchButton.BorderThickness = 2
$auditSearchButton.ToolTip = "Use the Search-UnifiedAuditLog cmdlet to search the unified audit log. This log contains events from Exchange Online, SharePoint Online, OneDrive for Business, Microsoft Entra ID, Microsoft Teams, Power BI, and other Microsoft 365 services. "
$auditSearchButton.Add_Click({

        $isUpdatingCheckBoxes = $true


        if (-not( Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
            Update-StatusBar -Message "The term 'Search-UnifiedAuditLog' is not recognized. Use the 'Connect EXO' button to connect. If it still fails, then it means you do not have permission to command" -TextColor "Red"
            $connectEXOButton.Content = "Connect EXO"
        }
        else {
            try {
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
                if ($sessionCommandComboBox.Text) {
                    $params['SessionCommand'] = $sessionCommandComboBox.Text
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
    
                # Call the Search-UnifiedAuditLog function
                $global:logDataArray = @()
                
                $global:logDataArray = Search-UnifiedAuditLog @params
    
                # Display the results
                if ($global:logDataArray ) {
                    Update-Filters
                    Update-TreeView
                }
                else {
                    Update-StatusBar -Message "No results found."
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                $statusBarText.Text = $_
                Update-StatusBar -Message "Error: $errorMessage"
                # Write-Host "Error: $_"
            }
        }

        $isUpdatingCheckBoxes = $false
    })

$searchButtonContainer.Children.Add($auditSearchButton) | out-null | Out-Null




# Load initial data if provided via -InputData
if ($PSBoundParameters.ContainsKey('InputData')) {
    if ($InputData -is [string] -and (Test-Path $InputData)) {
        # Load from CSV file
        $global:logDataArray = Import-DataFromFile -FilePath $InputData
    }
    elseif ($InputData -is [System.Collections.IEnumerable]) {
        # Use in-memory data
        $global:logDataArray = $InputData
    }
    else {
        [System.Windows.MessageBox]::Show("Error: Invalid input. Provide a valid CSV file path or in-memory data.") | Out-Null
        exit
    }

    if ($null -ne $global:logDataArray) {
        Update-Filters 
        Update-TreeView
        Update-StatusBar -Message "Data loaded successfully from input parameter."
    }
}

# Show Window
$window.child |Out-Null
$window.ShowDialog() | Out-Null