<#
.SYNOPSIS
    Search-UnifiedAuditLogCustom is an enhanced wrapper around the native Search-UnifiedAuditLog cmdlet, providing additional features such as a user-friendly GUI for constructing search queries, simplified output formatting, and integration with the Microsoft 365 audit operations catalog.

    .DESCRIPTION
    This function allows administrators and security professionals to perform more efficient and targeted searches of the Microsoft 365 Unified Audit Log. It includes a helper GUI that enables users to easily select operations from the official Microsoft Learn catalog, specify date ranges, user filters, and other parameters without needing to remember complex cmdlet syntax.
    The output can be returned in a simplified format that flattens nested JSON structures for easier analysis and export. This is particularly useful for security investigations, compliance audits, and general monitoring of activities across Microsoft 365 services.

    .PARAMETER StartDate
    The start date and time for the audit log search. If not specified, defaults to 24 hours ago.

    .PARAMETER EndDate
    The end date and time for the audit log search. If not specified, defaults to the current date and time.

    .PARAMETER Operations
    An array of operation names to filter the search. These can be selected from the helper GUI, which loads the catalog of operations from Microsoft Learn. Users can also enter raw cmdlet names (e.g., New-TransportRule) to filter by specific operations.

    .PARAMETER UserIds
    An array of user identifiers (e.g., email addresses) to filter the search results by specific users.

    .PARAMETER FreeText
    A free text string to search for within the audit log records.

    .PARAMETER ResultSize
    The maximum number of results to return from the search. Defaults to 5000.

    .PARAMETER SimpleView
    When specified, the output will be processed to flatten nested JSON structures into a simpler format. This is ideal for exporting to CSV or performing quick analysis without dealing with complex nested properties.

    .PARAMETER HelperGUI
    When specified, opens a graphical user interface to assist in constructing the search query with user-friendly controls and operation selection.
    The operations list is populated from the Microsoft Learn catalog of audit log activities, allowing users to easily find and select relevant operations for their search.
    Make sure to have access to the Microsoft Learn page for audit log activities to load the operations catalog successfully (https://learn.microsoft.com/en-us/purview/audit-log-activities).

    .EXAMPLE
    Search-UnifiedAuditLogCustom -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) -Operations "UserLoggedIn", "FileAccessed" -SimpleView

    This example searches the Unified Audit Log for "UserLoggedIn" and "FileAccessed" operations that occurred in the last 7 days, and returns the results in a simplified format.

    .EXAMPLE

#>

function Search-UnifiedAuditLogCustom {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [datetime]$StartDate,

        [Parameter(Mandatory = $false)]
        [datetime]$EndDate,

        [Parameter(Mandatory = $false)]
        [string[]]$Operations,

        [Parameter(Mandatory = $false)]
        [string[]]$UserIds,

        [Parameter(Mandatory = $false)]
        [string]$FreeText,

        [Parameter(Mandatory = $false)]
        [int]$ResultSize = 5000,

        [Parameter(Mandatory = $false)]
        [Alias('Simple')]
        [switch]$SimpleView,

        [Parameter(Mandatory = $false)]
        [switch]$HelperGUI
    )

    # function from HAWK module
    function Get-SimpleUnifiedAuditLog {
        <#
    .SYNOPSIS
        Flattens nested Microsoft 365 Unified Audit Log records into a simplified format.

    .DESCRIPTION
        This function processes Microsoft 365 Unified Audit Log records by converting nested JSON data
        (stored in the AuditData property) into a flat structure suitable for analysis and export.
        It handles complex nested objects, arrays, and special cases like parameter collections.

        The function:
        - Preserves base record properties
        - Flattens nested JSON structures
        - Provides special handling for Parameters collections
        - Creates human-readable command reconstructions
        - Supports type preservation for data analysis

    .PARAMETER Record
        A PowerShell object representing a unified audit log record. Typically, this is the output
        from Search-UnifiedAuditLog and should contain both base properties and an AuditData
        property containing a JSON string of additional audit information.

    .PARAMETER PreserveTypes
        When specified, maintains the original data types of values instead of converting them
        to strings. This is useful when the output will be used for further PowerShell processing
        rather than export to CSV/JSON.

    .EXAMPLE
        $auditLogs = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -RecordType ExchangeAdmin
        $auditLogs | Get-SimpleUnifiedAuditLog | Export-Csv -Path "AuditLogs.csv" -NoTypeInformation

        Processes Exchange admin audit logs and exports them to CSV with all nested properties flattened.

    .EXAMPLE
        $userChanges = Search-UnifiedAuditLog -UserIds user@domain.com -Operations "Add-*"
        $userChanges | Get-SimpleUnifiedAuditLog -PreserveTypes |
            Where-Object { $_.ResultStatus -eq $true } |
            Select-Object CreationTime, Operation, FullCommand

        Gets all "Add" operations for a specific user, preserves data types, filters for successful operations,
        and selects specific columns.

    .OUTPUTS
        Collection of PSCustomObjects with flattened properties from both the base record and AuditData.
        Properties include:
        - All base record properties (RecordType, CreationDate, etc.)
        - Flattened nested objects with property names using dot notation
        - Individual parameters as Param_* properties
        - ParameterString containing all parameters in a readable format
        - FullCommand showing reconstructed PowerShell command (when applicable)

    .NOTES
        Author: Jonathan Butler
        Version: 2.0
        Development Date: December 2024

        The function is designed to handle any RecordType from the Unified Audit Log and will
        automatically adapt to changes in the audit log schema. Special handling is implemented
        for common patterns like Parameters collections while maintaining flexibility for
        other nested structures.
    #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
            [PSObject]$Record,

            [Parameter(Mandatory = $false)]
            [switch]$PreserveTypes
        )

        begin {
            [System.Collections.Generic.List[PSCustomObject]]$resultsArray = @()

            function ConvertTo-FlatObject {
                param (
                    [Parameter(Mandatory = $true)]
                    [PSObject]$InputObject,

                    [Parameter(Mandatory = $false)]
                    [string]$Prefix = '',

                    [Parameter(Mandatory = $false)]
                    [switch]$PreserveTypes
                )

                $flatProperties = @{}

                foreach ($property in $InputObject.PSObject.Properties) {
                    $key = if ($Prefix) { "${Prefix}_$($property.Name)" } else { $property.Name }

                    if ($property.Name -eq 'Parameters' -and $property.Value -is [Array]) {
                        $parameterStrings = foreach ($parameter in $property.Value) {
                            "$($parameter.Name)=$($parameter.Value)"
                        }
                        $flatProperties['ParameterString'] = $parameterStrings -join ' | '

                        foreach ($parameter in $property.Value) {
                            $parameterKey = "Param_$($parameter.Name)"
                            $flatProperties[$parameterKey] = $parameter.Value
                        }

                        if ($InputObject.Operation) {
                            $parameterStrings = foreach ($parameter in $property.Value) {
                                $parameterValue = switch -Regex ($parameter.Value) {
                                    '\s' { "'$($parameter.Value)'" }
                                    '^True$|^False$' { "`$$($parameter.Value.ToLower())" }
                                    ';' { "'$($parameter.Value)'" }
                                    default { $parameter.Value }
                                }
                                "-$($parameter.Name) $parameterValue"
                            }
                            $flatProperties['FullCommand'] = "$($InputObject.Operation) $($parameterStrings -join ' ')"
                        }

                        continue
                    }

                    switch ($property.Value) {
                        { $_ -is [System.Collections.IDictionary] } {
                            $nestedProperties = ConvertTo-FlatObject -InputObject $_ -Prefix $key -PreserveTypes:$PreserveTypes
                            foreach ($nestedKey in $nestedProperties.Keys) {
                                $uniqueKey = if ($flatProperties.ContainsKey($nestedKey)) {
                                    $counter = 1
                                    while ($flatProperties.ContainsKey("${nestedKey}_$counter")) {
                                        $counter++
                                    }
                                    "${nestedKey}_$counter"
                                }
                                else {
                                    $nestedKey
                                }
                                $flatProperties[$uniqueKey] = $nestedProperties[$nestedKey]
                            }
                        }
                        { $_ -is [System.Collections.IList] -and $property.Name -ne 'Parameters' } {
                            if ($_.Count -gt 0) {
                                if ($_[0] -is [PSObject]) {
                                    for ($i = 0; $i -lt $_.Count; $i++) {
                                        $nestedProperties = ConvertTo-FlatObject -InputObject $_[$i] -Prefix "${key}_${i}" -PreserveTypes:$PreserveTypes
                                        foreach ($nestedKey in $nestedProperties.Keys) {
                                            $uniqueKey = if ($flatProperties.ContainsKey($nestedKey)) {
                                                $counter = 1
                                                while ($flatProperties.ContainsKey("${nestedKey}_$counter")) {
                                                    $counter++
                                                }
                                                "${nestedKey}_$counter"
                                            }
                                            else {
                                                $nestedKey
                                            }
                                            $flatProperties[$uniqueKey] = $nestedProperties[$nestedKey]
                                        }
                                    }
                                }
                                else {
                                    $flatProperties[$key] = $_ -join '|'
                                }
                            }
                            else {
                                $flatProperties[$key] = [string]::Empty
                            }
                        }
                        { $_ -is [PSObject] } {
                            $nestedProperties = ConvertTo-FlatObject -InputObject $_ -Prefix $key -PreserveTypes:$PreserveTypes
                            foreach ($nestedKey in $nestedProperties.Keys) {
                                $uniqueKey = if ($flatProperties.ContainsKey($nestedKey)) {
                                    $counter = 1
                                    while ($flatProperties.ContainsKey("${nestedKey}_$counter")) {
                                        $counter++
                                    }
                                    "${nestedKey}_$counter"
                                }
                                else {
                                    $nestedKey
                                }
                                $flatProperties[$uniqueKey] = $nestedProperties[$nestedKey]
                            }
                        }
                        default {
                            if ($PreserveTypes) {
                                $flatProperties[$key] = $_
                            }
                            else {
                                $flatProperties[$key] = switch ($_) {
                                    { $_ -is [datetime] } { $_ }
                                    { $_ -is [bool] } { $_ }
                                    { $_ -is [int] } { $_ }
                                    { $_ -is [long] } { $_ }
                                    { $_ -is [decimal] } { $_ }
                                    { $_ -is [double] } { $_ }
                                    default { [string]$_ }
                                }
                            }
                        }
                    }
                }

                return $flatProperties
            }
        }

        process {
            try {
                $baseProperties = $Record | Select-Object -Property * -ExcludeProperty AuditData

                $auditData = $Record.AuditData | ConvertFrom-Json
                if ($auditData) {
                    $flatAuditData = ConvertTo-FlatObject -InputObject $auditData -PreserveTypes:$PreserveTypes

                    $combinedProperties = @{}
                    $baseProperties.PSObject.Properties | ForEach-Object { $combinedProperties[$_.Name] = $_.Value }
                    $flatAuditData.GetEnumerator() | ForEach-Object { $combinedProperties[$_.Key] = $_.Value }

                    $null = $resultsArray.Add([PSCustomObject]$combinedProperties)
                }
            }
            catch {
                Write-Warning "Error processing record: $_"
                $errorProperties = @{
                    RecordType   = $Record.RecordType
                    CreationDate = Get-Date
                    Error        = $_.Exception.Message
                    Record       = $Record
                }
                $null = $resultsArray.Add([PSCustomObject]$errorProperties)
            }
        }

        end {
            $orderedProperties = @(
                'CreationTime',
                'Workload',
                'RecordType',
                'Operation',
                'ResultStatus',
                'ClientIP',
                'UserId',
                'Id',
                'OrganizationId',
                'UserType',
                'UserKey',
                'ObjectId',
                'Scope',
                'AppAccessContext'
            )

            $orderedResults = $resultsArray | ForEach-Object {
                $orderedObject = [ordered]@{}

                foreach ($prop in $orderedProperties) {
                    if ($_.PSObject.Properties.Name -contains $prop) {
                        $orderedObject[$prop] = $_.$prop
                    }
                }

                if ($_.PSObject.Properties.Name -contains 'ParameterString') {
                    $orderedObject['ParameterString'] = $_.ParameterString

                    $_.PSObject.Properties |
                    Where-Object { $_.Name -like 'Param_*' } |
                    Sort-Object -Property Name |
                    ForEach-Object {
                        $orderedObject[$_.Name] = $_.Value
                    }
                }

                $_.PSObject.Properties |
                Where-Object {
                    $_.Name -notin $orderedProperties -and
                    $_.Name -ne 'ParameterString' -and
                    $_.Name -notlike 'Param_*'
                } |
                ForEach-Object {
                    $orderedObject[$_.Name] = $_.Value
                }

                [PSCustomObject]$orderedObject
            }

            return $orderedResults
        }
    }

    if ($HelperGUI) {
        Invoke-SearchUnifiedAuditLogCustomHelperGUI
        return
    }

    if (-not $StartDate) {
        $StartDate = (Get-Date).AddDays(-1)
    }

    if (-not $EndDate) {
        $EndDate = Get-Date
    }

    $searchParams = @{
        StartDate  = $StartDate
        EndDate    = $EndDate
        ResultSize = $ResultSize
    }

    if ($Operations) {
        $searchParams['Operations'] = $Operations
    }

    if ($UserIds) {
        $searchParams['UserIds'] = $UserIds
    }

    if ($FreeText) {
        $searchParams['FreeText'] = $FreeText
    }

    [array]$auditLogs = Search-UnifiedAuditLog @searchParams

    if ($SimpleView) {
        return $auditLogs | Get-SimpleUnifiedAuditLog
    }

    return $auditLogs
}

function Invoke-SearchUnifiedAuditLogCustomHelperGUI {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    $moduleVersion = $null
    $loadedModule = Get-Module -Name 'PS365' -ErrorAction SilentlyContinue
    if ($loadedModule -and $loadedModule.Version) {
        $moduleVersion = "v$($loadedModule.Version)"
    }
    else {
        try {
            $manifestPath = Join-Path $PSScriptRoot '..\..\PS365.psd1'
            $manifest = Import-PowerShellDataFile -Path $manifestPath -ErrorAction Stop
            if ($manifest.ModuleVersion) { $moduleVersion = "v$($manifest.ModuleVersion)" }
        }
        catch {
            Write-Verbose "Could not read module version: $($_.Exception.Message)"
        }
    }

    $splashLogoPath = Join-Path $PSScriptRoot '..\..\Private\Assets\Search-UnifiedAuditLogCustom.png'
    $splash = Show-Splash `
        -Title 'Search-UnifiedAuditLogCustom' `
        -Subtitle 'Audit log search helper' `
        -InitialMessage 'Initializing...' `
        -Version $moduleVersion `
        -LogoPath $splashLogoPath

    [System.Collections.Generic.List[PSCustomObject]]$operationChoices = @()
    $operationLookupByDisplay = @{}

    # Operations catalog is loaded from Microsoft Learn with a local cache and an offline fallback
    # (fresh cache > live refresh > stale cache > bundled seed). See Get-UnifiedAuditLogOperationCatalog.
    if ($splash) { $splash.Update('Loading audit operations catalog...') }
    $catalog = Get-UnifiedAuditLogOperationCatalog
    if ($catalog -and $catalog.Operations) {
        foreach ($entry in $catalog.Operations) {
            $operationChoices.Add($entry)
        }
    }

    if ($splash) {
        $catalogSource = if ($catalog) { $catalog.Source } else { 'None' }
        switch ($catalogSource) {
            'Live' { $splash.Update("Loaded $($operationChoices.Count) operations from Microsoft Learn") }
            'Cache' { $splash.Update("Microsoft Learn unavailable or cached - loaded $($operationChoices.Count) operations from local cache") }
            'Seed' { $splash.Update("Microsoft Learn unavailable - loaded $($operationChoices.Count) operations from bundled list") }
            default { $splash.Update('Could not load operations catalog (you can still type raw cmdlets)') }
        }
    }

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Search-UnifiedAuditLogCustom Helper"
    Width="1200" Height="940"
    MinWidth="1200" MinHeight="940"
        WindowStartupLocation="CenterScreen"
        Background="#F3F5F8"
        FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="TextBox">
            <Setter Property="Margin" Value="0,4,0,0"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="0,6,6,0"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="MinHeight" Value="30"/>
            <Setter Property="MinWidth" Value="76"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="Foreground" Value="#0F172A"/>
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Normal"/>
        </Style>
        <Style x:Key="GhostButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Foreground" Value="#0F172A"/>
            <Setter Property="Background" Value="#F8FAFC"/>
        </Style>
        <Style x:Key="SuccessButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="Foreground" Value="#0F172A"/>
        </Style>
        <Style x:Key="NeutralButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="Foreground" Value="#0F172A"/>
        </Style>
        <Style x:Key="PrimaryButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="Foreground" Value="#0F172A"/>
            <Setter Property="FontWeight" Value="Normal"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="White" CornerRadius="10" Padding="16" BorderBrush="#E2E8F0" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="180"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="0,0,12,0">
                    <TextBlock Text="StartDate" FontWeight="SemiBold"/>
                    <DatePicker x:Name="StartDatePicker" Margin="0,4,0,0"/>
                    <TextBox x:Name="StartDateBox" ToolTip="Format: yyyy-MM-dd HH:mm"/>
                    <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                        <Button x:Name="StartYesterdayButton" Content="Yesterday 00:00" Style="{StaticResource GhostButtonStyle}"/>
                        <Button x:Name="StartNowButton" Content="Now" Style="{StaticResource GhostButtonStyle}"/>
                    </StackPanel>
                </StackPanel>

                <StackPanel Grid.Column="1" Margin="0,0,12,0">
                    <TextBlock Text="EndDate" FontWeight="SemiBold"/>
                    <DatePicker x:Name="EndDatePicker" Margin="0,4,0,0"/>
                    <TextBox x:Name="EndDateBox" ToolTip="Format: yyyy-MM-dd HH:mm"/>
                    <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                        <Button x:Name="EndTodayButton" Content="Today 23:59" Style="{StaticResource GhostButtonStyle}"/>
                        <Button x:Name="EndNowButton" Content="Now" Style="{StaticResource GhostButtonStyle}"/>
                    </StackPanel>
                </StackPanel>

                <StackPanel Grid.Column="2">
                    <TextBlock Text="ResultSize" FontWeight="SemiBold"/>
                    <TextBox x:Name="ResultSizeBox" Text="5000"/>
                    <CheckBox x:Name="SimpleViewCheckBox" Content="Simple view" Margin="0,10,0,0" VerticalAlignment="Center"/>
                    <Button x:Name="RetentionInfoButton" Content="How retention works?" Margin="0,10,0,0" HorizontalAlignment="Left" Style="{StaticResource GhostButtonStyle}"/>
                </StackPanel>
            </Grid>
        </Border>

        <Border Grid.Row="1" Margin="0,10,0,0" Background="White" CornerRadius="10" Padding="12" BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Presets" FontWeight="SemiBold" Margin="0,0,0,6"/>
                <StackPanel Orientation="Horizontal">
                    <Button x:Name="LoadSharingEventsButton" Content="Sharing activity (SPO/OneDrive)" Style="{StaticResource PrimaryButtonStyle}"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <Border Grid.Row="2" Margin="0,14,0,0" Background="White" CornerRadius="10" Padding="16" BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Operations (search friendly name OR type raw cmdlets like New-TransportRule, comma/semicolon-separated)" FontWeight="SemiBold" TextWrapping="Wrap"/>
                <Grid Margin="0,6,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="OperationsSearchBox" Grid.Column="0" ToolTip="Type to filter the list below, OR type a raw cmdlet (e.g. New-TransportRule) and press Enter / click 'Add as raw'"/>
                    <Button x:Name="AddCustomOperationButton" Grid.Column="1" Content="Add as raw" Margin="6,4,0,0" Style="{StaticResource PrimaryButtonStyle}"/>
                </Grid>

                <TextBlock Text="User IDs (comma or semicolon-separated, e.g. alice@contoso.com;bob@contoso.com)" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,12,0,0"/>
                <TextBox x:Name="UserIdsBox" Margin="0,6,0,0" ToolTip="Optional. Filter the search by one or more user principal names. Leave empty to search all users."/>
            </StackPanel>
        </Border>

        <Border Grid.Row="3" Margin="0,10,0,0" Background="White" CornerRadius="10" Padding="12" BorderBrush="#E2E8F0" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="1.3*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="1.3*"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0">
                    <TextBlock Text="Available operations" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ListBox x:Name="AvailableOperationsListBox" Height="190" SelectionMode="Extended" BorderThickness="1" BorderBrush="#CBD5E1"
                             ScrollViewer.HorizontalScrollBarVisibility="Auto"
                             ScrollViewer.CanContentScroll="True"
                             VirtualizingStackPanel.IsVirtualizing="True"
                             VirtualizingStackPanel.VirtualizationMode="Recycling"/>
                </StackPanel>

                <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="12,0">
                    <Button x:Name="AddOperationButton" Content="+ Add" Width="72" Style="{StaticResource PrimaryButtonStyle}"/>
                    <Button x:Name="RemoveOperationButton" Content="- Remove" Width="72" Style="{StaticResource NeutralButtonStyle}"/>
                    <Button x:Name="ClearOperationsButton" Content="Clear" Width="72" Style="{StaticResource GhostButtonStyle}"/>
                </StackPanel>

                <StackPanel Grid.Column="2">
                    <TextBlock Text="Selected operations" FontWeight="SemiBold" Margin="0,0,0,6"/>
                    <ListBox x:Name="SelectedOperationsListBox" Height="190" SelectionMode="Extended" BorderThickness="1" BorderBrush="#CBD5E1"
                             ScrollViewer.HorizontalScrollBarVisibility="Auto"
                             ScrollViewer.CanContentScroll="True"
                             VirtualizingStackPanel.IsVirtualizing="True"
                             VirtualizingStackPanel.VirtualizationMode="Recycling"/>
                </StackPanel>
            </Grid>
        </Border>

        <Border Grid.Row="4" Margin="0,10,0,0" Background="White" CornerRadius="10" Padding="16" BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Generated command" FontWeight="SemiBold"/>
                <TextBox x:Name="CommandBox" Margin="0,6,0,0" Height="120" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True" AcceptsReturn="True"/>
            </StackPanel>
        </Border>

        <Grid Grid.Row="5" Margin="0,14,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left">
                <Button x:Name="CopyButton" Content="Copy" Style="{StaticResource SuccessButtonStyle}"
                        ToolTip="Copies the generated command to the clipboard. Rarely needed here - use 'Run now' to execute the query and export the results directly."/>
                <Button x:Name="RunButton" Content="Run now" Style="{StaticResource PrimaryButtonStyle}"
                        ToolTip="Runs the query against the Unified Audit Log. A file dialog will open first to choose where to save the results - in this mode, results are always exported to an Excel file."/>
                <Button x:Name="CloseButton" Content="Close" Style="{StaticResource NeutralButtonStyle}"/>
            </StackPanel>
            <TextBlock x:Name="FooterText" Grid.Column="1" Foreground="#64748B" FontSize="11"
                       VerticalAlignment="Center" HorizontalAlignment="Right"
                       Text="by Clidsys - Bastien Perez"/>
        </Grid>
    </Grid>
</Window>
'@

    if ($splash) { $splash.Update('Building interface...') }
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $iconPath = Join-Path $PSScriptRoot '..\..\Private\Assets\Search-UnifiedAuditLogCustom.png'
    if (Test-Path -LiteralPath $iconPath) {
        try {
            $iconBitmap = New-Object System.Windows.Media.Imaging.BitmapImage
            $iconBitmap.BeginInit()
            $iconBitmap.UriSource = [Uri]$iconPath
            $iconBitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $iconBitmap.EndInit()
            $iconBitmap.Freeze()
            $window.Icon = $iconBitmap
        }
        catch {
            Write-Verbose "Could not load window icon: $($_.Exception.Message)"
        }
    }

    $footerText = $window.FindName('FooterText')
    if ($footerText) {
        $footerText.Text = if ($moduleVersion) {
            "by Clidsys - Bastien Perez  $moduleVersion"
        }
        else {
            'by Clidsys - Bastien Perez'
        }
    }

    $startDateBox = $window.FindName('StartDateBox')
    $endDateBox = $window.FindName('EndDateBox')
    $startDatePicker = $window.FindName('StartDatePicker')
    $endDatePicker = $window.FindName('EndDatePicker')
    $startYesterdayButton = $window.FindName('StartYesterdayButton')
    $startNowButton = $window.FindName('StartNowButton')
    $endTodayButton = $window.FindName('EndTodayButton')
    $endNowButton = $window.FindName('EndNowButton')
    $resultSizeBox = $window.FindName('ResultSizeBox')
    $simpleViewCheckBox = $window.FindName('SimpleViewCheckBox')
    $loadSharingEventsButton = $window.FindName('LoadSharingEventsButton')
    $operationsSearchBox = $window.FindName('OperationsSearchBox')
    $userIdsBox = $window.FindName('UserIdsBox')
    $addCustomOperationButton = $window.FindName('AddCustomOperationButton')
    $availableOperationsListBox = $window.FindName('AvailableOperationsListBox')
    $selectedOperationsListBox = $window.FindName('SelectedOperationsListBox')
    $addOperationButton = $window.FindName('AddOperationButton')
    $removeOperationButton = $window.FindName('RemoveOperationButton')
    $clearOperationsButton = $window.FindName('ClearOperationsButton')
    $commandBox = $window.FindName('CommandBox')
    $copyButton = $window.FindName('CopyButton')
    $runButton = $window.FindName('RunButton')
    $closeButton = $window.FindName('CloseButton')
    $retentionInfoButton = $window.FindName('RetentionInfoButton')

    $startDateBox.Text = (Get-Date).AddDays(-1).Date.ToString('yyyy-MM-dd 00:00')
    $endDateBox.Text = (Get-Date).Date.ToString('yyyy-MM-dd 23:59')
    $startDatePicker.SelectedDate = (Get-Date).AddDays(-1).Date
    $endDatePicker.SelectedDate = (Get-Date).Date

    foreach ($entry in $operationChoices) {
        if ([string]::IsNullOrWhiteSpace($entry.Operation)) {
            continue
        }

        $display = "$($entry.FriendlyName) [$($entry.Operation)]"
        $operationLookupByDisplay[$display] = $entry.Operation
    }

    $refreshOperationsList = {
        $availableOperationsListBox.Items.Clear()

        $searchValue = $operationsSearchBox.Text
        $filteredOperations = $operationChoices
        if (-not [string]::IsNullOrWhiteSpace($searchValue)) {
            $filteredOperations = $operationChoices | Where-Object {
                $_.FriendlyName -like "*$searchValue*" -or $_.Operation -like "*$searchValue*"
            }
        }

        foreach ($entry in $filteredOperations) {
            if ([string]::IsNullOrWhiteSpace($entry.Operation)) {
                continue
            }

            $display = "$($entry.FriendlyName) [$($entry.Operation)]"
            $null = $availableOperationsListBox.Items.Add($display)
        }
    }

    $parseDateInput = {
        param(
            [string]$InputText,
            [datetime]$Fallback
        )

        try {
            return [datetime]::ParseExact($InputText, 'yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        }
        catch {
            return $Fallback
        }
    }

    $buildCommand = {
        $startParsed = & $parseDateInput -InputText $startDateBox.Text -Fallback ((Get-Date).AddDays(-1).Date)
        $endParsed = & $parseDateInput -InputText $endDateBox.Text -Fallback ((Get-Date).Date.AddHours(23).AddMinutes(59))

        $startValue = $startParsed.ToString('yyyy-MM-dd HH:mm')
        $endValue = $endParsed.ToString('yyyy-MM-dd HH:mm')

        $sizeValue = 5000
        [void][int]::TryParse($resultSizeBox.Text, [ref]$sizeValue)
        if ($sizeValue -lt 1) { $sizeValue = 1 }

        $selectedOperations = @()
        foreach ($selectedDisplay in $selectedOperationsListBox.Items) {
            if ($operationLookupByDisplay.ContainsKey([string]$selectedDisplay)) {
                $selectedOperations += $operationLookupByDisplay[[string]$selectedDisplay]
            }
        }

        $command = "Search-UnifiedAuditLogCustom -StartDate '$startValue' -EndDate '$endValue' -ResultSize $sizeValue"
        if ($selectedOperations.Count -gt 0) {
            $operationsString = '"' + ($selectedOperations -join '","') + '"'
            $command += " -Operations @($operationsString)"
        }

        $userIds = @()
        if (-not [string]::IsNullOrWhiteSpace($userIdsBox.Text)) {
            $userIds = $userIdsBox.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
        if ($userIds.Count -gt 0) {
            $userIdsString = '"' + ($userIds -join '","') + '"'
            $command += " -UserIds @($userIdsString)"
        }

        if ($simpleViewCheckBox.IsChecked -eq $true) {
            $command += ' -SimpleView'
        }

        $commandBox.Text = $command
    }

    $operationsSearchBox.Add_TextChanged({
            & $refreshOperationsList
            & $buildCommand
        })

    $userIdsBox.Add_TextChanged({ & $buildCommand })

    $availableOperationsListBox.Add_MouseDoubleClick({
            if ($availableOperationsListBox.SelectedItem -and -not $selectedOperationsListBox.Items.Contains($availableOperationsListBox.SelectedItem)) {
                $null = $selectedOperationsListBox.Items.Add($availableOperationsListBox.SelectedItem)
                & $buildCommand
            }
        })

    $availableOperationsListBox.Add_SelectionChanged({
            & $buildCommand
        })

    $selectedOperationsListBox.Add_MouseDoubleClick({
            if ($selectedOperationsListBox.SelectedItem) {
                $selectedOperationsListBox.Items.Remove($selectedOperationsListBox.SelectedItem)
                & $buildCommand
            }
        })

    $addOperationButton.Add_Click({
            if ($availableOperationsListBox.SelectedItems.Count -gt 0) {
                foreach ($item in @($availableOperationsListBox.SelectedItems)) {
                    if (-not $selectedOperationsListBox.Items.Contains($item)) {
                        $null = $selectedOperationsListBox.Items.Add($item)
                    }
                }
                & $buildCommand
            }
        })

    $removeOperationButton.Add_Click({
            if ($selectedOperationsListBox.SelectedItems.Count -gt 0) {
                foreach ($item in @($selectedOperationsListBox.SelectedItems)) {
                    $selectedOperationsListBox.Items.Remove($item)
                }
                & $buildCommand
            }
        })

    $clearOperationsButton.Add_Click({
            $selectedOperationsListBox.Items.Clear()
            & $buildCommand
        })

    $addCustomOperation = {
        $raw = $operationsSearchBox.Text
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return
        }
        foreach ($entry in ($raw -split '[,;]')) {
            $opName = $entry.Trim()
            if ([string]::IsNullOrWhiteSpace($opName)) {
                continue
            }
            $display = "$opName [custom]"
            $operationLookupByDisplay[$display] = $opName
            if (-not $selectedOperationsListBox.Items.Contains($display)) {
                $null = $selectedOperationsListBox.Items.Add($display)
            }
        }
        $operationsSearchBox.Text = ''
        & $buildCommand
    }

    $addCustomOperationButton.Add_Click({ & $addCustomOperation })

    $operationsSearchBox.Add_KeyDown({
            param($source, $e)
            if ($e.Key -eq 'Return') {
                & $addCustomOperation
                $e.Handled = $true
            }
        })

    $selectedOperationsListBox.Add_SelectionChanged({
            & $buildCommand
        })

    $startDateBox.Add_TextChanged({
            & $buildCommand
        })

    $endDateBox.Add_TextChanged({
            & $buildCommand
        })

    $startDatePicker.Add_SelectedDateChanged({
            if ($startDatePicker.SelectedDate) {
                $existingTime = '00:00'
                if ($startDateBox.Text -match '^\d{4}-\d{2}-\d{2}\s(\d{2}:\d{2})$') {
                    $existingTime = $Matches[1]
                }
                $startDateBox.Text = ([datetime]$startDatePicker.SelectedDate).ToString("yyyy-MM-dd $existingTime")
            }
        })

    $endDatePicker.Add_SelectedDateChanged({
            if ($endDatePicker.SelectedDate) {
                $existingTime = '23:59'
                if ($endDateBox.Text -match '^\d{4}-\d{2}-\d{2}\s(\d{2}:\d{2})$') {
                    $existingTime = $Matches[1]
                }
                $endDateBox.Text = ([datetime]$endDatePicker.SelectedDate).ToString("yyyy-MM-dd $existingTime")
            }
        })

    $startYesterdayButton.Add_Click({
            $startDateBox.Text = (Get-Date).AddDays(-1).Date.ToString('yyyy-MM-dd 00:00')
            $startDatePicker.SelectedDate = (Get-Date).AddDays(-1).Date
        })

    $startNowButton.Add_Click({
            $now = Get-Date
            $startDateBox.Text = $now.ToString('yyyy-MM-dd HH:mm')
            $startDatePicker.SelectedDate = $now.Date
        })

    $endTodayButton.Add_Click({
            $endDateBox.Text = (Get-Date).Date.ToString('yyyy-MM-dd 23:59')
            $endDatePicker.SelectedDate = (Get-Date).Date
        })

    $endNowButton.Add_Click({
            $now = Get-Date
            $endDateBox.Text = $now.ToString('yyyy-MM-dd HH:mm')
            $endDatePicker.SelectedDate = $now.Date
        })

    $resultSizeBox.Add_TextChanged({
            & $buildCommand
        })

    $simpleViewCheckBox.Add_Click({
            & $buildCommand
        })

    $loadSharingEventsButton.Add_Click({
            # Load preset: Sharing & SharePoint/OneDrive Events from 2026-01-01 to today
            $startDateBox.Text = '2026-01-01 00:00'
            $endDateBox.Text = (Get-Date).ToString('yyyy-MM-dd 23:59')
            $startDatePicker.SelectedDate = [datetime]'2026-01-01'
            $endDatePicker.SelectedDate = (Get-Date).Date
            $resultSizeBox.Text = '5000'
            $simpleViewCheckBox.IsChecked = $true

            # Sharing operations documented in MS Learn (audit-log-sharing).
            # EmailAuthOTPAuthenticationSucceeded is added (as raw/custom) because external sharing
            # flows often rely on email OTP authentication.
            $sharingOperations = @(
                'SharingInvitationCreated',
                'SharingInvitationAccepted',
                'AnonymousLinkCreated',
                'AnonymousLinkUsed',
                'SecureLinkCreated',
                'AddedToSecureLink',
                'SharingSet',
                'AddedToGroup',
                'EmailAuthOTPAuthenticationSucceeded'
            )

            $selectedOperationsListBox.Items.Clear()
            foreach ($op in $sharingOperations) {
                $matchedDisplay = $null
                foreach ($display in $availableOperationsListBox.Items) {
                    if ($display.Contains("[$op]")) {
                        $matchedDisplay = $display
                        break
                    }
                }
                if ($matchedDisplay) {
                    $null = $selectedOperationsListBox.Items.Add($matchedDisplay)
                }
                else {
                    $customDisplay = "$op [custom]"
                    $operationLookupByDisplay[$customDisplay] = $op
                    $null = $selectedOperationsListBox.Items.Add($customDisplay)
                }
            }

            & $buildCommand
        })

    $copyButton.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($commandBox.Text)) {
                [System.Windows.Clipboard]::SetText($commandBox.Text)
            }
        })

    $runButton.Add_Click({
            $testExoConnection = {
                if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
                    return [bool](Get-ConnectionInformation -ErrorAction SilentlyContinue |
                        Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' })
                }
                return $false
            }

            $exoConnected = & $testExoConnection
            if (-not $exoConnected) {
                $answer = [System.Windows.MessageBox]::Show(
                    "Not connected to Exchange Online.`n`nDo you want to connect now? A sign-in prompt will open.",
                    'Exchange Online connection required',
                    'YesNo',
                    'Question')
                if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
                    return
                }

                $window.Cursor = [System.Windows.Input.Cursors]::Wait
                try {
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                }
                catch {
                    [System.Windows.MessageBox]::Show(
                        "Failed to connect to Exchange Online:`n$($_.Exception.Message)",
                        'Connection failed',
                        'OK',
                        'Error') | Out-Null
                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    return
                }
                $window.Cursor = [System.Windows.Input.Cursors]::Arrow

                $exoConnected = & $testExoConnection
                if (-not $exoConnected) {
                    [System.Windows.MessageBox]::Show(
                        'Exchange Online connection could not be verified. Please retry.',
                        'Connection not active',
                        'OK',
                        'Error') | Out-Null
                    return
                }
            }

            $startDate = & $parseDateInput -InputText $startDateBox.Text -Fallback ((Get-Date).AddDays(-1).Date)
            $endDate = & $parseDateInput -InputText $endDateBox.Text -Fallback ((Get-Date).Date.AddHours(23).AddMinutes(59))

            $sizeValue = 5000
            [void][int]::TryParse($resultSizeBox.Text, [ref]$sizeValue)
            if ($sizeValue -lt 1) { $sizeValue = 1 }

            $selectedOperations = @()
            foreach ($selectedDisplay in $selectedOperationsListBox.Items) {
                if ($operationLookupByDisplay.ContainsKey([string]$selectedDisplay)) {
                    $selectedOperations += $operationLookupByDisplay[[string]$selectedDisplay]
                }
            }

            $runParams = @{
                StartDate  = $startDate
                EndDate    = $endDate
                ResultSize = $sizeValue
            }

            if ($selectedOperations.Count -gt 0) {
                $runParams['Operations'] = $selectedOperations
            }

            $userIdsForRun = @()
            if (-not [string]::IsNullOrWhiteSpace($userIdsBox.Text)) {
                $userIdsForRun = $userIdsBox.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            }
            if ($userIdsForRun.Count -gt 0) {
                $runParams['UserIds'] = $userIdsForRun
            }

            if ($simpleViewCheckBox.IsChecked -eq $true) {
                $runParams['SimpleView'] = $true
            }

            # Resolve a short tenant name (e.g. 'contoso.onmicrosoft.com' -> 'contoso') for the filename
            $tenantName = 'tenant'
            try {
                $exoConn = Get-ConnectionInformation -ErrorAction Stop |
                    Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' } |
                    Select-Object -First 1
                if ($exoConn -and $exoConn.Organization) {
                    $tenantName = ($exoConn.Organization -split '\.')[0]
                }
            }
            catch {
                Write-Verbose "Could not resolve tenant name for filename: $($_.Exception.Message)"
            }
            $tenantSafe = ($tenantName -replace '[^A-Za-z0-9\-_]', '_')
            $executionStamp = (Get-Date).ToString('yyyy-MM-dd-HHmm')

            $defaultFileName = "${executionStamp}_UnifiedAuditLog_${tenantSafe}_$($startDate.ToString('yyyyMMdd-HHmm'))_to_$($endDate.ToString('yyyyMMdd-HHmm')).xlsx"
            $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
            $saveDialog.Title = 'Save audit log results as Excel'
            $saveDialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx'
            $saveDialog.FileName = $defaultFileName
            $saveDialog.InitialDirectory = [Environment]::GetFolderPath('UserProfile')
            if (-not $saveDialog.ShowDialog()) {
                return
            }
            $excelPath = $saveDialog.FileName

            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            try {
                $results = Search-UnifiedAuditLogCustom @runParams

                if (-not $results -or @($results).Count -eq 0) {
                    [System.Windows.MessageBox]::Show('No audit log entries returned for the selected filters.', 'No results', 'OK', 'Warning') | Out-Null
                    return
                }

                $rawCount = @($results).Count
                $dedupKey = $null
                foreach ($candidate in @('Identity', 'Id')) {
                    if ($results[0].PSObject.Properties.Name -contains $candidate) {
                        $dedupKey = $candidate
                        break
                    }
                }
                if ($dedupKey) {
                    $results = $results | Sort-Object -Property $dedupKey -Unique
                }
                $finalCount = @($results).Count
                $duplicatesRemoved = $rawCount - $finalCount

                if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                    [System.Windows.MessageBox]::Show('ImportExcel module is not installed. Run: Install-Module ImportExcel', 'Missing module', 'OK', 'Error') | Out-Null
                    return
                }

                $excelSplat = @{
                    Path          = $excelPath
                    WorksheetName = 'AuditLog'
                    AutoSize      = $true
                    AutoFilter    = $true
                    FreezeTopRow  = $true
                    TableName     = 'AuditLog'
                    ClearSheet    = $true
                }
                $results | Export-Excel @excelSplat

                $dedupMessage = if ($duplicatesRemoved -gt 0) { "`n($duplicatesRemoved duplicate row(s) removed)" } else { '' }
                $answer = [System.Windows.MessageBox]::Show("$finalCount row(s) exported to:`n$excelPath$dedupMessage`n`nOpen the file now?", 'Export complete', 'YesNo', 'Information')
                if ($answer -eq [System.Windows.MessageBoxResult]::Yes) {
                    Start-Process -FilePath $excelPath
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Error during export:`n$($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
            }
            finally {
                $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            }
        })

    $closeButton.Add_Click({
            $window.Close()
        })

    $retentionInfoButton.Add_Click({
            $retentionXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="How audit log retention works"
        Width="680" Height="500"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        Background="#F3F5F8"
        FontFamily="Segoe UI">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="0,6,8,0"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="MinHeight" Value="30"/>
            <Setter Property="MinWidth" Value="120"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="Foreground" Value="#0F172A"/>
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#FDEFD0" CornerRadius="10" Padding="14" BorderBrush="#F1C97A" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Audit log retention - what Microsoft says vs. what actually works"
                           FontWeight="SemiBold" FontSize="14" Foreground="#7A3A00" TextWrapping="Wrap"/>
            </StackPanel>
        </Border>

        <Border Grid.Row="1" Margin="0,12,0,0" Background="White" CornerRadius="10" Padding="16"
                BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock FontWeight="SemiBold" Foreground="#0F172A" Margin="0,0,0,4"
                           Text="Microsoft's official retention tiers"/>
                <TextBlock TextWrapping="Wrap" Foreground="#334155"
                           Text="- 180 days: default retention since October 2023 (no specific license required)."/>
                <TextBlock TextWrapping="Wrap" Foreground="#334155"
                           Text="- Up to 1 year: requires Office 365 E5 / Microsoft 365 E5 / Microsoft 365 E7, Microsoft Purview Suite (formerly Microsoft 365 E5 Compliance), or the E5 eDiscovery and Audit add-on."/>
                <TextBlock TextWrapping="Wrap" Foreground="#334155"
                           Text="- Up to 10 years: requires the above plus the 10-year audit log retention add-on."/>
                <TextBlock TextWrapping="Wrap" Foreground="#334155" Margin="0,4,0,0"
                           Text="Get-AdminAuditLogConfig nonetheless returns AdminAuditLogAgeLimit = 90 days."/>

                <TextBlock FontWeight="SemiBold" Foreground="#0F172A" Margin="0,12,0,4"
                           Text="In practice: up to 365 days on all tenants"/>
                <TextBlock TextWrapping="Wrap" Foreground="#334155"
                           Text="Search-UnifiedAuditLog accepts queries up to 365 days back even on non-E5 tenants, regardless of the values advertised by the documentation and the PowerShell configuration."/>

                <TextBlock Margin="0,12,0,0" TextWrapping="Wrap" Foreground="#64748B" FontStyle="Italic"
                           Text="Click 'Open article' to read the full write-up with PowerShell examples."/>
            </StackPanel>
        </Border>

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0">
            <Button x:Name="OpenArticleButton" Content="Open article"/>
            <Button x:Name="CloseRetentionButton" Content="Close"/>
        </StackPanel>
    </Grid>
</Window>
'@

            $retentionReader = New-Object System.Xml.XmlNodeReader ([xml]$retentionXaml)
            $retentionWindow = [Windows.Markup.XamlReader]::Load($retentionReader)
            $retentionWindow.Owner = $window

            $openArticleButton = $retentionWindow.FindName('OpenArticleButton')
            $closeRetentionButton = $retentionWindow.FindName('CloseRetentionButton')

            $openArticleButton.Add_Click({
                    Start-Process 'https://itpro-tips.com/microsoft-365-audit-logs-are-now-retained-for-365-days-for-all-tenants-with-powershell/'
                    $retentionWindow.Close()
                })

            $closeRetentionButton.Add_Click({ $retentionWindow.Close() })

            [void]$retentionWindow.ShowDialog()
        })

    & $refreshOperationsList
    & $buildCommand

    $window.Add_Closed({
            try {
                $availableOperationsListBox.ItemsSource = $null
                $availableOperationsListBox.Items.Clear()
                $selectedOperationsListBox.Items.Clear()
            }
            catch { }
        })

    try {
        if ($splash) { $splash.Close(); $splash = $null }
        [void]$window.ShowDialog()
    }
    finally {
        if ($splash) { try { $splash.Close() } catch { } }
        $window = $null
        $availableOperationsListBox = $null
        $selectedOperationsListBox = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
}
