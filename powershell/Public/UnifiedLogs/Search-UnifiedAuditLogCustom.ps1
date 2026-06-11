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

    .PARAMETER ChunkDays
    Size (in days) of each sub-window used to split the StartDate/EndDate range. Defaults to 7.
    The function loops over the full range one chunk at a time and uses session pagination inside each chunk,
    which avoids the server-side 'Search duration too long' error encountered on very wide windows.
    Lower this value (e.g. 1 or 3) if a chunk itself returns the 'too long' error.

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
        [switch]$HelperGUI,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 90)]
        [int]$ChunkDays = 7
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

    # Chunk the date range to avoid the server-side 'Search duration too long' error on wide windows.
    # Within each chunk we use session pagination (-SessionId + -SessionCommand ReturnLargeSet) so we can
    # gather more than 5000 records per sub-window. The global -ResultSize is enforced as the overall cap.
    [System.Collections.Generic.List[object]]$auditLogs = @()
    $chunkSpan = New-TimeSpan -Days $ChunkDays
    $cursor    = $StartDate
    $chunkIdx  = 0
    $totalSeconds = [math]::Max(1, ($EndDate - $StartDate).TotalSeconds)

    while ($cursor -lt $EndDate -and $auditLogs.Count -lt $ResultSize) {
        $chunkIdx++
        $chunkEnd = $cursor + $chunkSpan
        if ($chunkEnd -gt $EndDate) { $chunkEnd = $EndDate }

        $percent = [math]::Min(100, [math]::Max(0, [int]((($cursor - $StartDate).TotalSeconds / $totalSeconds) * 100)))
        Write-Progress -Activity 'Searching Unified Audit Log' -Status "Window $($cursor.ToString('yyyy-MM-dd')) -> $($chunkEnd.ToString('yyyy-MM-dd')) | $($auditLogs.Count)/$ResultSize records" -PercentComplete $percent

        Write-Verbose "Chunk $chunkIdx : $($cursor.ToString('yyyy-MM-dd HH:mm')) -> $($chunkEnd.ToString('yyyy-MM-dd HH:mm'))"

        $chunkParams = @{}
        foreach ($key in $searchParams.Keys) { $chunkParams[$key] = $searchParams[$key] }
        $chunkParams['StartDate'] = $cursor
        $chunkParams['EndDate']   = $chunkEnd
        # Per-call page size is capped server-side at 5000.
        $chunkParams['ResultSize'] = 5000

        $sessionId = [Guid]::NewGuid().ToString()
        $pageIndex = 0
        do {
            $pageIndex++
            try {
                $page = Search-UnifiedAuditLog @chunkParams -SessionId $sessionId -SessionCommand ReturnLargeSet -ErrorAction Stop
            }
            catch {
                Write-Warning "Search-UnifiedAuditLog failed on chunk $chunkIdx page $pageIndex ($($cursor.ToString('yyyy-MM-dd HH:mm')) -> $($chunkEnd.ToString('yyyy-MM-dd HH:mm'))): $($_.Exception.Message). Try a smaller -ChunkDays value (current: $ChunkDays)."
                $page = $null
            }

            if ($page) {
                foreach ($entry in $page) {
                    if ($auditLogs.Count -ge $ResultSize) { break }
                    $auditLogs.Add($entry)
                }
                Write-Verbose "Chunk $chunkIdx page $pageIndex : +$($page.Count) records (total $($auditLogs.Count)/$ResultSize)"
            }
        } while ($page -and $page.Count -gt 0 -and $auditLogs.Count -lt $ResultSize)

        # Advance the cursor by one second past chunkEnd to avoid re-fetching the boundary record.
        $cursor = $chunkEnd.AddSeconds(1)
    }

    Write-Progress -Activity 'Searching Unified Audit Log' -Completed

    if ($auditLogs.Count -eq 0) {
        Write-Warning 'Search-UnifiedAuditLog returned no records for the specified filters and time window.'
        return
    }

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
        Width="1200"
        MinWidth="1100" MinHeight="400" MaxHeight="900"
        SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        Background="#FAFAFA"
        FontFamily="Segoe UI"
        TextOptions.TextFormattingMode="Display">
    <Window.Resources>
        <SolidColorBrush x:Key="AccentBrush" Color="#0F6CBD"/>
        <SolidColorBrush x:Key="AccentHoverBrush" Color="#115EA3"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="StrokeBrush" Color="#E1E1E1"/>
        <SolidColorBrush x:Key="StrokeStrongBrush" Color="#C7C7C7"/>
        <SolidColorBrush x:Key="TextPrimaryBrush" Color="#242424"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#616161"/>
        <SolidColorBrush x:Key="TextHintBrush" Color="#8A8A8A"/>
        <SolidColorBrush x:Key="ChromeBrush" Color="#F5F5F5"/>
        <SolidColorBrush x:Key="ChromeHoverBrush" Color="#EBEBEB"/>

        <Style x:Key="SectionTitleStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="Margin" Value="0,0,0,4"/>
        </Style>
        <Style x:Key="FieldLabelStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
            <Setter Property="Margin" Value="0,0,0,2"/>
        </Style>
        <Style x:Key="HelperTextStyle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Foreground" Value="{StaticResource TextHintBrush}"/>
            <Setter Property="Margin" Value="0,2,0,0"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
        </Style>

        <Style x:Key="CardStyle" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource StrokeBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="4"/>
            <Setter Property="Padding" Value="10"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,3"/>
            <Setter Property="BorderBrush" Value="{StaticResource StrokeStrongBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="MinHeight" Value="24"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style x:Key="CompactInputStyle" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
            <Setter Property="Height" Value="24"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <SolidColorBrush x:Key="AccentSoftBrush" Color="#EAF2FB"/>

        <Style TargetType="Calendar">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="{StaticResource StrokeBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
        </Style>

        <Style TargetType="CalendarDayButton">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="MinWidth" Value="28"/>
            <Setter Property="MinHeight" Value="28"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CalendarDayButton">
                        <Border x:Name="bg" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bg" Property="Background" Value="{StaticResource AccentSoftBrush}"/>
                            </Trigger>
                            <Trigger Property="IsToday" Value="True">
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="bg" Property="Background" Value="{StaticResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsInactive" Value="True">
                                <Setter Property="Foreground" Value="{StaticResource TextHintBrush}"/>
                            </Trigger>
                            <Trigger Property="IsBlackedOut" Value="True">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="CalendarButton">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="MinWidth" Value="56"/>
            <Setter Property="MinHeight" Value="36"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Margin" Value="2"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CalendarButton">
                        <Border x:Name="bg" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bg" Property="Background" Value="{StaticResource AccentSoftBrush}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="bg" Property="Background" Value="{StaticResource AccentBrush}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="HasSelectedDays" Value="True">
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                            <Trigger Property="IsInactive" Value="True">
                                <Setter Property="Foreground" Value="{StaticResource TextHintBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="CalendarItem">
            <Setter Property="Margin" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CalendarItem">
                        <ControlTemplate.Resources>
                            <DataTemplate x:Key="{x:Static CalendarItem.DayTitleTemplateResourceKey}">
                                <TextBlock Text="{Binding}" FontSize="11" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextSecondaryBrush}"
                                           HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,4"/>
                            </DataTemplate>
                        </ControlTemplate.Resources>
                        <Border Background="White" Padding="8" CornerRadius="4">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>

                                <Grid Grid.Row="0">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <Button x:Name="PART_PreviousButton" Grid.Column="0"
                                            Width="28" Height="28" Style="{StaticResource CalendarToggleStyle}"
                                            Focusable="False">
                                        <Path Width="8" Height="10" Stretch="Uniform"
                                              Fill="{StaticResource TextSecondaryBrush}"
                                              Data="M 7,0 L 0,5 L 7,10 Z"/>
                                    </Button>
                                    <Button x:Name="PART_HeaderButton" Grid.Column="1"
                                            Style="{StaticResource CalendarToggleStyle}"
                                            HorizontalContentAlignment="Center"
                                            FontWeight="SemiBold" FontSize="12"
                                            Foreground="{StaticResource TextPrimaryBrush}"
                                            Focusable="False"/>
                                    <Button x:Name="PART_NextButton" Grid.Column="2"
                                            Width="28" Height="28" Style="{StaticResource CalendarToggleStyle}"
                                            Focusable="False">
                                        <Path Width="8" Height="10" Stretch="Uniform"
                                              Fill="{StaticResource TextSecondaryBrush}"
                                              Data="M 0,0 L 7,5 L 0,10 Z"/>
                                    </Button>
                                </Grid>

                                <Grid Grid.Row="1" Margin="0,6,0,0">
                                    <Grid x:Name="PART_MonthView" Visibility="Visible"/>
                                    <Grid x:Name="PART_YearView" Visibility="Hidden"/>
                                </Grid>
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DatePickerTextBox">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="6,0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DatePickerTextBox">
                        <Grid>
                            <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="Center" Focusable="False"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="CalendarToggleStyle" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="MinWidth" Value="0"/>
            <Setter Property="MinHeight" Value="0"/>
            <Setter Property="Foreground" Value="{StaticResource TextSecondaryBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bg" Background="{TemplateBinding Background}" CornerRadius="2" Padding="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bg" Property="Background" Value="#EAF2FB"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DatePicker">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="MinHeight" Value="26"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="{StaticResource StrokeStrongBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DatePicker">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="2"
                                SnapsToDevicePixels="True">
                            <Grid x:Name="PART_Root">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="26"/>
                                </Grid.ColumnDefinitions>
                                <DatePickerTextBox x:Name="PART_TextBox" Grid.Column="0"
                                                   Foreground="{TemplateBinding Foreground}"
                                                   VerticalAlignment="Center"/>
                                <Button x:Name="PART_Button" Grid.Column="1" Width="26"
                                        Style="{StaticResource CalendarToggleStyle}"
                                        Focusable="False">
                                    <Path Width="13" Height="13" Stretch="Uniform"
                                          Stroke="{StaticResource TextSecondaryBrush}" StrokeThickness="1.1"
                                          Fill="Transparent" SnapsToDevicePixels="True"
                                          Data="M0,3 L14,3 M3,0 L3,5 M11,0 L11,5 M0,7 L14,7 M0,3 L0,14 L14,14 L14,3 Z"/>
                                </Button>
                                <Popup x:Name="PART_Popup"
                                       PlacementTarget="{Binding ElementName=PART_Button}"
                                       Placement="Bottom" StaysOpen="False"/>
                                <Grid x:Name="PART_DisabledVisual" Background="#80FFFFFF" Visibility="Collapsed"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocusWithin" Value="True">
                                <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="PART_DisabledVisual" Property="Visibility" Value="Visible"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ListBox">
            <Setter Property="BorderBrush" Value="{StaticResource StrokeBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="0"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Padding" Value="10,4"/>
            <Setter Property="MinHeight" Value="26"/>
            <Setter Property="MinWidth" Value="72"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{StaticResource StrokeStrongBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimaryBrush}"/>
            <Setter Property="Background" Value="{StaticResource ChromeBrush}"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource ChromeHoverBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="PrimaryButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentHoverBrush}"/>
                    <Setter Property="BorderBrush" Value="{StaticResource AccentHoverBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="NeutralButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"/>
        <Style x:Key="SuccessButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"/>
        <Style x:Key="GhostButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
            <Setter Property="MinWidth" Value="0"/>
            <Setter Property="Padding" Value="6,3"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#EAF2FB"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="White" BorderBrush="{StaticResource StrokeBrush}" BorderThickness="0,0,0,1" Padding="20,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Unified Audit Log search" FontSize="15" FontWeight="SemiBold" Foreground="{StaticResource TextPrimaryBrush}"/>
                    <TextBlock Text="Build, preview and run a Microsoft 365 Unified Audit Log query, then export the results to Excel."
                               FontSize="11" Foreground="{StaticResource TextSecondaryBrush}" Margin="0,1,0,0"/>
                </StackPanel>
                <TextBlock x:Name="HeaderVersionText" Grid.Column="1" VerticalAlignment="Center"
                           FontSize="11" Foreground="{StaticResource TextHintBrush}"/>
            </Grid>
        </Border>

        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="20,12,20,12">
            <StackPanel>

                <Border Style="{StaticResource CardStyle}">
                    <StackPanel>
                        <TextBlock Text="Time range" Style="{StaticResource SectionTitleStyle}"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="220"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" Margin="0,0,16,0">
                                <TextBlock Text="Start date" Style="{StaticResource FieldLabelStyle}"/>
                                <DatePicker x:Name="StartDatePicker"/>
                                <TextBox x:Name="StartDateBox" Margin="0,6,0,0" Style="{StaticResource CompactInputStyle}" ToolTip="Format: yyyy-MM-dd HH:mm"/>
                                <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                                    <Button x:Name="StartYesterdayButton" Content="Yesterday 00:00" Style="{StaticResource GhostButtonStyle}"/>
                                    <Button x:Name="StartNowButton" Content="Now" Style="{StaticResource GhostButtonStyle}"/>
                                </StackPanel>
                            </StackPanel>

                            <StackPanel Grid.Column="1" Margin="0,0,16,0">
                                <TextBlock Text="End date" Style="{StaticResource FieldLabelStyle}"/>
                                <DatePicker x:Name="EndDatePicker"/>
                                <TextBox x:Name="EndDateBox" Margin="0,6,0,0" Style="{StaticResource CompactInputStyle}" ToolTip="Format: yyyy-MM-dd HH:mm"/>
                                <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                                    <Button x:Name="EndTodayButton" Content="Today 23:59" Style="{StaticResource GhostButtonStyle}"/>
                                    <Button x:Name="EndNowButton" Content="Now" Style="{StaticResource GhostButtonStyle}"/>
                                </StackPanel>
                            </StackPanel>

                            <StackPanel Grid.Column="2">
                                <TextBlock Text="Result size (max)" Style="{StaticResource FieldLabelStyle}"/>
                                <TextBox x:Name="ResultSizeBox" Text="5000" Style="{StaticResource CompactInputStyle}"/>
                                <CheckBox x:Name="SimpleViewCheckBox" Content="Flatten records (Simple view)" Margin="0,10,0,0" IsChecked="True"/>
                                <Button x:Name="RetentionInfoButton" Content="How does retention work?" Margin="-6,8,0,0" HorizontalAlignment="Left" Style="{StaticResource GhostButtonStyle}"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource CardStyle}" Margin="0,6,0,0">
                    <StackPanel>
                        <TextBlock Text="Presets" Style="{StaticResource SectionTitleStyle}"/>
                        <StackPanel Orientation="Horizontal">
                            <Button x:Name="LoadSharingEventsButton" Content="Sharing activity (SPO/OneDrive)"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource CardStyle}" Margin="0,6,0,0">
                    <StackPanel>
                        <TextBlock Text="Filters" Style="{StaticResource SectionTitleStyle}"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0" Margin="0,0,16,0">
                                <TextBlock Text="Operations" Style="{StaticResource FieldLabelStyle}"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="OperationsSearchBox" Grid.Column="0" Style="{StaticResource CompactInputStyle}"
                                             ToolTip="Type to filter the list below, OR type a raw cmdlet (e.g. New-TransportRule) and press Enter / click 'Add as raw'"/>
                                    <Button x:Name="AddCustomOperationButton" Grid.Column="1" Content="Add as raw" Margin="8,0,0,0"/>
                                </Grid>
                                <TextBlock Style="{StaticResource HelperTextStyle}"
                                           Text="Search a friendly name OR type raw cmdlets (comma/semicolon-separated)."/>
                            </StackPanel>

                            <StackPanel Grid.Column="1">
                                <TextBlock Text="User IDs" Style="{StaticResource FieldLabelStyle}"/>
                                <TextBox x:Name="UserIdsBox" Style="{StaticResource CompactInputStyle}"
                                         ToolTip="Optional. Filter the search by one or more user principal names. Leave empty to search all users."/>
                                <TextBlock Style="{StaticResource HelperTextStyle}"
                                           Text="Comma or semicolon-separated. Example: alice@contoso.com;bob@contoso.com"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource CardStyle}" Margin="0,6,0,0">
                    <StackPanel>
                        <TextBlock Text="Selected operations" Style="{StaticResource SectionTitleStyle}"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <StackPanel Grid.Column="0">
                                <TextBlock Text="Available" Style="{StaticResource FieldLabelStyle}"/>
                                <ListBox x:Name="AvailableOperationsListBox" Height="140" SelectionMode="Extended"
                                         ScrollViewer.HorizontalScrollBarVisibility="Auto"
                                         ScrollViewer.CanContentScroll="True"
                                         VirtualizingStackPanel.IsVirtualizing="True"
                                         VirtualizingStackPanel.VirtualizationMode="Recycling"/>
                            </StackPanel>

                            <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="14,0">
                                <Button x:Name="AddOperationButton" Content="Add &#x2192;" Width="92" Margin="0,0,0,8"/>
                                <Button x:Name="RemoveOperationButton" Content="&#x2190; Remove" Width="92" Margin="0,0,0,8"/>
                                <Button x:Name="ClearOperationsButton" Content="Clear" Width="92" Style="{StaticResource GhostButtonStyle}"/>
                            </StackPanel>

                            <StackPanel Grid.Column="2">
                                <TextBlock Text="Selected" Style="{StaticResource FieldLabelStyle}"/>
                                <ListBox x:Name="SelectedOperationsListBox" Height="140" SelectionMode="Extended"
                                         ScrollViewer.HorizontalScrollBarVisibility="Auto"
                                         ScrollViewer.CanContentScroll="True"
                                         VirtualizingStackPanel.IsVirtualizing="True"
                                         VirtualizingStackPanel.VirtualizationMode="Recycling"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </Border>

                <Border Style="{StaticResource CardStyle}" Margin="0,6,0,0">
                    <StackPanel>
                        <TextBlock Text="Generated command" Style="{StaticResource SectionTitleStyle}"/>
                        <TextBox x:Name="CommandBox" MinHeight="56" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                 IsReadOnly="True" AcceptsReturn="True" FontFamily="Consolas" FontSize="12"
                                 Background="#F9F9F9"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </ScrollViewer>

        <Border Grid.Row="2" Background="White" BorderBrush="{StaticResource StrokeBrush}" BorderThickness="0,1,0,0" Padding="20,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock x:Name="FooterText" Grid.Column="0" VerticalAlignment="Center"
                           Foreground="{StaticResource TextHintBrush}" FontSize="11">
                    <Run Text="by "/>
                    <Hyperlink x:Name="ClidsysLink" NavigateUri="https://clidsys.com"
                               Foreground="{StaticResource AccentBrush}" TextDecorations="None">
                        <Run Text="Clidsys"/>
                    </Hyperlink>
                    <Run Text=" - "/>
                    <Hyperlink x:Name="BastienLink" NavigateUri="https://www.linkedin.com/in/perez-bastien/"
                               Foreground="{StaticResource AccentBrush}" TextDecorations="None">
                        <Run Text="Bastien Perez"/>
                    </Hyperlink>
                </TextBlock>

                <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="CopyButton" Content="Copy command" Style="{StaticResource SuccessButtonStyle}"
                            ToolTip="Copies the generated command to the clipboard. Rarely needed here - use 'Run now' to execute the query and export the results directly."/>
                    <Button x:Name="CloseButton" Content="Close" Style="{StaticResource NeutralButtonStyle}"/>
                    <Button x:Name="RunButton" Content="Run now" Style="{StaticResource PrimaryButtonStyle}" Margin="8,0,0,0"
                            ToolTip="Runs the query against the Unified Audit Log. A file dialog will open first to choose where to save the results - in this mode, results are always exported to an Excel file."/>
                </StackPanel>
            </Grid>
        </Border>
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

    $clidsysLink = $window.FindName('ClidsysLink')
    if ($clidsysLink) {
        $clidsysLink.Add_RequestNavigate({
                param($source, $e)
                Start-Process $e.Uri.AbsoluteUri
                $e.Handled = $true
            })
    }

    $bastienLink = $window.FindName('BastienLink')
    if ($bastienLink) {
        $bastienLink.Add_RequestNavigate({
                param($source, $e)
                Start-Process $e.Uri.AbsoluteUri
                $e.Handled = $true
            })
    }

    $headerVersionText = $window.FindName('HeaderVersionText')
    if ($headerVersionText -and $moduleVersion) {
        $headerVersionText.Text = "PS365 $moduleVersion"
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

        [System.Collections.Generic.List[string]]$selectedOperations = @()
        foreach ($selectedDisplay in $selectedOperationsListBox.Items) {
            if ($operationLookupByDisplay.ContainsKey([string]$selectedDisplay)) {
                $selectedOperations.Add($operationLookupByDisplay[[string]$selectedDisplay])
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
                    $null = [System.Windows.MessageBox]::Show(
                        "Failed to connect to Exchange Online:`n$($_.Exception.Message)",
                        'Connection failed',
                        'OK',
                        'Error')
                    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                    return
                }
                $window.Cursor = [System.Windows.Input.Cursors]::Arrow

                $exoConnected = & $testExoConnection
                if (-not $exoConnected) {
                    $null = [System.Windows.MessageBox]::Show(
                        'Exchange Online connection could not be verified. Please retry.',
                        'Connection not active',
                        'OK',
                        'Error')
                    return
                }
            }

            $startDate = & $parseDateInput -InputText $startDateBox.Text -Fallback ((Get-Date).AddDays(-1).Date)
            $endDate = & $parseDateInput -InputText $endDateBox.Text -Fallback ((Get-Date).Date.AddHours(23).AddMinutes(59))

            $sizeValue = 5000
            [void][int]::TryParse($resultSizeBox.Text, [ref]$sizeValue)
            if ($sizeValue -lt 1) { $sizeValue = 1 }

            [System.Collections.Generic.List[string]]$selectedOperations = @()
            foreach ($selectedDisplay in $selectedOperationsListBox.Items) {
                if ($operationLookupByDisplay.ContainsKey([string]$selectedDisplay)) {
                    $selectedOperations.Add($operationLookupByDisplay[[string]$selectedDisplay])
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
            $executionStamp = (Get-Date).ToString('yyyy-MM-dd-HHmmss')

            $defaultFileName = "${executionStamp}_UnifiedAuditLog_${tenantSafe}_$($startDate.ToString('yyyyMMdd-HHmmss'))_to_$($endDate.ToString('yyyyMMdd-HHmmss')).xlsx"
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
                    $null = [System.Windows.MessageBox]::Show('No audit log entries returned for the selected filters.', 'No results', 'OK', 'Warning')
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
                    $null = [System.Windows.MessageBox]::Show('ImportExcel module is not installed. Run: Install-Module ImportExcel', 'Missing module', 'OK', 'Error')
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
                $null = [System.Windows.MessageBox]::Show("Error during export:`n$($_.Exception.Message)", 'Error', 'OK', 'Error')
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

        <Border Grid.Row="1" Margin="0,6,0,0" Background="White" CornerRadius="10" Padding="16"
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

                <TextBlock Margin="0,6,0,0" TextWrapping="Wrap" Foreground="#64748B" FontStyle="Italic"
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
