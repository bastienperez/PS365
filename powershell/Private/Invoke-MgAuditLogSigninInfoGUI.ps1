function Invoke-MgAuditLogSigninInfoGUI {
    <#
    .SYNOPSIS
        WPF helper interface for Get-MgAuditLogSigninInfo.

    .DESCRIPTION
        Lets the user pick a date/time range (with timezone), apply common
        filters (user, IP, top N, sign-in type, MFA, conditional access policy),
        choose Excel export, preview the equivalent command line and run it.

        Internal helper: called by Get-MgAuditLogSigninInfo -GUI.
    #>
    [CmdletBinding()]
    param()

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    $moduleVersion = $null
    $loadedModule = Get-Module -Name 'PS365' -ErrorAction SilentlyContinue
    if ($loadedModule -and $loadedModule.Version) {
        $moduleVersion = "v$($loadedModule.Version)"
    }

    $splash = Show-Splash `
        -Title 'Get-MgAuditLogSigninInfo' `
        -Subtitle 'Sign-in log query helper' `
        -InitialMessage 'Initializing...' `
        -Version $moduleVersion

    [System.Collections.Generic.List[PSObject]]$capList = @()

    try {
        $splash.Update('Connecting to Microsoft Graph (AuditLog.Read.All, Policy.Read.All)...')
        $null = Connect-MgGraph -Scopes 'AuditLog.Read.All', 'Policy.Read.All' -NoWelcome -ErrorAction Stop
    }
    catch {
        Write-Verbose "Connect-MgGraph failed: $($_.Exception.Message)"
    }

    try {
        $splash.Update('Loading Conditional Access Policies...')
        $caps = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop |
            Select-Object DisplayName, Id, State |
            Sort-Object DisplayName
        if ($caps) { $caps | ForEach-Object { $capList.Add($_) } }
    }
    catch {
        Write-Verbose "Could not load Conditional Access Policies: $($_.Exception.Message)"
    }

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Get-MgAuditLogSigninInfo Helper"
        Width="1080" Height="900"
        MinWidth="1000" MinHeight="800"
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
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="0,4,0,0"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,4,16,0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="0,6,6,0"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="MinWidth" Value="90"/>
            <Setter Property="BorderBrush" Value="#CBD5E1"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="GhostButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#0F172A"/>
        </Style>
        <Style x:Key="NeutralButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#F8FAFC"/>
            <Setter Property="Foreground" Value="#0F172A"/>
        </Style>
        <Style x:Key="PrimaryButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#DE761D"/>
            <Setter Property="BorderBrush" Value="#B85F12"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style x:Key="SuccessButtonStyle" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#0EA5E9"/>
            <Setter Property="BorderBrush" Value="#0284C7"/>
            <Setter Property="Foreground" Value="White"/>
        </Style>
    </Window.Resources>

    <ScrollViewer VerticalScrollBarVisibility="Auto">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Time range card -->
        <Border Grid.Row="0" Background="White" CornerRadius="10" Padding="16" BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Time range" FontWeight="SemiBold"/>
                <Grid Margin="0,6,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0" Margin="0,0,12,0">
                        <TextBlock Text="Preset" FontWeight="SemiBold" FontSize="11"/>
                        <ComboBox x:Name="TimeRangeCombo"/>
                    </StackPanel>

                    <StackPanel Grid.Column="1" Margin="0,0,12,0">
                        <TextBlock Text="StartDate" FontWeight="SemiBold" FontSize="11"/>
                        <DatePicker x:Name="StartDatePicker"/>
                        <TextBox x:Name="StartTimeBox" ToolTip="HH:mm"/>
                    </StackPanel>

                    <StackPanel Grid.Column="2" Margin="0,0,12,0">
                        <TextBlock Text="EndDate" FontWeight="SemiBold" FontSize="11"/>
                        <DatePicker x:Name="EndDatePicker"/>
                        <TextBox x:Name="EndTimeBox" ToolTip="HH:mm"/>
                    </StackPanel>

                    <StackPanel Grid.Column="3">
                        <TextBlock Text="Time zone" FontWeight="SemiBold" FontSize="11"/>
                        <ComboBox x:Name="TimeZoneCombo"/>
                    </StackPanel>
                </Grid>
            </StackPanel>
        </Border>

        <!-- Filters card -->
        <Border Grid.Row="1" Margin="0,12,0,0" Background="White" CornerRadius="10" Padding="16" BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Filters" FontWeight="SemiBold"/>
                <Grid Margin="0,6,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2*"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0" Margin="0,0,12,0">
                        <TextBlock Text="Users (UPNs, one per line or semicolon-separated)" FontWeight="SemiBold" FontSize="11"/>
                        <TextBox x:Name="UsersBox" Height="62" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                    </StackPanel>

                    <StackPanel Grid.Column="1" Margin="0,0,12,0">
                        <TextBlock Text="Top N (LastXSignIns)" FontWeight="SemiBold" FontSize="11"/>
                        <TextBox x:Name="TopNBox"/>
                        <TextBlock Text="IP filter (IPAddresses)" FontWeight="SemiBold" FontSize="11" Margin="0,8,0,0"/>
                        <TextBox x:Name="IpBox"/>
                    </StackPanel>

                    <StackPanel Grid.Column="2">
                        <CheckBox x:Name="ForceNewTokenCheck" Content="Force new token"/>
                        <CheckBox x:Name="ExportExcelCheck" Content="Export to Excel"/>
                    </StackPanel>
                </Grid>

                <TextBlock Text="Sign-in type" FontWeight="SemiBold" Margin="0,12,0,0"/>
                <WrapPanel Margin="0,6,0,0">
                    <CheckBox x:Name="ChkSuccessOnly" Content="Success only"/>
                    <CheckBox x:Name="ChkFailureOnly" Content="Failure only"/>
                    <CheckBox x:Name="ChkBadCredentialsOnly" Content="Bad credentials only"/>
                    <CheckBox x:Name="ChkLastLogonOnly" Content="Last logon only"/>
                    <CheckBox x:Name="ChkBasicAuthOnly" Content="Basic auth only"/>
                    <CheckBox x:Name="ChkMFAOnly" Content="MFA only"/>
                    <CheckBox x:Name="ChkNonMFAOnly" Content="Non-MFA only"/>
                    <CheckBox x:Name="ChkNonInteractive" Content="Non-interactive sign-ins"/>
                    <CheckBox x:Name="ChkServicePrincipal" Content="Service principal sign-ins"/>
                    <CheckBox x:Name="ChkManagedIdentity" Content="Managed identity sign-ins"/>
                </WrapPanel>
            </StackPanel>
        </Border>

        <!-- CAP card -->
        <Border Grid.Row="2" Margin="0,12,0,0" Background="White" CornerRadius="10" Padding="16" BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Conditional Access" FontWeight="SemiBold"/>
                <Grid Margin="0,6,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <ComboBox x:Name="CapCombo" Grid.Column="0" IsEditable="True"/>
                    <Button x:Name="CapRefreshButton" Grid.Column="1" Content="Refresh" Style="{StaticResource GhostButtonStyle}" Margin="6,4,0,0"/>
                </Grid>
                <WrapPanel Margin="0,8,0,0">
                    <CheckBox x:Name="ChkCapNotApplied" Content="Policy not applied"/>
                    <CheckBox x:Name="ChkCapSuccess" Content="CAP success only"/>
                    <CheckBox x:Name="ChkCapFailed" Content="CAP failed only"/>
                    <CheckBox x:Name="ChkAnalyzeReportOnly" Content="Analyze report-only CAP"/>
                </WrapPanel>
            </StackPanel>
        </Border>

        <!-- Generated command card -->
        <Border Grid.Row="3" Margin="0,12,0,0" Background="White" CornerRadius="10" Padding="16" BorderBrush="#E2E8F0" BorderThickness="1">
            <StackPanel>
                <TextBlock Text="Generated command" FontWeight="SemiBold"/>
                <TextBox x:Name="CommandBox" Margin="0,6,0,0" Height="100" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True" AcceptsReturn="True"/>
            </StackPanel>
        </Border>

        <!-- Action buttons -->
        <Grid Grid.Row="4" Margin="0,14,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left">
                <Button x:Name="CopyButton" Content="Copy" Style="{StaticResource SuccessButtonStyle}"/>
                <Button x:Name="RunButton" Content="Run now" Style="{StaticResource PrimaryButtonStyle}"/>
                <Button x:Name="CloseButton" Content="Close" Style="{StaticResource NeutralButtonStyle}"/>
            </StackPanel>
            <TextBlock x:Name="FooterText" Grid.Column="1" Foreground="#64748B" FontSize="11"
                       VerticalAlignment="Center" HorizontalAlignment="Right"
                       Text="by Clidsys - Bastien Perez"/>
        </Grid>
    </Grid>
    </ScrollViewer>
</Window>
'@

    if ($splash) { $splash.Update('Building interface...') }

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # FindName all controls
    $timeRangeCombo = $window.FindName('TimeRangeCombo')
    $startDatePicker = $window.FindName('StartDatePicker')
    $startTimeBox = $window.FindName('StartTimeBox')
    $endDatePicker = $window.FindName('EndDatePicker')
    $endTimeBox = $window.FindName('EndTimeBox')
    $timeZoneCombo = $window.FindName('TimeZoneCombo')
    $usersBox = $window.FindName('UsersBox')
    $topNBox = $window.FindName('TopNBox')
    $ipBox = $window.FindName('IpBox')
    $forceNewTokenCheck = $window.FindName('ForceNewTokenCheck')
    $exportExcelCheck = $window.FindName('ExportExcelCheck')
    $capCombo = $window.FindName('CapCombo')
    $capRefreshButton = $window.FindName('CapRefreshButton')
    $commandBox = $window.FindName('CommandBox')
    $copyButton = $window.FindName('CopyButton')
    $runButton = $window.FindName('RunButton')
    $closeButton = $window.FindName('CloseButton')
    $footerText = $window.FindName('FooterText')

    if ($footerText -and $moduleVersion) {
        $footerText.Text = "by Clidsys - Bastien Perez  $moduleVersion"
    }

    # Populate TimeRange preset
    $timeRangePresets = @('(custom)', 'Last2Minutes', 'Last10Minutes', 'LastHour', 'Last6Hours', 'Last12Hours', 'Last24Hours', 'Last3Days', 'Last7Days', 'Last15Days', 'Maximum')
    foreach ($preset in $timeRangePresets) { [void]$timeRangeCombo.Items.Add($preset) }
    $timeRangeCombo.SelectedIndex = 0

    # Populate TimeZone combo
    foreach ($tz in [System.TimeZoneInfo]::GetSystemTimeZones()) { [void]$timeZoneCombo.Items.Add($tz.Id) }
    $timeZoneCombo.SelectedItem = [System.TimeZoneInfo]::Local.Id

    # Default dates: last 7 days
    $startDatePicker.SelectedDate = (Get-Date).AddDays(-7).Date
    $startTimeBox.Text = '00:00'
    $endDatePicker.SelectedDate = (Get-Date).Date
    $endTimeBox.Text = (Get-Date).ToString('HH:mm')

    # Populate CAP combo
    $populateCap = {
        $capCombo.Items.Clear()
        [void]$capCombo.Items.Add('(none)')
        foreach ($cap in $capList) {
            [void]$capCombo.Items.Add($cap.DisplayName)
        }
        $capCombo.SelectedIndex = 0
    }
    & $populateCap

    # Build command preview
    $buildCommand = {
        $parts = @('Get-MgAuditLogSigninInfo')

        $tr = [string]$timeRangeCombo.SelectedItem
        if ($tr -and $tr -ne '(custom)') {
            $parts += "-TimeRange $tr"
        }
        else {
            if ($startDatePicker.SelectedDate) {
                $startTxt = $startDatePicker.SelectedDate.ToString('yyyy-MM-dd')
                if ($startTimeBox.Text) { $startTxt += " $($startTimeBox.Text)" }
                $parts += "-StartDate '$startTxt'"
            }
            if ($endDatePicker.SelectedDate) {
                $endTxt = $endDatePicker.SelectedDate.ToString('yyyy-MM-dd')
                if ($endTimeBox.Text) { $endTxt += " $($endTimeBox.Text)" }
                $parts += "-EndDate '$endTxt'"
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($usersBox.Text)) {
            $userList = $usersBox.Text -split '[;\r\n]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            if ($userList) {
                $usersJoined = ($userList | ForEach-Object { "'$_'" }) -join ','
                $parts += "-Users $usersJoined"
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($topNBox.Text)) { $parts += "-LastXSignIns $($topNBox.Text)" }
        if (-not [string]::IsNullOrWhiteSpace($ipBox.Text)) { $parts += "-IPAddresses $($ipBox.Text)" }

        $switchMap = [ordered]@{
            ChkSuccessOnly        = 'SuccessOnly'
            ChkFailureOnly        = 'FailureOnly'
            ChkBadCredentialsOnly = 'BadCredentialsOnly'
            ChkLastLogonOnly      = 'LastLogonOnly'
            ChkBasicAuthOnly      = 'BasicAuthenticationOnly'
            ChkMFAOnly            = 'MFASignInsOnly'
            ChkNonMFAOnly         = 'NonMFASignInsOnly'
            ChkNonInteractive     = 'NonInteractiveSignIns'
            ChkServicePrincipal   = 'ServicePrincipalSignIns'
            ChkManagedIdentity    = 'ManagedIdentitySignIns'
            ChkCapNotApplied      = 'ConditionalAccessPolicyNotApplied'
            ChkCapSuccess         = 'ConditionalAccessPolicySuccessOnly'
            ChkCapFailed          = 'ConditionalAccessPolicyFailedOnly'
            ChkAnalyzeReportOnly  = 'AnalyzeCAPInReportOnly'
        }
        foreach ($key in $switchMap.Keys) {
            $ctrl = $window.FindName($key)
            if ($ctrl -and $ctrl.IsChecked) {
                $parts += "-$($switchMap[$key])"
            }
        }

        $capName = [string]$capCombo.SelectedItem
        if (-not $capName) { $capName = $capCombo.Text }
        if ($capName -and $capName -ne '(none)') {
            $parts += "-ConditionalAccessPolicyName '$capName'"
        }

        if ($forceNewTokenCheck.IsChecked) { $parts += '-ForceNewToken' }
        if ($exportExcelCheck.IsChecked) { $parts += '-ExportToExcel' }

        $commandBox.Text = $parts -join ' '
    }

    # Wire change events
    $timeRangeCombo.Add_SelectionChanged({ & $buildCommand })
    $timeZoneCombo.Add_SelectionChanged({ & $buildCommand })
    $startDatePicker.Add_SelectedDateChanged({ & $buildCommand })
    $endDatePicker.Add_SelectedDateChanged({ & $buildCommand })
    $startTimeBox.Add_TextChanged({ & $buildCommand })
    $endTimeBox.Add_TextChanged({ & $buildCommand })
    $usersBox.Add_TextChanged({ & $buildCommand })
    $topNBox.Add_TextChanged({ & $buildCommand })
    $ipBox.Add_TextChanged({ & $buildCommand })
    $forceNewTokenCheck.Add_Click({ & $buildCommand })
    $exportExcelCheck.Add_Click({ & $buildCommand })
    $capCombo.Add_SelectionChanged({ & $buildCommand })

    foreach ($checkName in @('ChkSuccessOnly', 'ChkFailureOnly', 'ChkBadCredentialsOnly', 'ChkLastLogonOnly',
            'ChkBasicAuthOnly', 'ChkMFAOnly', 'ChkNonMFAOnly', 'ChkNonInteractive',
            'ChkServicePrincipal', 'ChkManagedIdentity',
            'ChkCapNotApplied', 'ChkCapSuccess', 'ChkCapFailed', 'ChkAnalyzeReportOnly')) {
        $c = $window.FindName($checkName)
        if ($c) { $c.Add_Click({ & $buildCommand }) }
    }

    & $buildCommand

    # CAP refresh button
    $capRefreshButton.Add_Click({
            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            try {
                $cps = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop |
                    Select-Object DisplayName, Id, State |
                    Sort-Object DisplayName
                $capList.Clear()
                foreach ($c in $cps) { $capList.Add($c) }
                & $populateCap
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Could not load Conditional Access Policies: $($_.Exception.Message)",
                    'Graph error', 'OK', 'Warning') | Out-Null
            }
            finally {
                $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            }
        })

    # Copy button
    $copyButton.Add_Click({
            if (-not [string]::IsNullOrWhiteSpace($commandBox.Text)) {
                [System.Windows.Clipboard]::SetText($commandBox.Text)
            }
        })

    # Close button
    $closeButton.Add_Click({ $window.Close() })

    # Run button - builds the parameter splat and invokes the function
    $runButton.Add_Click({
            $runParams = @{}

            $tr = [string]$timeRangeCombo.SelectedItem
            if ($tr -and $tr -ne '(custom)') {
                $runParams['TimeRange'] = $tr
            }
            else {
                $selectedTzId = [string]$timeZoneCombo.SelectedItem
                $selectedTz = [System.TimeZoneInfo]::FindSystemTimeZoneById($selectedTzId)
                if ($startDatePicker.SelectedDate) {
                    $startLocal = $startDatePicker.SelectedDate.Value.Date
                    if ($startTimeBox.Text -match '^(\d{1,2}):(\d{2})$') {
                        $startLocal = $startLocal.AddHours([int]$matches[1]).AddMinutes([int]$matches[2])
                    }
                    $runParams['StartDate'] = [System.TimeZoneInfo]::ConvertTimeToUtc(
                        [datetime]::SpecifyKind($startLocal, [System.DateTimeKind]::Unspecified), $selectedTz)
                }
                if ($endDatePicker.SelectedDate) {
                    $endLocal = $endDatePicker.SelectedDate.Value.Date
                    if ($endTimeBox.Text -match '^(\d{1,2}):(\d{2})$') {
                        $endLocal = $endLocal.AddHours([int]$matches[1]).AddMinutes([int]$matches[2])
                    }
                    $runParams['EndDate'] = [System.TimeZoneInfo]::ConvertTimeToUtc(
                        [datetime]::SpecifyKind($endLocal, [System.DateTimeKind]::Unspecified), $selectedTz)
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($usersBox.Text)) {
                $userList = $usersBox.Text -split '[;\r\n]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                if ($userList) { $runParams['Users'] = $userList }
            }
            if ($topNBox.Text -match '^\d+$') { $runParams['LastXSignIns'] = [int]$topNBox.Text }
            if (-not [string]::IsNullOrWhiteSpace($ipBox.Text)) { $runParams['IPAddresses'] = [int]$ipBox.Text }

            $switchMap = [ordered]@{
                ChkSuccessOnly        = 'SuccessOnly'
                ChkFailureOnly        = 'FailureOnly'
                ChkBadCredentialsOnly = 'BadCredentialsOnly'
                ChkLastLogonOnly      = 'LastLogonOnly'
                ChkBasicAuthOnly      = 'BasicAuthenticationOnly'
                ChkMFAOnly            = 'MFASignInsOnly'
                ChkNonMFAOnly         = 'NonMFASignInsOnly'
                ChkNonInteractive     = 'NonInteractiveSignIns'
                ChkServicePrincipal   = 'ServicePrincipalSignIns'
                ChkManagedIdentity    = 'ManagedIdentitySignIns'
                ChkCapNotApplied      = 'ConditionalAccessPolicyNotApplied'
                ChkCapSuccess         = 'ConditionalAccessPolicySuccessOnly'
                ChkCapFailed          = 'ConditionalAccessPolicyFailedOnly'
                ChkAnalyzeReportOnly  = 'AnalyzeCAPInReportOnly'
            }
            foreach ($key in $switchMap.Keys) {
                $ctrl = $window.FindName($key)
                if ($ctrl -and $ctrl.IsChecked) { $runParams[$switchMap[$key]] = $true }
            }

            $capName = [string]$capCombo.SelectedItem
            if (-not $capName) { $capName = $capCombo.Text }
            if ($capName -and $capName -ne '(none)') {
                $runParams['ConditionalAccessPolicyName'] = $capName
            }

            if ($forceNewTokenCheck.IsChecked) { $runParams['ForceNewToken'] = $true }
            if ($exportExcelCheck.IsChecked) { $runParams['ExportToExcel'] = $true }

            $window.Cursor = [System.Windows.Input.Cursors]::Wait
            try {
                $results = Get-MgAuditLogSigninInfo @runParams
                $count = if ($results) { @($results).Count } else { 0 }
                [System.Windows.MessageBox]::Show(
                    "Query completed. $count sign-in record(s) returned.",
                    'Done', 'OK', 'Information') | Out-Null
                $script:GuiResults = $results
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Error: $($_.Exception.Message)",
                    'Run failed', 'OK', 'Error') | Out-Null
            }
            finally {
                $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            }
        })

    try {
        if ($splash) { $splash.Close(); $splash = $null }
        [void]$window.ShowDialog()
    }
    finally {
        if ($splash) { try { $splash.Close() } catch { } }
    }

    if ($script:GuiResults) { return $script:GuiResults }
}
