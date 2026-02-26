<#
    .SYNOPSIS
    Get Microsoft Entra ID Audit Log Sign-In Details

    .DESCRIPTION
    Get Microsoft Entra ID Audit Log Sign-In Details with various filtering options.

    .PARAMETER StartDate
    The start date for filtering sign-in logs. Accepts either a DateTime object or a string in yyyy-MM-dd format.

    .PARAMETER EndDate  
    The end date for filtering sign-in logs. Accepts either a DateTime object or a string in yyyy-MM-dd format.

    .PARAMETER Users
    An array of user principal names to filter the sign-in logs.

    .PARAMETER LastXSignIns
    The number of most recent sign-ins to retrieve.
    The other filters (StartDate, EndDate, Users, etc.) will still apply.

    .PARAMETER IPAddresses
    A comma-separated list of IP addresses to filter the sign-in logs.

    .PARAMETER BasicAuthenticationOnly
    Switch to filter sign-ins using legacy authentication protocols.

    .PARAMETER SingleFactorAuthenticationOnly
    Switch to filter only single-factor authentication sign-in attempts.

    .PARAMETER FailureOnly
    Switch to filter only failed sign-in attempts.

    .PARAMETER SuccessOnly
    Switch to filter only successful sign-in attempts.

    .PARAMETER BadCredentialsOnly
    Switch to filter sign-ins with bad username or password (error code 50126).

    .PARAMETER LastLogonOnly
    Switch to get only the last logon details for each user.

    .PARAMETER NonMFASignInsOnly
    Switch to filter non-MFA sign-ins only.

    .PARAMETER MFASignInsOnly
    Switch to filter MFA sign-ins only.

    .PARAMETER NonInteractiveSignIns
    Switch to filter non-interactive sign-ins only.

    .PARAMETER ServicePrincipalSignIns
    Switch to filter service principal sign-ins only.

    .PARAMETER ManagedIdentitySignIns
    Switch to filter managed identity sign-ins only.

    .PARAMETER ConditionalAccessPolicyName
    Filter sign-ins by a specific Conditional Access Policy Name.

    .PARAMETER ConditionalAccessPolicyNotApplied
    Switch to filter sign-ins where the Conditional Access Policy was not applied.

    .PARAMETER ConditionalAccessPolicySuccessOnly
    Switch to filter sign-ins where the Conditional Access Policy evaluation was successful.

    .PARAMETER ConditionalAccessPolicyFailedOnly
    Switch to filter sign-ins where the Conditional Access Policy evaluation failed.

    .PARAMETER AnalyzeCAPInReportOnly 
    Switch to filter sign-ins with Conditional Access applied in ReportOnly mode.
    Only sign-ins where the policy was used (exclude 'NotApplied') are returned.

    .PARAMETER ExportToExcel
    Switch to export the sign-in details report to an Excel file.
    The file will be saved in the user's profile directory with a timestamped filename.

    .PARAMETER UseDatePicker
    Switch to open an interactive calendar dialog to select StartDate and EndDate visually.
    Requires Windows with .NET Windows Forms support (Windows PowerShell or pwsh on Windows).
    Any values already provided via -StartDate or -EndDate are used as the initial selection in the pickers.

    .PARAMETER ForceNewToken
    Switch to force the acquisition of a new authentication token.

    .EXAMPLE
    Get-MgAuditLogSigninInfo -StartDate '2024-01-01' -EndDate '2024-01-31' -Users 'user1@contoso.com', 'user2@contoso.com'

    Retrieves sign-in logs for specified users between January 1, 2024, and January 31, 2024.

    .EXAMPLE
    Get-MgAuditLogSigninInfo -LastXSignIns 100 -FailureOnly

    Retrieves the last 100 failed sign-in attempts.

    .EXAMPLE
    Get-MgAuditLogSigninInfo -AnalyzeCAPInReportOnly

    Retrieves sign-in logs with Conditional Access applied in ReportOnly mode.

    .EXAMPLE
    Get-MgAuditLogSigninInfo -StartDate (Get-Date).AddHours(-1) -NonMFASignInsOnly

    Retrieves non-MFA sign-ins from the last hour.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgAuditLogSigninInfo

    .NOTES
#>

function Get-MgAuditLogSigninInfo {
    param(
        [Parameter(Mandatory = $false)]    
        $StartDate,

        [Parameter(Mandatory = $false)]
        $EndDate,

        [Parameter(Mandatory = $false)]
        [string[]]$Users,

        [Parameter(Mandatory = $false)]
        [int]$LastXSignIns,

        [Parameter(Mandatory = $false)]
        [int]$IPAddresses,

        [Parameter(Mandatory = $false)]
        [switch]$BasicAuthenticationOnly,

        [Parameter(Mandatory = $false)]
        [switch]$SuccessOnly,

        [Parameter(Mandatory = $false)]
        [switch]$FailureOnly,

        [Parameter(Mandatory = $false)]
        [switch]$BadCredentialsOnly,

        [Parameter(Mandatory = $false)]
        [switch]$LastLogonOnly,

        [Parameter(Mandatory = $false)]
        [switch]$NonMFASignInsOnly,

        [Parameter(Mandatory = $false)]
        [switch]$MFASignInsOnly,

        [Parameter(Mandatory = $false)]
        [switch]$NonInteractiveSignIns,

        [Parameter(Mandatory = $false)]
        [switch]$ServicePrincipalSignIns,

        [Parameter(Mandatory = $false)]
        [switch]$ManagedIdentitySignIns,

        [Parameter(Mandatory = $false)]
        [string]$ConditionalAccessPolicyName,

        [Parameter(Mandatory = $false)]
        [switch]$ConditionalAccessPolicyNotApplied,

        [Parameter(Mandatory = $false)]
        [switch]$ConditionalAccessPolicySuccessOnly,

        [Parameter(Mandatory = $false)]
        [switch]$ConditionalAccessPolicyFailedOnly,   
        
        # Remplace plusieurs switches par un seul paramètre avec ValidateSet
        [Parameter(Mandatory = $false)]
        [ValidateSet(
            'Last2Minutes',
            'Last10Minutes',
            'LastHour',
            'Last6Hours',
            'Last12Hours',
            'Last24Hours',
            'Last3Days',
            'Last7Days',
            'Last15Days',
            'Maximum'
        )]
        [string]$TimeRange,

        [Parameter(Mandatory = $false)]
        [switch]$UseDatePicker,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$AnalyzeCAPInReportOnly,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )
    
    <# Excluded because can be used in PowerShell 5.1 or 7.x
    $modules = @(
        #'Microsoft.Graph.Reports',
        'Microsoft.Graph.Authentication'
    )

    foreach ($module in $modules) {
        
        try {
            $null = Get-InstalledModule $module -ErrorAction Stop
        }
        catch {
            Write-Warning "Please install $module first"
            return
        }
    }
#>
    Write-Verbose 'Connect MgGraph with AuditLog.Read.All scope'

    # Open a Windows Forms date+time-range picker when -UseDatePicker is requested
    if ($UseDatePicker.IsPresent) {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        # Resolve initial values (date + time)
        $initStart = if ($StartDate -is [DateTime]) { $StartDate } elseif ($StartDate) { [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', $null) } else { (Get-Date).Date.AddDays(-7) }
        $initEnd = if ($EndDate -is [DateTime]) { $EndDate } elseif ($EndDate) { [datetime]::ParseExact($EndDate, 'yyyy-MM-dd', $null).AddDays(1).AddSeconds(-1) } else { (Get-Date) }

        $colorAccent = [System.Drawing.Color]::FromArgb(222, 118, 29)   # PS365 orange #de761d
        $colorBg = [System.Drawing.Color]::FromArgb(245, 245, 248) # near-white
        $colorPanelBg = [System.Drawing.Color]::FromArgb(245, 245, 248) # same as form bg – suppresses white card
        $colorText = [System.Drawing.Color]::FromArgb(32, 32, 32)
        $colorSubText = [System.Drawing.Color]::FromArgb(96, 96, 96)
        $fontUI = [System.Drawing.Font]::new('Segoe UI', 9)
        $fontUIBold = [System.Drawing.Font]::new('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        $fontTitle = [System.Drawing.Font]::new('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)

        # --- Form ---
        $form = [System.Windows.Forms.Form]@{
            Text            = 'Sign-in log — Date & time range'
            Size            = [System.Drawing.Size]::new(440, 352)
            StartPosition   = 'CenterScreen'
            FormBorderStyle = 'FixedDialog'
            MaximizeBox     = $false
            MinimizeBox     = $false
            BackColor       = $colorBg
            Font            = $fontUI
        }

        # --- Header panel ---
        $pnlHeader = [System.Windows.Forms.Panel]@{
            Dock      = [System.Windows.Forms.DockStyle]::Top
            Height    = 52
            BackColor = $colorAccent
        }
        $lblTitle = [System.Windows.Forms.Label]@{
            Text        = 'Select date & time range'
            ForeColor   = [System.Drawing.Color]::White
            Font        = $fontTitle
            AutoSize    = $true
            Location    = [System.Drawing.Point]::new(16, 14)
            UseMnemonic = $false   # prevents '&' from being swallowed as a mnemonic accelerator
        }
        $pnlHeader.Controls.Add($lblTitle)

        # --- Content panel (white card) ---
        $pnlContent = [System.Windows.Forms.Panel]@{
            Location  = [System.Drawing.Point]::new(16, 68)
            Size      = [System.Drawing.Size]::new(400, 160)
            BackColor = $colorPanelBg
        }
        # No border needed – content panel blends with form background

        # Helper to build a label+date+time row inside a panel
        # Returns a Panel containing the row
        function New-DateTimeRow {
            param([string]$Caption, [datetime]$InitialValue, [int]$Top)
            $row = [System.Windows.Forms.Panel]@{
                Location  = [System.Drawing.Point]::new(0, $Top)
                Size      = [System.Drawing.Size]::new(400, 58)
                BackColor = $colorPanelBg
            }
            $lbl = [System.Windows.Forms.Label]@{
                Text      = $Caption
                Font      = $fontUIBold
                ForeColor = $colorText
                Location  = [System.Drawing.Point]::new(16, 8)
                AutoSize  = $true
            }
            $lblDateHint = [System.Windows.Forms.Label]@{
                Text      = 'Date'
                Font      = [System.Drawing.Font]::new('Segoe UI', 7.5)
                ForeColor = $colorSubText
                Location  = [System.Drawing.Point]::new(100, 6)
                AutoSize  = $true
            }
            $dtp = [System.Windows.Forms.DateTimePicker]@{
                Format   = [System.Windows.Forms.DateTimePickerFormat]::Short
                Location = [System.Drawing.Point]::new(100, 22)
                Size     = [System.Drawing.Size]::new(140, 24)
                Value    = $InitialValue
                Font     = $fontUI
            }
            $lblTimeHint = [System.Windows.Forms.Label]@{
                Text      = 'Time'
                Font      = [System.Drawing.Font]::new('Segoe UI', 7.5)
                ForeColor = $colorSubText
                Location  = [System.Drawing.Point]::new(256, 6)
                AutoSize  = $true
            }
            $dtpTime = [System.Windows.Forms.DateTimePicker]@{
                Format     = [System.Windows.Forms.DateTimePickerFormat]::Time
                ShowUpDown = $true
                Location   = [System.Drawing.Point]::new(256, 22)
                Size       = [System.Drawing.Size]::new(110, 24)
                Value      = $InitialValue
                Font       = $fontUI
            }
            $row.Controls.AddRange(@($lbl, $lblDateHint, $dtp, $lblTimeHint, $dtpTime))
            # Draw an explicit 1-px border around each DTP so all 4 sides are visible regardless
            # of background colour (the Windows-themed DTP border disappears on a same-colour bg)
            $row | Add-Member -NotePropertyName '_dtp' -NotePropertyValue $dtp
            $row | Add-Member -NotePropertyName '_dtpTime' -NotePropertyValue $dtpTime
            $row.Add_Paint({
                    param($s, $e)
                    $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(130, 135, 153))
                    foreach ($ctrl in @($s._dtp, $s._dtpTime)) {
                        $e.Graphics.DrawRectangle($pen, $ctrl.Left - 1, $ctrl.Top - 1, $ctrl.Width + 1, $ctrl.Height + 1)
                    }
                    $pen.Dispose()
                })
            # Return the row and both pickers as a hashtable
            return @{ Panel = $row; DatePicker = $dtp; TimePicker = $dtpTime }
        }

        $startRow = New-DateTimeRow -Caption 'From' -InitialValue $initStart -Top 10
        $endRow = New-DateTimeRow -Caption 'To' -InitialValue $initEnd -Top 88

        # Thin separator between rows
        $sep = [System.Windows.Forms.Panel]@{
            Location  = [System.Drawing.Point]::new(16, 74)
            Size      = [System.Drawing.Size]::new(368, 1)
            BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
        }

        $pnlContent.Controls.AddRange(@($startRow.Panel, $sep, $endRow.Panel))

        # --- Time zone selector ---
        $lblTZHint = [System.Windows.Forms.Label]@{
            Text      = 'Time zone'
            Font      = [System.Drawing.Font]::new('Segoe UI', 7.5)
            ForeColor = $colorSubText
            Location  = [System.Drawing.Point]::new(16, 236)
            AutoSize  = $true
        }
        $cmbTZ = [System.Windows.Forms.ComboBox]@{
            DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
            Location      = [System.Drawing.Point]::new(16, 252)
            Size          = [System.Drawing.Size]::new(400, 24)
            Font          = $fontUI
        }
        $cmbTZ.DisplayMember = 'DisplayName'
        foreach ($tz in [System.TimeZoneInfo]::GetSystemTimeZones()) { $null = $cmbTZ.Items.Add($tz) }
        $localId = [System.TimeZoneInfo]::Local.Id
        for ($i = 0; $i -lt $cmbTZ.Items.Count; $i++) {
            if ($cmbTZ.Items[$i].Id -eq $localId) { $cmbTZ.SelectedIndex = $i; break }
        }

        # --- Buttons ---
        $btnOK = [System.Windows.Forms.Button]@{
            Text         = 'Apply'
            Location     = [System.Drawing.Point]::new(222, 286)
            Size         = [System.Drawing.Size]::new(90, 32)
            FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
            BackColor    = $colorAccent
            ForeColor    = [System.Drawing.Color]::White
            Font         = $fontUIBold
            DialogResult = [System.Windows.Forms.DialogResult]::OK
            Cursor       = [System.Windows.Forms.Cursors]::Hand
        }
        $btnOK.FlatAppearance.BorderSize = 0

        $btnCancel = [System.Windows.Forms.Button]@{
            Text         = 'Cancel'
            Location     = [System.Drawing.Point]::new(326, 286)
            Size         = [System.Drawing.Size]::new(90, 32)
            FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
            BackColor    = [System.Drawing.Color]::FromArgb(225, 225, 225)
            ForeColor    = $colorText
            Font         = $fontUI
            DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            Cursor       = [System.Windows.Forms.Cursors]::Hand
        }
        $btnCancel.FlatAppearance.BorderSize = 0

        $lblPS365 = [System.Windows.Forms.Label]@{
            Text      = 'PS365 - ps365.clidsys.com'
            ForeColor = $colorSubText
            Font      = [System.Drawing.Font]::new('Segoe UI', 7.5)
            AutoSize  = $true
            Location  = [System.Drawing.Point]::new(16, 294)
        }

        $form.Controls.AddRange(@($pnlHeader, $pnlContent, $lblTZHint, $cmbTZ, $btnOK, $btnCancel, $lblPS365))
        $form.AcceptButton = $btnOK
        $form.CancelButton = $btnCancel

        if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Warning 'Date picker cancelled. No data retrieved.'
            return
        }

        # Merge date + time then convert from the selected timezone to UTC
        $selectedTz = $cmbTZ.SelectedItem
        $startLocal = $startRow.DatePicker.Value.Date.Add($startRow.TimePicker.Value.TimeOfDay)
        $endLocal = $endRow.DatePicker.Value.Date.Add($endRow.TimePicker.Value.TimeOfDay)
        $StartDate = [System.TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($startLocal, [System.DateTimeKind]::Unspecified), $selectedTz)
        $EndDate = [System.TimeZoneInfo]::ConvertTimeToUtc([datetime]::SpecifyKind($endLocal, [System.DateTimeKind]::Unspecified), $selectedTz)
    }

    if ($ForceNewToken) {
        Disconnect-MgGraph
        $null = Connect-MgGraph -Scopes AuditLog.Read.All
    }
    else {
        $null = Connect-MgGraph -Scopes AuditLog.Read.All
    }

    try {
        $null = Get-MgBetaAuditLogSignIn -Top 1 -ErrorAction stop
    }
    catch {
        if ($_.Exception.ErrorContent.Code) {
            Write-Warning "Unable to get MgBetaAuditLogSignIn: $($_.Exception.Message)"
            return
        }
    }

    [System.Collections.Generic.List[PSObject]]$signsInList = @()

    # Build StartDate/EndDate from TimeRange if provided (use full datetime when relevant)
    $endWasDayOnly = $false
    if ($TimeRange) {
        switch ($TimeRange) {
            'Last2Minutes' {
                $startDt = (Get-Date).AddMinutes(-2)
                break
            }
            'Last10Minutes' {
                $startDt = (Get-Date).AddMinutes(-10)
                break
            }
            'LastHour' {
                $startDt = (Get-Date).AddHours(-1)
                break
            }
            'Last6Hours' {
                $startDt = (Get-Date).AddHours(-6)
                break
            }
            'Last12Hours' {
                $startDt = (Get-Date).AddHours(-12)
                break
            }
            'Last24Hours' {
                $startDt = (Get-Date).AddHours(-24)
                break
            }
            'Last3Days' {
                $startDt = (Get-Date).AddDays(-3)
                break
            }
            'Last7Days' {
                $startDt = (Get-Date).AddDays(-7)
                break
            }
            'Last15Days' {
                $startDt = (Get-Date).AddDays(-14)
                break
            }
            'Maximum' {
                Write-Host -ForegroundColor Cyan "You have selected the 'Maximum' TimeRange option. The value corresponds to the maximum retention period of Microsoft Entra ID (Azure AD) sign-in logs, which is 30 days for Entra P1/P2 licence and 7 days otherwise."
                $startDt = (Get-Date).AddDays(-30)
                break
            }
        }

        # When using predefined TimeRange, default EndDate to now (preserve time resolution)
        $endDt = (Get-Date)
        $endWasDayOnly = $false
    }

    # Normalize and validate StartDate and EndDate once, then build a single UTC ISO date filter
    try {
        if (-not $startDt) {
            if ($StartDate) {
                # Handle both DateTime objects and string dates
                if ($StartDate -is [DateTime]) {
                    $parsedStart = $StartDate
                }
                else {
                    # Try to parse as string in yyyy-MM-dd format
                    $parsedStart = [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', $null)
                }
                
                # if startDate greater than 30 days ago, warn the user and set it to 30 days ago
                if ($parsedStart -lt (Get-Date).AddDays(-30)) {
                    Write-Warning 'StartDate is greater than 30 days ago. Microsoft Entra ID (Azure AD) sign-in logs are retained for 30 days only. Setting StartDate to 30 days ago.'
                    $startDt = (Get-Date).AddDays(-30)
                }
                else {
                    # use the provided datetime (preserving time if it's a DateTime object)
                    $startDt = $parsedStart
                }
            }
            else {
                # Default: 30 days ago (keeps previous behaviour)
                # https://learn.microsoft.com/en-us/entra/identity/monitoring-health/reference-reports-data-retention#how-long-does-microsoft-entra-id-store-the-data
                $startDt = (Get-Date).AddDays(-30)
            }
        }
        else {
            # If TimeRange produced a startDt older than retention, clamp it and warn
            if ($startDt -lt (Get-Date).AddDays(-30)) {
                Write-Warning 'Computed TimeRange start is greater than 30 days ago. Microsoft Entra ID sign-in logs are retained for 30 days only. Setting StartDate to 30 days ago.'
                $startDt = (Get-Date).AddDays(-30)
            }
        }
    }
    catch {
        Write-Warning "Unable to get date from StartDate. Please provide either a DateTime object or a string in yyyy-MM-dd format. $($_.Exception.Message)"
        return
    }

    try {
        if (-not $endDt) {
            if ($EndDate) {
                # Handle both DateTime objects and string dates
                if ($EndDate -is [DateTime]) {
                    $endDt = $EndDate
                    # DateTime objects preserve time, so don't treat as day-only
                    $endWasDayOnly = $false
                }
                else {
                    # Try to parse as string in yyyy-MM-dd format
                    $endDt = [datetime]::ParseExact($EndDate, 'yyyy-MM-dd', $null)
                    # mark that user provided day-only EndDate so we can make it exclusive later
                    $endWasDayOnly = $true
                }
            }
            else {
                # Default endDate: now (we will make it exclusive only if user provided day-only EndDate)
                $endDt = (Get-Date)
                $endWasDayOnly = $false
            }
        }
    }
    catch {
        Write-Warning "Unable to get date from EndDate. Please provide either a DateTime object or a string in yyyy-MM-dd format. $($_.Exception.Message)"
        return
    }

    # Make the end bound exclusive and include the full provided EndDate day when end was day-only
    if ($endWasDayOnly) {
        $endExclusive = $endDt.AddDays(1)
    }
    else {
        # endDt already includes a time component (e.g. TimeRange or now)
        $endExclusive = $endDt
    }

    # Convert to UTC ISO format accepted by Microsoft Graph filters
    $startUtc = $startDt.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $endUtc = $endExclusive.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

    $dateFilter = "createdDateTime gt $startUtc and createdDateTime lt $endUtc"

    $filter = $dateFilter
    
    if ($BadCredentialsOnly) {
        Write-Verbose 'Get All signs-in with bad username or password'

        $filter += ' and status/errorCode eq 50126'
    }

    if ($Users) {
        $userFilter = $null
        foreach ($user in $Users) {
            Write-Verbose "User: $user"

            if ($userFilter) {
                $userFilter += " or userPrincipalName eq '$user'"
            }
            else {
                $userFilter = " and (userPrincipalName eq '$user'"
            }
        }

        $userFilter += ')'
        $filter += "$userFilter"
    }    

    if ($LastLogonOnly) {
        $mgUsers = Get-MgUser -All -Property SignInActivity

        foreach ($mgUser in $mgUsers) {
            $mgUser.UserPrincipalName

            $object = [PSCustomObject][ordered]@{
                UserPrincipalName                = $mgUser.UserPrincipalName
                ID                               = $mgUser.Id
                LastNonInteractiveSignInDateTime = $mgUser.SignInActivity.LastNonInteractiveSignInDateTime
                LastSignInDateTime               = $mgUser.SignInActivity.LastSignInDateTime
                LastSuccessfulSignInDateTime     = $mgUser.SignInActivity.LastSuccessfulSignInDateTime
                AccountEnabled                   = $mgUser.AccountEnabled
                CreatedDateTime                  = $mgUser.CreatedDateTime
                CreationType                     = $mgUser.CreationType
                UserType                         = $mgUser.UserType
            }
            
            # Users not connected since > 120 days : $signsInList | Where-Object {$_.LastSignInDateTime -lt (Get-Date).AddDays(-120) -or ($_.LastSignInDateTime -eq $null)} | Select-Object UserPrincipalName, LastSignInDateTime
            $signsInList.Add($object)
        }
    }

    if ($FailureOnly) {
        Write-Verbose 'Get signs-in with status failure'
        
        # ignore null because Teams return null
        # ignore 50140 because it means "This occurred due to 'Keep me signed in' interrupt when the user was signing in"
        
        $filter += ' and status/errorCode ne 0 and status/errorCode ne 50140'
    }

    if ($SuccessOnly) {
        Write-Verbose 'Get signs-in with status success'
        
        $filter += ' and status/errorCode eq 0'
    }

    if ($BasicAuthenticationOnly) {
        Write-Verbose 'Get signs-in with legacy protocols'
        
        # ignore nul because Teams return null (?)
        $filter += " and clientAppUsed ne 'Mobile Apps and Desktop clients' and clientAppUsed ne 'Browser'"
    }

    if ($IPAddresses) {
        $ipFilter = $null

        foreach ($IPAddress in $IPAddresses -split ',') {
            Write-Verbose "IPAddress: $IPAddress"

            if ($ipFilter) {
                $ipFilter = " or ipaddress eq '$IPAddress'"
            }
            else {
                $ipFilter += " ipaddress eq '$IPAddress'"
            }
        }

        $filter += "$ipFilter"
    }

    if ($NonInteractiveSignIns.IsPresent) {
        Write-Verbose 'Filter non-interactive sign-ins only'
        $filter += " and (signInEventTypes/any(t: t eq 'noninteractiveUser'))"
    }

    if ($ServicePrincipalSignIns.IsPresent) {
        Write-Verbose 'Filter service principal sign-ins only'
        $filter += " and (signInEventTypes/any(t: t eq 'servicePrincipal'))"
    }

    if ($ManagedIdentitySignIns.IsPresent) {
        Write-Verbose 'Filter managed identity sign-ins only'
        $filter += " and (signInEventTypes/any(t: t eq 'managedIdentity'))"
    }

    if ($NonMFASignInsOnly.IsPresent) {
        Write-Verbose 'Filter non-MFA sign-ins only'
        $filter += " and (authenticationRequirement eq 'singleFactorAuthentication')"
    }

    if ($MFASignInsOnly.IsPresent) {
        Write-Verbose 'Filter MFA sign-ins only'
        $filter += " and (authenticationRequirement eq 'multiFactorAuthentication')"
    }

    if ($LastXSignIns) {
        Write-Verbose "Get-MgBetaAuditLogSignIn -Top: $LastXSignIns -Filter $filter"
        $signsIn = Get-MgBetaAuditLogSignIn -Top $LastXSignIns -Filter $filter
    }
    else {
        Write-Verbose "Get-MgBetaAuditLogSignIn -All:`$true -Filter $filter"

        $signsIn = Get-MgBetaAuditLogSignIn -All:$true -Filter $filter
    }
    
    Write-Verbose "Filter is $filter"

    if ($ConditionalAccessPolicyName) {
        Write-Verbose "Filter signs-in with Conditional Access Policy Name: $ConditionalAccessPolicyName"
        $signsIn = $signsIn | Where-Object {
            $_.Result.DisplayName -contains $ConditionalAccessPolicyName
        }
    }

    if ($ConditionalAccessPolicyNotApplied.IsPresent) {
        Write-Verbose 'Filter signs-in with Conditional Access Policy Not Applied'
        $signsIn = $signsIn | Where-Object { $_.ConditionalAccessStatus -eq 'notApplied' }
    }

    if ( $ConditionalAccessPolicySuccessOnly.IsPresent) {
        Write-Verbose 'Filter signs-in with Conditional Access Policy Success Only'
        $signsIn = $signsIn | Where-Object { $_.ConditionalAccessStatus -eq 'success' }
    }

    if ( $ConditionalAccessPolicyFailedOnly.IsPresent) {
        Write-Verbose 'Filter signs-in with Conditional Access Policy Failed Only'
        $signsIn = $signsIn | Where-Object { $_.ConditionalAccessStatus -eq 'failure' }
    }
    
    if ($AnalyzeCAPInReportOnly ) {
        Write-Verbose "Filter signs-in with Conditional Access applied in ReportOnly mode and used (exclude 'NotApplied')"
        
        foreach ($signIn in $signsIn) {

            # https://groovynerd.co.uk/how-to-gather-reportonly-conditional-access-sign-in-logs/
            $reportOnlyPolicies = $SignIn.AppliedConditionalAccessPolicies

            # Loop through each policy
            foreach ($policy in $reportOnlyPolicies) {
                # Check if the policy result is in ReportOnly mode (and not with 'Not Applied')
                # values can be any of the following: 'success', 'failure', 'notApplied', 'reportOnlySuccess', 'reportOnlyFailure', 'reportOnlyNotApplied' and 'notEnabled'
                # reportOnlySuccess: All configured policy conditions, required non-interactive grant controls, and session controls were satisfied. For example, a multifactor authentication requirement is satisfied by an MFA claim already present in the token, or a compliant device policy is satisfied by performing a device check on a compliant device.
                # reportOnlyFailure: All configured policy conditions were satisfied but not all the required non-interactive grant controls or session controls were satisfied. For example, a policy applies to a user where a block control is configured, or a device fails a compliant device policy.
                # reportOnlyInterrupted  = Report-only User action required : All configured policy conditions were satisfied but user action would be required to satisfy the required grant controls or session controls. With report-only mode, the user isn't prompted to satisfy the required controls. For example, users aren't prompted for multifactor authentication challenges or terms of use.
                # Report-only: Not applied : Not all configured policy conditions were satisfied. For example, the user is excluded from the policy or the policy only applies to certain trusted named locations.

                # source: https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-report-only#policy-evaluation-results
                
                if (($policy.Result -like 'reportOnly*') -and ($policy.Result -ne 'reportOnlyNotApplied')) {

                    Write-Verbose "Policy: $($policy.DisplayName) - Result: $($policy.Result)"
                    if ($null -eq $signIn.DeviceDetail.OperatingSystem) {
                        $os = 'Operating System not in logs'
                    }
                    else {
                        $os = $signIn.DeviceDetail.OperatingSystem
                    }

                    if ($null -eq $signIn.DeviceDetail.Browser) {
                        $browser = 'Browser not in logs'
                    }
                    else {
                        $browser = $signIn.DeviceDetail.Browser
                    }

                    if ($null -eq $signIn.Location.CountryOrRegion) {
                        $location = 'Unknown'
                    }
                    else {
                        $location = $signIn.Location.CountryOrRegion + '|' + $signIn.Location.State + '|' + $signIn.Location.City
                    }

                    # Create a structured report object for each matching log entry
                    $object = [PSCustomObject][ordered]@{
                        Time                             = $signIn.CreatedDateTime
                        UserDisplayName                  = $signIn.UserDisplayName
                        UserPrincipalName                = $signIn.UserPrincipalName
                        AppDisplayName                   = $signIn.AppDisplayName
                        IpAddress                        = $signIn.IpAddress
                        ClientAppUsed                    = $signIn.ClientAppUsed
                        ConditionalAccessStatus          = '-'
                        AppliedConditionalAccessPolicies = $policy.DisplayName
                        ErrorCode                        = '-'
                        AdditionalDetails                = '-'
                        FailureReason                    = '-'
                        DeviceName                       = $signIn.DeviceDetail.DisplayName
                        DeviceDetail                     = $os + '|' + $browser
                        DeviceIsManaged                  = $signIn.DeviceDetail.IsManaged
                        DeviceIsCompliant                = $signIn.DeviceDetail.IsCompliant
                        Location                         = $location
                        IsInteractive                    = $signIn.IsInteractive
                        PolicyResult                     = $policy.Result
                    }

                    # Add the report to the collection
                    $signsInList.Add($object)
                }
            }
        }
    }
    else {
        foreach ($signIn in $signsIn) {
            
            if ($null -eq $signIn.DeviceDetail.OperatingSystem) {
                $os = 'Operating System not in logs'
            }
            else {
                $os = $signIn.DeviceDetail.OperatingSystem
            }

            if ($null -eq $signIn.DeviceDetail.Browser) {
                $browser = 'Browser not in logs'
            }
            else {
                $browser = $signIn.DeviceDetail.Browser
            }

            if ($null -eq $signIn.Location.CountryOrRegion) {
                $location = 'Unknown'
            }
            else {
                $location = $signIn.Location.CountryOrRegion + '|' + $signIn.Location.State + '|' + $signIn.Location.City
            }

            $object = [PSCustomObject][ordered]@{
                DateTime                         = $signIn.CreatedDateTime
                UserDisplayName                  = $signIn.UserDisplayName
                UserPrincipalName                = $signIn.UserPrincipalName
                AppDisplayName                   = $signIn.AppDisplayName
                IpAddress                        = $signIn.IpAddress
                ClientAppUsed                    = $signIn.ClientAppUsed
                ConditionalAccessStatus          = $signIn.ConditionalAccessStatus
                AppliedConditionalAccessPolicies = ($signIn.AppliedConditionalAccessPolicies | Where-Object { $_.Result -ne 'NotApplied' } | ForEach-Object { "DisplayName=$($_.DisplayName);EnforcedGrantControls=$($_.EnforcedGrantControls -join '|');EnforcedSessionControls=$($_.EnforcedSessionControls -join '|');Result=$($_.Result)" }) -join '|'
                ErrorCode                        = $signIn.Status.ErrorCode
                AdditionalDetails                = if ($null -ne $signIn.Status.AdditionalDetails -ne '') { $signIn.Status.AdditionalDetails } else { 'no additional data' }
                FailureReason                    = $signIn.Status.FailureReason
                DeviceName                       = $signIn.DeviceDetail.DisplayName
                DeviceDetail                     = $os + '|' + $browser
                DeviceIsManaged                  = $signIn.DeviceDetail.IsManaged
                DeviceIsCompliant                = $signIn.DeviceDetail.IsCompliant
                Location                         = $location
                IsInteractive                    = $signIn.IsInteractive
                AuthenticationRequirement        = $signIn.AuthenticationRequirement
                MfaDetail                        = if ($signIn.MfaDetail.AuthMethod) { "$($signIn.MfaDetail.AuthMethod)|$($signIn.MfaDetail.AuthDetail)" } else { $null }
                AuthenticationDetails            = ($signIn.AuthenticationDetails | ForEach-Object {
                        "Method=$($_.AuthenticationMethod);Step=$($_.AuthenticationStepRequirement);Succeeded=$($_.Succeeded);Result=$($_.AuthenticationStepResultDetail)"
                    }) -join ' | '
            }

            $signsInList.Add($object)
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $ExcelFilePath = "$($env:userprofile)\$now-MgAuditLogSignInDetail_Report.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting sign-in details to Excel file: $ExcelFilePath"
        $signsInList | Export-Excel -Path $ExcelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-SignInLog'
    }
    else {
        return $signsInList
    }
}

#to do - organize properties better
<#
    $mgUsers = Get-MgUser -ConsistencyLevel eventual -All -Property @(
        'UserPrincipalName',
        'AccountEnabled',
        'UserType'
        'SignInActivity'
        'CreatedDateTime'        
        'DisplayName'
        'Mail'
        'OnPremisesImmutableId'
        'OnPremisesDistinguishedName'
        'OnPremisesLastSyncDateTime'
        'SignInSessionsValidFromDateTime'
        'RefreshTokensValidFromDateTime'
        'id',
        'ProxyAddresses',
        'OtherMails',
        'CreationType',
        'ExternalUserState',
        'ExternalUserStateChangeDateTime'
    ) | Select-Object @(
        'UserPrincipalName',
        'AccountEnabled',
        'UserType',
        @{Name = 'CreatedDateTime'; Expression = { ([datetime]$_.CreatedDateTime).ToString('yyyy-MM-dd HH:mm:ss') } }
        'DisplayName'
        'Mail'
        @{Name = 'ProxyAddresses'; Expression = { $_.ProxyAddresses -join '|' } }
        'OnPremisesImmutableId'
        'OnPremisesDistinguishedName'
        @{Name = 'OnPremisesLastSyncDateTime'; Expression = { ([datetime]$_.OnPremisesLastSyncDateTime).ToString('yyyy-MM-dd HH:mm:ss') } }
        @{Name = 'SignInSessionsValidFromDateTime'; Expression = { ([datetime]$_.SignInSessionsValidFromDateTime).ToString('yyyy-MM-dd HH:mm:ss') } }
        @{Name = 'RefreshTokensValidFromDateTime'; Expression = { ([datetime]$_.RefreshTokensValidFromDateTime).ToString('yyyy-MM-dd HH:mm:ss') } }
        'id'        
        @{Name = 'LastSignInDateTime'; Expression = { ([datetime]$_.SignInActivity.LastSignInDateTime).ToString('yyyy-MM-dd HH:mm:ss') } }
        @{Name = 'lastNonInteractiveSignInDateTime'; Expression = { ([datetime]$_.SignInActivity.AdditionalProperties.lastNonInteractiveSignInDateTime).ToString('yyyy-MM-dd HH:mm:ss') } }           
        @{Name = 'OtherMails'; Expression = { $_.OtherMails -join '|' } }
        'ExternalUserState',
        @{Name = 'ExternalUserStateChangeDateTime'; Expression = { ([datetime]$_.ExternalUserStateChangeDateTime).ToString('yyyy-MM-dd HH:mm:ss') } }
    )

    $mgUsers | Sort-Object UserPrincipalName 
    #>