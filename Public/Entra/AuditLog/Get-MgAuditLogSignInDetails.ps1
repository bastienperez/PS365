<#.SYNOPSIS
    Get Microsoft Entra ID (Azure AD) Audit Log Sign-In Details

.DESCRIPTION
    Get Microsoft Entra ID (Azure AD) Audit Log Sign-In Details with various filtering options.

    .PARAMETER StartDate
        The start date for filtering sign-in logs (format: yyyy-MM-dd).

    .PARAMETER EndDate  
        The end date for filtering sign-in logs (format: yyyy-MM-dd).

    .PARAMETER Users
        An array of user principal names to filter the sign-in logs.

    .PARAMETER LastXSignIns
        The number of most recent sign-ins to retrieve.
        The other filters (StartDate, EndDate, Users, etc.) will still apply.

    .PARAMETER IPAddresses
        A comma-separated list of IP addresses to filter the sign-in logs.

    .PARAMETER BasicAuthenticationOnly
        Switch to filter sign-ins using legacy authentication protocols.

    .PARAMETER FailuresOnly
        Switch to filter only failed sign-in attempts.

    .PARAMETER BadCredentialsOnly
        Switch to filter sign-ins with bad username or password (error code 50126).

    .PARAMETER LastLogonOnly
        Switch to get only the last logon details for each user.

    .PARAMETER ConditionalAccessPolicyName
        Filter sign-ins by a specific Conditional Access Policy Name.

    .PARAMETER AnalyzeReportOnlyCA 
        Switch to filter sign-ins with Conditional Access applied in ReportOnly mode.
        Only sign-ins where the policy was used (exclude 'NotApplied') are returned.

    .PARAMETER OutputFile
        The path to the output file where the results will be saved.

    .PARAMETER ForceNewToken
        Switch to force the acquisition of a new authentication token.

    .EXAMPLE
        Get-MgAuditLogSignInDetails -StartDate '2024-01-01' -EndDate '2024-01-31' -Users 'user1@contoso.com', 'user2@contoso.com'

        Retrieves sign-in logs for specified users between January 1, 2024, and January 31, 2024.

    .EXAMPLE
    Get-MgAuditLogSignInDetails -LastXSignIns 100 -FailuresOnly

    Retrieves the last 100 failed sign-in attempts.

    .EXAMPLE
    Get-MgAuditLogSignInDetails -AnalyzeReportOnlyCA

    Retrieves sign-in logs with Conditional Access applied in ReportOnly mode.

    .NOTES

#>

function Get-MgAuditLogSignInDetails {
    param(
        [Parameter(Mandatory = $false)]    
        [String]$StartDate,

        [Parameter(Mandatory = $false)]
        [String]$EndDate,

        [Parameter(Mandatory = $false)]
        [string[]]$Users,

        [Parameter(Mandatory = $false)]
        [int]$LastXSignIns,

        [Parameter(Mandatory = $false)]
        [int]$IPAddresses,

        [Parameter(Mandatory = $false)]
        [switch]$BasicAuthenticationOnly,

        [Parameter(Mandatory = $false)]
        [switch]$FailuresOnly,

        [Parameter(Mandatory = $false)]
        [switch]$BadCredentialsOnly,

        [Parameter(Mandatory = $false)]
        [switch]$LastLogonOnly,

        [Parameter(Mandatory = $false)]
        [string]$ConditionalAccessPolicyName,

        [Parameter(Mandatory = $false)]
        [switch]$AnalyzeReportOnlyCA,

        [Parameter(Mandatory = $false)]
        [string]$OutputFile,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken
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

    if ($ForceNewToken) {
        Disconnect-MgGraph
        $null = Connect-MgGraph -Scopes AuditLog.Read.All
    }
    else {
        $null = Connect-MgGraph -Scopes AuditLog.Read.All
    }

    try {
        $null = Get-MgAuditLogSignIn -Top 1 -ErrorAction stop
    }
    catch {
        if ($_.Exception.ErrorContent.Code) {
            Write-Warning "Unable to geet MgAuditLogSignIn: $($_.Exception.Message)"
            return
        }
    }

    [System.Collections.Generic.List[PSObject]]$signsInList = @()

    if ($StartDate) {
        try {
            $null = [datetime]::parseExact($StartDate, 'yyyy-MM-dd', $null)
        }
        catch {
            Write-Warning "Unable to get date from StartDate. Please add with the following format: yyyy-mm-dd $($_.Exception.Message)"
            return
        }

        $dateFilter = "createdDateTime gt $StartDate"
    }
    else {
        # define StartDate 31 days ago (30 days max) : https://docs.microsoft.com/en-us/azure/active-directory/reports-monitoring/reference-reports-data-retention#how-long-does-azure-ad-store-the-data
        $dateFilter = "createdDateTime gt $((Get-Date).AddDays(-31).Tostring('yyyy-MM-dd'))"
    }

    if ($EndDate) {
        try {
            $null = [datetime]::parseExact($EndDate, 'yyyy-MM-dd', $null)
        }
        catch {
            Write-Warning "Unable to get date from EndDate. Please add with the following format: yyyy-mm-dd $($_.Exception.Message)"
            return
        }

        $dateFilter += " and createdDateTime lt $EndDate"
    }
    else {
        # define endDate to tomorrow to be sure to take everything from today
        $dateFilter += " and createdDateTime lt $((Get-Date).AddDays(1).Tostring('yyyy-MM-dd'))"
    }

    
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

    if ($BadCredentialsOnly) {
        Write-Verbose 'Get signs-in with bad username or password'

        $filter += ' and status/errorCode eq 50126'
    }

    if ($FailuresOnly) {
        Write-Verbose 'Get signs-in with status failure'
        
        # ignore null because Teams return null
        # ignore 50140 because it means "This occurred due to 'Keep me signed in' interrupt when the user was signing in"
        
        $filter += ' and status/errorCode ne 0 and status/errorCode ne 50140'
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
                $userFilter = " or ipaddress eq '$IPAddress'"
            }
            else {
                $userFilter += " ipaddress eq '$IPAddress'"
            }
        }

        $filter += "$ipFilter"
    }


    if ($LastXSignIns) {
        Write-Verbose "Get-MgAuditLogSignIn -Top: $LastXSignIns -Filter $filter"
        $signsIn = Get-MgAuditLogSignIn -Top $LastXSignIns -Filter $filter
    }
    else {
        Write-Verbose "Get-MgAuditLogSignIn -All:`$true -Filter $filter"

        $signsIn = Get-MgAuditLogSignIn -All:$true -Filter $filter
    }
    
    Write-Verbose "Filter is $filter"

    if ($ConditionalAccessPolicyName) {
        Write-Verbose "Filter signs-in with Conditional Access Policy Name: $ConditionalAccessPolicyName"
        $signsIn = $signsIn | Where-Object {
            $_.Result.DisplayName -contains $ConditionalAccessPolicyName
        }
    }

    if ($AnalyzeReportOnlyCA ) {
        Write-Verbose "Filter signs-in with Conditional Access applied in ReportOnly mode and used (exclude 'NotApplied')"
        
        foreach ($signIn in $signsIn) {

            # https://groovynerd.co.uk/how-to-gather-reportonly-conditional-access-sign-in-logs/
            $reportOnlyPolicies = $SignIn.AppliedConditionalAccessPolicies

            # Loop through each policy
            foreach ($policy in $reportOnlyPolicies) {
                # Check if the policy result is in ReportOnly mode (and not 'Not Applied')
                # values can be any of the following: 'success', 'failure', 'notApplied', 'reportOnlySuccess', 'reportOnlyFailure', 'reportOnlyNotApplied' and 'notEnabled'
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
                Time                             = $signIn.CreatedDateTime
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
            }

            $signsInList.Add($object)
        }
    }

    return $signsInList

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
}