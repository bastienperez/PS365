<#
    .SYNOPSIS
    Retrieves all Entra ID applications configured for SAML SSO.

    .DESCRIPTION
    This function returns a list of all Entra ID applications configured for SAML Single Sign-On
    along with their SAML-related properties. Each row represents one SAML signing certificate
    (KeyCredential with Usage 'Sign'), so an application with multiple certificates will appear
    on multiple rows.

    The `SamlSigningCertificateIsPreferred` column identifies the currently active certificate:
    - True  : this is the active signing certificate
    - False : this certificate exists on the application but is not currently active

    .PARAMETER ObjectID
    (Optional) Retrieves the SAML configuration for a specific application by its ObjectID.

    .PARAMETER DisplayName
    (Optional) Retrieves the SAML configuration for a specific application by its DisplayName.
    Supports wildcards (* and ?) for partial name matching (e.g. "Azure*", "*Portal*").

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER ForceNewToken
    (Optional) Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .PARAMETER RunFromAzureAutomation
    (Optional) If specified, uses managed identity authentication instead of interactive authentication.
    This is useful when running the script in Azure environments like Azure Functions, Logic Apps, or VMs with managed identity enabled.
    When this parameter is used, ExpirationThresholdDays, NotificationRecipient and NotificationSender are required.

    PowerShell modules used in Azure Automation must be a MAXIMUM of version 2.25.0 when using PowerShell < 7.4.0, because starting from version 2.26.0, PowerShell 7.4.0 is required, and Azure Automation does not support it yet as of February 2026. For PowerShell 7.4.0+, there are no version restrictions.
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3147
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3151
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3166

    .PARAMETER ExpirationThresholdDays
    (Required when RunFromAzureAutomation is enabled) Number of days threshold for expiration notification. Default is 30 days.

    .PARAMETER NotificationRecipient
    (Required when RunFromAzureAutomation is enabled) Email address to receive expiration notifications.

    .PARAMETER NotificationSender
    (Required when RunFromAzureAutomation is enabled) Email address of the sender for expiration notifications.

    .PARAMETER IncludeSignInStats
    (Optional) If specified, includes sign-in statistics for the last 30 days for each application. Requires AuditLog.Read.All permission.
    Please be advised that this process is time-consuming.

    .PARAMETER DisableParallel
    (Optional) Forces sequential processing. By default, on PowerShell 7+ the function processes SAML applications in
    parallel (ForEach-Object -Parallel) to speed up discovery; on PowerShell 5.1 it always runs sequentially.

    .PARAMETER ThrottleLimit
    (Optional) Maximum number of concurrent runspaces when running in parallel. Default is 5.
    Keep this value moderate to avoid Microsoft Graph throttling (HTTP 429).

    .EXAMPLE
    Get-MgApplicationSAML

    Retrieves all Entra ID applications configured for SAML SSO.

    .EXAMPLE
    Get-MgApplicationSAML -DisableParallel

    Forces sequential processing even on PowerShell 7+ (useful for debugging or to avoid concurrent Graph calls).

    .EXAMPLE
    Get-MgApplicationSAML -IncludeSignInStats

    Retrieves all Entra ID applications configured for SAML SSO with sign-in statistics for the last 30 days.

    .EXAMPLE
    Get-MgApplicationSAML -ObjectID "xxx-xxx-xxx"

    Retrieves the SAML configuration for a specific application by its ObjectID.

    .EXAMPLE
    Get-MgApplicationSAML -DisplayName "My SAML App"

    Retrieves the SAML configuration for a specific application by its DisplayName.

    .EXAMPLE
    Get-MgApplicationSAML -DisplayName "Azure*"

    Retrieves the SAML configuration for all applications whose DisplayName starts with "Azure".

    .EXAMPLE
    Get-MgApplicationSAML -DisplayName "*Portal*"

    Retrieves the SAML configuration for all applications whose DisplayName contains "Portal".

    .EXAMPLE
    Get-MgApplicationSAML -ForceNewToken

    Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .EXAMPLE
    Get-MgApplicationSAML -ExportToExcel

    Gets all SAML applications and exports them to an Excel file.

    .EXAMPLE
    Get-MgApplicationSAML -RunFromAzureAutomation -ExpirationThresholdDays 30 -NotificationRecipient 'admin@company.com' -NotificationSender 'automation@company.com'

    Gets all SAML applications using managed identity authentication and sends notification for certificates expiring within 30 days.

    .EXAMPLE
    Get-MgApplicationSAML -RunFromAzureAutomation -ExpirationThresholdDays 7 -NotificationRecipient 'admin@company.com' -NotificationSender 'automation@company.com'

    Gets all SAML applications using managed identity and sends email notification for certificates expiring within 7 days.

    .NOTES
    More information on: https://itpro-tips.com/get-azure-ad-saml-certificate-details/

    This function requires the Microsoft.Graph.Beta.Applications module to be installed.

    .NOTES
    Limitations:
    The information about the SAML applications clams is not available in the Microsoft Graph API v1 but in https://main.iam.ad.ext.azure.com/api/ApplicationSso/<service-principal-id>/FederatedSsoV2 so we don't get them

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationSAML
#>

function Get-MgApplicationSAML {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [Parameter(Mandatory = $false, ParameterSetName = 'ByObjectId')]
        [string]$ObjectID,

        [Parameter(Mandatory = $false, ParameterSetName = 'ByDisplayName')]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath,

        [Parameter(Mandatory = $false)]
        [switch]$RunFromAzureAutomation,

        [Parameter(Mandatory = $false)]
        [int]$ExpirationThresholdDays = 30,

        [Parameter(Mandatory = $false)]
        [string]$NotificationRecipient,

        [Parameter(Mandatory = $false)]
        [string]$NotificationSender,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSignInStats,

        [Parameter(Mandatory = $false)]
        [switch]$DisableParallel,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 20)]
        [int]$ThrottleLimit = 5
    )

    # Run in parallel by default on PowerShell 7+ (ForEach-Object -Parallel); always sequential on PowerShell 5.1.
    $useParallel = ($PSVersionTable.PSVersion.Major -ge 7) -and -not $DisableParallel

    # Validate notification parameters
    if ($RunFromAzureAutomation.IsPresent) {
        if ([string]::IsNullOrWhiteSpace($NotificationRecipient)) {
            Write-Error 'NotificationRecipient parameter is required when RunFromAzureAutomation is enabled.'
            return
        }
        if ([string]::IsNullOrWhiteSpace($NotificationSender)) {
            Write-Error 'NotificationSender parameter is required when RunFromAzureAutomation is enabled.'
            return
        }
        if ($ExpirationThresholdDays -le 0) {
            Write-Error 'ExpirationThresholdDays must be greater than 0 when RunFromAzureAutomation is enabled.'
            return
        }

        try {
            Import-Module 'Microsoft.Graph.Users.Actions' -ErrorAction Stop -ErrorVariable mgGraphMailMissing
        }
        catch {
            if ($mgGraphMailMissing) {
                Write-Warning "Failed to import Microsoft.Graph.Users.Actions module: $($mgGraphMailMissing.Exception.Message)"
            }

            return
        }
    }

    try {
        # At the date of writing (december 2023), PreferredTokenSigningKeyEndDateTime parameter is only on Beta profile
        Import-Module 'Microsoft.Graph.Beta.Applications' -ErrorAction Stop -ErrorVariable mgGraphAppsMissing
    }
    catch {
        if ($mgGraphAppsMissing) {
            Write-Warning "Please install the Microsoft.Graph.Beta.Applications module: $($mgGraphAppsMissing.Exception.Message)"
        }

        return
    }

    $isConnected = $false

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)

    if ($ForceNewToken.IsPresent) {
        Write-Verbose 'Disconnecting from Microsoft Graph'
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        $isConnected = $false
    }

    $scopes = (Get-MgContext).Scopes

    $permissionsNeeded = @('Application.Read.All')
    if ($RunFromAzureAutomation.IsPresent) {
        $permissionsNeeded += 'Mail.Send'
    }
    if ($IncludeSignInStats.IsPresent) {
        $permissionsNeeded += 'AuditLog.Read.All'
    }

    $permissionMissing = $permissionsNeeded | Where-Object { $_ -notin $scopes }

    if ($permissionMissing) {
        Write-Verbose "You need to have the $($permissionsNeeded -join ',') permission in the current token, disconnect to force getting a new token with the right permissions"
    }

    # Version check for Azure Automation before connecting
    if ($RunFromAzureAutomation.IsPresent) {
        # Only check module version if PowerShell < 7.4 (Azure Automation limitation)
        if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
            $mgAuth = Get-Module 'Microsoft.Graph.Authentication' -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            if ($mgAuth -and [version]$mgAuth.Version -gt [version]'2.25.0') {
                Write-Error "Microsoft.Graph.Authentication v$($mgAuth.Version) is not compatible with Azure Automation on PowerShell $($PSVersionTable.PSVersion). Maximum supported version is 2.25.0. Script execution stopped."
                return
            }
        }
    }

    if (-not $isConnected) {
        if ($RunFromAzureAutomation.IsPresent) {
            Write-Verbose 'Connecting to Microsoft Graph using Managed Identity'
            $null = Connect-MgGraph -Identity -NoWelcome
        }
        else {
            Write-Verbose "Connecting to Microsoft Graph. Scopes: $($permissionsNeeded -join ',')"
            $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
        }
    }

    # When running from Azure Automation, verify that all required permissions are granted to the managed identity
    # AuditLog.Read.All is non-blocking: the script continues but sign-in stats will be unavailable
    $auditLogPermissionMissing = $false
    if ($RunFromAzureAutomation.IsPresent) {
        $currentScopes = (Get-MgContext).Scopes
        $missingPermissions = $permissionsNeeded | Where-Object { $_ -notin $currentScopes }
        if ($missingPermissions) {
            $criticalMissing = $missingPermissions | Where-Object { $_ -ne 'AuditLog.Read.All' }
            if ($criticalMissing) {
                Write-Error "The managed identity is missing the following required Graph permissions: $($criticalMissing -join ', '). Please grant these permissions to the managed identity and try again."
                return
            }
            if ($missingPermissions -contains 'AuditLog.Read.All') {
                $auditLogPermissionMissing = $true
                Write-Warning "The managed identity is missing the 'AuditLog.Read.All' permission. Sign-in statistics will not be available. The script will continue without sign-in data."
            }
        }
    }
    
    # Determine how to search for the Service Principal(s): by ObjectID (GUID), by DisplayName, or all
    if ($ObjectID) {
        $samlApplications = Get-MgBetaServicePrincipal -ServicePrincipalId $ObjectID
        # Verify it's a SAML application
        if ($samlApplications.PreferredSingleSignOnMode -ne 'saml') {
            Write-Warning "The application with ObjectID '$ObjectID' is not configured for SAML SSO (PreferredSingleSignOnMode: $($samlApplications.PreferredSingleSignOnMode))"
            return
        }
    }
    elseif ($DisplayName) {
        if ($DisplayName -match '[*?]') {
            # Wildcard mode: extract base search keyword for server-side pre-filter
            $searchKeyword = ($DisplayName -replace '[*?]', '').Trim()
            Write-Verbose "Wildcard detected. Using Graph \$search with keyword '$searchKeyword' then PowerShell -like '$DisplayName'"

            if ($searchKeyword) {
                $uri = "/beta/servicePrincipals?`$search=`"displayName:$searchKeyword`"&`$count=true&`$select=id,displayName,preferredSingleSignOnMode"
                $result = Invoke-MgGraphRequest -Uri $uri -Method GET -Headers @{ 'ConsistencyLevel' = 'eventual' }
            }
            else {
                # Pattern is only wildcards — fetch all SAML apps
                $uri = "/beta/servicePrincipals?`$filter=preferredSingleSignOnMode eq 'saml'&`$select=id,displayName,preferredSingleSignOnMode"
                $result = Invoke-MgGraphRequest -Uri $uri -Method GET
            }

            $candidateItems = if ($result.Value) { $result.Value } else { @() }
            # Post-filter: wildcard name match + SAML only
            $matchingIds = ($candidateItems | Where-Object { $_.displayName -like $DisplayName -and $_.preferredSingleSignOnMode -eq 'saml' }).id

            if ($matchingIds) {
                $samlApplications = @($matchingIds | ForEach-Object { Get-MgBetaServicePrincipal -ServicePrincipalId $_ })
                Write-Verbose "Found $($samlApplications.Count) SAML application(s) matching wildcard pattern '$DisplayName'"
            }
            else {
                $samlApplications = @()
            }
        }
        else {
            # Exact match mode (no wildcards)
            $escaped = $DisplayName -replace "'", "''"
            $filter = "DisplayName eq '$escaped' and PreferredSingleSignOnMode eq 'saml'"
            Write-Verbose "Filtering service principals with: $filter"
            $samlApplications = Get-MgBetaServicePrincipal -Filter $filter -All

            # If no exact match found, try to find apps where trimmed DisplayName matches
            if (-not $samlApplications) {
                Write-Verbose "No exact match found. Searching for apps with trimmed DisplayName matching '$DisplayName'..."
                $filter = "startswith(DisplayName, '$escaped') and PreferredSingleSignOnMode eq 'saml'"
                $candidateApps = Get-MgBetaServicePrincipal -Filter $filter -All

                # Filter in PowerShell to find apps where trimmed name matches
                $samlApplications = $candidateApps | Where-Object { $_.DisplayName.Trim() -eq $DisplayName }

                if ($samlApplications) {
                    Write-Verbose "Found $($samlApplications.Count) application(s) with trimmed DisplayName matching '$DisplayName'"
                }
            }
        }
    }
    else {
        $samlApplications = Get-MgBetaServicePrincipal -Filter "PreferredSingleSignOnMode eq 'saml'"
    }

    if (-not $samlApplications) {
        Write-Host 'No SAML applications found' -ForegroundColor Yellow
        return
    }

    Write-Host "$($samlApplications.Count) SAML application(s) found" -ForegroundColor Green

    [System.Collections.Generic.List[PSCustomObject]]$samlApplicationsArray = @()

    # Calculate date for 30 days ago for sign-in statistics
    $signInStartDate = (Get-Date).AddDays(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')

    if ($IncludeSignInStats.IsPresent) {
        Write-Host 'Retrieving sign-in statistics for each application - this may take several minutes...' -ForegroundColor Yellow
    }

    # Per-application processing. Emits one object per signing certificate (or a fallback object).
    # Shared by the sequential and parallel paths.
    $processSamlApp = {
        param($samlApp, $Prefix = '', $IncludeSignInStats = $false, $AuditLogPermissionMissing = $false, $SignInStartDate = $null)

        Write-Host "$Prefix$($samlApp.DisplayName)" -ForegroundColor Cyan
        # Reset to $null before each call: prevents previous iteration's value from bleeding through on silent errors
        $ownerObjects = $null
        $ownerString = $null
        $ownerObjects = Get-MgServicePrincipalOwner -ServicePrincipalId $samlApp.Id -ErrorAction SilentlyContinue

        # Build owners string: DisplayName for each owner, joined with '|'
        if ($ownerObjects) {
            $ownerString = ($ownerObjects | ForEach-Object {
                    $props = $_.AdditionalProperties
                    if ($props.ContainsKey('displayName') -and -not [string]::IsNullOrEmpty($props['displayName'])) { $props['displayName'] }
                    else { $_.Id }
                }) -join '|'
        }
        
        # Check for leading/trailing spaces in DisplayName
        $recommendation = $null
        if ($samlApp.DisplayName -ne $samlApp.DisplayName.Trim()) {
            $recommendation = 'DisplayName contains leading or trailing spaces - consider renaming'
            Write-Warning "Application '$($samlApp.DisplayName)' has leading or trailing spaces in the displayName"
        }
        
        # Get sign-in statistics if requested
        $signInCount = $null
        if ($AuditLogPermissionMissing) {
            $signInCount = 'N/A - AuditLog.Read.All permission missing'
        }
        elseif ($IncludeSignInStats) {
            try {
                $signInFilter = "appId eq '$($samlApp.AppId)' and createdDateTime ge $SignInStartDate"
                Write-Verbose "Sign-in filter: $signInFilter"
                $encodedFilter = [uri]::EscapeDataString($signInFilter)
                $uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=$encodedFilter&`$count=true&`$top=999"
                try{
                    $signInResponse = Invoke-MgGraphRequest -Uri $uri -Method GET -Headers @{ ConsistencyLevel = 'eventual' } -ErrorAction Stop -WarningAction Stop
                }
                catch {
                    $signInCount = "Problem to get sign-ins - $($_.Exception.Message)"
                }

                if ($null -ne $signInResponse.'@odata.count') {
                    $signInCount = [int]$signInResponse.'@odata.count'
                } else {
                    $allSignIns = [System.Collections.Generic.List[object]]@()
                    $signInResponse.value | Where-Object { $_.isInteractive -eq $true } | ForEach-Object { $null = $allSignIns.Add($_) }
                    $nextLink = $signInResponse.'@odata.nextLink'
                    while ($nextLink) {
                        $pageResponse = Invoke-MgGraphRequest -Uri $nextLink -Method GET -Headers @{ ConsistencyLevel = 'eventual' }
                        $pageResponse.value | Where-Object { $_.isInteractive -eq $true } | ForEach-Object { $null = $allSignIns.Add($_) }
                        $nextLink = $pageResponse.'@odata.nextLink'
                    }
                    $signInCount = $allSignIns.Count
                }
                
                Write-Verbose "Found $signInCount sign-ins in the last 30 days for $($samlApp.DisplayName)"
            }
            catch {
                Write-Warning "Could not retrieve sign-in statistics for '$($samlApp.DisplayName)': $($_.Exception.Message)"
                $signInCount = $null
            }
        }

        # Iterate over all signing certificates (KeyCredentials with Usage 'Sign')
        $signingCerts = $samlApp.KeyCredentials | Where-Object { $_.Usage -eq 'Sign' }

        if ($signingCerts) {
            foreach ($cert in $signingCerts) {
                # Identify the currently active/preferred certificate by matching its expiry date
                $isPreferred = $null -ne $cert.EndDateTime -and
                               $null -ne $samlApp.PreferredTokenSigningKeyEndDateTime -and
                               [datetime]$cert.EndDateTime -eq [datetime]$samlApp.PreferredTokenSigningKeyEndDateTime

                $object = [PSCustomObject][ordered]@{
                    DisplayName                         = $samlApp.DisplayName
                    Recommendation                      = $recommendation
                    Id                                  = $samlApp.Id
                    AppId                               = $samlApp.AppId
                    EntraUrl                            = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($samlApp.Id)/appId/$($samlApp.AppId)"
                    LoginUrl                            = $samlApp.LoginUrl
                    LogoutUrl                           = $samlApp.LogoutUrl
                    NotificationEmailAddresses          = $samlApp.NotificationEmailAddresses -join '|'
                    AppRoleAssignmentRequired           = $samlApp.AppRoleAssignmentRequired
                    PreferredSingleSignOnMode           = $samlApp.PreferredSingleSignOnMode
                    SamlSigningCertificateIsPreferred   = $isPreferred
                    SamlSigningCertificateDisplayName   = $cert.DisplayName
                    SamlSigningCertificateKeyId         = $cert.KeyId
                    SamlSigningCertificateStartTime     = $cert.StartDateTime
                    SamlSigningCertificateEndTime       = $cert.EndDateTime
                    # EndDateTime compared to now to check validity
                    SamlSigningCertificateValid         = $cert.EndDateTime -gt (Get-Date)
                    SamlSigningCertificateExpiresInDays = if ($cert.EndDateTime) { [int](New-TimeSpan -Start (Get-Date) -End $cert.EndDateTime).TotalDays } else { $null }
                    ReplyUrls                           = $samlApp.ReplyUrls -join '|'
                    SignInAudience                      = $samlApp.SignInAudience
                    Owners                              = $ownerString
                }

                if ($IncludeSignInStats -or $AuditLogPermissionMissing) {
                    $object | Add-Member -MemberType NoteProperty -Name InteractiveSignInsLast30Days -Value $signInCount
                }

                $object
            }
        }
        else {
            # No signing KeyCredentials found - fall back to PreferredTokenSigningKeyEndDateTime
            $object = [PSCustomObject][ordered]@{
                DisplayName                         = $samlApp.DisplayName
                Recommendation                      = $recommendation
                Id                                  = $samlApp.Id
                AppId                               = $samlApp.AppId
                EntraUrl                            = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($samlApp.Id)/appId/$($samlApp.AppId)"
                LoginUrl                            = $samlApp.LoginUrl
                LogoutUrl                           = $samlApp.LogoutUrl
                NotificationEmailAddresses          = $samlApp.NotificationEmailAddresses -join '|'
                AppRoleAssignmentRequired           = $samlApp.AppRoleAssignmentRequired
                PreferredSingleSignOnMode           = $samlApp.PreferredSingleSignOnMode
                SamlSigningCertificateIsPreferred   = $null
                SamlSigningCertificateDisplayName   = $null
                SamlSigningCertificateKeyId         = $null
                SamlSigningCertificateStartTime     = $null
                SamlSigningCertificateEndTime       = $samlApp.PreferredTokenSigningKeyEndDateTime
                # PreferredTokenSigningKeyEndDateTime compared to now to check validity
                SamlSigningCertificateValid         = $samlApp.PreferredTokenSigningKeyEndDateTime -gt (Get-Date)
                SamlSigningCertificateExpiresInDays = if ($samlApp.PreferredTokenSigningKeyEndDateTime) { [int](New-TimeSpan -Start (Get-Date) -End $samlApp.PreferredTokenSigningKeyEndDateTime).TotalDays } else { $null }
                ReplyUrls                           = $samlApp.ReplyUrls -join '|'
                SignInAudience                      = $samlApp.SignInAudience
                Owners                              = $ownerString
            }

            if ($IncludeSignInStats -or $AuditLogPermissionMissing) {
                $object | Add-Member -MemberType NoteProperty -Name InteractiveSignInsLast30Days -Value $signInCount
            }

            $object
        }
    }

    if ($useParallel) {
        Write-Verbose "Processing SAML applications in parallel (ThrottleLimit: $ThrottleLimit)..."
        $processText = $processSamlApp.ToString()
        $inclStats = $IncludeSignInStats.IsPresent
        $parallelResults = $samlApplications | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
            Import-Module Microsoft.Graph.Beta.Applications -ErrorAction SilentlyContinue
            $sb = [scriptblock]::Create($using:processText)
            & $sb $_ '' $using:inclStats $using:auditLogPermissionMissing $using:signInStartDate
        }
        foreach ($result in $parallelResults) {
            if ($result) { $samlApplicationsArray.Add($result) }
        }
    }
    else {
        $appCounter = 0
        foreach ($samlApp in $samlApplications) {
            $appCounter++
            $results = & $processSamlApp $samlApp "Processing $appCounter/$($samlApplications.Count): " $IncludeSignInStats.IsPresent $auditLogPermissionMissing $signInStartDate
            foreach ($result in $results) {
                if ($result) { $samlApplicationsArray.Add($result) }
            }
        }
    }

    # Check for expiring certificates and send notification if enabled
    if ($RunFromAzureAutomation.IsPresent) {
        $expiringCertificates = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and
            $_.SamlSigningCertificateExpiresInDays -le $ExpirationThresholdDays
        }

        # Calculate statistics for different expiration categories
        $halfThreshold = [int]($ExpirationThresholdDays / 2)
        $expiredCertificates = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and $_.SamlSigningCertificateExpiresInDays -le 0
        }
        $certificatesExpiring15Days = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and $_.SamlSigningCertificateExpiresInDays -le $halfThreshold -and $_.SamlSigningCertificateExpiresInDays -gt 0
        }
        $certificatesExpiring30Days = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and $_.SamlSigningCertificateExpiresInDays -le $ExpirationThresholdDays -and $_.SamlSigningCertificateExpiresInDays -gt 0
        }

        Write-Verbose "Sending notification email. Found $($expiringCertificates.Count) SAML certificates expiring within $ExpirationThresholdDays days."

        $expiringCertificates = $expiringCertificates | Sort-Object SamlSigningCertificateExpiresInDays

        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
<title>Microsoft Entra ID SAML Application Certificates Expiration Alert</title>
<style>
    body { 
        font-family: Segoe UI, SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif; 
        margin: 0; 
        padding: 20px; 
        color: #11100f; 
        font-size: 14px;
        line-height: 20px;
        background-color: #ffffff;
    }
    
    h2 { 
        padding-top: 0; 
        margin: 0 0 16px 0; 
        font-family: "Segoe UI Semibold", SegoeUISemibold, "Segoe UI", SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif;
        font-weight: 600; 
        font-size: 20px; 
        line-height: 28px;
        color: #323130;
    }
    
    table { 
        border-spacing: 0; 
        border-collapse: collapse; 
        width: 100%; 
        margin-bottom: 20px;
        background-color: #ffffff;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    th { 
        vertical-align: middle;
        color: #ffffff;
        background-color: #323130;
        padding: 3px 8px;
        text-align: left;
        font-family: "Segoe UI Semibold", SegoeUISemibold, "Segoe UI", SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif;
        font-weight: 600;
        font-size: 12px;
        line-height: 16px;
        word-wrap: break-word;
    }
    
    td { 
        vertical-align: middle;
        color: #11100f;
        padding: 3px 8px;
        border-bottom: solid 1px #c8c6c4;
        word-wrap: break-word;
        font-size: 12px;
        line-height: 16px;
    }
    
    .critical { background-color: #FFF0F0; color: #A80000; }
    .warning { background-color: #FDEFD0; color: #7A3A00; }
    .caution { background-color: #CCE4FF; color: #003882; }
    
    .footer {
        margin-top: 30px;
        padding: 20px;
        background-color: #faf9f8;
        border-radius: 8px;
        border-top: 3px solid #0078d4;
    }
    
    .footer p {
        margin: 8px 0;
        font-size: 13px;
        color: #605e5c;
    }
    
    .action-required {
        font-weight: 600;
        color: #d73502;
    }
</style>
</head>
<body>
    <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:12px;background:transparent;box-shadow:none;" role="presentation">
        <tr>
            <td width="33%" valign="top" style="width:33%;padding:4pt 3pt 4pt 5pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#FFF0F0;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#A80000;font-weight:600;">Expired certificates</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#A80000;font-weight:bold;">$($expiredCertificates.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#A80000;border-bottom:none;vertical-align:middle;">
                                        certificates already expired
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="34%" valign="top" style="width:34%;padding:4pt 3pt 4pt 3pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#FDEFD0;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#7A3A00;font-weight:600;">Expiring within $halfThreshold days</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#7A3A00;font-weight:bold;">$($certificatesExpiring15Days.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#7A3A00;border-bottom:none;vertical-align:middle;">
                                        certificates expire within $halfThreshold days
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="33%" valign="top" style="width:33%;padding:4pt 5pt 4pt 3pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#CCE4FF;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#003882;font-weight:600;">Expiring within $ExpirationThresholdDays days</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#003882;font-weight:bold;">$($certificatesExpiring30Days.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#003882;border-bottom:none;vertical-align:middle;">
                                        certificates expire within $ExpirationThresholdDays days
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    
    <h2>SAML Certificates requiring attention</h2>
    <table>
        <tr>
            <th>Application</th>
            <th>App ID</th>
            <th>Expires In (Days)</th>
            <th>Expiry Date</th>
            <th>Login URL</th>
            <th>Owners</th>
            $(if ($IncludeSignInStats.IsPresent -or $auditLogPermissionMissing) { '<th>Sign-ins (30d)</th>' })
        </tr>
"@

        foreach ($cert in $expiringCertificates) {
            $rowClass = if ($cert.SamlSigningCertificateExpiresInDays -le 0) { 'critical' } elseif ($cert.SamlSigningCertificateExpiresInDays -le $halfThreshold) { 'warning' } else { 'caution' }
            $expiresInDaysDisplay = if ($cert.SamlSigningCertificateExpiresInDays -lt 0) { "$($cert.SamlSigningCertificateExpiresInDays) (already expired)" } else { $cert.SamlSigningCertificateExpiresInDays }
            $appLink = "<strong style=`"color:#11100f;font-size:12px;line-height:16px;`">$($cert.DisplayName)</strong> <a href=`"$($cert.EntraUrl)`" style=`"text-decoration:none;font-size:14px;line-height:16px;`" title=`"Open in Entra`">&#x1F517;</a>"
            $ownersList = @($cert.Owners -split '\|' | Where-Object { $_ })
            $ownersHtml = if ($ownersList.Count -gt 1) { ($ownersList | ForEach-Object { "- $_" }) -join '<br>' } else { $ownersList -join '' }
            $signInsHtml = if ($IncludeSignInStats.IsPresent -or $auditLogPermissionMissing) { "<td>$($cert.InteractiveSignInsLast30Days)</td>" } else { '' }
            $emailBody += "<tr class=`"$rowClass`"><td>$appLink</td><td>$($cert.AppId)</td><td><strong>$expiresInDaysDisplay</strong></td><td>$($cert.SamlSigningCertificateEndTime)</td><td>$($cert.LoginUrl)</td><td>$ownersHtml</td>$signInsHtml</tr>"
        }

        if ($expiringCertificates.Count -eq 0) {
            $colSpan = if ($IncludeSignInStats.IsPresent -or $auditLogPermissionMissing) { 7 } else { 6 }
            $emailBody += "<tr><td colspan=`"$colSpan`" style=`"text-align:center;padding:12px 8px;color:#605e5c;font-style:italic;`">No certificates requiring attention - all SAML signing certificates are healthy.</td></tr>"
        }

        $emailFooter = if ($expiringCertificates.Count -gt 0) {
            '<p class="action-required">Action Required:</p><p>Please review and renew these SAML signing certificates before they expire to avoid authentication disruptions.</p>'
        }
        else {
            '<p style="color:#107C10;font-weight:600;">All SAML signing certificates are healthy. No action required at this time.</p>'
        }

        $emailSubject = if ($expiringCertificates.Count -gt 0) {
            "Microsoft Entra ID SAML Certificates Expiring ($($expiringCertificates.Count) certificates)"
        }
        else {
            'Microsoft Entra ID SAML Certificates - All Healthy'
        }

        $emailBody += @"
    </table>

    <div class="footer">
        $emailFooter
        <hr style="border: none; border-top: 1px solid #d2d0ce; margin: 15px 0;">
        <p><em>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Get-MgApplicationSAML v0.1.9</em></p>
    </div>
</body>
</html>
"@

        try {
            $params = @{
                Message         = @{
                    Subject      = $emailSubject
                    Body         = @{
                        ContentType = 'HTML'
                        Content     = $emailBody
                    }
                    ToRecipients = @(
                        @{
                            EmailAddress = @{
                                Address = $NotificationRecipient
                            }
                        }
                    )
                }
                SaveToSentItems = $false
            }

            Send-MgUserMail -UserId $NotificationSender -BodyParameter $params
            Write-Host -ForegroundColor Green "Expiration notification email sent successfully to $NotificationRecipient"
        }
        catch {
            Write-Warning "Failed to send notification email: $($_.Exception.Message)"
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$(if ($ExportPath) { $ExportPath } else { $env:userprofile })\$now-MgApplicationSAML.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting SAML applications to Excel file: $excelFilePath"
        $samlApplicationsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraSAMLApplications'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif (-not $RunFromAzureAutomation.IsPresent) {
        return $samlApplicationsArray
    }
}