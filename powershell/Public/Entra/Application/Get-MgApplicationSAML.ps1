<#
    .SYNOPSIS
    Retrieves all Entra ID applications configured for SAML SSO.

    .DESCRIPTION
    This function returns a list of all Entra ID applications configured for SAML Single Sign-On
    along with their SAML-related properties. Each row represents one certificate, so an application
    with several certificates appears on several rows. The `CertificateType` column says which:
    - SAML signing      : the certificate Entra signs the SAML tokens with (service principal,
                          KeyCredential with Usage 'Sign')
    - SAML verification : the service provider's public key Entra checks signed SAML requests with
                          (app registration, KeyCredential with Usage 'Verify')
    - Token encryption  : the service provider's certificate Entra encrypts the SAML tokens with
                          (app registration, KeyCredential with Usage 'Encrypt')

    The `CertificateIsActive` column says whether the certificate is in use:
    - SAML signing      : True for the certificate of preferredTokenSigningKeyThumbprint (or, without
                          it, of preferredTokenSigningKeyEndDateTime). With neither, every one is True.
    - SAML verification : True while signed requests are required (requestSignatureVerification)
    - Token encryption  : True for the certificate named by tokenEncryptionKeyId

    Every Graph call goes through Invoke-MgGraphRequest: only the Microsoft.Graph.Authentication
    module is needed. An owner lookup that fails says so in the Owners column (it is never left empty).

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
    (Required when RunFromAzureAutomation is enabled) Number of days threshold for expiration notification. Default is 30 days. Only the active signing certificate of each application is counted.

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

    .PARAMETER UseBatchRequest
    (Optional) Uses the Microsoft Graph JSON $batch endpoint (via Invoke-MgGraphBatchRequest, 20 requests per
    HTTP call with automatic HTTP 429 retry) to retrieve the application owners, instead of ForEach-Object -Parallel.
    Requires the PS365 module (Invoke-MgGraphBatchRequest); when the command is not available (e.g. script
    deployed standalone), the function warns and falls back to the default behavior.
    Note: with -IncludeSignInStats, the sign-in statistics are still retrieved per application (not batched).

    .EXAMPLE
    Get-MgApplicationSAML

    Retrieves all Entra ID applications configured for SAML SSO.

    .EXAMPLE
    Get-MgApplicationSAML -UseBatchRequest

    Uses the Graph JSON $batch endpoint to retrieve the application owners (20 requests per HTTP call, automatic 429 retry).

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

    This function requires only the Microsoft.Graph.Authentication module (Invoke-MgGraphRequest).

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
        [Alias('Identity')]
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
        [int]$ThrottleLimit = 5,

        [Parameter(Mandatory = $false)]
        [switch]$UseBatchRequest
    )

    # Execution duration, displayed in the Azure Automation HTML report footer
    $executionStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Batch mode requires Invoke-MgGraphBatchRequest (PS365 module). When the script is deployed
    # standalone (e.g. copied to a client environment or an Azure Automation runbook), fall back
    # to the default parallel/sequential behavior instead of failing.
    $useBatch = $UseBatchRequest.IsPresent
    if ($useBatch -and -not (Get-Command 'Invoke-MgGraphBatchRequest' -ErrorAction SilentlyContinue)) {
        Write-Warning 'UseBatchRequest requires the Invoke-MgGraphBatchRequest command (PS365 module), which is not available. Falling back to the default discovery method.'
        $useBatch = $false
    }

    # Run in parallel by default on PowerShell 7+ (ForEach-Object -Parallel); always sequential on PowerShell 5.1.
    $useParallel = (-not $useBatch) -and ($PSVersionTable.PSVersion.Major -ge 7) -and -not $DisableParallel

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
    }

    try {
        # Only Invoke-MgGraphRequest and Connect-MgGraph: no other Microsoft.Graph module is loaded
        Import-Module 'Microsoft.Graph.Authentication' -ErrorAction Stop -ErrorVariable mgGraphAuthMissing
    }
    catch {
        if ($mgGraphAuthMissing) {
            Write-Warning "Please install the Microsoft.Graph.Authentication module: $($mgGraphAuthMissing.Exception.Message)"
        }

        return
    }

    # The SAML properties of a service principal. preferredTokenSigningKeyEndDateTime is only on the beta endpoint.
    $spSelect = 'id,appId,displayName,loginUrl,logoutUrl,notificationEmailAddresses,appRoleAssignmentRequired,preferredSingleSignOnMode,preferredTokenSigningKeyThumbprint,preferredTokenSigningKeyEndDateTime,keyCredentials,replyUrls,signInAudience'

    # GET with paging, the results as PSCustomObject (property names are then case-insensitive)
    $getGraphItems = {
        param([string]$Uri, [hashtable]$Headers = @{})
        $items = [System.Collections.Generic.List[object]]::new()
        $nextLink = $Uri
        while ($nextLink) {
            $page = Invoke-MgGraphRequest -Uri $nextLink -Method GET -Headers $Headers -OutputType PSObject -ErrorAction Stop
            if ($null -ne $page.PSObject.Properties['value']) {
                foreach ($item in $page.value) { $items.Add($item) }
                $nextLink = $page.'@odata.nextLink'
            }
            else {
                $items.Add($page)
                $nextLink = $null
            }
        }
        return , $items
    }

    $isConnected = $false

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)

    if ($ForceNewToken.IsPresent) {
        Write-Verbose 'Disconnecting from Microsoft Graph'
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        $isConnected = $false
    }

    $permissionsNeeded = @('Application.Read.All')
    if ($RunFromAzureAutomation.IsPresent) {
        $permissionsNeeded += 'Mail.Send'
    }
    if ($IncludeSignInStats.IsPresent) {
        $permissionsNeeded += 'AuditLog.Read.All'
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

    # Skip scope check in Azure Automation: Managed Identity uses fixed app-registration scopes,
    # and Test-MgGraphPermission lives in Private/ which isn't available when the script is deployed standalone in a runbook.
    if (-not $RunFromAzureAutomation.IsPresent) {
        if (-not (Test-MgGraphPermission -RequiredScopes $permissionsNeeded -CallerName $MyInvocation.MyCommand.Name)) {
            return
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
        $samlApplications = & $getGraphItems "/beta/servicePrincipals/$($ObjectID)?`$select=$spSelect"
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

            $candidateItems = @(& { $page = $result; while ($page) { $page.value; $nextPage = $page.'@odata.nextLink'; $page = if ($nextPage) { Invoke-MgGraphRequest -Uri $nextPage -Method GET -Headers @{ ConsistencyLevel = 'eventual' } -ErrorAction Stop } else { $null } } })
            # Post-filter: wildcard name match + SAML only
            $matchingIds = ($candidateItems | Where-Object { $_.displayName -like $DisplayName -and $_.preferredSingleSignOnMode -eq 'saml' }).id

            if ($matchingIds) {
                $samlApplications = @($matchingIds | ForEach-Object { & $getGraphItems "/beta/servicePrincipals/$($_)?`$select=$spSelect" })
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
            $samlApplications = & $getGraphItems "/beta/servicePrincipals?`$filter=$([uri]::EscapeDataString($filter))&`$select=$spSelect"

            # If no exact match found, try to find apps where trimmed DisplayName matches
            if (-not $samlApplications) {
                Write-Verbose "No exact match found. Searching for apps with trimmed DisplayName matching '$DisplayName'..."
                $filter = "startswith(DisplayName, '$escaped') and PreferredSingleSignOnMode eq 'saml'"
                $candidateApps = & $getGraphItems "/beta/servicePrincipals?`$filter=$([uri]::EscapeDataString($filter))&`$select=$spSelect"

                # Filter in PowerShell to find apps where trimmed name matches
                $samlApplications = $candidateApps | Where-Object { $_.DisplayName.Trim() -eq $DisplayName }

                if ($samlApplications) {
                    Write-Verbose "Found $($samlApplications.Count) application(s) with trimmed DisplayName matching '$DisplayName'"
                }
            }
        }
    }
    else {
        $samlApplications = & $getGraphItems "/beta/servicePrincipals?`$filter=$([uri]::EscapeDataString("preferredSingleSignOnMode eq 'saml'"))&`$select=$spSelect"
    }

    $samlApplications = @($samlApplications | Where-Object { $_ })
    if (-not $samlApplications) {
        Write-Host 'No SAML applications found' -ForegroundColor Yellow
        return
    }

    Write-Host "$(@($samlApplications).Count) SAML application(s) found" -ForegroundColor Green

    $samlApplicationsArray = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Calculate date for 30 days ago for sign-in statistics
    $signInStartDate = (Get-Date).AddDays(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')

    if ($IncludeSignInStats.IsPresent) {
        Write-Host 'Retrieving sign-in statistics for each application - this may take several minutes...' -ForegroundColor Yellow
    }

    # Per-application processing. Emits one object per signing certificate (or a fallback object).
    # Shared by the sequential, parallel and batch paths: the batch path fetches the owners upfront
    # via the $batch endpoint and passes them as PreFetchedOwners.
    $processSamlApp = {
        param($samlApp, $Prefix = '', $IncludeSignInStats = $false, $AuditLogPermissionMissing = $false, $SignInStartDate = $null, $PreFetchedOwners = $null)

        Write-Host "$Prefix$($samlApp.DisplayName)" -ForegroundColor Cyan
        # Reset to $null before each call: prevents previous iteration's value from bleeding through on silent errors
        $ownerObjects = $null
        $ownerString = $null
        # A failed lookup is written in the Owners column: an empty value must mean "no owner"
        if ($null -ne $PreFetchedOwners) {
            $ownerObjects = $PreFetchedOwners
        }
        else {
            try {
                $ownerObjects = @(& {
                        $ownersUri = "/v1.0/servicePrincipals/$($samlApp.Id)/owners?`$select=id,displayName"
                        while ($ownersUri) {
                            $ownersPage = Invoke-MgGraphRequest -Uri $ownersUri -Method GET -OutputType PSObject -ErrorAction Stop
                            $ownersPage.value
                            $ownersUri = $ownersPage.'@odata.nextLink'
                        }
                    })
            }
            catch {
                $ownerString = "Error: $($_.Exception.Message)"
                Write-Warning "Could not read the owners of '$($samlApp.DisplayName)': $($_.Exception.Message)"
            }
        }

        # Build owners string: DisplayName for each owner (Id when it has none), joined with '|'
        if ($ownerObjects) {
            $ownerString = ($ownerObjects | ForEach-Object {
                    $ownerDisplayName = if ($_ -is [System.Collections.IDictionary]) { $_['displayName'] } else { $_.displayName }
                    $ownerId = if ($_ -is [System.Collections.IDictionary]) { $_['id'] } else { $_.id }
                    if ([string]::IsNullOrEmpty($ownerDisplayName)) { $ownerId } else { $ownerDisplayName }
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
                    $allSignIns = [System.Collections.Generic.List[object]]::new()
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

        # The service provider's certificates are on the app registration: verification (Usage 'Verify')
        # and token encryption (Usage 'Encrypt'). Not found (404): the app registration lives in another
        # tenant, there is nothing to read.
        $appRegistration = $null
        try {
            $appRegistration = Invoke-MgGraphRequest -Uri "/v1.0/applications(appId='$($samlApp.AppId)')?`$select=id,keyCredentials,tokenEncryptionKeyId,requestSignatureVerification" -Method GET -OutputType PSObject -ErrorAction Stop
        }
        catch {
            if ($_.Exception.Message -notmatch 'NotFound|404|does not exist') {
                Write-Warning "Could not read the app registration of '$($samlApp.DisplayName)': $($_.Exception.Message)"
                $recommendation = (@($recommendation, "App registration not read: $($_.Exception.Message)") | Where-Object { $_ }) -join ' | '
            }
        }
        $signedRequestsRequired = if ($appRegistration) { $appRegistration.requestSignatureVerification.isSignedRequestRequired -eq $true } else { $null }

        $now = Get-Date
        $newCertificateRow = {
            param($Type, $IsActive, $Certificate, $EndDate)
            $endTime = if ($EndDate) { [datetime]$EndDate } else { $null }
            $startTime = if ($Certificate -and $Certificate.startDateTime) { [datetime]$Certificate.startDateTime } else { $null }
            $row = [PSCustomObject][ordered]@{
                DisplayName                = $samlApp.DisplayName
                Recommendation             = $recommendation
                Id                         = $samlApp.Id
                AppId                      = $samlApp.AppId
                EntraUrl                   = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($samlApp.Id)/appId/$($samlApp.AppId)"
                LoginUrl                   = $samlApp.LoginUrl
                LogoutUrl                  = $samlApp.LogoutUrl
                NotificationEmailAddresses = $samlApp.NotificationEmailAddresses -join '|'
                AppRoleAssignmentRequired  = $samlApp.AppRoleAssignmentRequired
                PreferredSingleSignOnMode  = $samlApp.PreferredSingleSignOnMode
                SignedRequestsRequired     = $signedRequestsRequired
                CertificateType            = $Type
                CertificateIsActive        = $IsActive
                CertificateDisplayName     = if ($Certificate) { $Certificate.displayName } else { $null }
                CertificateKeyId           = if ($Certificate) { $Certificate.keyId } else { $null }
                CertificateStartTime       = $startTime
                CertificateEndTime         = $endTime
                CertificateValid           = if ($endTime) { $endTime -gt $now } else { $null }
                CertificateExpiresInDays   = if ($endTime) { [int](New-TimeSpan -Start $now -End $endTime).TotalDays } else { $null }
                ReplyUrls                  = $samlApp.ReplyUrls -join '|'
                SignInAudience             = $samlApp.SignInAudience
                Owners                     = $ownerString
            }
            if ($IncludeSignInStats -or $AuditLogPermissionMissing) {
                $row | Add-Member -MemberType NoteProperty -Name InteractiveSignInsLast30Days -Value $signInCount
            }
            $row
        }

        # SAML signing certificates (service principal, Usage 'Sign'). The active one is the one of
        # preferredTokenSigningKeyThumbprint (customKeyIdentifier is that thumbprint in base64), or else
        # the one ending at preferredTokenSigningKeyEndDateTime. With neither, every one counts as active.
        $signingCerts = @($samlApp.KeyCredentials | Where-Object { $_.Usage -eq 'Sign' })
        $preferredThumbprint = "$($samlApp.PreferredTokenSigningKeyThumbprint)".ToUpperInvariant()
        $preferredEnd = if ($samlApp.PreferredTokenSigningKeyEndDateTime) { [datetime]$samlApp.PreferredTokenSigningKeyEndDateTime } else { $null }

        if ($signingCerts) {
            foreach ($cert in $signingCerts) {
                $certThumbprint = $null
                try { $certThumbprint = [System.BitConverter]::ToString([Convert]::FromBase64String("$($cert.CustomKeyIdentifier)")).Replace('-', '') } catch { $certThumbprint = $null }
                $isActive = if ($preferredThumbprint) {
                    $certThumbprint -eq $preferredThumbprint
                }
                elseif ($preferredEnd) {
                    $null -ne $cert.EndDateTime -and [datetime]$cert.EndDateTime -eq $preferredEnd
                }
                else {
                    $true
                }
                & $newCertificateRow 'SAML signing' $isActive $cert $cert.EndDateTime
            }
        }
        else {
            # No signing KeyCredentials found - fall back to PreferredTokenSigningKeyEndDateTime
            & $newCertificateRow 'SAML signing' $null $null $samlApp.PreferredTokenSigningKeyEndDateTime
        }

        if ($appRegistration) {
            foreach ($cert in @($appRegistration.keyCredentials | Where-Object { $_.usage -eq 'Verify' })) {
                & $newCertificateRow 'SAML verification' $signedRequestsRequired $cert $cert.endDateTime
            }
            foreach ($cert in @($appRegistration.keyCredentials | Where-Object { $_.usage -eq 'Encrypt' })) {
                $isActive = -not [string]::IsNullOrEmpty($appRegistration.tokenEncryptionKeyId) -and $cert.keyId -eq $appRegistration.tokenEncryptionKeyId
                & $newCertificateRow 'Token encryption' $isActive $cert $cert.endDateTime
            }
        }
    }

    if ($useBatch) {
        Write-Verbose 'Retrieving service principal owners with the Graph $batch endpoint (20 requests per HTTP call)...'
        $ownersRequests = [System.Collections.Generic.List[hashtable]]::new()
        foreach ($samlApp in $samlApplications) {
            $ownersRequests.Add(@{
                    id     = "$($samlApp.Id)"
                    method = 'GET'
                    url    = "/servicePrincipals/$($samlApp.Id)/owners"
                })
        }

        $ownersResponses = Invoke-MgGraphBatchRequest -Requests $ownersRequests -GraphVersion 'v1.0' -Activity 'Getting service principal owners'

        $appCounter = 0
        $samlApplicationsCount = @($samlApplications).Count
        foreach ($samlApp in $samlApplications) {
            $appCounter++
            # A failed/missing batch response passes $null: the scriptblock then reads the owners itself
            $ownersResponse = $ownersResponses["$($samlApp.Id)"]
            # A response with more pages (@odata.nextLink) is read again in full by the scriptblock
            $preFetchedOwners = if ($ownersResponse -and $ownersResponse.status -eq 200 -and -not $ownersResponse.body.'@odata.nextLink') { @($ownersResponse.body.value) } else { $null }

            $results = & $processSamlApp $samlApp "Processing $appCounter/${samlApplicationsCount}: " $IncludeSignInStats.IsPresent $auditLogPermissionMissing $signInStartDate $preFetchedOwners
            foreach ($result in $results) {
                if ($result) { $samlApplicationsArray.Add($result) }
            }
        }
    }
    elseif ($useParallel) {
        Write-Verbose "Processing SAML applications in parallel (ThrottleLimit: $ThrottleLimit)..."
        $processText = $processSamlApp.ToString()
        $inclStats = $IncludeSignInStats.IsPresent
        $parallelResults = $samlApplications | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
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
        # Report footer version: resolved from the loaded PS365 module, with a hardcoded fallback for
        # standalone deployments (e.g. Azure Automation runbook) where the module is not available.
        # Keep the fallback in sync with ModuleVersion in PS365.psd1.
        $ps365Module = Get-Module -Name 'PS365' -ErrorAction SilentlyContinue
        $ps365Version = if ($ps365Module) { $ps365Module.Version } else { '0.5.5' }
        $ps365VersionDisplay = " (PS365 v$ps365Version)"

        # Human-readable execution duration for the report footer (e.g. '4 min 32 s' or '45 s')
        $executionElapsed = $executionStopwatch.Elapsed
        $executionTimeDisplay = if ($executionElapsed.TotalMinutes -ge 1) {
            '{0} min {1} s' -f [int][Math]::Floor($executionElapsed.TotalMinutes), $executionElapsed.Seconds
        }
        else {
            '{0} s' -f [int][Math]::Ceiling($executionElapsed.TotalSeconds)
        }

        # Only the certificates in use raise an alert. An inactive one (an old signing
        # certificate left on the app, an encryption certificate not selected, a verification
        # certificate while signed requests are not required) breaks nothing when it expires.
        $activeCertificates = @($samlApplicationsArray | Where-Object { $_.CertificateIsActive -ne $false })

        $expiringCertificates = $activeCertificates | Where-Object {
            $null -ne $_.CertificateExpiresInDays -and
            $_.CertificateExpiresInDays -le $ExpirationThresholdDays
        }

        # Calculate statistics for different expiration categories
        $halfThreshold = [int]($ExpirationThresholdDays / 2)
        $expiredCertificates = $activeCertificates | Where-Object {
            $null -ne $_.CertificateExpiresInDays -and $_.CertificateExpiresInDays -le 0
        }
        $certificatesExpiringCritical = $activeCertificates | Where-Object {
            $null -ne $_.CertificateExpiresInDays -and $_.CertificateExpiresInDays -le $halfThreshold -and $_.CertificateExpiresInDays -gt 0
        }
        $certificatesExpiringWarning = $activeCertificates | Where-Object {
            $null -ne $_.CertificateExpiresInDays -and $_.CertificateExpiresInDays -le $ExpirationThresholdDays -and $_.CertificateExpiresInDays -gt 0
        }

        Write-Verbose "Sending notification email. Found $($expiringCertificates.Count) SAML certificates expiring within $ExpirationThresholdDays days."

        $expiringCertificates = $expiringCertificates | Sort-Object CertificateExpiresInDays

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
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#7A3A00;font-weight:bold;">$($certificatesExpiringCritical.Count)</span>
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
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#003882;font-weight:bold;">$($certificatesExpiringWarning.Count)</span>
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
            <th>Certificate</th>
            <th>Expires In (Days)</th>
            <th>Expiry Date</th>
            <th>Login URL</th>
            <th>Owners</th>
            $(if ($IncludeSignInStats.IsPresent -or $auditLogPermissionMissing) { '<th>Sign-ins (30d)</th>' })
        </tr>
"@

        foreach ($cert in $expiringCertificates) {
            $rowClass = if ($cert.CertificateExpiresInDays -le 0) { 'critical' } elseif ($cert.CertificateExpiresInDays -le $halfThreshold) { 'warning' } else { 'caution' }
            $expiresInDaysDisplay = if ($cert.CertificateExpiresInDays -lt 0) { "$($cert.CertificateExpiresInDays) (already expired)" } else { $cert.CertificateExpiresInDays }
            $appLink = "<strong style=`"color:#11100f;font-size:12px;line-height:16px;`">$($cert.DisplayName)</strong> <a href=`"$($cert.EntraUrl)`" style=`"text-decoration:none;font-size:14px;line-height:16px;`" title=`"Open in Entra`">&#x1F517;</a>"
            $ownersList = @($cert.Owners -split '\|' | Where-Object { $_ })
            $ownersHtml = if ($ownersList.Count -gt 1) { ($ownersList | ForEach-Object { "- $_" }) -join '<br>' } else { $ownersList -join '' }
            $signInsHtml = if ($IncludeSignInStats.IsPresent -or $auditLogPermissionMissing) { "<td>$($cert.InteractiveSignInsLast30Days)</td>" } else { '' }
            $emailBody += "<tr class=`"$rowClass`"><td>$appLink</td><td>$($cert.AppId)</td><td>$($cert.CertificateType)</td><td><strong>$expiresInDaysDisplay</strong></td><td>$($cert.CertificateEndTime)</td><td>$($cert.LoginUrl)</td><td>$ownersHtml</td>$signInsHtml</tr>"
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
        <p><em>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Get-MgApplicationSAML$ps365VersionDisplay - execution time: $executionTimeDisplay</em></p>
    </div>
</body>
</html>
"@

        try {
            $params = @{
                message         = @{
                    subject      = $emailSubject
                    body         = @{
                        contentType = 'HTML'
                        content     = $emailBody
                    }
                    toRecipients = @(
                        @{
                            emailAddress = @{
                                address = $NotificationRecipient
                            }
                        }
                    )
                }
                saveToSentItems = $false
            }

            $mailParams = @{
                Method      = 'POST'
                Uri         = "/v1.0/users/$NotificationSender/sendMail"
                Body        = ($params | ConvertTo-Json -Depth 10)
                ContentType = 'application/json'
                ErrorAction = 'Stop'
            }
            $null = Invoke-MgGraphRequest @mailParams
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
        $samlApplicationsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraSAMLApplications' -TableStyle Light9
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif (-not $RunFromAzureAutomation.IsPresent) {
        return $samlApplicationsArray
    }
}
