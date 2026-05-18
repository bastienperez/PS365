<#
    .SYNOPSIS
    Restarts synchronization jobs for one or more Entra ID applications.

    .DESCRIPTION
    Resolves the target service principal(s) by ObjectID or DisplayName (wildcards supported),
    retrieves all associated synchronization jobs, then restarts them via the Microsoft Graph API.

    Two authentication modes are available:
    - Interactive  : prompts for sign-in with delegated permissions.
    - Managed Identity (RunFromAzureAutomation) : uses the managed identity assigned to the Azure resource (Automation Account, Function App, VM, etc.) - no credentials required.

    .PARAMETER ObjectID
    ObjectID (GUID) of the target service principal.
    Cannot be combined with -DisplayName.

    .PARAMETER DisplayName
    Display name of the target service principal(s).
    Supports wildcards (* and ?):
    - "Workday*"          matches all apps whose name starts with "Workday"
    - "*Provisioning*"    matches all apps whose name contains "Provisioning"
    - "Workday to Entra"  exact match (no wildcard)
    Cannot be combined with -ObjectID.

    .PARAMETER ForceNewToken
    Disconnects the existing Microsoft Graph session before connecting, forcing a fresh
    token to be acquired. Useful when the current token has expired or lacks required scopes.

    .PARAMETER RunFromAzureAutomation
    Authenticates using the Managed Identity of the Azure resource.
    Use this switch when the function runs inside:
    - Azure Automation runbooks
    - Azure Functions
    - Azure VMs / Container Apps with a managed identity enabled

    When active, authentication uses the Managed Identity (Out-GridView is unavailable).

    .PARAMETER ResetScope
    Comma-separated combination of synchronizationJobRestartScope values controlling what is reset when
    the job restarts. Values can be combined (e.g. "Escrows, QuarantineState").

    Supported values:
    `None` Starts a paused or quarantined provisioning job. DO NOT USE. Use the Start synchronizationJob API instead.
    `ConnectorDataStore` Clears the underlying cache for all users. DO NOT USE. Contact Microsoft Support for guidance.
    `Escrows` Provisioning failures are marked as escrows and retried. Clearing escrows will stop the service from retrying failures.
    `Watermark` Removing the watermark causes the service to re-evaluate all the users again, rather than just processing changes.
    `QuarantineState` Temporarily lifts the quarantine.
    `Full` Use this if you want all of the options (Escrows + Watermark + QuarantineState).
    `ForceDeletes` Forces the system to delete the pending deleted users when using the accidental deletions prevention feature and the deletion threshold is exceeded.

    An empty string ("") emulates the "Restart provisioning" button in the Microsoft Entra admin
    center. It is similar to setting resetScope to include QuarantineState, Watermark, and Escrows,
    and meets most customer needs. If you use `RunFromAzureAutomation`, you need to explicitly choose a reset scope, as the portal default relies on delegated permissions which are not available with Managed Identity authentication.

    Default : 'Escrows'  (most common use case - retry errored objects without reprocessing
    the entire directory).

    Reference: https://learn.microsoft.com/en-us/graph/api/resources/synchronization-synchronizationjobrestartcriteria

    .EXAMPLE
    Restart-MgSynchronizationJob -ObjectID 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'

    Restarts all synchronization jobs for the application identified by its ObjectID.

    .EXAMPLE
    Restart-MgSynchronizationJob -DisplayName 'Workday to Entra'

    Restarts all synchronization jobs for the application named exactly 'Workday to Entra'.

    .EXAMPLE
    Restart-MgSynchronizationJob -DisplayName '*Provisioning*'

    Restarts all jobs for every application whose name contains 'Provisioning'.

    .EXAMPLE
    Restart-MgSynchronizationJob -DisplayName 'Workday to Entra' -RunFromAzureAutomation

    Restarts all jobs for the matching application using Managed Identity authentication.
    Suitable for use in Azure Automation runbooks.

    .EXAMPLE
    Restart-MgSynchronizationJob -DisplayName 'Workday*' -ForceNewToken

    Forces a new Graph token before resolving applications and restarting jobs.

    .EXAMPLE
    Restart-MgSynchronizationJob -DisplayName 'Workday to Entra' -ResetScope 'Escrows'

    Retries only the objects currently in escrow (errored), without reprocessing the full directory.

    .EXAMPLE
    Restart-MgSynchronizationJob -DisplayName 'Workday to Entra' -ResetScope 'Watermark, Escrows'

    Combines a full directory re-evaluation with an escrow retry.

    .EXAMPLE
    Restart-MgSynchronizationJob -DisplayName 'Workday to Entra' -ResetScope ''

    Emulates the "Restart provisioning" button in the Entra portal
    (equivalent to QuarantineState + Watermark + Escrows with no explicit criteria body).

    .NOTES
    In the logs, you can filter on `configurationCategory`: `ProvisioningManagement` and `Activity Type`: `Enable/restart provisioning`.

    Required Microsoft Graph permissions:
    - Application.Read.All
    - Synchronization.ReadWrite.All

    Required PowerShell modules:
    - Microsoft.Graph.Authentication
    - Microsoft.Graph.Applications

    API version used: v1.0

    In the log, you can configurationCategory: ProvisioningManagement and Activity Type: Enable/restart provisioning/

    .LINK
    https://learn.microsoft.com/en-us/graph/api/synchronization-synchronizationjob-restart
#>
function Restart-MgSynchronizationJob {
    [CmdletBinding(DefaultParameterSetName = 'ByDisplayName', SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByObjectId')]
        [string]$ObjectID,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByDisplayName')]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$RunFromAzureAutomation,

        # Scope of the reset applied when restarting the synchronization job.
        # Accepted values (combinable with commas):
        #   None                - DO NOT USE. Starts a paused or quarantined provisioning job. Use the Start synchronizationJob API instead.
        #   ConnectorDataStore  - DO NOT USE. Clears the underlying cache for all users. Contact Microsoft Support for guidance
        #
        #   Escrows         - Provisioning failures are marked as escrows and retried. Clearing escrows will stop the service from retrying failures.
        #   Watermark       - Removing the watermark causes the service to re-evaluate all the users again, rather than just processing changes.
        #   QuarantineState - Temporarily lifts the quarantine.
        #   Full            - Use this if you want all of the options (Escrows + Watermark + QuarantineState).
        #   ForceDeletes    - Forces the system to delete the pending deleted users when using the accidental deletions prevention feature and the deletion threshold is exceeded.
        # Empty string emulates the Entra portal "Restart provisioning" button
        #   (equivalent to QuarantineState + Watermark + Escrows with no explicit criteria body).
        #   Note: if you use -RunFromAzureAutomation, explicitly choose a reset scope - the portal default relies on delegated permissions unavailable with Managed Identity authentication.
        # Reference: https://learn.microsoft.com/en-us/graph/api/resources/synchronization-synchronizationjobrestartcriteria?view=graph-rest-1.0
        [Parameter(Mandatory = $false)]
        [ValidateScript({
            $validValues = @('Escrows', 'Watermark', 'QuarantineState', 'Full', 'ForceDeletes')
            if ([string]::IsNullOrWhiteSpace($_)) { return $true }
            $parts = $_ -split '\s*,\s*'
            foreach ($part in $parts) {
                if ($part -notin $validValues) {
                    throw "Invalid ResetScope value: '$part'. Accepted values: $($validValues -join ', '), or an empty string."
                }
            }
            return $true
        })]
        [string]$ResetScope = 'Escrows'
    )

    # In Azure Automation, Write-Host colors are useless and Write-Output is the conventional stream for status messages.
    function Write-StatusMessage {
        param(
            [Parameter(Mandatory = $true, Position = 0)]
            [string]$Message,

            [Parameter(Mandatory = $false, Position = 1)]
            [string]$Color = 'Cyan'
        )

        if ($RunFromAzureAutomation) {
            Write-Output $Message
        }
        else {
            Write-Host $Message -ForegroundColor $Color
        }
    }

    # --- ResetScope / RunFromAzureAutomation compatibility check ---
    if ($RunFromAzureAutomation -and [string]::IsNullOrWhiteSpace($ResetScope)) {
        Write-Error "When using -RunFromAzureAutomation, -ResetScope cannot be empty. The portal default relies on delegated permissions unavailable with Managed Identity. Specify an explicit value (e.g. 'Escrows', 'Full')."
        return
    }

    # --- Authentication ---
    if ($ForceNewToken) {
        if (Get-MgContext -ErrorAction SilentlyContinue) {
            Write-StatusMessage '[i] Existing Microsoft Graph connection detected. Disconnecting to force new token...' 'Yellow'
            $null = Disconnect-MgGraph
        }
    }

    $permissionsNeeded = @('Application.Read.All', 'Synchronization.ReadWrite.All')
    Write-Verbose "Required permissions: $($permissionsNeeded -join ', ')"

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        if ($RunFromAzureAutomation) {
            Write-Verbose 'Connecting to Microsoft Graph using Managed Identity.'
            $null = Connect-MgGraph -Identity -NoWelcome
        }
        else {
            $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
        }
    }

    # --- Permission check ---
    # Synchronization.ReadWrite.All is non-blocking when running as Managed Identity:
    # the managed identity may be an owner of the target service principal, which grants
    # implicit restart rights even without the Synchronization.ReadWrite.All scope.
    $currentScopes = (Get-MgContext).Scopes
    Write-Verbose "Current token scopes: $($currentScopes -join ', ')"
    $missingPermissions = $permissionsNeeded | Where-Object { $_ -notin $currentScopes }
    if ($missingPermissions) {
        if ($RunFromAzureAutomation) {
            $criticalMissing = $missingPermissions | Where-Object { $_ -ne 'Synchronization.ReadWrite.All' }
            if ($criticalMissing) {
                Write-Error "The managed identity is missing the following required Graph permissions: $($criticalMissing -join ', '). Please grant these permissions to the managed identity and try again."
                return
            }
            if ($missingPermissions -contains 'Synchronization.ReadWrite.All') {
                Write-Warning "The managed identity is missing 'Synchronization.ReadWrite.All'. The script will continue, restart may still succeed if the managed identity is an owner of the target application."
            }
        }
        else {
            Write-Error "The current token is missing the following required Graph permissions: $($missingPermissions -join ', '). Run with -ForceNewToken to reconnect with the correct scopes."
            return
        }
    }
    else {
        Write-Verbose 'All required permissions are present.'
    }

    # --- Resolve service principal(s) ---
    Write-Verbose "Parameter set: $($PSCmdlet.ParameterSetName)"
    [System.Collections.Generic.List[PSCustomObject]]$servicePrincipals = @()

    if ($PSCmdlet.ParameterSetName -eq 'ByObjectId') {
        Write-Verbose "Resolving service principal by ObjectID: $ObjectID"
        $sp = Get-MgServicePrincipal -ServicePrincipalId $ObjectID -Property DisplayName, Id -ErrorAction SilentlyContinue
        if ($sp) {
            $servicePrincipals.Add([PSCustomObject]@{ DisplayName = $sp.DisplayName; Id = $sp.Id })
        }
        else {
            Write-Error "[X] No service principal found with ObjectID '$ObjectID'."
            return
        }
    }
    else {
        if ($DisplayName -match '[*?]') {
            # Wildcard mode: server-side pre-filter with a keyword, then client-side -like
            $searchKeyword = ($DisplayName -replace '[*?]', '').Trim()
            Write-Verbose "Wildcard detected. Using Graph search with keyword '$searchKeyword', then filtering with -like '$DisplayName'."

            if ($searchKeyword) {
                $searchParams = @{
                    Method     = 'GET'
                    Uri        = "https://graph.microsoft.com/v1.0/servicePrincipals?`$search=`"displayName:$searchKeyword`"&`$select=displayName,id&`$top=999"
                    Headers    = @{ ConsistencyLevel = 'eventual' }
                    OutputType = 'PSObject'
                }
                $result = Invoke-MgGraphRequest @searchParams
            }
            else {
                $result = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=displayName,id&$top=999' -OutputType PSObject
            }

            $candidates = if ($result.Value) { $result.Value } else { @() }
            Write-Verbose "Server-side pre-filter returned $($candidates.Count) candidate(s)."
            $matches_ = $candidates | Where-Object { $_.displayName -like $DisplayName }
            Write-Verbose "Client-side -like filter matched $($matches_.Count) service principal(s)."
            foreach ($item in $matches_) {
                $servicePrincipals.Add([PSCustomObject]@{ DisplayName = $item.displayName; Id = $item.id })
            }

            Write-Verbose "Found $($servicePrincipals.Count) service principal(s) matching '$DisplayName'."
        }
        else {
            # Exact match
            $escaped = $DisplayName -replace "'", "''"
            $exactMatches = Get-MgServicePrincipal -Filter "displayName eq '$escaped'" -All -Property DisplayName, Id
            Write-Verbose "Exact filter returned $(@($exactMatches).Count) result(s) for '$DisplayName'."

            if (-not $exactMatches) {
                # Fallback: trimmed display name comparison via client-side filter
                Write-Verbose "No exact match for '$DisplayName'. Attempting trimmed client-side comparison."
                $allSps = Get-MgServicePrincipal -All -Property DisplayName, Id
                $exactMatches = $allSps | Where-Object { $_.DisplayName.Trim() -eq $DisplayName.Trim() }
                Write-Verbose "Trimmed client-side comparison matched $(@($exactMatches).Count) result(s)."
            }

            foreach ($sp in $exactMatches) {
                $servicePrincipals.Add([PSCustomObject]@{ DisplayName = $sp.DisplayName; Id = $sp.Id })
            }
        }

        if ($servicePrincipals.Count -eq 0) {
            Write-Error "[X] No service principal found matching DisplayName '$DisplayName'."
            return
        }
    }

    Write-StatusMessage "[i] $($servicePrincipals.Count) service principal(s) resolved." 'Cyan'

    # --- Collect synchronization jobs ---
    [System.Collections.Generic.List[PSCustomObject]]$allSyncJobs = @()

    foreach ($sp in $servicePrincipals) {
        Write-StatusMessage "[...] Retrieving sync jobs for: $($sp.DisplayName)" 'Cyan'

        $jobsResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/synchronization/jobs" -OutputType PSObject
        $syncJobs = if ($jobsResponse.Value) { $jobsResponse.Value } else { @() }
        Write-Verbose "Found $($syncJobs.Count) sync job(s) for '$($sp.DisplayName)'."

        if (-not $syncJobs -or $syncJobs.Count -eq 0) {
            Write-Warning "[*] No synchronization jobs found for '$($sp.DisplayName)'."
            continue
        }

        foreach ($job in $syncJobs) {
            Write-Verbose "  Job: $($job.id) | Template: $($job.templateId) | Schedule: $(if ($job.schedule) { $job.schedule.state } else { 'N/A' }) | Status: $(if ($job.status) { $job.status.code } else { 'N/A' })"
            $allSyncJobs.Add([PSCustomObject][ordered]@{
                ServicePrincipalId   = $sp.Id
                ServicePrincipalName = $sp.DisplayName
                SyncJobId            = $job.id
                TemplateId           = $job.templateId
                ScheduleState        = if ($job.schedule) { $job.schedule.state } else { 'N/A' }
                StatusCode           = if ($job.status) { $job.status.code } else { 'N/A' }
            })
        }
    }

    if ($allSyncJobs.Count -eq 0) {
        Write-StatusMessage '[i] No synchronization jobs to restart.' 'Yellow'
        return
    }

    # --- Determine which jobs to restart ---
    [System.Collections.Generic.List[PSCustomObject]]$jobsToRestart = @()

    foreach ($job in $allSyncJobs) {
        $jobsToRestart.Add($job)
    }
    Write-StatusMessage "[i] Restarting all $($jobsToRestart.Count) synchronization job(s)..." 'Yellow'

    # --- Restart jobs ---
    $successCount = 0
    $failureCount = 0

    foreach ($job in $jobsToRestart) {
        $restartUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($job.ServicePrincipalId)/synchronization/jobs/$($job.SyncJobId)/restart"
        Write-Verbose "POST $restartUri"
        Write-StatusMessage "[...] Restarting: $($job.ServicePrincipalName) / $($job.SyncJobId)" 'Cyan'

        if ($PSCmdlet.ShouldProcess("$($job.ServicePrincipalName) / $($job.SyncJobId)", 'Restart synchronization job')) {
            # Build request body based on ResetScope parameter
            $body = @{ criteria = @{} }
            if (-not [string]::IsNullOrWhiteSpace($ResetScope)) {
                $body.criteria.resetScope = $ResetScope
            }
            
            $restartBodyJson = $body | ConvertTo-Json -Depth 3 -Compress
            Write-Verbose "  Request body: $restartBodyJson"

            try {
                $statusCode = $null
                $null = Invoke-MgGraphRequest -Method POST -Uri $restartUri -Body $restartBodyJson -ContentType 'application/json' -ErrorAction Stop -StatusCodeVariable statusCode -Verbose
                $scopeLabel = if ([string]::IsNullOrWhiteSpace($ResetScope)) { '(portal default)' } else { $ResetScope }

                # A successful restart returns HTTP 204 No Content (empty body).
                if ($statusCode -eq 204) {
                    Write-StatusMessage "  [OK] Restart initiated successfully. (ResetScope: $scopeLabel)" 'Green'
                    $successCount++
                }
                else {
                    Write-StatusMessage "  [X] Unexpected response: HTTP $statusCode (expected 204 No Content). (ResetScope: $scopeLabel)" 'Red'
                    $failureCount++
                }
            }
            catch {
                Write-StatusMessage "  [X] Restart failed: $($_.Exception.Message)" 'Red'
                $failureCount++
            }
        }
    }

    Write-StatusMessage "[i] Done. Success: $successCount | Failed: $failureCount" 'Cyan'
    return
}