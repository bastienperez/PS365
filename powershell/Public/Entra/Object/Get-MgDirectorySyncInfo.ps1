<#
    .SYNOPSIS
    Reports the identities that can synchronize Active Directory to Entra ID: the Connect Sync applications and the legacy synchronization accounts.

    .DESCRIPTION
    Microsoft Entra Connect Sync used to authenticate with a directory synchronization account, the Sync_<server>_<id>
    user holding the Directory Synchronization Accounts role. It now authenticates with an application identity: a
    certificate on an application that carries the ADSynchronization.ReadWrite.All application role on the first-party
    Microsoft Entra AD Synchronization Service (AppId 6bf85cfa-ac8a-4be5-b5de-425a0d0dc016).

    Moving to the application does not remove the account. Microsoft documents that removal as a separate manual step,
    so a tenant that has migrated usually still carries the old account, still in a privileged role, used by nobody.

    This function reports both mechanisms in one list, because it is their coexistence that has to be read:
    - which application actually synchronizes today, and when its certificate expires. An expired certificate stops
      synchronization entirely, so this is an availability matter as much as a security one;
    - which synchronization accounts remain, and whether they still sign in. An account that has not signed in since
      the migration is the leftover to remove.

    A tenant is limited to 20 synchronization accounts, and leftovers count against that limit.

    .PARAMETER DaysUntilExpiry
    Number of days below which a certificate is reported as expiring. Default 30.

    .PARAMETER StaleAfterDays
    Number of days without a sign-in above which a synchronization account is reported as stale. Default 30.

    .PARAMETER ForceNewToken
    Forces a new token to be requested from Microsoft Graph.

    .PARAMETER ExportToExcel
    Exports the result to an Excel file in the user's profile directory instead of returning it.

    .PARAMETER ExportPath
    Optional output directory for the Excel export. Defaults to the user profile.

    .EXAMPLE
    Get-MgDirectorySyncInfo

    Lists the synchronization applications and the synchronization accounts, with their status.

    .EXAMPLE
    Get-MgDirectorySyncInfo -DaysUntilExpiry 60

    Widens the certificate warning to 60 days, to catch the next rollover before the maintenance window closes.

    .EXAMPLE
    Get-MgDirectorySyncInfo -ExportToExcel

    Exports the report to an Excel file in the user's profile directory.

    .NOTES
    Required Microsoft Graph permissions:
        - Application.Read.All
        - Directory.Read.All

    Optional:
        - AuditLog.Read.All

    AuditLog.Read.All is only needed for the last sign-in of the synchronization accounts, so it is requested but not
    required. Without it the accounts are still listed, with the Unknown status rather than presented as stale: an
    account that cannot be shown to be idle must not be offered for deletion.

    .LINK
    https://learn.microsoft.com/entra/identity/hybrid/connect/authenticate-application-id

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgDirectorySyncInfo
#>

function Get-MgDirectorySyncInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$DaysUntilExpiry = 30,

        [Parameter(Mandatory = $false)]
        [int]$StaleAfterDays = 30,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath
    )

    $modules = @(
        'Microsoft.Graph.Authentication'
    )

    foreach ($module in $modules) {
        try {
            $null = Import-Module $module -ErrorAction Stop
        }
        catch {
            Write-Warning "Please install $module first"
            return
        }
    }

    # AuditLog.Read.All is requested but not required: it only adds the last sign-in of the
    # synchronization accounts. Requiring it would deny the whole report to a caller who can
    # perfectly well list the identities, which is the part that matters.
    $permissionsNeeded = @(
        'Application.Read.All'
        'Directory.Read.All'
    )
    $optionalPermissions = @(
        'AuditLog.Read.All'
    )

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
    if ($ForceNewToken.IsPresent) {
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        $isConnected = $false
    }
    if (-not $isConnected) {
        Write-Host -ForegroundColor Cyan 'Connecting to Microsoft Graph'
        $null = Connect-MgGraph -Scopes ($permissionsNeeded + $optionalPermissions) -NoWelcome
    }

    if (-not (Test-MgGraphPermission -RequiredScopes $permissionsNeeded -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    # First-party application that Connect Sync authenticates against, same identifier in every cloud.
    $syncServiceAppId = '6bf85cfa-ac8a-4be5-b5de-425a0d0dc016'
    $syncRoleTemplateId = 'd29b2b05-8046-44ba-8758-1e26182fcf32'

    function Get-GraphCollection {
        param(
            [Parameter(Mandatory = $true)] [string]$Uri
        )

        $items = [System.Collections.Generic.List[PSCustomObject]]::new()
        $next = $Uri

        # Through the retry wrapper: a throttled page would otherwise land in a catch block that
        # continues with partial discovery, which is exactly the silence this report must not have.
        do {
            $response = Invoke-MgGraphRequestWithRetry -Method GET -Uri $next
            foreach ($item in $response.value) { $items.Add([PSCustomObject]$item) }
            $next = $response.'@odata.nextLink'
        } while ($next)

        return $items
    }

    $resultsArray = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Tenant state first: without it the rest cannot be read. A tenant that no longer synchronizes
    # should have neither a synchronization application nor a synchronization account left.
    $syncEnabled = $null
    $lastSyncDate = $null
    $discoveryIsComplete = $true
    try {
        $organization = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime'
        $syncEnabled = $organization[0].onPremisesSyncEnabled
        $lastSyncDate = $organization[0].onPremisesLastSyncDateTime
    }
    catch {
        $discoveryIsComplete = $false
        Write-Warning "Unable to read the tenant synchronization state: $_"
    }

    Write-Host -ForegroundColor Cyan 'Retrieving the synchronization applications'
    $assignments = @()
    try {
        $syncServicePrincipals = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$syncServiceAppId'&`$select=id,appId,displayName,appRoles"
        if ($syncServicePrincipals.Count -eq 0) {
            # Expected on a cloud-only tenant, and on a tenant still on a Connect version predating
            # application-based authentication: the first-party application is not provisioned there.
            Write-Host -ForegroundColor Yellow 'The Microsoft Entra AD Synchronization Service application is not present in this tenant, so no application-based synchronization is configured.'
        }
        else {
            $assignments = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($syncServicePrincipals[0].id)/appRoleAssignedTo"
        }
    }
    catch {
        $discoveryIsComplete = $false
        Write-Warning "Unable to read the synchronization application assignments: $_"
    }

    foreach ($assignment in $assignments) {
        $servicePrincipal = $null
        $credentialsReadable = $true
        try {
            $servicePrincipal = [PSCustomObject](Invoke-MgGraphRequestWithRetry -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($assignment.principalId)?`$select=id,appId,displayName,accountEnabled,createdDateTime,keyCredentials")
        }
        catch {
            # Without the service principal there is no way to reach the credentials, so the
            # certificate state is unknown. Reporting it as NoCredential would read as a broken
            # synchronization that is not broken.
            $credentialsReadable = $false
            Write-Warning "Unable to read the service principal '$($assignment.principalDisplayName)': $_"
        }

        # The certificate is uploaded on the application registration; the service principal carries a
        # copy. The application is read first because it is the object an administrator has to fix.
        $application = $null
        if ($servicePrincipal -and $servicePrincipal.appId) {
            try {
                $applications = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$($servicePrincipal.appId)'&`$select=id,appId,displayName,createdDateTime,keyCredentials"
                $application = $applications | Select-Object -First 1
            }
            catch {
                $credentialsReadable = $false
                Write-Warning "Unable to read the application '$($assignment.principalDisplayName)': $_"
            }
        }

        $credentials = @(if ($application) { $application.keyCredentials } elseif ($servicePrincipal) { $servicePrincipal.keyCredentials })
        $latestExpiry = $credentials | Where-Object { $_.endDateTime } | Sort-Object { [datetime]$_.endDateTime } -Descending | Select-Object -First 1
        $daysLeft = if ($latestExpiry) { [math]::Floor(([datetime]$latestExpiry.endDateTime - (Get-Date)).TotalDays) } else { $null }

        $status = if (-not $credentialsReadable) { 'Unknown' }
        elseif (-not $credentials -or $credentials.Count -eq 0) { 'NoCredential' }
        elseif ($null -eq $daysLeft) { 'Unknown' }
        elseif ($daysLeft -lt 0) { 'CertificateExpired' }
        elseif ($daysLeft -le $DaysUntilExpiry) { 'CertificateExpiring' }
        else { 'Active' }

        $resultsArray.Add([PSCustomObject][ordered]@{
                IdentityType        = 'Application'
                DisplayName         = $assignment.principalDisplayName
                Identifier          = if ($servicePrincipal) { $servicePrincipal.appId } else { $assignment.principalId }
                ObjectId            = $assignment.principalId
                Enabled             = if ($servicePrincipal) { $servicePrincipal.accountEnabled } else { $null }
                CreatedDateTime     = if ($application) { $application.createdDateTime } elseif ($servicePrincipal) { $servicePrincipal.createdDateTime } else { $null }
                LastSignIn          = $null
                DaysSinceLastSignIn = $null
                CertificateExpiry   = if ($latestExpiry) { $latestExpiry.endDateTime } else { $null }
                CertificateDaysLeft = $daysLeft
                CertificateCount    = @($credentials).Count
                AssignedThrough     = 'ADSynchronization application role'
                Status              = $status
            })
    }

    Write-Host -ForegroundColor Cyan 'Retrieving the synchronization accounts'
    $roleMembers = @()
    try {
        $roleMembers = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/directoryRoles(roleTemplateId='$syncRoleTemplateId')/members"
    }
    catch {
        # The role is only instantiated once a synchronization account exists, so a 404 here means
        # there is none, which is the expected state after a completed migration. Any other failure
        # is a read that did not happen and must not be presented as an empty result.
        if ($_.Exception.Message -match '404|not found|Resource .* does not exist') {
            Write-Host -ForegroundColor Green 'No synchronization account holds the Directory Synchronization Accounts role in this tenant.'
        }
        else {
            $discoveryIsComplete = $false
            Write-Warning "Unable to read the Directory Synchronization Accounts role members: $_"
        }
    }

    foreach ($member in $roleMembers) {
        $user = $null
        $signInReadable = $true
        try {
            $user = [PSCustomObject](Invoke-MgGraphRequestWithRetry -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($member.id)?`$select=id,displayName,userPrincipalName,accountEnabled,createdDateTime,signInActivity")
        }
        catch {
            # Losing signInActivity must not turn into a wrong verdict: the account is reported as
            # Unknown rather than stale, since deleting a synchronization account stops the sync.
            $signInReadable = $false
            Write-Warning "Unable to read the sign-in activity of '$($member.userPrincipalName)': $_"
        }

        $lastSignIn = if ($user) {
            @($user.signInActivity.lastSignInDateTime, $user.signInActivity.lastNonInteractiveSignInDateTime) |
                Where-Object { $_ } | Sort-Object { [datetime]$_ } -Descending | Select-Object -First 1
        }
        $daysSince = if ($lastSignIn) { [math]::Floor(((Get-Date) - [datetime]$lastSignIn).TotalDays) } else { $null }

        $status = if (-not $signInReadable) { 'Unknown' }
        elseif ($null -eq $daysSince) { 'LegacyAccountStale' }
        elseif ($daysSince -gt $StaleAfterDays) { 'LegacyAccountStale' }
        else { 'LegacyAccountActive' }

        $resultsArray.Add([PSCustomObject][ordered]@{
                IdentityType        = 'SyncAccount'
                DisplayName         = if ($user) { $user.displayName } else { $member.displayName }
                Identifier          = if ($user) { $user.userPrincipalName } else { $member.userPrincipalName }
                ObjectId            = $member.id
                Enabled             = if ($user) { $user.accountEnabled } else { $null }
                CreatedDateTime     = if ($user) { $user.createdDateTime } else { $null }
                LastSignIn          = $lastSignIn
                DaysSinceLastSignIn = $daysSince
                CertificateExpiry   = $null
                CertificateDaysLeft = $null
                CertificateCount    = $null
                AssignedThrough     = 'Directory Synchronization Accounts role'
                Status              = $status
            })
    }

    if ($resultsArray.Count -eq 0) {
        if (-not $discoveryIsComplete) {
            Write-Warning 'No synchronization identity was found, but at least one read failed, so the tenant cannot be called cloud-only on this result. Review the warnings above and rerun.'
        }
        elseif ($syncEnabled -eq $true) {
            Write-Warning 'Directory synchronization is enabled on this tenant but no synchronization identity was found, which means synchronization is running on an identity this report cannot see, or is broken.'
        }
        else {
            Write-Host -ForegroundColor Green 'No synchronization identity found, which matches a cloud-only tenant.'
        }
        return
    }

    if (-not $discoveryIsComplete) {
        Write-Warning 'At least one read failed, so this list may be incomplete. An identity missing from it is not proof that it does not exist.'
    }

    $syncStateLabel = if ($syncEnabled -eq $true) { "enabled, last sync $lastSyncDate" } elseif ($null -eq $syncEnabled) { 'unknown' } else { 'disabled' }
    Write-Host -ForegroundColor Cyan "Directory synchronization is $syncStateLabel"

    $accounts = @($resultsArray | Where-Object { $_.IdentityType -eq 'SyncAccount' })
    $stale = @($accounts | Where-Object { $_.Status -eq 'LegacyAccountStale' })
    $expiring = @($resultsArray | Where-Object { $_.Status -in @('CertificateExpiring', 'CertificateExpired') })

    # 20 synchronization accounts per tenant is a hard limit, and leftovers count against it: a
    # tenant that reaches it cannot install another synchronization server.
    Write-Host -ForegroundColor $(if ($accounts.Count -ge 20) { 'Red' } elseif ($accounts.Count -gt 0) { 'Yellow' } else { 'Green' }) "$($accounts.Count) synchronization account(s) out of the 20 allowed per tenant"

    if ($stale.Count -gt 0) {
        Write-Host -ForegroundColor Yellow "$($stale.Count) synchronization account(s) show no sign-in for more than $StaleAfterDays day(s). Confirm the migration to application-based authentication before removing them."
    }
    if ($expiring.Count -gt 0) {
        Write-Host -ForegroundColor Red "$($expiring.Count) synchronization application certificate(s) expire within $DaysUntilExpiry day(s). Synchronization stops when the last one expires."
    }

    $resultsArray = $resultsArray | Sort-Object IdentityType, DisplayName

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$(if ($ExportPath) { $ExportPath } else { $env:userprofile })\$now-DirectorySyncInfo.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $resultsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-DirectorySync' -TableStyle Light9
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $resultsArray
    }
}
