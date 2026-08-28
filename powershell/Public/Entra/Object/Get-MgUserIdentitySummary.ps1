<#
    .SYNOPSIS
    Summarizes how every user of a Microsoft Entra ID tenant authenticates, internal and external alike.

    .DESCRIPTION
    Returns one row per distinct combination of (SignInType, Issuer, UserType, CreationType,
    InvitationState, IsExternal) with the number of identities matching it. It gives a one-page
    picture of the authentication landscape of the tenant.

    The granularity is the identity, not the user. A cloud-only internal user carries a single
    identity ('userPrincipalName' issued by this tenant). An external account that redeemed its
    invitation carries two: the local 'userPrincipalName' one plus a 'federated' one issued by its
    identity provider. An account that never redeemed its invitation only carries the local one,
    which is exactly what this view makes visible.

    Because internal and external accounts can share the very same (SignInType, Issuer) pair, the
    IsExternal column tells them apart. It uses the same detection as Get-MgExternalUser, so both
    reports always agree.

    .PARAMETER ExternalOnly
    (Optional) Restricts the summary to external accounts. Without it, every user of the tenant is
    included, which is the point of this function.

    .PARAMETER ForceNewToken
    Switch parameter to force getting a new token from Microsoft Graph.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER ExportPath
    (Optional) Output directory for the Excel export. Defaults to the user profile.

    .EXAMPLE
    Get-MgUserIdentitySummary | Format-Table -AutoSize

    Displays the identity summary of the whole tenant as a table. Format-Table is needed because the
    objects carry more than four properties, which PowerShell would otherwise render as a list.

    .EXAMPLE
    Get-MgUserIdentitySummary -ExternalOnly | Format-Table -AutoSize

    Same summary restricted to the external accounts.

    .EXAMPLE
    Get-MgUserIdentitySummary -ExportToExcel

    Exports the identity summary to an Excel file in the user's profile directory.

    .OUTPUTS
    System.Collections.Generic.List[PSCustomObject]

    .NOTES
    OUTPUT PROPERTIES
    - Count           : number of identities matching the combination
    - SignInType      : signInType of the identity ('userPrincipalName', 'federated', ...)
                        '<no identities>' when Graph returned no identity at all for the account
    - Issuer          : issuer of the identity (a tenant domain, ExternalAzureAD, MSA, mail, ...)
    - UserType        : userType of the account holding the identity (Member or Guest)
    - CreationType    : creationType of the account ('Invitation' when it comes from a B2B invitation)
    - InvitationState : Graph 'externalUserState' property (PendingAcceptance, Accepted, or empty)
    - IsExternal      : whether the account holding the identity is external (see Get-MgExternalUser)

    Required Microsoft Graph permissions:
        - User.Read.All

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgUserIdentitySummary
#>

function Get-MgUserIdentitySummary {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ExternalOnly,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath
    )

    try {
        $null = Import-Module 'Microsoft.Graph.Authentication' -ErrorAction Stop
    }
    catch {
        Write-Warning 'Please install Microsoft.Graph.Authentication first'
        return
    }

    $permissionsNeeded = @('User.Read.All')

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
    if ($ForceNewToken.IsPresent) {
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        $isConnected = $false
    }
    if (-not $isConnected) {
        Write-Host -ForegroundColor Cyan 'Connecting to Microsoft Graph'
        $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
    }

    if (-not (Test-MgGraphPermission -RequiredScopes $permissionsNeeded -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    $selectProperties = 'id,userPrincipalName,userType,creationType,externalUserState,identities'
    $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectProperties&`$top=999"

    Write-Host -ForegroundColor Cyan 'Retrieving users from Microsoft Entra ID'

    [System.Collections.Generic.List[PSCustomObject]]$identityRows = @()
    $processedCount = 0
    $externalCount = 0

    do {
        try {
            $response = Invoke-MgGraphRequestWithRetry -Method GET -Uri $uri -ErrorAction Stop
        }
        catch {
            Write-Warning "Unable to retrieve users: $_"
            return
        }

        foreach ($user in $response.value) {
            $processedCount++

            $isExternal = Test-MgUserIsExternal -User $user
            if ($isExternal) {
                $externalCount++
            }
            elseif ($ExternalOnly.IsPresent) {
                continue
            }

            # An account with no identity at all still deserves a row, otherwise it silently vanishes
            $identities = if ($user.identities) {
                $user.identities
            }
            else {
                @(@{ signInType = '<no identities>'; issuer = $null })
            }

            foreach ($identity in $identities) {
                $identityRows.Add([PSCustomObject][ordered]@{
                        SignInType      = $identity.signInType
                        Issuer          = $identity.issuer
                        UserType        = $user.userType
                        CreationType    = $user.creationType
                        InvitationState = $user.externalUserState
                        IsExternal      = $isExternal
                    })
            }
        }

        $uri = $response.'@odata.nextLink'
        Write-Progress -Activity 'Retrieving users' -Status "$processedCount user(s) scanned - $externalCount external user(s) found"
    } while ($uri)

    Write-Progress -Activity 'Retrieving users' -Completed

    if ($identityRows.Count -eq 0) {
        Write-Host -ForegroundColor Yellow "No user found among the $processedCount user(s) scanned."
        return
    }

    Write-Host -ForegroundColor Green "Scanned $processedCount user(s): $externalCount external, $($processedCount - $externalCount) internal, for $($identityRows.Count) identity(ies) summarized."

    [System.Collections.Generic.List[PSCustomObject]]$summaryArray = @()
    $groups = $identityRows |
        Group-Object SignInType, Issuer, UserType, CreationType, InvitationState, IsExternal |
        Sort-Object Count -Descending

    foreach ($group in $groups) {
        $sample = $group.Group[0]
        $summaryArray.Add([PSCustomObject][ordered]@{
                Count           = $group.Count
                SignInType      = $sample.SignInType
                Issuer          = $sample.Issuer
                UserType        = $sample.UserType
                CreationType    = $sample.CreationType
                InvitationState = $sample.InvitationState
                IsExternal      = $sample.IsExternal
            })
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
        $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-MgUserIdentitySummary.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting identity summary to Excel file: $excelFilePath"
        $summaryArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-IdentitySummary' -TableStyle Light9
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $summaryArray
    }
}
