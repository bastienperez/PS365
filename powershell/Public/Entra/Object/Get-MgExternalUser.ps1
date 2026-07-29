<#
    .SYNOPSIS
    Lists external users in Microsoft Entra ID based on their identities and invitation trace, not on userType.

    .DESCRIPTION
    Most scripts identify external users with userType -eq 'Guest'. This is unreliable:
    users onboarded through cross-tenant synchronization, or guests converted to members,
    keep userType 'Member' while still being backed by an account from another directory.

    Get-MgExternalUser returns every account that is external by any of these criteria:
    - it holds at least one identity with signInType 'federated' (the account authenticates elsewhere)
    - it was created through an invitation (creationType 'Invitation')
    - it carries an invitation state (externalUserState)
    - its userPrincipalName contains the #EXT# pattern
    - its userType is 'Guest'

    The invitation state is deliberately NOT part of the detection: a guest who never redeemed
    its invitation is reported like any other external account.

    Each result is classified through the AuthSource property, which tells whether the account
    relies on its own identity provider or exists only thanks to this tenant:

    - ExternalTenant    : issuer ExternalAzureAD - account from another Microsoft Entra ID tenant (B2B / cross-tenant sync)
    - MicrosoftAccount  : issuer MSA - personal Microsoft account
    - SocialIdP         : google.com, facebook.com, ... - social identity provider
    - ExternalIdP       : issuer is a third-party domain - direct federation (SAML / WS-Fed)
    - LocalOTP          : issuer 'mail' or one of this tenant's own domains - email one-time passcode,
                          the account has NO external identity provider and depends entirely on this tenant
    - PendingRedemption : no federated identity yet - the invitation was never redeemed, so Entra ID
                          has not determined which identity provider the account will use
    - Other             : anything else

    Accounts with their own identity provider are ExternalTenant, MicrosoftAccount, SocialIdP and
    ExternalIdP. Accounts depending on this tenant are LocalOTP. PendingRedemption is undetermined
    until the invitation is redeemed.

    Filtering is done client-side because Microsoft Graph does not support server-side
    filtering on identities/signInType.

    .PARAMETER Issuer
    (Optional) Restricts the results to one or more identity issuers (for example 'ExternalAzureAD').
    Because the filter applies to the federated identities, accounts classified PendingRedemption
    are excluded when this parameter is used: they do not carry an issuer yet.
    When omitted, every external account is returned whatever its issuer or invitation state.

    .PARAMETER UserType
    (Optional) Restricts the results to a given userType. Valid values: All, Member, Guest. Default is All.
    Use 'Member' to surface external accounts that are NOT flagged as guests - typically
    cross-tenant synchronized users, which a userType-based script would miss.

    .PARAMETER IncludeSignInActivity
    (Optional) Adds the last sign-in information (interactive, non-interactive and last successful sign-in)
    plus the number of days since the last interactive sign-in.
    Requires the AuditLog.Read.All permission and a Microsoft Entra ID P1 or P2 license.
    When the tenant does not meet those requirements, the report falls back to the standard
    properties and a warning is displayed.

    .PARAMETER ForceNewToken
    Switch parameter to force getting a new token from Microsoft Graph.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER ExportPath
    (Optional) Output directory for the Excel export. Defaults to the user profile.

    .EXAMPLE
    Get-MgExternalUser

    Returns every external account of the tenant, whatever its issuer, userType and invitation state.

    .EXAMPLE
    Get-MgExternalUser | Where-Object { $_.AuthSource -eq 'LocalOTP' }

    Returns the external accounts that have no external identity provider and exist only through
    this tenant (email one-time passcode).

    .EXAMPLE
    Get-MgExternalUser | Where-Object { $_.AuthSource -eq 'PendingRedemption' }

    Returns the invited accounts that never redeemed their invitation.

    .EXAMPLE
    Get-MgExternalUser -Issuer 'ExternalAzureAD' -UserType Member

    Returns external users from another tenant that are NOT flagged as guests, which is the
    typical footprint of cross-tenant synchronization.

    .EXAMPLE
    Get-MgExternalUser -IncludeSignInActivity | Where-Object { $_.DaysSinceLastSignIn -gt 90 }

    Returns external users that have not signed in interactively for more than 90 days.

    .EXAMPLE
    Get-MgExternalUser -ExportToExcel

    Exports the external users report to an Excel file in the user's profile directory.

    .OUTPUTS
    System.Collections.Generic.List[PSCustomObject]

    .NOTES
    OUTPUT PROPERTIES
    Identity        : DisplayName, UserPrincipalName, Mail, ExternalDomain
                      ExternalDomain is the domain the account really belongs to, derived from the UPN
                      (#EXT# pattern) or from the mail attribute.
    Origin          : UserType, CreationType, AuthSource, Issuer, SignInType, IssuerAssignedId
                      Issuer, SignInType and IssuerAssignedId come from the federated identities and
                      are empty for PendingRedemption accounts.
                      IssuerAssignedId is the identifier assigned by the issuer (typically the email
                      address for social or OTP accounts). It is usually empty for ExternalAzureAD.
    Account state   : AccountEnabled, InvitationState, InvitationStateChangeDateTime
                      InvitationState is the Graph 'externalUserState' property: PendingAcceptance until
                      the invitee redeems the invitation, then Accepted. It stays empty for accounts
                      that were never invited (cross-tenant sync, direct creation).
    Sign-in activity: LastSignInDateTime, LastNonInteractiveSignInDateTime, LastSuccessfulSignInDateTime,
                      DaysSinceLastSignIn (only with -IncludeSignInActivity)
    Metadata        : CompanyName, CreatedDateTime, OnPremisesSyncEnabled, Id

    Required Microsoft Graph permissions:
        - User.Read.All
        - Domain.Read.All (to recognize an issuer that is one of this tenant's own domains)
        - AuditLog.Read.All (only with -IncludeSignInActivity)

    An external account that redeemed its invitation carries two identities: one with
    signInType 'userPrincipalName' issued by this tenant, and one with signInType 'federated'
    issued by its identity provider. Only the federated ones drive the AuthSource classification.

    With -IncludeSignInActivity the page size is capped at 120 users per request: Microsoft Graph
    enforces this limit when signInActivity is part of the $select clause. The report is therefore
    slower on large tenants.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgExternalUser
#>

function Get-MgExternalUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Issuer,

        [Parameter(Mandatory = $false)]
        [ValidateSet('All', 'Member', 'Guest')]
        [string]$UserType = 'All',

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSignInActivity,

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

    $permissionsNeeded = @('User.Read.All', 'Domain.Read.All')
    if ($IncludeSignInActivity.IsPresent) {
        # signInActivity is only returned when the token carries AuditLog.Read.All
        $permissionsNeeded += 'AuditLog.Read.All'
    }

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

    # Tenant domains: an issuer matching one of them means the account has no external identity
    # provider (email one-time passcode), so it depends entirely on this tenant.
    Write-Host -ForegroundColor Cyan 'Retrieving tenant domains'
    [System.Collections.Generic.List[string]]$tenantDomains = @()
    $domainsUri = 'https://graph.microsoft.com/v1.0/domains?$select=id&$top=999'
    try {
        do {
            $domainsResponse = Invoke-MgGraphRequestWithRetry -Method GET -Uri $domainsUri
            foreach ($domain in $domainsResponse.value) {
                $tenantDomains.Add($domain.id)
            }
            $domainsUri = $domainsResponse.'@odata.nextLink'
        } while ($domainsUri)
    }
    catch {
        Write-Warning "Unable to retrieve the tenant domains: $_"
        Write-Warning 'Identities issued by one of this tenant domains will not be classified as LocalOTP.'
    }

    $socialIssuers = @('google.com', 'facebook.com', 'twitter.com', 'linkedin.com', 'github.com', 'apple.com', 'amazon.com')

    # identities is not returned by default and cannot be filtered server-side on signInType,
    # so the whole user set is retrieved and filtered client-side.
    $selectBaseProperties = 'id,displayName,userPrincipalName,mail,userType,creationType,accountEnabled,companyName,createdDateTime,externalUserState,externalUserStateChangeDateTime,onPremisesSyncEnabled,identities'

    # Not an alias of the switch: this one is turned off at runtime when the tenant cannot
    # serve signInActivity (missing Entra ID P1/P2 license), so the report degrades instead of failing.
    $signInActivityAvailable = $IncludeSignInActivity.IsPresent

    if ($IncludeSignInActivity.IsPresent) {
        # Graph caps the page size at 120 users per request when signInActivity is selected
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectBaseProperties,signInActivity&`$top=120"
    }
    else {
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectBaseProperties&`$top=999"
    }

    Write-Host -ForegroundColor Cyan 'Retrieving users from Microsoft Entra ID'

    [System.Collections.Generic.List[PSCustomObject]]$externalUsersArray = @()
    $processedCount = 0
    $referenceDate = Get-Date

    do {
        $response = $null
        try {
            $response = Invoke-MgGraphRequestWithRetry -Method GET -Uri $uri -ErrorAction Stop
        }
        catch {
            # signInActivity also requires an Entra ID P1/P2 license: degrade gracefully instead of failing
            if ($signInActivityAvailable -and ($processedCount -eq 0)) {
                Write-Warning "Unable to retrieve signInActivity (an Entra ID P1/P2 license and AuditLog.Read.All are required): $_"
                Write-Warning 'Falling back to the report without sign-in activity.'
                $signInActivityAvailable = $false
                $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectBaseProperties&`$top=999"
                continue
            }

            Write-Warning "Unable to retrieve users: $_"
            return
        }

        foreach ($user in $response.value) {
            $processedCount++

            if (-not (Test-MgUserIsExternal -User $user)) {
                continue
            }

            $userPrincipalName = $user.userPrincipalName
            $federatedIdentities = @($user.identities | Where-Object { $_.signInType -eq 'federated' })

            if ($Issuer) {
                # Filtering on the issuer necessarily drops the accounts that do not have one yet
                $federatedIdentities = @($federatedIdentities | Where-Object { $_.issuer -in $Issuer })
                if ($federatedIdentities.Count -eq 0) {
                    continue
                }
            }

            if (($UserType -ne 'All') -and ($user.userType -ne $UserType)) {
                continue
            }

            # Where the account really authenticates
            if ($federatedIdentities.Count -eq 0) {
                $authSource = 'PendingRedemption'
            }
            else {
                $authSourceList = foreach ($identity in $federatedIdentities) {
                    $issuerValue = $identity.issuer
                    if (-not $issuerValue) { 'Other' }
                    elseif ($issuerValue -eq 'ExternalAzureAD') { 'ExternalTenant' }
                    elseif ($issuerValue -eq 'MSA') { 'MicrosoftAccount' }
                    elseif ($issuerValue -eq 'mail') { 'LocalOTP' }
                    elseif ($issuerValue -in $tenantDomains) { 'LocalOTP' }
                    elseif ($issuerValue -in $socialIssuers) { 'SocialIdP' }
                    elseif ($issuerValue -match '\.') { 'ExternalIdP' }
                    else { 'Other' }
                }
                $authSource = (@($authSourceList) | Sort-Object -Unique) -join '; '
            }

            # Resolve the domain the account really belongs to
            $externalDomain = $null
            if ($userPrincipalName -match '#EXT#') {
                $externalPart = ($userPrincipalName -split '#EXT#')[0]
                if ($externalPart -match '^.+_(?<domain>[^_]+\.[^_]+)$') {
                    $externalDomain = $matches['domain']
                }
            }
            if ((-not $externalDomain) -and $user.mail) {
                $externalDomain = ($user.mail -split '@')[-1]
            }
            if ((-not $externalDomain) -and ($userPrincipalName -match '@')) {
                $externalDomain = ($userPrincipalName -split '@')[-1]
            }

            $issuerList = ($federatedIdentities.issuer | Sort-Object -Unique) -join '; '
            $signInTypeList = ($federatedIdentities.signInType | Sort-Object -Unique) -join '; '
            $issuerAssignedIdList = ($federatedIdentities | Where-Object { $_.issuerAssignedId } | ForEach-Object { $_.issuerAssignedId }) -join '; '

            # Who / where the account comes from / account state
            $properties = [ordered]@{
                DisplayName                   = $user.displayName
                UserPrincipalName             = $userPrincipalName
                Mail                          = $user.mail
                ExternalDomain                = $externalDomain
                UserType                      = $user.userType
                CreationType                  = $user.creationType
                AuthSource                    = $authSource
                Issuer                        = $issuerList
                SignInType                    = $signInTypeList
                IssuerAssignedId              = $issuerAssignedIdList
                AccountEnabled                = $user.accountEnabled
                InvitationState               = $user.externalUserState
                InvitationStateChangeDateTime = $user.externalUserStateChangeDateTime
            }

            if ($signInActivityAvailable) {
                $signInActivity = $user.signInActivity
                $lastSignInDateTime = $signInActivity.lastSignInDateTime

                $daysSinceLastSignIn = $null
                if ($lastSignInDateTime) {
                    try {
                        $lastSignInDate = if ($lastSignInDateTime -is [datetime]) {
                            $lastSignInDateTime
                        }
                        else {
                            [datetime]::Parse($lastSignInDateTime, [cultureinfo]::InvariantCulture)
                        }
                        $daysSinceLastSignIn = [int]($referenceDate - $lastSignInDate).TotalDays
                    }
                    catch {
                        Write-Verbose "Unable to parse lastSignInDateTime '$lastSignInDateTime' for $userPrincipalName"
                    }
                }

                $properties['LastSignInDateTime'] = $lastSignInDateTime
                $properties['LastNonInteractiveSignInDateTime'] = $signInActivity.lastNonInteractiveSignInDateTime
                $properties['LastSuccessfulSignInDateTime'] = $signInActivity.lastSuccessfulSignInDateTime
                $properties['DaysSinceLastSignIn'] = $daysSinceLastSignIn
            }

            # Metadata
            $properties['CompanyName'] = $user.companyName
            $properties['CreatedDateTime'] = $user.createdDateTime
            $properties['OnPremisesSyncEnabled'] = if ($null -eq $user.onPremisesSyncEnabled) { $false } else { $user.onPremisesSyncEnabled }
            $properties['Id'] = $user.id

            $externalUsersArray.Add([PSCustomObject]$properties)
        }

        $uri = $response.'@odata.nextLink'
        Write-Progress -Activity 'Retrieving users' -Status "$processedCount user(s) scanned - $($externalUsersArray.Count) external user(s) found"
    } while ($uri)

    Write-Progress -Activity 'Retrieving users' -Completed

    if ($externalUsersArray.Count -eq 0) {
        Write-Host -ForegroundColor Yellow "No external user found among the $processedCount user(s) scanned."
        return
    }

    Write-Host -ForegroundColor Green "Found $($externalUsersArray.Count) external user(s) among $processedCount user(s) scanned."

    Write-Host -ForegroundColor Cyan 'Breakdown by AuthSource:'
    foreach ($group in ($externalUsersArray | Group-Object AuthSource | Sort-Object Count -Descending)) {
        Write-Host "  $($group.Name): $($group.Count)"
    }

    $externalMembersCount = @($externalUsersArray | Where-Object { $_.UserType -eq 'Member' }).Count
    if ($externalMembersCount -gt 0) {
        Write-Host -ForegroundColor Yellow "$externalMembersCount external user(s) have userType 'Member' and would be missed by a userType-based script."
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
        $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-MgExternalUser.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting external users to Excel file: $excelFilePath"
        $externalUsersArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-ExternalUsers'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $externalUsersArray
    }
}
