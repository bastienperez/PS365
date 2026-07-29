<#
    .SYNOPSIS
    Tells whether a Microsoft Graph user object is an external account.

    .DESCRIPTION
    Shared detection logic used by Get-MgExternalUser and Get-MgExternalUserSummary, so both
    reports always cover exactly the same population.

    An account is external as soon as it authenticates somewhere else (at least one identity with
    signInType 'federated') OR carries a trace of invitation.

    The invitation state is deliberately NOT part of the test: a guest who never redeemed its
    invitation has no federated identity yet (Entra ID has not determined its identity provider),
    but it is an external account nonetheless.

    .PARAMETER User
    The user object as returned by Microsoft Graph. The following properties must be part of the
    $select clause: userPrincipalName, userType, creationType, externalUserState, identities.
#>
function Test-MgUserIsExternal {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNull()]
        $User
    )

    # The account authenticates through another identity provider
    if (@($User.identities | Where-Object { $_.signInType -eq 'federated' }).Count -gt 0) {
        return $true
    }

    # No federated identity, but something says the account did not originate from this tenant
    return ($User.creationType -eq 'Invitation') -or
           ([bool]$User.externalUserState) -or
           ($User.userType -eq 'Guest') -or
           ($User.userPrincipalName -match '#EXT#')
}
