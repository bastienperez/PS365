<#
    .SYNOPSIS
    Verifies that the current Microsoft Graph token has the required scopes.

    .DESCRIPTION
    Reads the scopes from Get-MgContext and checks that every requirement listed in
    -RequiredScopes is present. A stronger scope satisfies a weaker requirement
    following the hierarchy ReadBasic < Read < ReadWrite (e.g. Application.ReadWrite.All
    covers Application.Read.All, and BitlockerKey.Read.All covers BitlockerKey.ReadBasic.All).

    A requirement may list alternatives separated by '|'
    (e.g. 'Policy.Read.All|Policy.ReadWrite.MobilityManagement'): the requirement is
    satisfied when any one of the alternatives is granted, mirroring the permission
    lists of the Microsoft Graph documentation.

    Returns $true when every requirement is satisfied, otherwise writes a
    warning listing the missing scopes and returns $false.

    .PARAMETER RequiredScopes
    One or more Microsoft Graph scope requirements that the calling function needs.
    Each entry is either a single scope or several alternatives separated by '|'.

    .PARAMETER CallerName
    (Optional) Name of the caller, used to make the warning message clearer.
#>
function Test-MgGraphPermission {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$RequiredScopes,

        [Parameter(Mandatory = $false)]
        [string]$CallerName = (Get-PSCallStack)[1].Command
    )

    # Scopes suggested in warning messages: the first alternative of each requirement
    $suggestedScopes = foreach ($requirement in $RequiredScopes) {
        ($requirement -split '\|')[0].Trim()
    }

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Warning "[$CallerName] No active Microsoft Graph session. Run Connect-MgGraph -Scopes $($suggestedScopes -join ',')"
        return $false
    }

    $grantedScopes = @($context.Scopes)

    $missingScopes = foreach ($requirement in $RequiredScopes) {
        [System.Collections.Generic.List[string]]$acceptedScopes = @()
        foreach ($alternative in ($requirement -split '\|')) {
            $alternative = $alternative.Trim()
            $acceptedScopes.Add($alternative)
            # A stronger scope satisfies a weaker requirement: ReadBasic < Read < ReadWrite
            if ($alternative -match '\.ReadBasic\.') {
                $acceptedScopes.Add(($alternative -replace '\.ReadBasic\.', '.Read.'))
                $acceptedScopes.Add(($alternative -replace '\.ReadBasic\.', '.ReadWrite.'))
            }
            elseif ($alternative -match '\.Read\.') {
                $acceptedScopes.Add(($alternative -replace '\.Read\.', '.ReadWrite.'))
            }
        }

        $satisfied = $false
        foreach ($acceptedScope in $acceptedScopes) {
            if ($acceptedScope -in $grantedScopes) {
                $satisfied = $true
                break
            }
        }

        if (-not $satisfied) {
            $requirement -replace '\|', ' or '
        }
    }

    if ($missingScopes) {
        Write-Warning "[$CallerName] The current Microsoft Graph token is missing the following scope(s): $($missingScopes -join ', ').`n`nReconnect with:`nConnect-MgGraph -NoWelcome -Scopes $($suggestedScopes -join ',')"
        return $false
    }

    return $true
}
