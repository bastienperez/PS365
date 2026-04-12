<#
    .SYNOPSIS
    Verifies that the current Microsoft Graph token has the required scopes.

    .DESCRIPTION
    Reads the scopes from Get-MgContext and checks that every scope listed in
    -RequiredScopes is present. A Read scope is considered satisfied when the
    corresponding ReadWrite scope is granted (e.g. Application.ReadWrite.All
    covers Application.Read.All).

    Returns $true when every required scope is present, otherwise writes a
    warning listing the missing scopes and returns $false.

    .PARAMETER RequiredScopes
    One or more Microsoft Graph scopes that the calling function needs.

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

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Warning "[$CallerName] No active Microsoft Graph session. Run Connect-MgGraph -Scopes $($RequiredScopes -join ',')"
        return $false
    }

    $grantedScopes = @($context.Scopes)

    $missing = foreach ($scope in $RequiredScopes) {
        if ($scope -in $grantedScopes) { continue }
        # Accept the ReadWrite equivalent as a superset of a Read scope.
        $readWriteEquivalent = $scope -replace '\.Read\.', '.ReadWrite.'
        if ($readWriteEquivalent -ne $scope -and $readWriteEquivalent -in $grantedScopes) { continue }
        $scope
    }

    if ($missing) {
        Write-Warning "[$CallerName] The current Microsoft Graph token is missing the following scope(s): $($missing -join ', '). Reconnect with: Connect-MgGraph -Scopes $($RequiredScopes -join ',') (or bypass this check with -NoPermissionCheck)."
        return $false
    }

    return $true
}
