function Get-GraphLocation {
    [CmdletBinding()]
    param (
        [Parameter()]
        $Id
    )
    if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PS365Refresh }
    $RestSplat = @{
        Uri     = 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/{0}' -f $Id
        Headers = @{ "Authorization" = "Bearer $Token" }
        Method  = 'Get'
    }
    Invoke-RestMethod @RestSplat -Verbose:$false

}
