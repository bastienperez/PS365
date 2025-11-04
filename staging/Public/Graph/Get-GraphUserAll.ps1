function Get-GraphUserAll {
    [CmdletBinding()]
    param (
        [Parameter()]
        [switch]
        $IncludeGuests
    )
    if (-not $IncludeGuests) {
        $uri = "https://graph.microsoft.com/beta/users/?`$filter=userType eq 'Member'"
    }
    else { $Uri = 'https://graph.microsoft.com/beta/users' }

    $RestSplat = @{
        Uri     = $Uri
        Headers = @{ "Authorization" = "Bearer $Token" }
        Method  = 'Get'
    }
    do {
        if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PS365Refresh }
        $Response = Invoke-RestMethod @RestSplat -Verbose:$false
        $Response.value
        if ($Response.'@odata.nextLink' -match 'skip') { $Next = $Response.'@odata.nextLink' }
        else { $Next = $null }
        $RestSplat = @{
            Uri     = $Next
            Headers = @{ "Authorization" = "Bearer $Token" }
            Method  = 'Get'
        }

    } until (-not $next)
}
