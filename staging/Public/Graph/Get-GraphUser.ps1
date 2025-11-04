function Get-GraphUser {
    [CmdletBinding()]
    param (
        [Parameter()]
        $UserId
    )
    if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PS365Refresh }
    $RestSplat = @{
        Uri     = 'https://graph.microsoft.com/beta/users/{0}' -f $UserId
        Headers = @{ "Authorization" = "Bearer $Token" }
        Method  = 'Get'
    }
    Invoke-RestMethod @RestSplat -Verbose:$false

}
