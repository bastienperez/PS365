function Get-MemAuthenticationMethods {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $UserPrincipalName
    )
    process {
        foreach ($UPN in $UserPrincipalName) {
            if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PS365Refresh }
            $RestSplat = @{
                Uri     = 'https://graph.microsoft.com/beta/users/{0}/authentication/methods' -f $UPN
                Headers = @{ "Authorization" = "Bearer $Token" }
                Method  = 'Get'
            }
            Invoke-RestMethod @RestSplat -Verbose:$false
        }
    }
}
