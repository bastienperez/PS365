function Get-AzureADSP {
    [CmdletBinding()]
    param (

    )
    if ([datetime]::UtcNow -ge $TimeToRefresh) {
        Connect-PS365Refresh
    }
    $RestSplat = @{
        Uri     = 'https://graph.microsoft.com/beta/servicePrincipals/'
        Headers = @{ "Authorization" = "Bearer $Token" }
        Method  = 'Get'
    }
    Invoke-RestMethod @RestSplat -Verbose:$false | Select-Object -ExpandProperty Value
}
