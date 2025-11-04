function Get-MemMobileDeviceConfigData {
    [CmdletBinding()]
    param (

    )
    if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PS365Refresh }
    $RestSplat = @{
        Uri     = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$filter=isof('microsoft.graph.windowsUpdateForBusinessConfiguration')&`$expand=assignments"
        Headers = @{ 'Authorization' = "Bearer $Token" }
        Method  = 'Get'
    }

    Invoke-RestMethod @RestSplat -Verbose:$false | Select-Object -ExpandProperty value
}