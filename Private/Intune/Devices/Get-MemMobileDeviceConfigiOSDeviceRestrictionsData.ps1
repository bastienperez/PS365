function Get-MemMobileDeviceConfigiOSDeviceRestrictionsData {
    [CmdletBinding()]
    param (

    )
    if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PoshGraphRefresh }
    $RestSplat = @{
        Uri     = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$filter=isof('microsoft.graph.iosGeneralDeviceConfiguration')&`$expand=assignments"
        Headers = @{ 'Authorization' = "Bearer $Token" }
        Method  = 'Get'
    }

    Invoke-RestMethod @RestSplat -Verbose:$false | Select-Object -ExpandProperty value
}