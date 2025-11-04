function Get-MemMobileAppProtectionAndroidData {
    [CmdletBinding()]
    param (

    )
    if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PS365Refresh }
    $RestSplat = @{
        Uri     = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections?`$expand=deploymentSummary,apps,assignments"
        Headers = @{ 'Authorization' = "Bearer $Token" }
        Method  = 'Get'
    }
    Invoke-RestMethod @RestSplat -Verbose:$false | Select-Object -ExpandProperty Value

}