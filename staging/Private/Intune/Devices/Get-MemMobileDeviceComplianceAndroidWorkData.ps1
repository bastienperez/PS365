function Get-MemMobileDeviceComplianceAndroidWorkData {
    [CmdletBinding()]
    param (

    )
    if ([datetime]::UtcNow -ge $TimeToRefresh) { Connect-PS365Refresh }
    $RestSplat = @{
        Uri     = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?`$filter=isof('microsoft.graph.androidWorkProfileCompliancePolicy')&`$expand=assignments,scheduledActionsForRule(`$expand=scheduledActionConfigurations)"
        Headers = @{ 'Authorization' = "Bearer $Token" }
        Method  = 'Get'
    }
    
    Invoke-RestMethod @RestSplat -Verbose:$false | Select-Object -ExpandProperty value
}