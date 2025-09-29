function Get-AzureTrafficManagerEndpointReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        $TrafficMgrProfile
    )
    begin {

    }
    process {

        foreach ($CurTrafficMgrProfile in $TrafficMgrProfile.Endpoints) {
            [PSCustomObject][ordered]@{
                Priority              = $CurTrafficMgrProfile.Priority
                Name                  = $CurTrafficMgrProfile.Name
                Weight                = $CurTrafficMgrProfile.Weight
                Location              = $CurTrafficMgrProfile.Location
                Target                = $CurTrafficMgrProfile.Target
                ProfileName           = $CurTrafficMgrProfile.ProfileName
                ResourceGroupName     = $CurTrafficMgrProfile.ResourceGroupName
                Type                  = $CurTrafficMgrProfile.Type
                EndpointStatus        = $CurTrafficMgrProfile.EndpointStatus
                Id                    = $CurTrafficMgrProfile.Id -replace '.*\/'
                TargetResourceId      = $CurTrafficMgrProfile.TargetResourceId -replace '.*\/'
                EndpointMonitorStatus = $CurTrafficMgrProfile.EndpointMonitorStatus
                MinChildEndpoints     = $CurTrafficMgrProfile.MinChildEndpoints
                GeoMapping            = $CurTrafficMgrProfile.GeoMapping
            }
        }
    }
    end {

    }
}