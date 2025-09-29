function Get-AzureTrafficManagerReport {
    param (
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        $TrafficMgrProfile
    )
    begin {

    }
    process {

        foreach ($CurTrafficMgrProfile in $TrafficMgrProfile) {

            [PSCustomObject][ordered]@{
                ResourceGroupName      = $CurTrafficMgrProfile.ResourceGroupName
                ProfileName            = $CurTrafficMgrProfile.Name
                ProfileDnsName         = $CurTrafficMgrProfile.RelativeDnsName + '.trafficmanager.net'
                ProfileTtl             = $CurTrafficMgrProfile.Ttl
                ProfileStatus          = $CurTrafficMgrProfile.ProfileStatus
                ProfileRoutingMethod   = $CurTrafficMgrProfile.TrafficRoutingMethod
                ProfileMonitorProtocol = $CurTrafficMgrProfile.MonitorProtocol
                ProfileMonitorPort     = $CurTrafficMgrProfile.MonitorPort
                ProfileMonitorPath     = $CurTrafficMgrProfile.MonitorPath
                ProfileMonitorInterval = $CurTrafficMgrProfile.MonitorIntervalInSeconds
                ProfileMonitorTimeout  = $CurTrafficMgrProfile.MonitorTimeoutInSeconds
                ProfileMonitorFailures = $CurTrafficMgrProfile.MonitorToleratedNumberOfFailures
                ProfileEndpoints       = ($CurTrafficMgrProfile.Endpoints.Name | Where-Object {$_ -ne $null}) -join "`r`n"

            }
        }
    }
    end {

    }
}