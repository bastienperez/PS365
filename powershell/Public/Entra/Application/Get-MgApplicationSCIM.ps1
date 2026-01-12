<#
    .SYNOPSIS
    Retrieves all Entra ID applications configured for SCIM provisioning.

    .DESCRIPTION
    This function returns a list of all Entra ID applications with SCIM provisioning enabled,
    along with their synchronization job details and settings.

    .EXAMPLE
    $scimApps = Get-MgApplicationSCIM

    Retrieves all Entra ID applications with SCIM provisioning enabled.

    .EXAMPLE
    Get-MgApplicationSCIM -ForceNewToken

    Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .EXAMPLE
    Get-MgApplicationSCIM -Export

    Exports the SCIM configuration details to a CSV file.

    .EXAMPLE
    Get-MgApplicationSCIM -ObjectID "xxx-xxx-xxx"

    Retrieves the SCIM configuration for a specific application by its ObjectID.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationSCIM

    .NOTES
    LIMITATIONS
    The groups assignments are not retrieved because based on https://main.iam.ad.ext.azure.com

    This function requires the Microsoft.Graph.Applications and Microsoft.Graph.Authentication modules.
#>

function Get-MgApplicationSCIM {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [Parameter(Mandatory = $false, ParameterSetName = 'ByObjectId')]
        [string]$ObjectID,
        [Parameter(Mandatory = $false, ParameterSetName = 'ByDisplayName')]
        [string]$DisplayName,
        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    [System.Collections.Generic.List[PSCustomObject]]$synchronizationJobsArray = @()
    [System.Collections.Generic.List[PSCustomObject]]$synchronizationJobsDetailsArray = @()

    if ($ForceNewToken.IsPresent) {
        if (Get-MgContext) {
            $null = Disconnect-MgGraph
        }

        $scopes = @(
            'Directory.Read.All'
        )

        Connect-MgGraph -Scopes $scopes -NoWelcome
    }
    # Determine how to search for the Service Principal(s): by ObjectID (GUID), by DisplayName, or all
    if ($ObjectID) {
        # If ObjectID looks like a GUID, use it directly
        $servicePrincipals = Get-MgServicePrincipal -ServicePrincipalId $ObjectID
    }
    elseif ($DisplayName) {
        # Use OData filter to search by DisplayName. Escape single quotes by doubling them.
        $escaped = $DisplayName -replace "'", "''"
        $filter = "DisplayName eq '$escaped'"
        Write-Verbose "Filtering service principals with: $filter"
        $servicePrincipals = Get-MgServicePrincipal -Filter $filter -All -Property DisplayName, Id
    }
    else {
        # Default: get all service principals (existing behavior)
        $servicePrincipals = Get-MgServicePrincipal -All -Property DisplayName, Id
    }

    Write-Host "$($servicePrincipals.Count) service principals found"

    $i = 0
    foreach ($servicePrincipal in $servicePrincipals) {
        $i++
        Write-Host "($i/$($servicePrincipals.Count)) - $($servicePrincipal.DisplayName): check for synchronization jobs " -ForegroundColor Cyan -NoNewline
        # Service principal is the ObjectID in Microsoft Entra ID
        $job = Get-MgServicePrincipalSynchronizationJob -ServicePrincipalId $servicePrincipal.Id -All
        
        if ($job) {
            Write-Host "$($servicePrincipal.DisplayName) - Synchronization job found" -ForegroundColor Green
            
            # We need to keep $servicePrincipal.Id in a new property ServicePrincipalID

            $job | Add-Member -MemberType NoteProperty -Name ServicePrincipalId -Value $servicePrincipal.Id
            $job | Add-Member -MemberType NoteProperty -Name DisplayName -Value $servicePrincipal.DisplayName
        
            $provisioningSettings = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($job.ServicePrincipalId)/synchronization/secrets" -Method Get -OutputType PSObject

            $job | Add-Member -MemberType NoteProperty -Name ProvisioningBaseAddress -Value $($provisioningSettings.Value | Where-Object { $_.key -eq 'BaseAddress' }).Value
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningSyncAll -Value $($provisioningSettings.Value | Where-Object { $_.key -eq 'SyncAll' }).Value
            $SyncNotificationSettings = ($provisioningSettings.Value | Where-Object { $_.key -eq 'SyncNotificationSettings' }).Value | ConvertFrom-Json
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationEnabled -Value $SyncNotificationSettings.Enabled
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationRecipientAddress -Value $SyncNotificationSettings.Recipients
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationDeleteThresholdEnabled -Value $SyncNotificationSettings.DeleteThresholdEnabled
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationDeleteThresholdValue -Value $SyncNotificationSettings.DeleteThresholdValue
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationHumanResourcesLookaheadQueryEnabled -Value $SyncNotificationSettings.HumanResourcesLookaheadQueryEnabled    

            # status Value
            $job | Add-Member -MemberType NoteProperty -Name StatusCode -Value $job.Status.Code
            $job | Add-Member -MemberType NoteProperty -Name CountSuccessiveCompleteFailures -Value $job.Status.CountSuccessiveCompleteFailures
            $job | Add-Member -MemberType NoteProperty -Name EscrowsPruned -Value $job.Status.EscrowsPruned
            $job | Add-Member -MemberType NoteProperty -Name SteadyStateFirstAchievedTime -Value $job.Status.SteadyStateFirstAchievedTime
            $job | Add-Member -MemberType NoteProperty -Name SteadyStateLastAchievedTime -Value $job.Status.SteadyStateLastAchievedTime
            $job | Add-Member -MemberType NoteProperty -Name TroubleshootingUrl -Value $job.Status.TroubleshootingUrl

            # synchronizationJobSettings Value
            foreach ($property in $job.SynchronizationJobSettings) {
                $job | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
            }

            # we exclude the Status and SynchronizationJobSettings properties because we already got its values
            $job = $job | Select-Object -ExcludeProperty Status, SynchronizationJobSettings
            $synchronizationJobsArray.Add($job)
        }
        else {
            Write-Host "$($servicePrincipal.DisplayName) - No synchronization job found" -ForegroundColor Yellow
        }
    }

    $j = 0
    foreach ($job in $synchronizationJobsArray) {
        $j++

        # Get information about

        Write-Host "Get synchronization settings $($job.DisplayName) ($j/$($synchronizationJobsArray.Count))"

        $jobSchema = Get-MgServicePrincipalSynchronizationJobSchema -ServicePrincipalId $job.ServicePrincipalId -SynchronizationJobId $job.Id

        $job | Add-Member -MemberType NoteProperty -Name Scheduling -Value $job.Schedule.Interval
        $job | Add-Member -MemberType NoteProperty -Name SchedulingState -Value $job.Schedule.State
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionDate -Value $job.Status.LastSuccessfulExecution.TimeEnded
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionState -Value $job.Status.LastSuccessfulExecution.State
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionWithExportsDate -Value $job.Status.LastSuccessfulExecutionWithExports.TimeEnded
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionWithExportsState -Value $job.Status.LastSuccessfulExecutionWithExports.State

        if ($job.Status.Quarantine.CurrentBegan) {
            $job | Add-Member -MemberType NoteProperty -Name Quarantined -Value $true
        }
        else {
            $job | Add-Member -MemberType NoteProperty -Name Quarantined -Value $false
        }

        foreach ($type in $job.Status.SynchronizedEntryCountByType) {
            # it's not an hashtable but an array
            $key = $type.Key
            $count = $type.Value

            $job | Add-Member -MemberType NoteProperty -Name "SynchronizedEntryCountByType_$key" -Value $count
        }

        [System.Collections.Generic.List[PSCustomObject]]$attributesArray = @()
    
        foreach ($objectMapping in $jobSchema.SynchronizationRules.ObjectMappings) {

            foreach ($mapping in $objectMapping) {
                $res = $null

                $targetObjectType = $mapping.TargetObjectName

                foreach ($attribute in $mapping.AttributeMappings) {
                
                    $object = [PSCustomObject][ordered]@{
                        Type                    = $targetObjectType
                        DefaultValue            = $attribute.DefaultValue
                        ExportMissingReferences = $attribute.ExportMissingReferences
                        FlowBehavior            = $attribute.FlowBehavior
                        FlowType                = $attribute.FlowType
                        MatchingPriority        = $attribute.MatchingPriority
                        SourceExpression        = $attribute.Source.Expression
                        SourceAttributeName     = $attribute.Source.AttributeName
                        TargetAttributeName     = $attribute.TargetAttributeName
                    }

                    $attributesArray.Add($object)

                    # add delimiter if not the first attribute
                    if ($null -ne $res) {
                        $res = "$res # "
                    }

                    if ([string]::IsNullOrWhitespace($object.SourceAttributeName)) {
                        $res = "$res$($object.SourceExpression) --> $($object.TargetAttributeName)"
                    }
                    else {
                        $res = "$res$($object.SourceAttributeName) --> $($object.TargetAttributeName)"
                    }
                }
                try {
                    $job | Add-Member -MemberType NoteProperty -Name "Attributes-$($object.Type)" -Value $res
                }
                catch {
                    Write-Host "Error adding member to job: $($_.Exception.Message)" -ForegroundColor Red
                }           
            }
        }
    
        # exclude properties because they are not needed
        $job = $job | Select-Object * -ExcludeProperty Schedule, Schema

        $synchronizationJobsDetailsArray.Add($job)
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationSCIM-SynchronizationJobsInfo.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting SCIM synchronization jobs to Excel file: $excelFilePath"
        $synchronizationJobsDetailsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-ApplicationSCIM'
        Write-Host -ForegroundColor Green "Export completed successfully!"
    }
    else {
        return $synchronizationJobsDetailsArray
    }
}