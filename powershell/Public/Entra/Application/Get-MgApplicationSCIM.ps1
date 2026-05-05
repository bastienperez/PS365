<#
    .SYNOPSIS
    Retrieves all Entra ID applications configured for SCIM provisioning.

    .DESCRIPTION
    This function returns a list of all Entra ID applications with SCIM provisioning enabled,
    along with their synchronization job details and settings.

    .PARAMETER ExcludeAttributeMappings
    (Optional) If specified, skips the retrieval of the attribute mapping schema (ObjectMappings). This speeds up execution significantly when mapping details are not needed.

    .PARAMETER IncludeFailedObjects
    (Optional) If specified, fetches the list of objects currently in error (escrowed) for each synchronization job via the provisioning audit logs API.
    Requires the AuditLog.Read.All permission. If connecting interactively, this scope is added automatically.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER ForceNewToken
    (Optional) Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .PARAMETER ObjectID
    (Optional) Retrieves the SCIM configuration for a specific application by its ObjectID.

    .PARAMETER DisplayName
    (Optional) Retrieves the SCIM configuration for a specific application by its DisplayName.
    Supports wildcards (* and ?) for partial name matching (e.g. "Azure*", "*Portal*").

    .PARAMETER RunFromAzureAutomation
    (Optional) If specified, uses managed identity authentication instead of interactive authentication.
    This is useful when running the script in Azure environments like Azure Functions, Logic Apps, or VMs with managed identity enabled.
    When this parameter is used, NotificationRecipient and NotificationSender are required.

    PowerShell modules used in Azure Automation must be a MAXIMUM of version 2.25.0 when using PowerShell < 7.4.0, because starting from version 2.26.0, PowerShell 7.4.0 is required, and Azure Automation does not support it yet as of February 2026. For PowerShell 7.4.0+, there are no version restrictions.
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3147
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3151
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3166

    .PARAMETER NotificationRecipient
    (Required when RunFromAzureAutomation is enabled) Email address to receive synchronization health notifications.

    .PARAMETER NotificationSender
    (Required when RunFromAzureAutomation is enabled) Email address of the sender for synchronization health notifications.

    .EXAMPLE
    $scimApps = Get-MgApplicationSCIM

    Retrieves all Entra ID applications with SCIM provisioning enabled.

    .EXAMPLE
    Get-MgApplicationSCIM -ForceNewToken

    Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .EXAMPLE
    Get-MgApplicationSCIM -ExportToExcel

    Exports the SCIM configuration details to an Excel file.

    .EXAMPLE
    Get-MgApplicationSCIM -ObjectID "xxx-xxx-xxx"

    Retrieves the SCIM configuration for a specific application by its ObjectID.

    .EXAMPLE
    Get-MgApplicationSCIM -DisplayName "My App"

    Retrieves the SCIM configuration for a specific application by its DisplayName.

    .EXAMPLE
    Get-MgApplicationSCIM -DisplayName "Azure*"

    Retrieves the SCIM configuration for all applications whose DisplayName starts with "Azure".

    .EXAMPLE
    Get-MgApplicationSCIM -DisplayName "*Provisioning*"

    Retrieves the SCIM configuration for all applications whose DisplayName contains "Provisioning".

    .EXAMPLE
    Get-MgApplicationSCIM -RunFromAzureAutomation -NotificationRecipient 'admin@company.com' -NotificationSender 'automation@company.com'

    Gets all SCIM provisioning jobs using managed identity and sends a health report for apps with synchronization issues.

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
        [switch]$ExcludeAttributeMappings,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeFailedObjects,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [switch]$RunFromAzureAutomation,

        [Parameter(Mandatory = $false)]
        [string]$NotificationRecipient,

        [Parameter(Mandatory = $false)]
        [string]$NotificationSender
    )

    # Validate notification parameters
    if ($RunFromAzureAutomation.IsPresent) {
        if ([string]::IsNullOrWhiteSpace($NotificationRecipient)) {
            Write-Error 'NotificationRecipient parameter is required when RunFromAzureAutomation is enabled.'
            return
        }
        if ([string]::IsNullOrWhiteSpace($NotificationSender)) {
            Write-Error 'NotificationSender parameter is required when RunFromAzureAutomation is enabled.'
            return
        }

        try {
            Import-Module 'Microsoft.Graph.Users.Actions' -ErrorAction Stop -ErrorVariable mgGraphMailMissing
        }
        catch {
            if ($mgGraphMailMissing) {
                Write-Warning "Failed to import Microsoft.Graph.Users.Actions module: $($mgGraphMailMissing.Exception.Message)"
            }

            return
        }

        # Only check module version if PowerShell < 7.4 (Azure Automation limitation)
        if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
            $mgAuth = Get-Module 'Microsoft.Graph.Authentication' -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            if ($mgAuth -and [version]$mgAuth.Version -gt [version]'2.25.0') {
                Write-Error "Microsoft.Graph.Authentication v$($mgAuth.Version) is not compatible with Azure Automation on PowerShell $($PSVersionTable.PSVersion). Maximum supported version is 2.25.0. Script execution stopped."
                return
            }
        }
    }

    [System.Collections.Generic.List[PSCustomObject]]$synchronizationJobsArray = @()
    [System.Collections.Generic.List[PSCustomObject]]$synchronizationJobsDetailsArray = @()

    if ($ForceNewToken.IsPresent) {
        if (Get-MgContext) {
            $null = Disconnect-MgGraph
        }
    }

    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        if ($RunFromAzureAutomation.IsPresent) {
            Write-Verbose 'Connecting to Microsoft Graph using Managed Identity'
            $null = Connect-MgGraph -Identity -NoWelcome
        }
        else {
            $scopes = @('Directory.Read.All')
            if ($IncludeFailedObjects) { $scopes += 'AuditLog.Read.All' }
            Write-Verbose "Connecting to Microsoft Graph. Scopes: $($scopes -join ',')"
            $null = Connect-MgGraph -Scopes $scopes -NoWelcome
        }
    }

    # Verify AuditLog.Read.All permission when IncludeFailedObjects is requested
    if ($IncludeFailedObjects) {
        $currentScopes = (Get-MgContext).Scopes
        if ($currentScopes -notcontains 'AuditLog.Read.All') {
            Write-Error "AuditLog.Read.All permission is required when -IncludeFailedObjects is specified. Current scopes: $($currentScopes -join ', '). Please use -ForceNewToken to reconnect with the required scope."
            return
        }
        Write-Verbose 'AuditLog.Read.All permission verified.'
    }

    # Determine how to search for the Service Principal(s): by ObjectID (GUID), by DisplayName, or all
    if ($ObjectID) {
        $servicePrincipals = Get-MgServicePrincipal -ServicePrincipalId $ObjectID
    }
    elseif ($DisplayName) {
        if ($DisplayName -match '[*?]') {
            # Wildcard mode: extract base search keyword for server-side pre-filter
            $searchKeyword = ($DisplayName -replace '[*?]', '').Trim()
            Write-Verbose "Wildcard detected. Using Graph \$search with keyword '$searchKeyword' then PowerShell -like '$DisplayName'"

            if ($searchKeyword) {
                $uri = "/v1.0/servicePrincipals?`$search=`"displayName:$searchKeyword`"&`$count=true&`$select=displayName,id"
                $result = Invoke-MgGraphRequest -Uri $uri -Method GET -Headers @{ 'ConsistencyLevel' = 'eventual' }
            }
            else {
                # Pattern is only wildcards — fetch all
                $uri = '/v1.0/servicePrincipals?$select=displayName,id'
                $result = Invoke-MgGraphRequest -Uri $uri -Method GET
            }

            $candidateItems = if ($result.Value) { $result.Value } else { @() }
            $matchingItems = $candidateItems | Where-Object { $_.displayName -like $DisplayName }
            $servicePrincipals = @($matchingItems | ForEach-Object {
                [PSCustomObject]@{ DisplayName = $_.displayName; Id = $_.id }
            })
            Write-Verbose "Found $($servicePrincipals.Count) service principal(s) matching wildcard pattern '$DisplayName'"
        }
        else {
            # Exact match mode (no wildcards)
            $escaped = $DisplayName -replace "'", "''"
            $filter = "DisplayName eq '$escaped'"
            Write-Verbose "Filtering service principals with: $filter"
            $servicePrincipals = Get-MgServicePrincipal -Filter $filter -All -Property DisplayName, Id

            # If no exact match found, try to find apps where trimmed DisplayName matches
            if (-not $servicePrincipals) {
                Write-Verbose "No exact match found. Searching for service principals with trimmed DisplayName matching '$DisplayName'..."
                $filter = "startswith(DisplayName, '$escaped')"
                $candidateSPs = Get-MgServicePrincipal -Filter $filter -All -Property DisplayName, Id
                $servicePrincipals = $candidateSPs | Where-Object { $_.DisplayName.Trim() -eq $DisplayName }

                if ($servicePrincipals) {
                    Write-Verbose "Found $($servicePrincipals.Count) service principal(s) with trimmed DisplayName matching '$DisplayName'"
                }
            }
        }
    }
    else {
        $servicePrincipals = Get-MgServicePrincipal -All -Property DisplayName, Id
    }

    Write-Host "$($servicePrincipals.Count) service principals found"

    $i = 0
    foreach ($servicePrincipal in $servicePrincipals) {
        $i++
        Write-Host "($i/$($servicePrincipals.Count)) - $($servicePrincipal.DisplayName): check for synchronization jobs " -ForegroundColor Cyan -NoNewline
        $job = Get-MgServicePrincipalSynchronizationJob -ServicePrincipalId $servicePrincipal.Id -All

        if ($job) {
            Write-Host "$($servicePrincipal.DisplayName) - Synchronization job found" -ForegroundColor Green

            $job | Add-Member -MemberType NoteProperty -Name ServicePrincipalId -Value $servicePrincipal.Id
            $job | Add-Member -MemberType NoteProperty -Name DisplayName -Value $servicePrincipal.DisplayName
            $job | Add-Member -MemberType NoteProperty -Name EntraUrl -Value "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/ProvisioningActivity/objectId/$($servicePrincipal.Id)"

            $provisioningSettings = Invoke-GraphRequest -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($job.ServicePrincipalId)/synchronization/secrets" -Method Get -OutputType PSObject

            $job | Add-Member -MemberType NoteProperty -Name ProvisioningBaseAddress -Value $($provisioningSettings.Value | Where-Object { $_.key -eq 'BaseAddress' }).Value
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningSyncAll -Value $($provisioningSettings.Value | Where-Object { $_.key -eq 'SyncAll' }).Value
            $SyncNotificationSettings = ($provisioningSettings.Value | Where-Object { $_.key -eq 'SyncNotificationSettings' }).Value | ConvertFrom-Json
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationEnabled -Value $SyncNotificationSettings.Enabled
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationRecipientAddress -Value $SyncNotificationSettings.Recipients
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationDeleteThresholdEnabled -Value $SyncNotificationSettings.DeleteThresholdEnabled
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationDeleteThresholdValue -Value $SyncNotificationSettings.DeleteThresholdValue
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningNotificationHumanResourcesLookaheadQueryEnabled -Value $SyncNotificationSettings.HumanResourcesLookaheadQueryEnabled
            $job | Add-Member -MemberType NoteProperty -Name ProvisioningSkipOutOfScopeDeletions -Value $($provisioningSettings.Value | Where-Object { $_.key -eq 'SkipOutOfScopeDeletions' }).Value

            # Check for leading/trailing spaces in DisplayName
            $recommendation = $null
            if ($servicePrincipal.DisplayName -ne $servicePrincipal.DisplayName.Trim()) {
                $recommendation = 'DisplayName contains leading or trailing spaces - consider renaming'
                Write-Warning "Application '$($servicePrincipal.DisplayName)' has leading or trailing spaces in the displayName"
            }
            $job | Add-Member -MemberType NoteProperty -Name Recommendation -Value $recommendation

            # status Value
            $job | Add-Member -MemberType NoteProperty -Name StatusCode -Value $job.Status.Code
            $job | Add-Member -MemberType NoteProperty -Name StatusCountSuccessiveCompleteFailures -Value $job.Status.CountSuccessiveCompleteFailures
            $job | Add-Member -MemberType NoteProperty -Name StatusEscrowsPruned -Value $job.Status.EscrowsPruned
            $job | Add-Member -MemberType NoteProperty -Name StatusSteadyStateFirstAchievedTime -Value $job.Status.SteadyStateFirstAchievedTime
            $job | Add-Member -MemberType NoteProperty -Name StatusSteadyStateLastAchievedTime -Value $job.Status.SteadyStateLastAchievedTime
            $job | Add-Member -MemberType NoteProperty -Name StatusTroubleshootingUrl -Value $job.Status.TroubleshootingUrl

            # synchronizationJobSettings Value
            foreach ($property in $job.SynchronizationJobSettings) {
                $job | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
            }

            <#
            `$job.SynchronizationJobSettings` not useful
            Name                                Value
            ----                                -----
            AzureIngestionAttributeOptimization True
            LookaheadQueryEnabled               False
            Domain                              {"DomainDiscoveredAt":null,"DomainFQDN":null,"DomainNetBios":null,"ForestFQDN":null,"ForestNetBios":null}
            #>
            $job = $job | Select-Object -ExcludeProperty SynchronizationJobSettings
            $synchronizationJobsArray.Add($job)
        }
        else {
            Write-Host "$($servicePrincipal.DisplayName) - No synchronization job found" -ForegroundColor Yellow

            # No sync job: add a fallback row so the service principal always appears in the output
            $fallbackJob = [PSCustomObject][ordered]@{
                DisplayName       = $servicePrincipal.DisplayName
                ServicePrincipalId = $servicePrincipal.Id
                EntraUrl          = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/ProvisioningActivity/objectId/$($servicePrincipal.Id)"
                StatusCode        = '-'
            }
            $synchronizationJobsDetailsArray.Add($fallbackJob)
        }
    }

    $j = 0
    foreach ($job in $synchronizationJobsArray) {
        $j++

        Write-Host "Get synchronization settings $($job.DisplayName) ($j/$($synchronizationJobsArray.Count))"

        if (-not $ExcludeAttributeMappings) {
            $jobSchema = Get-MgServicePrincipalSynchronizationJobSchema -ServicePrincipalId $job.ServicePrincipalId -SynchronizationJobId $job.Id
        }

        $job | Add-Member -MemberType NoteProperty -Name Scheduling -Value $job.Schedule.Interval
        $job | Add-Member -MemberType NoteProperty -Name SchedulingState -Value $job.Schedule.State

        if ($job.Status.Quarantine.CurrentBegan) {
            $job | Add-Member -MemberType NoteProperty -Name Quarantined -Value $true
            $quarantineSeriesBegan = if ($job.Status.Quarantine.SeriesBegan) { " (series started: $($job.Status.Quarantine.SeriesBegan))" } else { '' }
            $quarantineReason = if ($job.Status.Quarantine.Reason) { " | Reason: $($job.Status.Quarantine.Reason)" } else { '' }
            $quarantineNextAttempt = if ($job.Status.Quarantine.NextAttempt) { " | Next attempt: $($job.Status.Quarantine.NextAttempt)" } else { '' }
            $quarantineRecommendation = "Application is in quarantine since $($job.Status.Quarantine.CurrentBegan)$quarantineSeriesBegan$quarantineReason$quarantineNextAttempt - please review provisioning logs in Entra"
            if (-not [string]::IsNullOrWhiteSpace($job.Recommendation)) {
                $job.Recommendation += " | $quarantineRecommendation"
            }
            else {
                $job.Recommendation = $quarantineRecommendation
            }
        }
        else {
            $job | Add-Member -MemberType NoteProperty -Name Quarantined -Value $false
        }

        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionDate -Value $job.Status.LastSuccessfulExecution.TimeEnded
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionState -Value $job.Status.LastSuccessfulExecution.State
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionCountEscrowed -Value $job.Status.LastSuccessfulExecution.CountEscrowed
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionWithExportsDate -Value $job.Status.LastSuccessfulExecutionWithExports.TimeEnded
        $job | Add-Member -MemberType NoteProperty -Name LastSuccessfulExecutionWithExportsState -Value $job.Status.LastSuccessfulExecutionWithExports.State
        $job | Add-Member -MemberType NoteProperty -Name LastExecutionCountEscrowed -Value $job.Status.LastExecution.CountEscrowed
        $job | Add-Member -MemberType NoteProperty -Name LastExecutionCountEscrowedRaw -Value $job.Status.LastExecution.CountEscrowedRaw

        if ($IncludeFailedObjects) {
            Write-Verbose "Fetching failed objects for $($job.DisplayName)..."
            [System.Collections.Generic.List[PSCustomObject]]$escrowedObjectsList = @()
            $dateFrom    = (Get-Date).AddDays(-30).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
            $dateTo      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
            $spId        = $job.ServicePrincipalId
            $spName      = [Uri]::EscapeDataString($job.DisplayName)
            $escrowFilter = "(activityDateTime ge $dateFrom and activityDateTime le $dateTo and (provisioningStatusInfo/status eq microsoft.graph.provisioningResult'failure') and (contains(tolower(servicePrincipal/id), '$spId') or contains(tolower(servicePrincipal/displayName), '$spName')))"
            $escrowUri   = "https://graph.microsoft.com/v1.0/auditLogs/provisioning?`$filter=$escrowFilter&`$top=500&`$orderby=activityDateTime desc"
            do {
                $escrowResponse = Invoke-MgGraphRequest -Method GET -Uri $escrowUri
                foreach ($item in $escrowResponse['value']) {
                    $failedStep = $item['provisioningSteps'] | Where-Object { $_['status'] -eq 'failure' } | Select-Object -First 1
                    $failedStepDetails = if ($failedStep -and $failedStep['details']) {
                        ($failedStep['details'].GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }) -join ' | '
                    } else { $null }

                    $errorInfo    = $item['provisioningStatusInfo']['errorInformation']
                    $errorCode    = if ($errorInfo) { $errorInfo['errorCode'] } elseif ($failedStep) { $failedStep['name'] } else { $null }
                    $reason       = if ($errorInfo) { $errorInfo['reason'] } elseif ($failedStep) { $failedStep['description'] } else { $null }

                    $srcDetails   = $item['sourceIdentity']['details']
                    $dn           = if ($srcDetails -is [System.Collections.IDictionary]) {
                                        $srcDetails['DistinguishedName']
                                    } elseif ($srcDetails) {
                                        ($srcDetails | Where-Object { $_['key'] -eq 'DistinguishedName' })['value']
                                    } else { $null }

                    $escrowedObjectsList.Add([PSCustomObject]@{
                        ActivityDateTime      = $item['activityDateTime']
                        SourceDisplayName     = $item['sourceIdentity']['displayName']
                        SourceId              = $item['sourceIdentity']['id']
                        DistinguishedName     = $dn
                        TargetDisplayName     = $item['targetIdentity']['displayName']
                        FailedStep            = if ($failedStep) { $failedStep['name'] } else { $null }
                        FailedStepType        = if ($failedStep) { $failedStep['provisioningStepType'] } else { $null }
                        FailedStepDescription = if ($failedStep) { $failedStep['description'] } else { $null }
                        FailedStepDetails     = $failedStepDetails
                        ErrorCode             = $errorCode
                        Reason                = $reason
                    })
                }
                $escrowUri = $escrowResponse['@odata.nextLink']
            } while ($escrowUri)

            $failedObjectsStats = $escrowedObjectsList |
                Group-Object FailedStep, ErrorCode |
                Sort-Object Count -Descending |
                ForEach-Object {
                    $parts = $_.Name -split ', ', 2
                    [PSCustomObject]@{
                        FailedStep = $parts[0]
                        ErrorCode  = $parts[1]
                        Count      = $_.Count
                    }
                }

            $latestFailedObjects = $escrowedObjectsList |
                Group-Object SourceId |
                ForEach-Object {
                    $_.Group | Sort-Object ActivityDateTime -Descending | Select-Object -First 1
                }

            $job | Add-Member -MemberType NoteProperty -Name FailedObjectsStats -Value $failedObjectsStats
            $job | Add-Member -MemberType NoteProperty -Name LatestFailedObjects -Value $latestFailedObjects
            $job | Add-Member -MemberType NoteProperty -Name FailedObjects -Value $escrowedObjectsList
        }

        foreach ($type in $job.Status.SynchronizedEntryCountByType) {
            $key = $type.Key -replace 'urn:ietf:params:scim:schemas:extension:enterprise:2\.0:', '' -replace 'urn:ietf:params:scim:schemas:core:2\.0:', ''
            $count = $type.Value

            $job | Add-Member -MemberType NoteProperty -Name "SynchronizedEntryCountByType_$key" -Value $count
        }

        [System.Collections.Generic.List[PSCustomObject]]$attributesArray = @()

        if (-not $ExcludeAttributeMappings) {
        foreach ($objectMapping in $jobSchema.SynchronizationRules.ObjectMappings) {

            foreach ($mapping in $objectMapping) {
                $res = $null

                if($mapping.TargetObjectName -like 'urn:ietf:params:scim:schemas:extension:enterprise:2.0:*') {
                    Write-Warning "Object type '$($mapping.TargetObjectName)' includes 'urn:ietf:params:scim:schemas:extension:enterprise:2.0:'. It is removed to improve column name readability and allow reuse of the same column for other SCIM types using only user/group instead of urn:ietf:params:scim:schemas:extension:enterprise:2.0:x"
                    $targetObjectType = $mapping.TargetObjectName -replace 'urn:ietf:params:scim:schemas:extension:enterprise:2.0:', ''
                }
                elseif($mapping.TargetObjectName -like 'urn:ietf:params:scim:schemas:core:2.0:*') {
                    Write-Warning "Object type '$($mapping.TargetObjectName)' includes 'urn:ietf:params:scim:schemas:core:2.0:'. It is removed to improve column name readability and allow reuse of the same column for other SCIM types using only user/group instead of urn:ietf:params:scim:schemas:core:2.0:x"
                    $targetObjectType = $mapping.TargetObjectName -replace 'urn:ietf:params:scim:schemas:core:2.0:', ''
                }
                else {
                    $targetObjectType = $mapping.TargetObjectName
                }

                $flowTypes = ($mapping.FlowTypes -join ',')
                $mappingMeta = "Enabled: $($mapping.Enabled) | FlowTypes: $flowTypes | Name: $($mapping.Name) | Source: $($mapping.SourceObjectName)"
                $job | Add-Member -MemberType NoteProperty -Name "ObjectMapping-$targetObjectType" -Value $mappingMeta -Force

                if ($flowTypes -ne 'Add,Update,Delete') {
                    $flowTypeRecommendation = "Object mapping for ``$targetObjectType`` has FlowType ``$flowTypes`` expected ``Add,Update,Delete``, please review to be sure it's intentional"
                    if (-not [string]::IsNullOrWhiteSpace($job.Recommendation)) {
                        $job.Recommendation += " | $flowTypeRecommendation"
                    }
                    else {
                        $job.Recommendation = $flowTypeRecommendation
                    }
                }

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

                    if ($null -ne $res) {
                        $res = "$res | "
                    }

                    if ([string]::IsNullOrWhiteSpace($object.SourceAttributeName)) {
                        $res = "$res$($object.SourceExpression) --> $($object.TargetAttributeName)"
                    }
                    else {
                        $res = "$res$($object.SourceAttributeName) --> $($object.TargetAttributeName)"
                    }
                }

                if($objectType -like 'urn:ietf:params:scim:schemas:extension:enterprise:2.0:*') {
                    Write-Warning "Object type '$objectType' includes 'urn:ietf:params:scim:schemas:extension:enterprise:2.0:'. It is removed to improve column name readability and allow reuse of the same column for other SCIM types using only user/group instead of urn:ietf:params:scim:schemas:extension:enterprise:2.0:x"
                    $object.Type = $object.Type -replace 'urn:ietf:params:scim:schemas:extension:enterprise:2.0:', ''
                }
                $job | Add-Member -MemberType NoteProperty -Name "Attributes-$($object.Type)" -Value $res -Force
            }
        }
        } # end if -not ExcludeAttributeMappings

        $job = $job | Select-Object * -ExcludeProperty Schedule, Schema

        $synchronizationJobsDetailsArray.Add($job)
    }

    # Send health report if running from Azure Automation
    if ($RunFromAzureAutomation.IsPresent) {
        $unhealthyJobs = $synchronizationJobsDetailsArray | Where-Object {
            $_.StatusCode -ne 'Steady' -or
            $_.Quarantined -eq $true -or
            $_.CountSuccessiveCompleteFailures -gt 0
        }

        $quarantinedJobs = $synchronizationJobsDetailsArray | Where-Object { $_.Quarantined -eq $true }
        $failingJobs = $synchronizationJobsDetailsArray | Where-Object { $_.CountSuccessiveCompleteFailures -gt 0 }
        $nonSteadyJobs = $synchronizationJobsDetailsArray | Where-Object { $_.StatusCode -ne 'Steady' -and $_.Quarantined -ne $true }

        Write-Verbose "Sending SCIM health report email ($($unhealthyJobs.Count) issues found out of $($synchronizationJobsDetailsArray.Count) jobs)."

        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
<title>Microsoft Entra ID SCIM Provisioning Health Report</title>
<style>
    body { 
        font-family: Segoe UI, SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif; 
        margin: 0; 
        padding: 20px; 
        color: #11100f; 
        font-size: 14px;
        line-height: 20px;
        background-color: #ffffff;
    }
    
    h2 { 
        padding-top: 0; 
        margin: 0 0 16px 0; 
        font-family: "Segoe UI Semibold", SegoeUISemibold, "Segoe UI", SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif;
        font-weight: 600; 
        font-size: 20px; 
        line-height: 28px;
        color: #323130;
    }
    
    table { 
        border-spacing: 0; 
        border-collapse: collapse; 
        width: 100%; 
        margin-bottom: 20px;
        background-color: #ffffff;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    th { 
        vertical-align: middle;
        color: #ffffff;
        background-color: #323130;
        padding: 3px 8px;
        text-align: left;
        font-family: "Segoe UI Semibold", SegoeUISemibold, "Segoe UI", SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif;
        font-weight: 600;
        font-size: 12px;
        line-height: 16px;
        word-wrap: break-word;
    }
    
    td { 
        vertical-align: middle;
        color: #11100f;
        padding: 3px 8px;
        border-bottom: solid 1px #c8c6c4;
        word-wrap: break-word;
        font-size: 12px;
        line-height: 16px;
    }
    
    .critical { background-color: #FFF0F0; color: #A80000; }
    .warning { background-color: #FDEFD0; color: #7A3A00; }
    .caution { background-color: #CCE4FF; color: #003882; }
    
    .footer {
        margin-top: 30px;
        padding: 20px;
        background-color: #faf9f8;
        border-radius: 8px;
        border-top: 3px solid #0078d4;
    }
    
    .footer p {
        margin: 8px 0;
        font-size: 13px;
        color: #605e5c;
    }
    
    .action-required {
        font-weight: 600;
        color: #d73502;
    }
</style>
</head>
<body>
    <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:12px;background:transparent;box-shadow:none;" role="presentation">
        <tr>
            <td width="25%" valign="top" style="width:25%;padding:4pt 3pt 4pt 5pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#FFF0F0;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#A80000;font-weight:600;">Quarantined</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#A80000;font-weight:bold;">$($quarantinedJobs.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#A80000;border-bottom:none;vertical-align:middle;">
                                        apps in quarantine
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="25%" valign="top" style="width:25%;padding:4pt 3pt 4pt 3pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#FDEFD0;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#7A3A00;font-weight:600;">Successive Failures</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#7A3A00;font-weight:bold;">$($failingJobs.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#7A3A00;border-bottom:none;vertical-align:middle;">
                                        apps with successive failures
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="25%" valign="top" style="width:25%;padding:4pt 3pt 4pt 3pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#CCE4FF;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#003882;font-weight:600;">Not Steady</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#003882;font-weight:bold;">$($nonSteadyJobs.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#003882;border-bottom:none;vertical-align:middle;">
                                        apps not in steady state
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="25%" valign="top" style="width:25%;padding:4pt 5pt 4pt 3pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#DFF6DD;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#107C10;font-weight:600;">Healthy</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#107C10;font-weight:bold;">$($synchronizationJobsDetailsArray.Count - $unhealthyJobs.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#107C10;border-bottom:none;vertical-align:middle;">
                                        apps healthy
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

"@

        if ($unhealthyJobs.Count -gt 0) {
            $emailBody += @'
    <h2>SCIM Provisioning Jobs Requiring Attention</h2>
    <table>
        <tr>
            <th>Application</th>
            <th>Status</th>
            <th>Quarantined</th>
            <th>Successive Failures</th>
            <th>Escrowed (last run)</th>
            <th>Last Successful Sync</th>
            <th>Scheduling State</th>
        </tr>
'@
            $unhealthyJobs = $unhealthyJobs | Sort-Object Quarantined -Descending

            foreach ($job in $unhealthyJobs) {
                $rowClass = if ($job.Quarantined) { 'critical' } elseif ($job.CountSuccessiveCompleteFailures -gt 0) { 'warning' } else { 'caution' }
                $appLink = "<strong style=`"color:#11100f;font-size:12px;line-height:16px;`">$($job.DisplayName)</strong> <a href=`"$($job.EntraUrl)`" style=`"text-decoration:none;font-size:14px;line-height:16px;`" title=`"Open in Entra`">&#x1F517;</a>"
                $quarantinedDisplay = if ($job.Quarantined) { '<strong>Yes</strong>' } else { 'No' }
                $escrowedCount = $job.LastExecutionCountEscrowed
                $escrowedDisplay = if ($escrowedCount -gt 0) { "<strong style=`"color:#A80000;`">$escrowedCount</strong>" } else { "$escrowedCount" }
                $emailBody += "<tr class=`"$rowClass`"><td>$appLink</td><td>$($job.StatusCode)</td><td>$quarantinedDisplay</td><td>$($job.CountSuccessiveCompleteFailures)</td><td>$escrowedDisplay</td><td>$($job.LastSuccessfulExecutionDate)</td><td>$($job.SchedulingState)</td></tr>"
            }

            $emailBody += '    </table>'

            if ($IncludeFailedObjects) {
                $jobsWithFailedObjects = $unhealthyJobs | Where-Object { $_.FailedObjects -and $_.FailedObjects.Count -gt 0 }
                if ($jobsWithFailedObjects.Count -gt 0) {
                    $emailBody += '    <h2>Failed Objects Detail (Last 30 Days)</h2>'
                    foreach ($failedJob in $jobsWithFailedObjects) {
                        $emailBody += "    <h3 style=`"font-size:14px;margin:16px 0 6px 0;color:#323130;`">$($failedJob.DisplayName) <span style=`"color:#A80000;`">($($failedJob.FailedObjects.Count) failed objects)</span></h3>"

                        # Stats grouped by FailedStep + ErrorCode
                        $statsGroups = $failedJob.FailedObjects | Group-Object FailedStep, ErrorCode | Sort-Object Count -Descending
                        $emailBody += @'
    <table>
        <tr>
            <th>Failed Step</th>
            <th>Error Code</th>
            <th>Count</th>
        </tr>
'@
                        foreach ($grp in $statsGroups) {
                            $parts = $grp.Name -split ', ', 2
                            $emailBody += "<tr class=`"warning`"><td>$($parts[0])</td><td>$($parts[1])</td><td><strong>$($grp.Count)</strong></td></tr>"
                        }
                        $emailBody += '    </table>'

                        # Detail table for each object
                        $emailBody += "    <p style=`"font-size:12px;margin:12px 0 4px 0;color:#605e5c;`"><em>Object detail:</em></p>"
                        $emailBody += @'
    <table>
        <tr>
            <th>Activity Date</th>
            <th>Source Name</th>
            <th>Source ID</th>
            <th>Target Name</th>
            <th>Failed Step</th>
            <th>Step Type</th>
            <th>Step Description</th>
            <th>Step Details</th>
            <th>Error Code</th>
            <th>Reason</th>
        </tr>
'@
                        foreach ($obj in $failedJob.FailedObjects) {
                            $emailBody += "<tr class=`"critical`"><td>$($obj.ActivityDateTime)</td><td>$($obj.SourceDisplayName)</td><td>$($obj.SourceId)</td><td>$($obj.TargetDisplayName)</td><td>$($obj.FailedStep)</td><td>$($obj.FailedStepType)</td><td>$($obj.FailedStepDescription)</td><td>$($obj.FailedStepDetails)</td><td>$($obj.ErrorCode)</td><td>$($obj.Reason)</td></tr>"
                        }
                        $emailBody += '    </table>'
                    }
                }
            }
        }
        else {
            $emailBody += "    <p style=`"color:#107c10;font-weight:600;`">All SCIM provisioning jobs are healthy.</p>"
        }

        $emailBody += @"

    <div class="footer">
        $(if ($unhealthyJobs.Count -gt 0) { '<p class="action-required">Action Required:</p><p>Please review these SCIM provisioning jobs to avoid user provisioning disruptions.</p>' } else { '<p>No action required. All provisioning jobs are running as expected.</p>' })
        <hr style="border: none; border-top: 1px solid #d2d0ce; margin: 15px 0;">
        <p><em>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Get-MgApplicationSCIM v0.0.96</em></p>
    </div>
</body>
</html>
"@

        try {
            $params = @{
                Message         = @{
                    Subject      = if ($unhealthyJobs.Count -gt 0) { "Microsoft Entra ID SCIM Provisioning Issues Detected ($($unhealthyJobs.Count) apps)" } else { "Microsoft Entra ID SCIM Provisioning Health Report - All Healthy ($($synchronizationJobsDetailsArray.Count) apps)" }
                    Body         = @{
                        ContentType = 'HTML'
                        Content     = $emailBody
                    }
                    ToRecipients = @(
                        @{
                            EmailAddress = @{
                                Address = $NotificationRecipient
                            }
                        }
                    )
                }
                SaveToSentItems = $false
            }

            Send-MgUserMail -UserId $NotificationSender -BodyParameter $params
            Write-Host -ForegroundColor Green "SCIM health notification email sent successfully to $NotificationRecipient"
        }
        catch {
            Write-Warning "Failed to send notification email: $($_.Exception.Message)"
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationSCIM-SynchronizationJobsInfo.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting SCIM synchronization jobs to Excel file: $excelFilePath"
        $synchronizationJobsDetailsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-ApplicationSCIM'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif (-not $RunFromAzureAutomation.IsPresent) {
        return $synchronizationJobsDetailsArray
    }
}