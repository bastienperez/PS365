<#
    .SYNOPSIS
    Retrieves all Entra ID applications configured for SCIM provisioning.

    .DESCRIPTION
    This function returns a list of all Entra ID applications with SCIM provisioning enabled,
    along with their synchronization job details and settings.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER ForceNewToken
    (Optional) Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .PARAMETER ObjectID
    (Optional) Retrieves the SCIM configuration for a specific application by its ObjectID.

    .PARAMETER DisplayName
    (Optional) Retrieves the SCIM configuration for a specific application by its DisplayName.

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
            Connect-MgGraph -Identity -NoWelcome
        }
        else {
            $scopes = @('Directory.Read.All')
            Write-Verbose "Connecting to Microsoft Graph. Scopes: $($scopes -join ',')"
            Connect-MgGraph -Scopes $scopes -NoWelcome
        }
    }

    # Determine how to search for the Service Principal(s): by ObjectID (GUID), by DisplayName, or all
    if ($ObjectID) {
        $servicePrincipals = Get-MgServicePrincipal -ServicePrincipalId $ObjectID
    }
    elseif ($DisplayName) {
        $escaped = $DisplayName -replace "'", "''"
        $filter = "DisplayName eq '$escaped'"
        Write-Verbose "Filtering service principals with: $filter"
        $servicePrincipals = Get-MgServicePrincipal -Filter $filter -All -Property DisplayName, Id
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
            <th>Last Successful Sync</th>
            <th>Scheduling State</th>
        </tr>
'@
            $unhealthyJobs = $unhealthyJobs | Sort-Object Quarantined -Descending

            foreach ($job in $unhealthyJobs) {
                $rowClass = if ($job.Quarantined) { 'critical' } elseif ($job.CountSuccessiveCompleteFailures -gt 0) { 'warning' } else { 'caution' }
                $appLink = "<strong style=`"color:#11100f;font-size:12px;line-height:16px;`">$($job.DisplayName)</strong> <a href=`"$($job.EntraUrl)`" style=`"text-decoration:none;font-size:14px;line-height:16px;`" title=`"Open in Entra`">&#x1F517;</a>"
                $quarantinedDisplay = if ($job.Quarantined) { '<strong>Yes</strong>' } else { 'No' }
                $emailBody += "<tr class=`"$rowClass`"><td>$appLink</td><td>$($job.StatusCode)</td><td>$quarantinedDisplay</td><td>$($job.CountSuccessiveCompleteFailures)</td><td>$($job.LastSuccessfulExecutionDate)</td><td>$($job.SchedulingState)</td></tr>"
            }

            $emailBody += '    </table>'
        }
        else {
            $emailBody += "    <p style=`"color:#107c10;font-weight:600;`">All SCIM provisioning jobs are healthy.</p>"
        }

        $emailBody += @"

    <div class="footer">
        $(if ($unhealthyJobs.Count -gt 0) { '<p class="action-required">Action Required:</p><p>Please review these SCIM provisioning jobs to avoid user provisioning disruptions.</p>' } else { '<p>No action required. All provisioning jobs are running as expected.</p>' })
        <hr style="border: none; border-top: 1px solid #d2d0ce; margin: 15px 0;">
        <p><em>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Get-MgApplicationSCIM v0.71.0</em></p>
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
                SaveToSentItems = 'false'
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
