<#
    .SYNOPSIS
    Estimates the Microsoft 365 SLA service credit for one or more Service Health incidents and prepares the claim.

    .DESCRIPTION
    Reads the incidents from Microsoft 365 Service Health, merges the incident windows of each service (two
    overlapping incidents on the same service are counted once), splits them per calendar month (the SLA
    measures uptime per month) and returns, for each service and month, the downtime, the uptime percentage,
    the credit tier and the estimated credit.

    The credit tiers are the ones of the Microsoft Online Services SLA for the 99.9% services:
    below 99.9% uptime 25%, below 99% 50%, below 95% 100% of the monthly fee of the affected service.

    The affected users are the enabled users with an active service plan of the affected service
    (Exchange Online, SharePoint Online and OneDrive, Microsoft Teams). Use -AffectedUserCount when only part
    of the tenant was affected, or for a service the function cannot map to a service plan.

    The credit is computed on the price of the affected service, not on the whole licence: for a suite licence
    (Microsoft 365 E3, Business Premium...), pass the share of the price that Microsoft assigns to the service
    with -ServicePrice. Without a price for a service, the credit percentage is returned and EstimatedCredit
    stays empty.

    With -OutputFolder, the function also writes the incidents, the estimate, the affected users per country,
    the post-incident reports that are already published and a claim text to paste into the support request
    (Microsoft 365 admin center > Support > New service request).

    The function assumes every affected user was affected for the whole incident window: the estimate is a
    ceiling, Microsoft computes the final credit from its own user-minute measurements.

    .PARAMETER IncidentId
    One or more Service Health issue IDs (e.g. EX1550543, MO1550821).

    .PARAMETER ServicePrice
    Monthly price per user of each affected service, keyed by the Service Health service name
    (e.g. @{ 'Exchange Online' = 4; 'Microsoft Teams' = 3.2 }).

    .PARAMETER AffectedUserCount
    Number of affected users per service, keyed by the Service Health service name, when it differs from the
    licensed users (e.g. @{ 'Exchange Online' = 250 }).

    .PARAMETER Currency
    Currency of the prices, shown in the output and in the claim text. Default: EUR.

    .PARAMETER OutputFolder
    Folder for the CSV files, the post-incident reports and the claim text. Created when missing.

    .PARAMETER ExportToExcel
    Exports the estimate to an Excel file in the user profile instead of returning it.

    .EXAMPLE
    Get-MgSlaCreditEstimate -IncidentId 'EX1550543', 'MO1550821' -ServicePrice @{ 'Exchange Online' = 4 }

    Returns the downtime, uptime, credit tier and estimated credit per service and month.

    .EXAMPLE
    Get-MgSlaCreditEstimate -IncidentId 'EX1550543' -ServicePrice @{ 'Exchange Online' = 4 } -OutputFolder 'C:\Temp\SLA-claim'

    Same estimate, plus the CSV files, the post-incident report when published and claim-request.txt in C:\Temp\SLA-claim.

    .EXAMPLE
    Get-MgSlaCreditEstimate -IncidentId 'TM1551002' -ServicePrice @{ 'Microsoft Teams' = 3.2 } -AffectedUserCount @{ 'Microsoft Teams' = 120 }

    Estimate for a Teams incident that affected 120 users only.

    .EXAMPLE
    Get-MgSlaCreditEstimate -IncidentId 'EX1550543' -ServicePrice @{ 'Exchange Online' = 4 } -ExportToExcel

    Exports the estimate to an Excel file in the user profile.

    .NOTES
    Required Microsoft Graph scopes: ServiceHealth.Read.All, User.Read.All.
    Inspired by LazyAdmin (Rudy Mens): https://lazyadmin.nl/office-365/claim-microsoft-365-sla-credit-after-an-outage/

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgSlaCreditEstimate
#>

function Get-MgSlaCreditEstimate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$IncidentId,

        [Parameter(Mandatory = $false)]
        [hashtable]$ServicePrice = @{},

        [Parameter(Mandatory = $false)]
        [hashtable]$AffectedUserCount = @{},

        [Parameter(Mandatory = $false)]
        [string]$Currency = 'EUR',

        [Parameter(Mandatory = $false)]
        [string]$OutputFolder,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    $requiredScopes = @('ServiceHealth.Read.All', 'User.Read.All')
    if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    # Service Health service name -> 'service' value of the users' assignedPlans
    $servicePlanMap = @{
        'Exchange Online'      = 'exchange'
        'SharePoint Online'    = 'SharePoint'
        'OneDrive for Business' = 'SharePoint'
        'Microsoft Teams'      = 'TeamspaceAPI'
    }

    # Dates come back as DateTime or string depending on the PowerShell version
    function ConvertTo-SlaUtcDate {
        param ($Value)

        if (($null -eq $Value) -or ([string]$Value -eq '')) { return $null }
        if ($Value -is [datetime]) { return $Value.ToUniversalTime() }

        $styles = [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal
        return [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, $styles)
    }

    function Get-SlaCreditPercent {
        param ([double]$Uptime)

        if ($Uptime -lt 95) { return 100 }
        if ($Uptime -lt 99) { return 50 }
        if ($Uptime -lt 99.9) { return 25 }
        return 0
    }

    if ($OutputFolder) {
        $null = New-Item -Path $OutputFolder -ItemType Directory -Force
    }

    #region Incidents
    [System.Collections.Generic.List[PSCustomObject]]$incidents = @()
    $nowUtc = (Get-Date).ToUniversalTime()

    foreach ($id in $IncidentId) {
        $issue = $null
        try {
            $issue = Invoke-MgGraphRequestWithRetry -Method GET -Uri "/v1.0/admin/serviceAnnouncement/issues/$id" -OutputType PSObject
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Warning "Incident $id not read from Service Health: $errorMessage"
        }
        if (-not $issue) {
            continue
        }

        $start = ConvertTo-SlaUtcDate $issue.startDateTime
        $end = ConvertTo-SlaUtcDate $issue.endDateTime
        $isOngoing = -not $end
        if ($isOngoing) {
            Write-Warning "Incident $id has no end date yet, the estimate stops at the current time"
            $end = $nowUtc
        }

        $reportFile = $null
        if ($OutputFolder -and ($issue.status -eq 'postIncidentReviewPublished')) {
            $tempFile = Join-Path $OutputFolder "$id-post-incident-report"
            try {
                $null = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/admin/serviceAnnouncement/issues/$id/incidentReport" -OutputFilePath $tempFile -ErrorAction Stop

                # The report comes as a PDF or a Word document: the first bytes tell which
                $header = [System.IO.File]::ReadAllBytes($tempFile)[0..3]
                $extension = '.bin'
                if ([System.Text.Encoding]::ASCII.GetString($header) -eq '%PDF') { $extension = '.pdf' }
                elseif (($header[0] -eq 0x50) -and ($header[1] -eq 0x4B)) { $extension = '.docx' }

                $reportFile = "$tempFile$extension"
                Move-Item -Path $tempFile -Destination $reportFile -Force
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Warning "Post-incident report of $id not downloaded: $errorMessage"
            }
        }
        elseif ($OutputFolder) {
            Write-Verbose "No post-incident report published yet for $id"
        }

        $incidents.Add([PSCustomObject][ordered]@{
                IncidentId       = $issue.id
                Title            = $issue.title
                Service          = $issue.service
                Classification   = $issue.classification
                Status           = $issue.status
                StartUtc         = $start
                EndUtc           = $end
                Ongoing          = $isOngoing
                DurationHours    = [math]::Round(($end - $start).TotalHours, 1)
                PostIncidentFile = $reportFile
            })
    }

    if ($incidents.Count -eq 0) {
        Write-Warning 'No incident found, nothing to estimate'
        return
    }
    #endregion

    #region Affected users per service (licensed users unless -AffectedUserCount says otherwise)
    $services = @($incidents | Select-Object -ExpandProperty Service -Unique)
    $servicesToCount = @($services | Where-Object { (-not $AffectedUserCount.ContainsKey($_)) -and $servicePlanMap.ContainsKey($_) })

    $usersPerService = @{}
    [System.Collections.Generic.List[PSCustomObject]]$usersPerCountry = @()

    if ($servicesToCount.Count -gt 0) {
        $users = @(Invoke-MgGraphPagedRequest -Uri '/v1.0/users?$filter=accountEnabled eq true&$select=id,usageLocation,assignedPlans&$top=999')

        foreach ($service in $servicesToCount) {
            $planService = $servicePlanMap[$service]
            $serviceUsers = @($users | Where-Object {
                    @($_.assignedPlans | Where-Object { ($_.service -eq $planService) -and ($_.capabilityStatus -eq 'Enabled') }).Count -gt 0
                })
            $usersPerService[$service] = $serviceUsers.Count

            foreach ($group in ($serviceUsers | Group-Object -Property usageLocation)) {
                $country = if ($group.Name) { $group.Name } else { 'Not set' }
                $usersPerCountry.Add([PSCustomObject][ordered]@{
                        Service = $service
                        Country = $country
                        Users   = $group.Count
                    })
            }
        }
    }

    foreach ($service in $services) {
        if ($AffectedUserCount.ContainsKey($service)) {
            $usersPerService[$service] = [int]$AffectedUserCount[$service]
        }
        elseif (-not $usersPerService.ContainsKey($service)) {
            Write-Warning "No service plan known for '$service': pass its user count with -AffectedUserCount"
            $usersPerService[$service] = $null
        }
    }
    #endregion

    #region Downtime per service and month (overlapping incidents merged)
    [System.Collections.Generic.List[PSCustomObject]]$estimate = @()

    foreach ($service in $services) {
        $windows = @($incidents | Where-Object { $_.Service -eq $service } | Sort-Object StartUtc)

        # Merge the overlapping windows so a minute is never counted twice
        [System.Collections.Generic.List[PSCustomObject]]$merged = @()
        foreach ($window in $windows) {
            $last = if ($merged.Count -gt 0) { $merged[$merged.Count - 1] } else { $null }
            if ($last -and ($window.StartUtc -le $last.EndUtc)) {
                if ($window.EndUtc -gt $last.EndUtc) { $last.EndUtc = $window.EndUtc }
                $last.IncidentIds += $window.IncidentId
            }
            else {
                $merged.Add([PSCustomObject][ordered]@{
                        StartUtc    = $window.StartUtc
                        EndUtc      = $window.EndUtc
                        IncidentIds = @($window.IncidentId)
                    })
            }
        }

        # Downtime minutes per calendar month (UTC)
        $monthly = [ordered]@{}
        foreach ($window in $merged) {
            $monthStart = [datetime]::new($window.StartUtc.Year, $window.StartUtc.Month, 1, 0, 0, 0, [DateTimeKind]::Utc)
            while ($monthStart -lt $window.EndUtc) {
                $monthEnd = $monthStart.AddMonths(1)
                $from = if ($window.StartUtc -gt $monthStart) { $window.StartUtc } else { $monthStart }
                $to = if ($window.EndUtc -lt $monthEnd) { $window.EndUtc } else { $monthEnd }
                $key = $monthStart.ToString('yyyy-MM')

                if (-not $monthly.Contains($key)) {
                    $monthly[$key] = [PSCustomObject][ordered]@{
                        MonthStart  = $monthStart
                        Minutes     = 0
                        IncidentIds = @()
                    }
                }
                $monthly[$key].Minutes += ($to - $from).TotalMinutes
                $monthly[$key].IncidentIds += $window.IncidentIds
                $monthStart = $monthEnd
            }
        }

        $price = if ($ServicePrice.ContainsKey($service)) { [double]$ServicePrice[$service] } else { $null }
        $userCount = $usersPerService[$service]

        foreach ($key in $monthly.Keys) {
            $month = $monthly[$key]
            $monthMinutes = ($month.MonthStart.AddMonths(1) - $month.MonthStart).TotalMinutes
            $uptime = [math]::Round(($monthMinutes - $month.Minutes) / $monthMinutes * 100, 3)
            $creditPercent = Get-SlaCreditPercent -Uptime $uptime

            $credit = $null
            if (($null -ne $price) -and ($null -ne $userCount)) {
                $credit = [math]::Round($userCount * $price * $creditPercent / 100, 2)
            }

            $estimate.Add([PSCustomObject][ordered]@{
                    Service         = $service
                    Month           = $key
                    IncidentIds     = (@($month.IncidentIds | Select-Object -Unique) -join ', ')
                    DowntimeHours   = [math]::Round($month.Minutes / 60, 1)
                    UptimePercent   = $uptime
                    CreditPercent   = $creditPercent
                    AffectedUsers   = $userCount
                    PricePerUser    = $price
                    EstimatedCredit = $credit
                    Currency        = $Currency
                })
        }

        if ($null -eq $price) {
            Write-Warning "No price for '$service': pass it with -ServicePrice to get the credit amount"
        }
    }
    #endregion

    #region Files and claim text
    if ($OutputFolder) {
        $csvParams = @{ NoTypeInformation = $true; Encoding = 'UTF8' }

        $incidents | Select-Object IncidentId, Title, Service, Classification, Status,
        @{ Name = 'StartUtc'; Expression = { $_.StartUtc.ToString('yyyy-MM-dd HH:mm') } },
        @{ Name = 'EndUtc'; Expression = { $_.EndUtc.ToString('yyyy-MM-dd HH:mm') } }, Ongoing, DurationHours |
        Export-Csv -Path (Join-Path $OutputFolder 'incidents.csv') @csvParams

        $estimate | Export-Csv -Path (Join-Path $OutputFolder 'credit-estimate.csv') @csvParams

        if ($usersPerCountry.Count -gt 0) {
            $usersPerCountry | Sort-Object Service, @{ Expression = 'Users'; Descending = $true } |
            Export-Csv -Path (Join-Path $OutputFolder 'affected-users-per-country.csv') @csvParams
        }

        $tenantId = (Get-MgContext).TenantId
        [System.Collections.Generic.List[String]]$lines = @()
        $lines.Add('Request for an SLA service credit')
        $lines.Add('')
        $lines.Add("Tenant ID: $tenantId")
        $lines.Add('')
        $lines.Add('Service Health incidents:')
        foreach ($incident in $incidents) {
            $startText = $incident.StartUtc.ToString('yyyy-MM-dd HH:mm')
            $endText = $incident.EndUtc.ToString('yyyy-MM-dd HH:mm')
            $ongoingText = if ($incident.Ongoing) { ' (still open)' } else { '' }
            $lines.Add("- $($incident.IncidentId), $($incident.Service): $startText to $endText UTC, $($incident.DurationHours) hours$ongoingText")
        }
        $lines.Add('')
        $lines.Add('Impact per service and month:')
        foreach ($row in $estimate) {
            $usersText = if ($null -ne $row.AffectedUsers) { "$($row.AffectedUsers) affected users" } else { 'affected users to complete' }
            $lines.Add("- $($row.Service), $($row.Month): $($row.DowntimeHours) hours of downtime, monthly uptime $($row.UptimePercent)%, $usersText")
        }
        if ($usersPerCountry.Count -gt 0) {
            $lines.Add('')
            $lines.Add('Licensed users per country:')
            foreach ($row in ($usersPerCountry | Sort-Object Service, @{ Expression = 'Users'; Descending = $true })) {
                $lines.Add("- $($row.Service), $($row.Country): $($row.Users)")
            }
        }
        $lines.Add('')
        $lines.Add('What the users could not do:')
        $lines.Add('[To complete: the features that failed, e.g. no access to mailboxes in Outlook, Outlook on the web and mobile]')
        $lines.Add('')
        $lines.Add('What we did during the incident:')
        $lines.Add('[To complete: workarounds, communication to users, support ticket numbers]')

        $lines -join [Environment]::NewLine | Out-File -FilePath (Join-Path $OutputFolder 'claim-request.txt') -Encoding UTF8
        Write-Verbose "Files written to $OutputFolder"
    }
    #endregion

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgSlaCreditEstimate.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $estimate | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'MgSlaCreditEstimate'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $estimate
    }
}
