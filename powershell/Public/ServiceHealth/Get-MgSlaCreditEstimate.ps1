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
    with -ServicePrice. Without -ServicePrice for a service, the public list price of its standalone plan is used
    (Private/Assets/Get-MgSlaCreditEstimate-list-prices.json, excluding VAT, annual commitment: Exchange Online
    Plan 1 and Plan 2 in EUR and USD, Teams Essentials in EUR) when its currency matches
    -Currency. For Exchange Online, each licensed user is priced at the plan of their service plan
    (EXCHANGE_S_STANDARD = Plan 1, EXCHANGE_S_ENTERPRISE = Plan 2, also held by suites such as Microsoft 365 E3),
    one row per plan; other Exchange plans (Kiosk, F3...) are priced as Plan 1. The Plan and PriceSource columns
    say which price was used. Without any price, the credit percentage is returned and
    EstimatedCredit stays empty.

    With -OutputFolder, the function also writes the incidents, the estimate, the affected users per country,
    the post-incident reports that are already published, an incident report in Markdown (incident-report.md:
    timeline, impact, Microsoft updates and estimate) and a claim text to paste into the support request
    (claim-request.txt, Microsoft 365 admin center > Support > New service request).

    The function assumes every affected user was affected for the whole incident window: the estimate is a
    ceiling, Microsoft computes the final credit from its own user-minute measurements.

    .PARAMETER IncidentId
    One or more Service Health issue IDs (e.g. EX1550543, MO1550821).

    .PARAMETER KnownIncident
    Name of a major outage listed in the module (e.g. Exchange-2026-11-12): its Service Health IDs are used,
    they are the same in every tenant. Press Tab after -KnownIncident to see the names. Can be combined with
    -IncidentId.

    .PARAMETER ListIncidents
    Lists the Service Health incidents that started since -Since, with their service and duration, to pick the
    IDs to pass to -IncidentId. Advisories are left out unless -IncludeAdvisory is set.

    .PARAMETER Since
    With -ListIncidents, the date from which incidents are listed. Default: 30 days ago.

    .PARAMETER IncludeAdvisory
    With -ListIncidents, also lists the advisories (minor issues, usually not eligible for a credit).

    .PARAMETER ServicePrice
    Monthly price per user of each affected service, keyed by the Service Health service name
    (e.g. @{ 'Exchange Online' = 4; 'Microsoft Teams' = 3.2 }). Overrides the public list price of the service.

    .PARAMETER AffectedUserCount
    Number of affected users per service, keyed by the Service Health service name, when it differs from the
    licensed users (e.g. @{ 'Exchange Online' = 250 }).

    .PARAMETER Currency
    Currency of the prices, shown in the output and in the claim text. Default: EUR.

    .PARAMETER OutputFolder
    Folder for the CSV files, the post-incident reports and the claim text. Created when missing.

    .PARAMETER ExportToExcel
    Exports the estimate to an Excel file in the user profile instead of returning it, with a second worksheet
    for the licensed users per country (when the users are counted from the tenant, not from -AffectedUserCount).

    .EXAMPLE
    Get-MgSlaCreditEstimate -KnownIncident 'Exchange-2026-11-12' -ServicePrice @{ 'Exchange Online' = 4 }

    Estimate for the Exchange Online outage of November 12 to 13, 2026, without looking up its incident IDs.

    .EXAMPLE
    Get-MgSlaCreditEstimate -ListIncidents -Since '2026-11-01'

    Lists the incidents that started since November 1, 2026, longest first.

    .EXAMPLE
    Get-MgSlaCreditEstimate -IncidentId 'EX1550543', 'MO1550821' -ServicePrice @{ 'Exchange Online' = 4 }

    Returns the downtime, uptime, credit tier and estimated credit per service and month.

    .EXAMPLE
    Get-MgSlaCreditEstimate -IncidentId 'EX1550543' -ServicePrice @{ 'Exchange Online' = 4 } -OutputFolder 'C:\Temp\SLA-claim'

    Same estimate, plus the CSV files, the post-incident report when published, incident-report.md and claim-request.txt in C:\Temp\SLA-claim.

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
    [CmdletBinding(DefaultParameterSetName = 'Estimate')]
    param (
        [Parameter(Mandatory = $false, ParameterSetName = 'Estimate')]
        [string[]]$IncidentId,

        [Parameter(Mandatory = $false, ParameterSetName = 'Estimate')]
        [ArgumentCompleter({
                param ($commandName, $parameterName, $wordToComplete)
                $catalogPath = Join-Path (Get-Module PS365).ModuleBase 'Private' 'Assets' 'Get-MgSlaCreditEstimate-known-incidents.json'
                if (Test-Path $catalogPath) {
                    foreach ($entry in (Get-Content $catalogPath -Raw | ConvertFrom-Json)) {
                        if ($entry.Name -like "$wordToComplete*") {
                            [System.Management.Automation.CompletionResult]::new($entry.Name, $entry.Name, 'ParameterValue', $entry.Description)
                        }
                    }
                }
            })]
        [string[]]$KnownIncident,

        [Parameter(Mandatory = $true, ParameterSetName = 'List')]
        [switch]$ListIncidents,

        [Parameter(Mandatory = $false, ParameterSetName = 'List')]
        [datetime]$Since = (Get-Date).AddDays(-30),

        [Parameter(Mandatory = $false, ParameterSetName = 'List')]
        [switch]$IncludeAdvisory,

        [Parameter(Mandatory = $false, ParameterSetName = 'Estimate')]
        [hashtable]$ServicePrice = @{},

        [Parameter(Mandatory = $false, ParameterSetName = 'Estimate')]
        [hashtable]$AffectedUserCount = @{},

        [Parameter(Mandatory = $false, ParameterSetName = 'Estimate')]
        [string]$Currency = 'EUR',

        [Parameter(Mandatory = $false, ParameterSetName = 'Estimate')]
        [string]$OutputFolder,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    $requiredScopes = @('ServiceHealth.Read.All', 'User.Read.All')
    if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    # Dates come back as DateTime or string depending on the PowerShell version
    function ConvertTo-SlaUtcDate {
        param ($Value)

        if (($null -eq $Value) -or ([string]$Value -eq '')) { return $null }
        if ($Value -is [datetime]) { return $Value.ToUniversalTime() }

        $styles = [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal
        return [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, $styles)
    }

    #region List mode: the incidents to choose from
    if ($ListIncidents.IsPresent) {
        $sinceUtc = $Since.ToUniversalTime()
        $nowUtc = (Get-Date).ToUniversalTime()
        [System.Collections.Generic.List[PSCustomObject]]$candidates = @()

        foreach ($issue in (Invoke-MgGraphPagedRequest -Uri '/v1.0/admin/serviceAnnouncement/issues?$select=id,title,service,classification,status,startDateTime,endDateTime')) {
            if ((-not $IncludeAdvisory.IsPresent) -and ($issue.classification -ne 'incident')) { continue }

            $start = ConvertTo-SlaUtcDate $issue.startDateTime
            if ((-not $start) -or ($start -lt $sinceUtc)) { continue }

            $end = ConvertTo-SlaUtcDate $issue.endDateTime
            $isOngoing = -not $end
            $durationEnd = if ($isOngoing) { $nowUtc } else { $end }

            $candidates.Add([PSCustomObject][ordered]@{
                    IncidentId     = $issue.id
                    Service        = $issue.service
                    Title          = $issue.title
                    Classification = $issue.classification
                    Status         = $issue.status
                    StartUtc       = $start
                    EndUtc         = $end
                    Ongoing        = $isOngoing
                    DurationHours  = [math]::Round(($durationEnd - $start).TotalHours, 1)
                })
        }

        $candidates = [System.Collections.Generic.List[PSCustomObject]]@($candidates | Sort-Object DurationHours -Descending)

        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $excelFilePath = "$($env:userprofile)\$now-MgSlaIncidents.xlsx"
            Write-Host -ForegroundColor Cyan 'Exporting to Excel...'
            $candidates | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -TableStyle Light9 -WorksheetName 'MgSlaIncidents'
            Add-PS365ExcelAboutSheet -Path $excelFilePath
            Write-Host -ForegroundColor Green 'Excel report exported to:'
            Write-Host -ForegroundColor White "  $excelFilePath"
            return
        }
        return $candidates
    }
    #endregion

    #region Known incidents: names resolved to their Service Health IDs
    [System.Collections.Generic.List[String]]$allIncidentIds = @()
    foreach ($id in $IncidentId) { $allIncidentIds.Add($id) }

    if ($KnownIncident) {
        $catalogPath = Join-Path $PSScriptRoot '..' '..' 'Private' 'Assets' 'Get-MgSlaCreditEstimate-known-incidents.json'
        $catalog = @(Get-Content -Path $catalogPath -Raw | ConvertFrom-Json)

        foreach ($name in $KnownIncident) {
            $entry = $catalog | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            if (-not $entry) {
                $knownNames = ($catalog.Name -join ', ')
                Write-Warning "Unknown incident name '$name'. Known names: $knownNames"
                continue
            }
            foreach ($id in $entry.IncidentIds) {
                if (-not $allIncidentIds.Contains($id)) { $allIncidentIds.Add($id) }
            }
        }
    }

    if ($allIncidentIds.Count -eq 0) {
        Write-Warning 'Pass -IncidentId, -KnownIncident or -ListIncidents'
        return
    }
    #endregion

    # Service Health service name -> 'service' value of the users' assignedPlans
    $servicePlanMap = @{
        'Exchange Online'      = 'exchange'
        'SharePoint Online'    = 'SharePoint'
        'OneDrive for Business' = 'SharePoint'
        'Microsoft Teams'      = 'TeamspaceAPI'
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

    # Public list prices of the standalone plans, used when -ServicePrice has no price for a service.
    # A service can have several plans (Exchange Online Plan 1 and Plan 2): each one names its service plan ID,
    # matched against the users' assignedPlans, and one of them is the Default (users with another plan, or a
    # user count given with -AffectedUserCount).
    $listPrices = @{}
    $listPricePath = Join-Path $PSScriptRoot '..' '..' 'Private' 'Assets' 'Get-MgSlaCreditEstimate-list-prices.json'
    if (Test-Path -Path $listPricePath) {
        foreach ($entry in @(Get-Content -Path $listPricePath -Raw | ConvertFrom-Json)) {
            if ($entry.Currency -eq $Currency) {
                if (-not $listPrices.ContainsKey($entry.Service)) {
                    $listPrices[$entry.Service] = [System.Collections.Generic.List[PSCustomObject]]::new()
                }
                $listPrices[$entry.Service].Add($entry)
            }
        }
    }

    function Get-SlaListPriceSource {
        param ($Entry)
        return "List price $($Entry.Plan), $($Entry.Market), checked $($Entry.CheckedOn)"
    }

    function Get-SlaDefaultListPrice {
        param ([string]$Service)
        $default = @($listPrices[$Service] | Where-Object { $_.Default }) | Select-Object -First 1
        if (-not $default) {
            $default = $listPrices[$Service] | Select-Object -First 1
        }
        return $default
    }

    #region Incidents
    [System.Collections.Generic.List[PSCustomObject]]$incidents = @()
    $issueById = @{}
    $nowUtc = (Get-Date).ToUniversalTime()

    foreach ($id in $allIncidentIds) {
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

        $issueById[$issue.id] = $issue

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
    # Service -> users per priced plan, when the list prices of the service distinguish several plans
    $pricingPerService = @{}
    [System.Collections.Generic.List[PSCustomObject]]$usersPerCountry = @()

    if ($servicesToCount.Count -gt 0) {
        $users = @(Invoke-MgGraphPagedRequest -Uri '/v1.0/users?$filter=accountEnabled eq true&$select=id,usageLocation,assignedPlans&$top=999')

        foreach ($service in $servicesToCount) {
            $planService = $servicePlanMap[$service]
            $serviceUsers = @($users | Where-Object {
                    @($_.assignedPlans | Where-Object { ($_.service -eq $planService) -and ($_.capabilityStatus -eq 'Enabled') }).Count -gt 0
                })
            $usersPerService[$service] = $serviceUsers.Count

            # Several priced plans for this service (Exchange Online Plan 1 / Plan 2): each user is priced at
            # the most expensive plan they hold (a suite such as Microsoft 365 E3 holds the Plan 2 service plan)
            $planEntries = @($listPrices[$service] | Where-Object { $_.ServicePlanId })
            if ((-not $ServicePrice.ContainsKey($service)) -and ($planEntries.Count -gt 0)) {
                $defaultEntry = Get-SlaDefaultListPrice -Service $service
                $planCounts = [ordered]@{}

                foreach ($serviceUser in $serviceUsers) {
                    $enabledPlanIds = @($serviceUser.assignedPlans | Where-Object { ($_.service -eq $planService) -and ($_.capabilityStatus -eq 'Enabled') } | ForEach-Object { [string]$_.servicePlanId })
                    $matchedEntry = @($planEntries | Where-Object { $_.ServicePlanId -in $enabledPlanIds }) | Sort-Object -Property { [double]$_.Price } -Descending | Select-Object -First 1
                    $planKey = if ($matchedEntry) { $matchedEntry.Plan } else { '' }

                    if (-not $planCounts.Contains($planKey)) {
                        $planCounts[$planKey] = [PSCustomObject]@{ Entry = $matchedEntry; Users = 0 }
                    }
                    $planCounts[$planKey].Users++
                }

                [System.Collections.Generic.List[PSCustomObject]]$pricing = @()
                foreach ($planKey in $planCounts.Keys) {
                    $planCount = $planCounts[$planKey]
                    if ($planCount.Entry) {
                        $pricingItem = [PSCustomObject]@{
                            Plan        = $planCount.Entry.Plan
                            Price       = [double]$planCount.Entry.Price
                            PriceSource = Get-SlaListPriceSource -Entry $planCount.Entry
                            Users       = $planCount.Users
                        }
                    }
                    else {
                        $pricingItem = [PSCustomObject]@{
                            Plan        = "Other $service plans"
                            Price       = [double]$defaultEntry.Price
                            PriceSource = "$(Get-SlaListPriceSource -Entry $defaultEntry) (no priced plan matched)"
                            Users       = $planCount.Users
                        }
                    }
                    $pricing.Add($pricingItem)
                }
                $pricingPerService[$service] = $pricing
            }

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

        # One pricing line per priced plan: a single line with -ServicePrice, or when the list price has one plan
        $userCount = $usersPerService[$service]
        if ($ServicePrice.ContainsKey($service)) {
            $pricing = @([PSCustomObject]@{ Plan = $null; Price = [double]$ServicePrice[$service]; PriceSource = '-ServicePrice'; Users = $userCount })
        }
        elseif ($pricingPerService.ContainsKey($service)) {
            $pricing = @($pricingPerService[$service])
        }
        elseif ($listPrices.ContainsKey($service)) {
            $defaultEntry = Get-SlaDefaultListPrice -Service $service
            $pricing = @([PSCustomObject]@{ Plan = $defaultEntry.Plan; Price = [double]$defaultEntry.Price; PriceSource = Get-SlaListPriceSource -Entry $defaultEntry; Users = $userCount })
        }
        else {
            $pricing = @([PSCustomObject]@{ Plan = $null; Price = $null; PriceSource = $null; Users = $userCount })
        }

        $listPricedPlans = @($pricing | Where-Object { $_.PriceSource -like 'List price*' } | ForEach-Object { "$($_.Plan) $($_.Price) $Currency" })
        if ($listPricedPlans.Count -gt 0) {
            Write-Warning "'$service': public list price used ($($listPricedPlans -join ', '), excluding VAT). For a suite licence or a negotiated price, pass -ServicePrice."
        }

        foreach ($key in $monthly.Keys) {
            $month = $monthly[$key]
            $monthMinutes = ($month.MonthStart.AddMonths(1) - $month.MonthStart).TotalMinutes
            $uptime = [math]::Round(($monthMinutes - $month.Minutes) / $monthMinutes * 100, 3)
            $creditPercent = Get-SlaCreditPercent -Uptime $uptime

            foreach ($pricingItem in $pricing) {
                $credit = $null
                if (($null -ne $pricingItem.Price) -and ($null -ne $pricingItem.Users)) {
                    $credit = [math]::Round($pricingItem.Users * $pricingItem.Price * $creditPercent / 100, 2)
                }

                $estimateRow = [PSCustomObject][ordered]@{
                    Service         = $service
                    Plan            = $pricingItem.Plan
                    Month           = $key
                    IncidentIds     = (@($month.IncidentIds | Select-Object -Unique) -join ', ')
                    DowntimeHours   = [math]::Round($month.Minutes / 60, 1)
                    UptimePercent   = $uptime
                    CreditPercent   = $creditPercent
                    AffectedUsers   = $pricingItem.Users
                    PricePerUser    = $pricingItem.Price
                    PriceSource     = $pricingItem.PriceSource
                    EstimatedCredit = $credit
                    Currency        = $Currency
                }
                $estimate.Add($estimateRow)
            }
        }

        if ($null -eq $pricing[0].Price) {
            Write-Warning "No price for '$service': pass it with -ServicePrice to get the credit amount"
        }
    }
    #endregion

    #region Files and claim text
    if ($OutputFolder) {
        # pwsh 7 never writes the type line: -NoTypeInformation is obsolete there and only prints a warning
        $csvParams = @{ Encoding = 'UTF8' }

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
        # One line per service and month: the per-plan rows only matter for the price, not for the claim
        foreach ($group in ($estimate | Group-Object -Property Service, Month)) {
            $row = $group.Group[0]
            $knownUsers = @($group.Group | Where-Object { $null -ne $_.AffectedUsers })
            $usersText = if ($knownUsers.Count -gt 0) { "$(($knownUsers | Measure-Object -Property AffectedUsers -Sum).Sum) affected users" } else { 'affected users to complete' }
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

        # Incident report in Markdown: timeline, impact, Microsoft updates and estimate, for internal follow-up
        $toText = {
            param ($Html)
            if (-not $Html) { return '' }
            $text = [string]$Html -replace '(?i)<br\s*/?>', "`n" -replace '(?i)</(p|div|li)>', "`n" -replace '<[^>]+>', ''
            $text = [System.Net.WebUtility]::HtmlDecode($text)
            return (($text -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join "`n`n")
        }

        [System.Collections.Generic.List[String]]$markdown = @()
        $markdown.Add('# Microsoft 365 incident report')
        $markdown.Add('')
        $markdown.Add("Tenant ID: $tenantId  ")
        $markdown.Add("Generated: $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm')) UTC")
        $markdown.Add('')
        $markdown.Add('## Incidents')
        $markdown.Add('')
        $markdown.Add('| ID | Service | Title | Start (UTC) | End (UTC) | Hours | Status |')
        $markdown.Add('|---|---|---|---|---|---|---|')
        foreach ($incident in $incidents) {
            $endText = if ($incident.Ongoing) { 'still open' } else { $incident.EndUtc.ToString('yyyy-MM-dd HH:mm') }
            $titleText = ([string]$incident.Title) -replace '\|', '\|'
            $markdown.Add("| $($incident.IncidentId) | $($incident.Service) | $titleText | $($incident.StartUtc.ToString('yyyy-MM-dd HH:mm')) | $endText | $($incident.DurationHours) | $($incident.Status) |")
        }

        $markdown.Add('')
        $markdown.Add('## Estimated credit')
        $markdown.Add('')
        $markdown.Add('| Service | Plan | Month | Downtime (h) | Uptime % | Credit % | Users | Price/user | Credit |')
        $markdown.Add('|---|---|---|---|---|---|---|---|---|')
        foreach ($row in $estimate) {
            $markdown.Add("| $($row.Service) | $($row.Plan) | $($row.Month) | $($row.DowntimeHours) | $($row.UptimePercent) | $($row.CreditPercent) | $($row.AffectedUsers) | $($row.PricePerUser) | $($row.EstimatedCredit) $($row.Currency) |")
        }
        $totalCredit = ($estimate | Where-Object { $null -ne $_.EstimatedCredit } | Measure-Object -Property EstimatedCredit -Sum).Sum
        $markdown.Add('')
        $markdown.Add("Total estimated credit: **$([math]::Round([double]$totalCredit, 2)) $Currency** (ceiling: every user counted as affected for the whole window).")
        $priceSources = @($estimate | Where-Object { $_.PriceSource } | ForEach-Object { "- $($_.Service), $($_.Plan): $($_.PriceSource)" } | Select-Object -Unique)
        if ($priceSources.Count -gt 0) {
            $markdown.Add('')
            $markdown.Add('Prices used:')
            $markdown.Add('')
            foreach ($priceLine in $priceSources) { $markdown.Add($priceLine) }
        }

        foreach ($incident in $incidents) {
            $issue = $issueById[$incident.IncidentId]
            $markdown.Add('')
            $markdown.Add("## $($incident.IncidentId) - $($incident.Title)")
            $markdown.Add('')
            $markdown.Add("- Service: $($incident.Service)")
            if ($issue.feature) { $markdown.Add("- Feature: $($issue.featureGroup) / $($issue.feature)") }
            $markdown.Add("- Classification: $($incident.Classification), status: $($incident.Status)")
            if ($incident.PostIncidentFile) { $markdown.Add("- Post-incident report: $(Split-Path -Path $incident.PostIncidentFile -Leaf)") }
            if ($issue.impactDescription) {
                $markdown.Add('')
                $markdown.Add('### Impact')
                $markdown.Add('')
                $markdown.Add((& $toText $issue.impactDescription))
            }

            $posts = @($issue.posts | Sort-Object { ConvertTo-SlaUtcDate $_.createdDateTime })
            if ($posts.Count -gt 0) {
                $markdown.Add('')
                $markdown.Add('### Microsoft updates')
                foreach ($post in $posts) {
                    $postDate = ConvertTo-SlaUtcDate $post.createdDateTime
                    $markdown.Add('')
                    $markdown.Add("#### $($postDate.ToString('yyyy-MM-dd HH:mm')) UTC")
                    $markdown.Add('')
                    $markdown.Add((& $toText $post.description.content))
                }
            }
        }

        $markdown -join [Environment]::NewLine | Out-File -FilePath (Join-Path $OutputFolder 'incident-report.md') -Encoding UTF8
        Write-Verbose "Files written to $OutputFolder"
    }
    #endregion

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgSlaCreditEstimate.xlsx"
        Write-Host -ForegroundColor Cyan 'Exporting to Excel...'
        $estimate | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -TableStyle Light9 -WorksheetName 'MgSlaCreditEstimate' -TableName 'SlaCreditEstimate'
        if ($usersPerCountry.Count -gt 0) {
            $usersPerCountry | Sort-Object Service, @{ Expression = 'Users'; Descending = $true } |
            Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -TableStyle Light9 -WorksheetName 'MgSlaUsersPerCountry' -TableName 'SlaUsersPerCountry'
        }
        Add-PS365ExcelAboutSheet -Path $excelFilePath
        Write-Host -ForegroundColor Green 'Excel report exported to:'
        Write-Host -ForegroundColor White "  $excelFilePath"
    }
    else {
        return $estimate
    }
}
