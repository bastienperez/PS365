<#
    .SYNOPSIS
    Checks the health of Microsoft Entra Connect Sync on the sync server: service, scheduler, last runs, objects in error, pending exports, password hash sync and event log.

    .DESCRIPTION
    Runs a set of local checks on the Entra Connect Sync server and returns one row per check, with a status
    (OK, Info, Warning or Error), the observed value and a detail. Nothing is read from Microsoft Entra ID: every
    check uses the ADSync module, csexport.exe and the local event log.

    Checks:
    - Service: ADSync service state, Entra Connect version, auto-upgrade state.
    - Scheduler: sync cycle enabled, staging mode, maintenance, cycle in progress, next cycle overdue.
    - Run: result of the last run of each run profile per connector (Delta Import, Delta Synchronization,
      Export...) and age of the last run of each connector. A failing import (connector unreachable, expired
      service account password) shows up here: it is a run result, not an object error.
    - Objects: number of objects in import/synchronization error and in export error per connector, with the
      most frequent error types. Get-EntraConnectSyncObjectReport -ErrorsOnly lists them one by one.
    - Pending export: number of objects waiting to be exported per connector.
    - Password hash sync: enabled on the tenant and per AD connector, sync failures (event 611).
    - Event log: errors of the ADSync and Directory Synchronization sources.

    With -LastRun, the function returns the run history instead: for each connector, the last run of each run
    profile, as shown in the Operations tab of the Synchronization Service Manager, with its counters.

    Must be run on the Entra Connect Sync server (ADSync module) with a member of the ADSyncAdmins group.

    .PARAMETER ConnectorName
    Name of one or more connectors to check. By default, all connectors.

    .PARAMETER LastRun
    Returns the last run of each run profile per connector, with its result, duration and counters, instead of the health checks.

    .PARAMETER RunHistoryCount
    Number of recent runs read per connector to find the last run of each profile. Default 50.

    .PARAMETER MaxRunAgeHours
    A connector whose last run is older than this number of hours is reported as Warning. Default 3.

    .PARAMETER PendingExportThreshold
    A connector with at least this number of pending exports is reported as Warning. Default 500.

    .PARAMETER EventLogHours
    Number of hours of event log read for the event log and password hash sync checks. Default 24.

    .PARAMETER ExportToExcel
    Exports the result to an Excel file in the user profile instead of returning it.

    .EXAMPLE
    Get-EntraConnectSyncHealth

    Runs every check on all connectors.

    .EXAMPLE
    Get-EntraConnectSyncHealth | Where-Object { $_.Status -in 'Warning', 'Error' }

    Shows only the checks that need attention.

    .EXAMPLE
    Get-EntraConnectSyncHealth -LastRun | Format-Table ConnectorName, RunProfileName, Result, StartDate, DurationSeconds, TotalErrors

    Shows the last run of each run profile on every connector.

    .EXAMPLE
    Get-EntraConnectSyncHealth -ExportToExcel

    Exports the health checks to an Excel file.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-EntraConnectSyncHealth

    .NOTES
    Requires the ADSync module, installed with Microsoft Entra Connect Sync.
#>

function Get-EntraConnectSyncHealth {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$ConnectorName,

        [Parameter(Mandatory = $false)]
        [switch]$LastRun,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 1000)]
        [int]$RunHistoryCount = 50,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 720)]
        [int]$MaxRunAgeHours = 3,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$PendingExportThreshold = 500,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 720)]
        [int]$EventLogHours = 24,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    if (-not (Get-Command -Name 'Get-ADSyncScheduler' -ErrorAction SilentlyContinue)) {
        throw 'The ADSync module is not available. Run this function on the Microsoft Entra Connect Sync server.'
    }

    [System.Collections.Generic.List[PSCustomObject]]$healthChecks = @()
    [System.Collections.Generic.List[PSCustomObject]]$runRows = @()

    function Add-HealthCheck {
        param ([string]$Category, [string]$Check, [string]$Target, [string]$Status, $Value, [string]$Detail)

        $object = [PSCustomObject][ordered]@{
            Category = $Category
            Check    = $Check
            Target   = $Target
            Status   = $Status
            Value    = $Value
            Detail   = $Detail
        }
        $healthChecks.Add($object)
    }

    # Property names of the run results are not documented: read them defensively (missing = $null)
    function Get-RunValue {
        param ($InputObject, [string]$Name)
        if (($null -ne $InputObject) -and ($InputObject.PSObject.Properties[$Name])) {
            return $InputObject.$Name
        }
        return $null
    }

    # The engine may return local or UTC dates: compare with "now" of the same kind
    function Get-AgeHours {
        param ([datetime]$Date)
        $now = if ($Date.Kind -eq [System.DateTimeKind]::Utc) { [datetime]::UtcNow } else { Get-Date }
        return [math]::Round(($now - $Date).TotalHours, 1)
    }

    #region Service
    if (-not $LastRun) {
        $service = Get-Service -Name 'ADSync' -ErrorAction SilentlyContinue
        if ($null -eq $service) {
            Add-HealthCheck -Category 'Service' -Check 'ADSync service' -Status 'Error' -Value 'Not found' -Detail 'The Microsoft Azure AD Sync service is not installed on this server.'
        }
        else {
            $serviceStatus = if ($service.Status -eq 'Running') { 'OK' } else { 'Error' }
            Add-HealthCheck -Category 'Service' -Check 'ADSync service' -Status $serviceStatus -Value ([string]$service.Status) -Detail "Start type: $($service.StartType)"
        }

        $adSyncModule = @(Get-Module -Name ADSync) + @(Get-Module -Name ADSync -ListAvailable) | Select-Object -First 1
        if ($adSyncModule) {
            $miiServerPath = Join-Path -Path (Split-Path -Path $adSyncModule.ModuleBase -Parent) -ChildPath 'miiserver.exe'
            if (Test-Path -Path $miiServerPath) {
                $version = (Get-Item -Path $miiServerPath).VersionInfo.ProductVersion
                Add-HealthCheck -Category 'Service' -Check 'Entra Connect version' -Status 'Info' -Value $version -Detail 'Compare with the version history: https://learn.microsoft.com/entra/identity/hybrid/connect/reference-connect-version-history'
            }
        }

        try {
            $autoUpgrade = [string](Get-ADSyncAutoUpgrade -ErrorAction Stop)
            $autoUpgradeStatus = switch ($autoUpgrade) {
                'Enabled' { 'OK' }
                'Suspended' { 'Warning' }
                default { 'Info' }
            }
            $autoUpgradeDetail = switch ($autoUpgrade) {
                'Suspended' { 'Auto-upgrade was suspended by the engine or by an administrator: upgrade manually or check Get-ADSyncAutoUpgrade -Detail.' }
                'Disabled' { 'Upgrades are manual.' }
                default { $null }
            }
            Add-HealthCheck -Category 'Service' -Check 'Auto-upgrade' -Status $autoUpgradeStatus -Value $autoUpgrade -Detail $autoUpgradeDetail
        }
        catch {
            Add-HealthCheck -Category 'Service' -Check 'Auto-upgrade' -Status 'Warning' -Value 'Not readable' -Detail $_.Exception.Message
        }
    }
    #endregion

    #region Scheduler
    $scheduler = $null
    if (-not $LastRun) {
        try {
            $scheduler = Get-ADSyncScheduler -ErrorAction Stop
        }
        catch {
            Add-HealthCheck -Category 'Scheduler' -Check 'Scheduler' -Status 'Error' -Value 'Not readable' -Detail $_.Exception.Message
        }

        if ($scheduler) {
            if ($scheduler.SyncCycleEnabled) {
                Add-HealthCheck -Category 'Scheduler' -Check 'Sync cycle enabled' -Status 'OK' -Value $true
            }
            else {
                Add-HealthCheck -Category 'Scheduler' -Check 'Sync cycle enabled' -Status 'Error' -Value $false -Detail 'No automatic sync. Re-enable with Set-ADSyncScheduler -SyncCycleEnabled $true.'
            }

            if ($scheduler.StagingModeEnabled) {
                Add-HealthCheck -Category 'Scheduler' -Check 'Staging mode' -Status 'Warning' -Value $true -Detail 'This server imports and synchronizes but exports nothing: expected only on a standby server.'
            }
            else {
                Add-HealthCheck -Category 'Scheduler' -Check 'Staging mode' -Status 'OK' -Value $false
            }

            if ($scheduler.MaintenanceEnabled) {
                Add-HealthCheck -Category 'Scheduler' -Check 'Maintenance enabled' -Status 'OK' -Value $true
            }
            else {
                Add-HealthCheck -Category 'Scheduler' -Check 'Maintenance enabled' -Status 'Warning' -Value $false -Detail 'Maintenance tasks (run history cleanup, password sync key rollover...) are disabled.'
            }

            Add-HealthCheck -Category 'Scheduler' -Check 'Sync cycle in progress' -Status 'Info' -Value ([bool]$scheduler.SyncCycleInProgress)

            $interval = $scheduler.CurrentlyEffectiveSyncCycleInterval
            $nextCycleUtc = $scheduler.NextSyncCycleStartTimeInUTC
            if ($nextCycleUtc -is [datetime]) {
                $nextCycleLocal = [datetime]::SpecifyKind($nextCycleUtc, [System.DateTimeKind]::Utc).ToLocalTime()
                $nextCycleStatus = 'OK'
                $nextCycleDetail = "Interval: $interval, next policy: $($scheduler.NextSyncCyclePolicyType)"

                # A next cycle far in the past means the scheduler is not running cycles any more
                $tolerance = if ($interval -is [timespan]) { $interval + $interval } else { [timespan]::FromHours(1) }
                if (($scheduler.SyncCycleEnabled) -and (-not $scheduler.SyncCycleInProgress) -and ($nextCycleUtc -lt [datetime]::UtcNow.Subtract($tolerance))) {
                    $nextCycleStatus = 'Warning'
                    $nextCycleDetail = "Next cycle overdue. $nextCycleDetail"
                }
                Add-HealthCheck -Category 'Scheduler' -Check 'Next sync cycle' -Status $nextCycleStatus -Value $nextCycleLocal -Detail $nextCycleDetail
            }
        }
    }
    #endregion

    #region Connectors
    try {
        $connectors = @(Get-ADSyncConnector -ErrorAction Stop)
    }
    catch {
        Add-HealthCheck -Category 'Connector' -Check 'Connectors' -Status 'Error' -Value 'Not readable' -Detail $_.Exception.Message
        $connectors = @()
    }

    if ($ConnectorName) {
        $unknownNames = @($ConnectorName | Where-Object { $_ -notin $connectors.Name })
        if ($unknownNames.Count -gt 0) {
            Write-Warning "[!] Connector not found: $($unknownNames -join ', ')"
        }
        $connectors = @($connectors | Where-Object { $_.Name -in $ConnectorName })
    }
    #endregion

    #region Runs
    $stepCounters = @(
        'StageNoChange', 'StageAdd', 'StageUpdate', 'StageRename', 'StageDelete', 'StageDeleteAdd', 'StageFailure',
        'ExportAdd', 'ExportUpdate', 'ExportRename', 'ExportDelete', 'ExportDeleteAdd', 'ExportFailure',
        'FlowFailure'
    )

    foreach ($connector in $connectors) {
        Write-Verbose "[i] Reading run history of $($connector.Name)"

        try {
            $runs = @(Get-ADSyncRunProfileResult -ConnectorId $connector.Identifier -NumberRequested $RunHistoryCount -RunStepDetails -ErrorAction Stop)
        }
        catch {
            Add-HealthCheck -Category 'Run' -Check 'Run history' -Target $connector.Name -Status 'Warning' -Value 'Not readable' -Detail $_.Exception.Message
            continue
        }

        if ($runs.Count -eq 0) {
            Add-HealthCheck -Category 'Run' -Check 'Run history' -Target $connector.Name -Status 'Warning' -Value 'No run' -Detail "No run in the last $RunHistoryCount runs of the history."
            continue
        }

        # The last run of each run profile
        $lastRuns = foreach ($profileGroup in ($runs | Group-Object -Property RunProfileName)) {
            $profileGroup.Group | Sort-Object -Property StartDate -Descending | Select-Object -First 1
        }

        foreach ($run in ($lastRuns | Sort-Object -Property StartDate -Descending)) {
            $startDate = Get-RunValue -InputObject $run -Name 'StartDate'
            $endDate = Get-RunValue -InputObject $run -Name 'EndDate'
            $durationSeconds = $null
            if (($startDate -is [datetime]) -and ($endDate -is [datetime])) {
                $durationSeconds = [math]::Round(($endDate - $startDate).TotalSeconds)
            }

            $runProfileName = Get-RunValue -InputObject $run -Name 'RunProfileName'
            $runResult = [string](Get-RunValue -InputObject $run -Name 'Result')
            $steps = @(Get-RunValue -InputObject $run -Name 'RunStepResults')

            $row = [ordered]@{
                ConnectorName   = $connector.Name
                ConnectorType   = $connector.Type
                RunProfileName  = $runProfileName
                RunNumber       = Get-RunValue -InputObject $run -Name 'RunNumber'
                Result          = $runResult
                StartDate       = $startDate
                EndDate         = $endDate
                DurationSeconds = $durationSeconds
                StepResults     = (@($steps | ForEach-Object { Get-RunValue -InputObject $_ -Name 'StepResult' } | Where-Object { $_ }) -join ', ')
            }

            $totalErrors = 0
            foreach ($counter in $stepCounters) {
                $sum = 0
                foreach ($step in $steps) {
                    $value = Get-RunValue -InputObject $step -Name $counter
                    if ($value -as [int]) {
                        $sum += [int]$value
                    }
                }
                $row[$counter] = $sum
                if ($counter -match 'Failure$') {
                    $totalErrors += $sum
                }
            }
            $row.TotalErrors = $totalErrors

            $runRows.Add([PSCustomObject]$row)

            # success: OK. completed-*-errors / completed-*-warnings: the run ended but some objects failed.
            # stopped-* (connectivity, server-down, extension-dll-exception...): the run itself failed.
            $runStatus = switch -Regex ($runResult) {
                '^success$' { 'OK' }
                '^in-progress$' { 'Info' }
                '^completed-' { 'Warning' }
                '^(stopped-|failed)' { 'Error' }
                default { 'Warning' }
            }

            $failureCounters = @($stepCounters | Where-Object { ($_ -match 'Failure$') -and ($row[$_] -gt 0) } | ForEach-Object { "$_ $($row[$_])" })
            $runDetail = "Started $startDate"
            if ($null -ne $durationSeconds) {
                $runDetail += ", $durationSeconds s"
            }
            if ($failureCounters.Count -gt 0) {
                $runDetail += ". $($failureCounters -join ', ')"
            }

            Add-HealthCheck -Category 'Run' -Check "Last $runProfileName" -Target $connector.Name -Status $runStatus -Value $runResult -Detail $runDetail
        }

        $lastStart = ($runs | Where-Object { $_.StartDate -is [datetime] } | Sort-Object -Property StartDate -Descending | Select-Object -First 1).StartDate
        if ($lastStart -is [datetime]) {
            $ageHours = Get-AgeHours -Date $lastStart
            $ageStatus = if ($ageHours -gt $MaxRunAgeHours) { 'Warning' } else { 'OK' }
            Add-HealthCheck -Category 'Run' -Check 'Last run age (hours)' -Target $connector.Name -Status $ageStatus -Value $ageHours -Detail "Last run started $lastStart (threshold $MaxRunAgeHours h)"
        }
    }
    #endregion

    if ($LastRun) {
        $result = $runRows
    }
    else {
        #region Objects in error and pending exports
        foreach ($connector in $connectors) {
            Write-Verbose "[i] Reading objects in error and pending exports of $($connector.Name)"

            $errorRows = @(Get-EntraConnectSyncObjectReport -ErrorsOnly -ConnectorName $connector.Name)

            foreach ($errorKind in 'Synchronization', 'Export') {
                $kindRows = if ($errorKind -eq 'Export') {
                    @($errorRows | Where-Object { $_.HasExportError })
                }
                else {
                    @($errorRows | Where-Object { $_.HasSyncError })
                }
                $typeProperty = if ($errorKind -eq 'Export') { 'ExportErrorType' } else { 'SynchronizationErrorType' }

                $kindStatus = if ($kindRows.Count -gt 0) { 'Warning' } else { 'OK' }
                $kindDetail = $null
                if ($kindRows.Count -gt 0) {
                    $topTypes = @($kindRows | Group-Object -Property $typeProperty | Sort-Object -Property Count -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name) ($($_.Count))" })
                    $kindDetail = "$($topTypes -join ', '). Details: Get-EntraConnectSyncObjectReport -ErrorsOnly -ConnectorName '$($connector.Name)'"
                }

                $checkName = if ($errorKind -eq 'Export') { 'Objects in export error' } else { 'Objects in import/sync error' }
                Add-HealthCheck -Category 'Objects' -Check $checkName -Target $connector.Name -Status $kindStatus -Value $kindRows.Count -Detail $kindDetail
            }

            $pendingExports = Invoke-ADSyncCsExport -ConnectorName $connector.Name -Filter 'x' -CountOnly
            if ($null -eq $pendingExports) {
                Add-HealthCheck -Category 'Pending export' -Check 'Pending exports' -Target $connector.Name -Status 'Warning' -Value 'Not readable' -Detail 'csexport failed: see the warning above.'
            }
            else {
                $pendingStatus = if ($pendingExports -ge $PendingExportThreshold) { 'Warning' } else { 'OK' }
                $pendingDetail = if ($pendingExports -ge $PendingExportThreshold) {
                    "At least $PendingExportThreshold objects waiting for the next export: check that the export runs (staging mode, deletion threshold, stopped run)."
                }
                elseif ($pendingExports -gt 0) {
                    'Objects waiting for the next export run: expected between two cycles.'
                }
                else {
                    $null
                }
                Add-HealthCheck -Category 'Pending export' -Check 'Pending exports' -Target $connector.Name -Status $pendingStatus -Value $pendingExports -Detail $pendingDetail
            }
        }
        #endregion

        #region Event log
        $eventStart = (Get-Date).AddHours(-$EventLogHours)
        [System.Collections.Generic.List[object]]$syncEvents = @()
        foreach ($providerName in 'ADSync', 'Directory Synchronization') {
            # Get-WinEvent fails when nothing matches or when the source is not registered: neither is a sync problem
            try {
                $providerEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = $providerName; Level = 1, 2; StartTime = $eventStart } -ErrorAction Stop)
            }
            catch {
                Write-Verbose "[i] No error event read for $($providerName): $($_.Exception.Message)"
                $providerEvents = @()
            }
            foreach ($providerEvent in $providerEvents) {
                $syncEvents.Add($providerEvent)
            }
        }
        #endregion

        #region Password hash sync
        try {
            $passwordHashSyncCloud = [bool](Get-ADSyncAADCompanyFeature -ErrorAction Stop).PasswordHashSync
            $phsStatus = if ($passwordHashSyncCloud) { 'OK' } else { 'Info' }
            Add-HealthCheck -Category 'Password hash sync' -Check 'Password hash sync enabled' -Status $phsStatus -Value $passwordHashSyncCloud

            if ($passwordHashSyncCloud) {
                foreach ($connector in ($connectors | Where-Object { $_.Type -eq 'AD' })) {
                    $connectorPhs = [bool](Get-ADSyncAADPasswordSyncConfiguration -SourceConnector $connector.Name -ErrorAction Stop).Enabled
                    $connectorPhsStatus = if ($connectorPhs) { 'OK' } else { 'Warning' }
                    Add-HealthCheck -Category 'Password hash sync' -Check 'Password hash sync on connector' -Target $connector.Name -Status $connectorPhsStatus -Value $connectorPhs
                }

                # Event 611: password hash synchronization failed for a domain
                $phsFailures = @($syncEvents | Where-Object { ($_.ProviderName -eq 'Directory Synchronization') -and ($_.Id -eq 611) })
                $phsFailureStatus = if ($phsFailures.Count -gt 0) { 'Error' } else { 'OK' }
                $phsFailureDetail = if ($phsFailures.Count -gt 0) { "Last on $($phsFailures[0].TimeCreated): $((($phsFailures[0].Message) -split '\r?\n')[0])" } else { $null }
                Add-HealthCheck -Category 'Password hash sync' -Check "Sync failures (event 611, last $EventLogHours h)" -Status $phsFailureStatus -Value $phsFailures.Count -Detail $phsFailureDetail
            }
        }
        catch {
            Add-HealthCheck -Category 'Password hash sync' -Check 'Password hash sync' -Status 'Warning' -Value 'Not readable' -Detail $_.Exception.Message
        }
        #endregion

        #region Event log summary
        $eventStatus = if ($syncEvents.Count -gt 0) { 'Warning' } else { 'OK' }
        $eventDetail = $null
        if ($syncEvents.Count -gt 0) {
            $topEvents = @($syncEvents | Group-Object -Property ProviderName, Id | Sort-Object -Property Count -Descending | Select-Object -First 5 | ForEach-Object { "$($_.Name) ($($_.Count))" })
            $latestEvent = $syncEvents | Sort-Object -Property TimeCreated -Descending | Select-Object -First 1
            $eventDetail = "$($topEvents -join '; '). Last: $($latestEvent.TimeCreated) $((($latestEvent.Message) -split '\r?\n')[0])"
        }
        Add-HealthCheck -Category 'Event log' -Check "Errors (last $EventLogHours h)" -Status $eventStatus -Value $syncEvents.Count -Detail $eventDetail
        #endregion

        $result = $healthChecks
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $suffix = if ($LastRun) { 'EntraConnectSyncLastRun' } else { 'EntraConnectSyncHealth' }
        $excelFilePath = "$($env:userprofile)\$now-$suffix.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $result | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -TableStyle Light9 -WorksheetName $suffix
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $result
    }
}
