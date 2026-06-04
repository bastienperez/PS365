function Get-UnifiedAuditLogOperationCatalog {
    <#
    .SYNOPSIS
        Returns the catalog of Unified Audit Log operations used by the Search-UnifiedAuditLogCustom helper GUI.

    .DESCRIPTION
        Loads the list of audit log operations (Operation + friendly name) from the Microsoft Learn
        audit-log-activities page, with a resilient caching strategy:

        1. If a local cache exists and is younger than MaxCacheAgeDays, it is used directly (no network call).
        2. Otherwise the page is fetched from Microsoft Learn and the parsed list is written to the cache.
        3. If the page is unreachable, the (possibly stale) local cache is used as a fallback.
        4. If no cache exists, the bundled seed shipped with the module is used as a last resort.

        The cache lives in the per-user local application data folder (LOCALAPPDATA on Windows,
        ~/Library/Application Support on macOS, ~/.local/share on Linux) under a PS365 subfolder, so it
        is writable per-user and works with PowerShell 7 on any platform. The seed lives in
        Private\Assets and ships with the module, so the helper still works fully offline on first run.

    .PARAMETER Url
        The Microsoft Learn page to parse. Defaults to the audit-log-activities catalog.

    .PARAMETER MaxCacheAgeDays
        Maximum age (in days) of the local cache before a refresh from Microsoft Learn is attempted. Default is 1.

    .PARAMETER ForceRefresh
        Ignore the cache freshness check and always attempt to refresh from Microsoft Learn first.

    .OUTPUTS
        PSCustomObject with:
        - Operations : a List[PSCustomObject] of { FriendlyName, Operation }
        - Source     : 'Live', 'Cache', 'Seed' or 'None'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Url = 'https://learn.microsoft.com/en-us/purview/audit-log-activities',

        [Parameter(Mandatory = $false)]
        [int]$MaxCacheAgeDays = 1,

        [Parameter(Mandatory = $false)]
        [switch]$ForceRefresh
    )

    # Resolve a writable per-user cache folder cross-platform: LOCALAPPDATA on Windows,
    # ~/Library/Application Support on macOS, ~/.local/share on Linux, temp as a last resort.
    $localAppData = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::LocalApplicationData)
    if ([string]::IsNullOrWhiteSpace($localAppData)) {
        $localAppData = [System.IO.Path]::GetTempPath()
    }
    $cacheDirectory = Join-Path $localAppData 'PS365'
    $cachePath = Join-Path $cacheDirectory 'Search-UnifiedAuditLogCustom-operations.json'
    $seedPath = Join-Path $PSScriptRoot 'Assets\Search-UnifiedAuditLogCustom-operations.json'

    # Normalize any collection of objects into a clean List[PSCustomObject] of { FriendlyName, Operation }.
    $normalize = {
        param($Items)

        [System.Collections.Generic.List[PSCustomObject]]$list = @()
        foreach ($item in @($Items)) {
            if ($null -eq $item) {
                continue
            }
            $operation = [string]$item.Operation
            if ([string]::IsNullOrWhiteSpace($operation)) {
                continue
            }
            $friendly = [string]$item.FriendlyName
            if ([string]::IsNullOrWhiteSpace($friendly)) {
                $friendly = $operation
            }
            $list.Add([PSCustomObject]@{
                    FriendlyName = $friendly
                    Operation    = $operation
                })
        }

        return , $list
    }

    # Build the catalog from raw HTML table rows returned by Get-HTMLTables.
    $buildFromRows = {
        param($Rows)

        $seen = @{}
        [System.Collections.Generic.List[PSCustomObject]]$list = @()

        foreach ($row in ($Rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Operation) })) {
            $operationName = ([string]$row.Operation).Trim().TrimEnd('.')
            if ([string]::IsNullOrWhiteSpace($operationName)) {
                continue
            }

            $friendlyName = $null
            if ($row.PSObject.Properties.Match('Friendly name').Count -gt 0) {
                $friendlyName = $row.'Friendly name'
            }
            elseif ($row.PSObject.Properties.Match('FriendlyName').Count -gt 0) {
                $friendlyName = $row.FriendlyName
            }

            if ([string]::IsNullOrWhiteSpace($friendlyName)) {
                $friendlyName = $operationName
            }
            $friendlyName = ([string]$friendlyName).Trim().TrimEnd('.')

            if (-not $seen.ContainsKey($operationName)) {
                $seen[$operationName] = $true
                $list.Add([PSCustomObject]@{
                        FriendlyName = $friendlyName
                        Operation    = $operationName
                    })
            }
        }

        return , $list
    }

    $loadFromJsonFile = {
        param($Path)

        if (-not (Test-Path -LiteralPath $Path)) {
            return $null
        }

        try {
            $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($content)) {
                return $null
            }
            $parsed = $content | ConvertFrom-Json -ErrorAction Stop
            return (& $normalize $parsed)
        }
        catch {
            Write-Verbose "Unable to read operations catalog from '$Path'. $($_.Exception.Message)"
            return $null
        }
    }

    # 1. Fresh cache - use it directly, no network call.
    if (-not $ForceRefresh.IsPresent -and (Test-Path -LiteralPath $cachePath)) {
        try {
            $cacheAgeDays = ((Get-Date) - (Get-Item -LiteralPath $cachePath).LastWriteTime).TotalDays
            if ($cacheAgeDays -le $MaxCacheAgeDays) {
                $cachedOperations = & $loadFromJsonFile $cachePath
                if ($cachedOperations -and $cachedOperations.Count -gt 0) {
                    Write-Verbose "Loaded $($cachedOperations.Count) operations from fresh cache ($([math]::Round($cacheAgeDays, 1)) days old)."
                    return [PSCustomObject]@{ Operations = $cachedOperations; Source = 'Cache' }
                }
            }
        }
        catch {
            Write-Verbose "Cache freshness check failed. $($_.Exception.Message)"
        }
    }

    # 2. Refresh from Microsoft Learn and update the cache.
    try {
        $tableRows = Get-HTMLTables -URL $Url
        if ($tableRows) {
            $liveOperations = & $buildFromRows $tableRows
            if ($liveOperations -and $liveOperations.Count -gt 0) {
                try {
                    if (-not (Test-Path -LiteralPath $cacheDirectory)) {
                        $null = New-Item -ItemType Directory -Path $cacheDirectory -Force -ErrorAction Stop
                    }
                    $liveOperations | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $cachePath -Encoding UTF8 -ErrorAction Stop
                    Write-Verbose "Cached $($liveOperations.Count) operations to '$cachePath'."
                }
                catch {
                    Write-Verbose "Unable to write operations cache. $($_.Exception.Message)"
                }

                return [PSCustomObject]@{ Operations = $liveOperations; Source = 'Live' }
            }
        }
    }
    catch {
        Write-Verbose "Failed to load operation catalog from Microsoft Learn. $($_.Exception.Message)"
    }

    # 3. Fallback - stale cache.
    $staleOperations = & $loadFromJsonFile $cachePath
    if ($staleOperations -and $staleOperations.Count -gt 0) {
        Write-Verbose "Microsoft Learn unreachable. Using stale cache ($($staleOperations.Count) operations)."
        return [PSCustomObject]@{ Operations = $staleOperations; Source = 'Cache' }
    }

    # 4. Fallback - bundled seed.
    $seedOperations = & $loadFromJsonFile $seedPath
    if ($seedOperations -and $seedOperations.Count -gt 0) {
        Write-Verbose "Microsoft Learn and cache unavailable. Using bundled seed ($($seedOperations.Count) operations)."
        return [PSCustomObject]@{ Operations = $seedOperations; Source = 'Seed' }
    }

    Write-Verbose 'No operations catalog could be loaded (live, cache and seed all unavailable).'
    return [PSCustomObject]@{ Operations = ([System.Collections.Generic.List[PSCustomObject]]@()); Source = 'None' }
}
