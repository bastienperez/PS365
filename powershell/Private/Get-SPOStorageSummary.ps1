function Get-SPOStorageSummary {
    <#
    .SYNOPSIS
    Computes the storage summary (quota, volume per service, reclaimable volume, estimated gain) shared by
    Get-SPOStorageReport and Get-SPOSiteReport -IncludeStorageSummary.

    .DESCRIPTION
    Pure computation, no call to SharePoint or Graph: the callers collect the data and pass it in a
    normalized shape, so both functions produce exactly the same figures.

    Only the SharePoint sites count against the tenant storage quota. OneDrive (per-user storage) is
    reported for the volume picture but is never part of the quota percentage nor of the reclaimable volume.

    Reclaimable volume, SharePoint sites only, without double counting:
    - Version history: the VersionSize of every site. An upper bound - the part actually freed depends on
      the version limit applied.
    - Dormant sites: sites without any content change for DormantDays, current content only (their versions
      are already counted in the line above).
    - Dormant sites without owner or active admin: subset of the dormant sites, shown for prioritization
      only (already included above), when the caller knows which sites are orphaned.

    .PARAMETER Sites
    Normalized site objects: Title, Url, Template, UsedGB, VersionGB, LastContentModified and IsOrphan
    ($true, $false or $null when unknown).

    .PARAMETER TenantQuotaGB
    Tenant SharePoint storage quota in GB, $null when it could not be read.

    .PARAMETER OneDriveCollected
    $false when the OneDrive sites were not enumerated (ExcludeOneDrive): the OneDrive volume is then
    reported as not collected instead of 0.

    .PARAMETER DormantDays
    Number of days without content change after which a site is dormant.

    .PARAMETER PricePerGBMonth
    Price of one GB of extra SharePoint storage per month. When omitted, no gain is computed.

    .PARAMETER Currency
    Currency label of PricePerGBMonth.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Sites,

        [Parameter(Mandatory = $false)]
        [Nullable[double]]$TenantQuotaGB,

        [Parameter(Mandatory = $false)]
        [bool]$OneDriveCollected = $true,

        [Parameter(Mandatory = $true)]
        [int]$DormantDays,

        [Parameter(Mandatory = $false)]
        [Nullable[double]]$PricePerGBMonth,

        [Parameter(Mandatory = $false)]
        [string]$Currency = 'EUR'
    )

    $dormantSince = (Get-Date).ToUniversalTime().AddDays(-$DormantDays)
    $siteRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $versionSizeReported = $false
    $orphanInfoAvailable = $false

    foreach ($site in $Sites) {
        $isOneDrive = "$($site.Template)" -like 'SPSPERS*' -or "$($site.Url)" -like '*-my.sharepoint.com/personal/*'
        $usedGB = if ($null -ne $site.UsedGB) { [double]$site.UsedGB } else { 0 }
        $versionGB = if ($null -ne $site.VersionGB) { [double]$site.VersionGB } else { 0 }

        if ($null -ne $site.VersionGB) {
            $versionSizeReported = $true
        }
        if ($null -ne $site.IsOrphan) {
            $orphanInfoAvailable = $true
        }

        $lastModified = $null
        if ($site.LastContentModified) {
            $lastModified = ([datetime]$site.LastContentModified).ToUniversalTime()
        }

        # A site whose last change date is unknown is never called dormant
        $isDormant = -not $isOneDrive -and $null -ne $lastModified -and $lastModified -lt $dormantSince
        # Versions are reclaimable on every SharePoint site; on a dormant site the current content is too
        $reclaimableGB = if ($isOneDrive) { 0 } elseif ($isDormant) { $usedGB } else { [math]::Min($versionGB, $usedGB) }

        $siteRow = [PSCustomObject][ordered]@{
            Title               = $site.Title
            Url                 = $site.Url
            Kind                = if ($isOneDrive) { 'OneDrive' } else { 'SharePoint' }
            UsedGB              = [math]::Round($usedGB, 2)
            VersionGB           = if ($null -ne $site.VersionGB) { [math]::Round($versionGB, 2) } else { $null }
            LastContentModified = $lastModified
            IsDormant           = $isDormant
            IsOrphan            = $site.IsOrphan
            ReclaimableGB       = [math]::Round($reclaimableGB, 2)
        }
        $siteRows.Add($siteRow)
    }

    $sharePointRows = @($siteRows | Where-Object { $_.Kind -eq 'SharePoint' })
    $oneDriveRows = @($siteRows | Where-Object { $_.Kind -eq 'OneDrive' })
    $dormantRows = @($sharePointRows | Where-Object { $_.IsDormant })

    $sharePointUsedGB = [double](($sharePointRows | Measure-Object -Property UsedGB -Sum).Sum)
    $oneDriveUsedGB = if ($OneDriveCollected) { [double](($oneDriveRows | Measure-Object -Property UsedGB -Sum).Sum) } else { $null }

    # Version history of every SharePoint site, capped at the site usage (VersionSize is part of it)
    $versionsGB = [double](($sharePointRows | ForEach-Object { [math]::Min([double]$_.VersionGB, [double]$_.UsedGB) } | Measure-Object -Sum).Sum)
    # Current content of the dormant sites: their versions are already in the line above
    # [double]0, not 0: with an int literal [math]::Max picks its integer overload and truncates 0.04 to 0
    $dormantGB = [double](($dormantRows | ForEach-Object { [math]::Max([double]0, [double]$_.UsedGB - [double]$_.VersionGB) } | Measure-Object -Sum).Sum)
    $reclaimableGB = $versionsGB + $dormantGB

    $getYearlyGain = {
        param([double]$VolumeGB)
        if ($null -eq $PricePerGBMonth) {
            return $null
        }
        return [math]::Round($VolumeGB * $PricePerGBMonth * 12, 0)
    }

    $items = [System.Collections.Generic.List[PSCustomObject]]::new()

    $versionItem = [PSCustomObject][ordered]@{
        Item                = 'Version history'
        Sites               = @($sharePointRows | Where-Object { $_.VersionGB -gt 0 }).Count
        VolumeGB            = [math]::Round($versionsGB, 2)
        Treatment           = 'Trim the versions (version count limit or automatic expiration)'
        EstimatedYearlyGain = & $getYearlyGain $versionsGB
        Note                = if ($versionSizeReported) { 'Upper bound: the part actually freed depends on the version limit applied.' } else { 'VersionSize not reported by the tenant: volume unknown.' }
    }
    $items.Add($versionItem)

    $dormantItem = [PSCustomObject][ordered]@{
        Item                = "Dormant sites (no content change for $DormantDays days)"
        Sites               = $dormantRows.Count
        VolumeGB            = [math]::Round($dormantGB, 2)
        Treatment           = 'Archive (Microsoft 365 Archive) or delete'
        EstimatedYearlyGain = & $getYearlyGain $dormantGB
        Note                = 'Current content only: the versions of these sites are already counted above.'
    }
    $items.Add($dormantItem)

    if ($orphanInfoAvailable) {
        $orphanRows = @($dormantRows | Where-Object { $_.IsOrphan -eq $true })
        $orphanItem = [PSCustomObject][ordered]@{
            Item                = 'Dormant sites without owner or active admin'
            Sites               = $orphanRows.Count
            VolumeGB            = [math]::Round([double](($orphanRows | Measure-Object -Property UsedGB -Sum).Sum), 2)
            Treatment           = 'Archive, then delete'
            EstimatedYearlyGain = $null
            Note                = 'Already included above: the first sites to process, nobody is left to object.'
        }
        $items.Add($orphanItem)
    }

    $totalItem = [PSCustomObject][ordered]@{
        Item                = 'Total reclaimable, overlap removed'
        Sites               = @($sharePointRows | Where-Object { $_.ReclaimableGB -gt 0 }).Count
        VolumeGB            = [math]::Round($reclaimableGB, 2)
        Treatment           = $null
        EstimatedYearlyGain = & $getYearlyGain $reclaimableGB
        Note                = 'SharePoint sites only: OneDrive does not count against the tenant storage quota.'
    }
    $items.Add($totalItem)

    $quotaUsedPercent = if ($TenantQuotaGB -gt 0) { [math]::Round($sharePointUsedGB / $TenantQuotaGB * 100, 1) } else { $null }

    return [PSCustomObject][ordered]@{
        GeneratedOn         = Get-Date
        TenantQuotaGB       = if ($null -ne $TenantQuotaGB) { [math]::Round($TenantQuotaGB, 2) } else { $null }
        SharePointUsedGB    = [math]::Round($sharePointUsedGB, 2)
        QuotaUsedPercent    = $quotaUsedPercent
        QuotaFreeGB         = if ($null -ne $TenantQuotaGB) { [math]::Round($TenantQuotaGB - $sharePointUsedGB, 2) } else { $null }
        OneDriveUsedGB      = if ($null -ne $oneDriveUsedGB) { [math]::Round($oneDriveUsedGB, 2) } else { $null }
        ReclaimableGB       = [math]::Round($reclaimableGB, 2)
        ReclaimablePercent  = if ($sharePointUsedGB -gt 0) { [math]::Round($reclaimableGB / $sharePointUsedGB * 100, 1) } else { $null }
        EstimatedYearlyGain = & $getYearlyGain $reclaimableGB
        PricePerGBMonth     = $PricePerGBMonth
        Currency            = $Currency
        DormantDays         = $DormantDays
        Items               = $items
        Sites               = $siteRows
    }
}
