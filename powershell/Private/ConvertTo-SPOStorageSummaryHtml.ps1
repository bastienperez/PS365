function ConvertTo-SPOStorageSummaryHtml {
    <#
    .SYNOPSIS
    Renders the storage summary built by Get-SPOStorageSummary as a self-contained HTML section.

    .DESCRIPTION
    Returns a style block plus a storage section, so the same markup can be embedded in the
    Get-SPOSiteReport HTML export or wrapped into its own page by Get-SPOStorageReport. The styles reuse the
    CSS variables of Get-PS365HtmlBranding (--surface, --border, --brand...) and fall back to fixed values
    otherwise.

    .PARAMETER Summary
    Object returned by Get-SPOStorageSummary.

    .PARAMETER TopSites
    Number of sites listed in the 'largest reclaimable volumes' table. Default is 20.

    .PARAMETER CopyButtonHtml
    (Optional) Copy button markup from Get-PS365HtmlBranding, shown next to the section title. The page must
    also carry the branding script (ScriptHtml) for the button to work.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Summary,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 500)]
        [int]$TopSites = 20,

        [Parameter(Mandatory = $false)]
        [string]$CopyButtonHtml
    )

    $encode = { param($Text) if ([string]::IsNullOrEmpty("$Text")) { '' } else { [System.Net.WebUtility]::HtmlEncode("$Text") } }

    # MB below 1 GB, GB below 1 TB, TB above (a small tenant would otherwise read '0.0 GB' everywhere),
    # invariant culture so the page reads the same whatever the machine locale
    $formatSize = {
        param($SizeGB)
        if ($null -eq $SizeGB) { return 'n/a' }
        $culture = [System.Globalization.CultureInfo]::InvariantCulture
        $size = [double]$SizeGB
        if ([math]::Abs($size) -ge 1024) { return ($size / 1024).ToString('N1', $culture) + ' TB' }
        if ([math]::Abs($size) -ge 1 -or $size -eq 0) { return $size.ToString('N1', $culture) + ' GB' }
        return ($size * 1024).ToString('N0', $culture) + ' MB'
    }

    $formatMoney = {
        param($Amount)
        if ($null -eq $Amount) { return $null }
        return ([double]$Amount).ToString('N0', [System.Globalization.CultureInfo]::InvariantCulture) + " $($Summary.Currency)"
    }

    $sb = [System.Text.StringBuilder]::new()

    $null = $sb.Append(@'
<style>
  /* 365Ops look, on the tokens of Get-PS365HtmlBranding (fallback values when embedded elsewhere) */
  section.storage { margin-bottom: 24px; }
  section.storage .section-head { margin-bottom: 4px; }
  section.storage .storage-sub { color: var(--muted, #4B5160); margin: 0 0 14px; }
  .storage-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; margin-bottom: 12px; }
  .storage-tile { background: var(--surface, #fff); border: 1px solid var(--border, #E3E5EA); border-radius: 12px; padding: 18px 20px; }
  .storage-tile .lbl { font-size: 13px; color: var(--muted, #4B5160); font-weight: 600; }
  .storage-tile .val { font-size: 28px; font-weight: 700; letter-spacing: -0.02em; margin: 6px 0 4px; color: var(--text, #16181D); }
  .storage-tile .val.warn { color: #B45309; }
  .storage-tile .val.crit { color: #C2410C; }
  .storage-tile .val.good { color: #1A7F4B; }
  .storage-tile .hint { font-size: 12.5px; color: var(--sub, #5B6270); }
  .storage-bars { background: var(--surface, #fff); border: 1px solid var(--border, #E3E5EA); border-radius: 12px; padding: 18px 20px; margin-bottom: 12px; }
  .storage-bars .bar-row { display: grid; grid-template-columns: 110px 1fr 100px; align-items: center; gap: 12px; padding: 5px 0; }
  .storage-bars .bar-track { background: var(--divider, #ECEEF1); border-radius: 5px; height: 10px; overflow: hidden; }
  .storage-bars .bar-fill { height: 100%; border-radius: 5px; background: var(--brand, #de761d); }
  .storage-bars .bar-fill.outside { background: var(--skip, #C9CCD3); }
  .storage-bars .bar-val { text-align: right; font-weight: 600; }
  .storage-bars .bar-note { font-size: 12.5px; color: var(--sub, #5B6270); margin-top: 8px; }
  .storage-card { background: var(--surface, #fff); border: 1px solid var(--border, #E3E5EA); border-radius: 12px; overflow: hidden; margin-bottom: 12px; }
  table.storage-table { width: 100%; border-collapse: collapse; }
  .storage-table th { text-align: left; font-size: 12px; font-weight: 600; color: var(--muted, #4B5160); padding: 10px 14px; border-bottom: 1px solid var(--border, #E3E5EA); background: var(--surface2, #FAFAFB); }
  .storage-table td { padding: 12px 14px; border-bottom: 1px solid var(--divider, #ECEEF1); vertical-align: top; }
  .storage-table tr:last-child td { border-bottom: none; }
  .storage-table td.num { white-space: nowrap; font-weight: 600; }
  .storage-table td.note { color: var(--sub, #5B6270); font-size: 12.5px; }
  .storage-table tr.total td { background: var(--surface2, #FAFAFB); font-weight: 700; }
  .storage-table .url { color: var(--sub, #5B6270); font-size: 12px; font-family: var(--mono, Consolas, monospace); }
  details.storage-top > summary { cursor: pointer; color: var(--accent, #9A4A0A); font-weight: 600; margin-bottom: 10px; }
</style>
'@)

    $quotaClass = if ($null -eq $Summary.QuotaUsedPercent) { '' } elseif ($Summary.QuotaUsedPercent -ge 90) { 'crit' } elseif ($Summary.QuotaUsedPercent -ge 75) { 'warn' } else { 'good' }
    $quotaValue = if ($null -ne $Summary.QuotaUsedPercent) { "$($Summary.QuotaUsedPercent) %" } else { 'n/a' }
    $quotaHint = if ($null -ne $Summary.TenantQuotaGB) { "$(& $formatSize $Summary.QuotaFreeGB) still available" } else { 'Tenant quota could not be read' }
    $gainValue = & $formatMoney $Summary.EstimatedYearlyGain
    $gainTile = if ($gainValue) {
        "<div class='storage-tile'><div class='lbl'>Potential gain</div><div class='val'>&#8776; $(& $encode $gainValue)</div><div class='hint'>per year, at $($Summary.PricePerGBMonth) $(& $encode $Summary.Currency)/GB/month - estimate to confirm with your Microsoft agreement</div></div>"
    }
    else {
        "<div class='storage-tile'><div class='lbl'>Potential gain</div><div class='val'>-</div><div class='hint'>Pass -PricePerGBMonth (price of 1 GB of extra SharePoint storage) to estimate it</div></div>"
    }

    $null = $sb.Append("<section class='storage' data-copy-root><div class='section-head'><h2>Storage summary</h2>$CopyButtonHtml</div>")
    $null = $sb.Append("<div class='storage-sub'>Volumes at collection time ($($Summary.GeneratedOn.ToString('yyyy-MM-dd'))). Only SharePoint sites count against the tenant storage quota.</div>")
    $null = $sb.Append("<div class='storage-grid'>")
    $null = $sb.Append("<div class='storage-tile'><div class='lbl'>SharePoint volume</div><div class='val'>$(& $formatSize $Summary.SharePointUsedGB)</div><div class='hint'>of $(& $formatSize $Summary.TenantQuotaGB) tenant storage quota</div></div>")
    $null = $sb.Append("<div class='storage-tile'><div class='lbl'>Quota used</div><div class='val $quotaClass'>$quotaValue</div><div class='hint'>$(& $encode $quotaHint)</div></div>")
    $reclaimHint = if ($null -ne $Summary.ReclaimablePercent) { "$($Summary.ReclaimablePercent) % of the SharePoint volume" } else { '' }
    $null = $sb.Append("<div class='storage-tile'><div class='lbl'>Reclaimable</div><div class='val good'>$(& $formatSize $Summary.ReclaimableGB)</div><div class='hint'>$reclaimHint</div></div>")
    $null = $sb.Append($gainTile)
    $null = $sb.Append('</div>')

    # Volume per service: the bar widths share one scale, so the services compare visually
    $services = [System.Collections.Generic.List[PSCustomObject]]::new()
    $services.Add([PSCustomObject]@{ Name = 'SharePoint'; SizeGB = $Summary.SharePointUsedGB; InQuota = $true })
    $services.Add([PSCustomObject]@{ Name = 'OneDrive'; SizeGB = $Summary.OneDriveUsedGB; InQuota = $false })
    $maxSize = [double](($services | Where-Object { $null -ne $_.SizeGB } | Measure-Object -Property SizeGB -Maximum).Maximum)

    $null = $sb.Append("<div class='storage-bars'>")
    foreach ($service in $services) {
        if ($null -eq $service.SizeGB) {
            $null = $sb.Append("<div class='bar-row'><span>$($service.Name)</span><span class='bar-track'></span><span class='bar-val'>not collected</span></div>")
            continue
        }
        $widthPercent = if ($maxSize -gt 0) { [math]::Round([double]$service.SizeGB / $maxSize * 100, 1) } else { 0 }
        $fillClass = if ($service.InQuota) { 'bar-fill' } else { 'bar-fill outside' }
        $widthText = $widthPercent.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        $null = $sb.Append("<div class='bar-row'><span>$($service.Name)</span><span class='bar-track'><span class='$fillClass' style='display:block;width:$widthText%'></span></span><span class='bar-val'>$(& $formatSize $service.SizeGB)</span></div>")
    }
    $null = $sb.Append("<div class='bar-note'>Grey bar: OneDrive (per-user storage) does not consume the tenant SharePoint quota.</div></div>")

    $null = $sb.Append("<div class='storage-card'><table class='storage-table'><thead><tr><th>Item</th><th>Sites</th><th>Volume</th><th>Treatment</th><th>Estimated gain / year</th><th>Note</th></tr></thead><tbody>")
    foreach ($item in $Summary.Items) {
        $rowClass = if ($item.Item -like 'Total*') { " class='total'" } else { '' }
        $itemGain = & $formatMoney $item.EstimatedYearlyGain
        if (-not $itemGain) {
            $itemGain = if ($item.Item -like '*without owner*') { 'included above' } else { '-' }
        }
        $null = $sb.Append("<tr$rowClass><td>$(& $encode $item.Item)</td><td class='num'>$($item.Sites)</td><td class='num'>$(& $formatSize $item.VolumeGB)</td><td>$(& $encode $item.Treatment)</td><td class='num'>$(& $encode $itemGain)</td><td class='note'>$(& $encode $item.Note)</td></tr>")
    }
    $null = $sb.Append('</tbody></table></div>')

    $topRows = @($Summary.Sites | Where-Object { $_.ReclaimableGB -gt 0 } | Sort-Object -Property ReclaimableGB -Descending | Select-Object -First $TopSites)
    if ($topRows.Count -gt 0) {
        $null = $sb.Append("<details class='storage-top'><summary>Largest reclaimable volumes (top $($topRows.Count) sites)</summary>")
        $null = $sb.Append("<div class='storage-card'><table class='storage-table'><thead><tr><th>Site</th><th>Used</th><th>Versions</th><th>Last content change</th><th>Reclaimable</th></tr></thead><tbody>")
        foreach ($topRow in $topRows) {
            $flags = @()
            if ($topRow.IsDormant) { $flags += 'dormant' }
            if ($topRow.IsOrphan -eq $true) { $flags += 'no owner or active admin' }
            $flagText = if ($flags.Count -gt 0) { " <span class='url'>($($flags -join ', '))</span>" } else { '' }
            $lastChange = if ($topRow.LastContentModified) { $topRow.LastContentModified.ToString('yyyy-MM-dd') } else { 'unknown' }
            $null = $sb.Append("<tr><td>$(& $encode $topRow.Title)$flagText<div class='url'>$(& $encode $topRow.Url)</div></td><td class='num'>$(& $formatSize $topRow.UsedGB)</td><td class='num'>$(& $formatSize $topRow.VersionGB)</td><td>$lastChange</td><td class='num'>$(& $formatSize $topRow.ReclaimableGB)</td></tr>")
        }
        $null = $sb.Append('</tbody></table></div></details>')
    }

    $null = $sb.Append('</section>')
    return $sb.ToString()
}
