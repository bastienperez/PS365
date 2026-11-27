function ConvertTo-SPOStorageSummaryRow {
    <#
    .SYNOPSIS
    Flattens the storage summary built by Get-SPOStorageSummary into rows for an Excel worksheet.

    .DESCRIPTION
    One table, two sections: 'Overview' (quota, volume per service, totals) then 'Reclaimable' (one row per
    reclaimable item). Shared by Get-SPOStorageReport and Get-SPOSiteReport -IncludeStorageSummary so both
    Excel exports carry the same layout.

    .PARAMETER Summary
    Object returned by Get-SPOStorageSummary.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Summary
    )

    $rows = [System.Collections.Generic.List[PSCustomObject]]::new()
    $gainColumn = "EstimatedYearlyGain$($Summary.Currency)"

    $addRow = {
        param($Section, $Item, $Sites, $VolumeGB, $Treatment, $Gain, $Note)
        $row = [PSCustomObject][ordered]@{
            Section     = $Section
            Item        = $Item
            Sites       = $Sites
            VolumeGB    = $VolumeGB
            Treatment   = $Treatment
            $gainColumn = $Gain
            Note        = $Note
        }
        $rows.Add($row)
    }

    $quotaNote = if ($null -ne $Summary.QuotaUsedPercent) { "$($Summary.QuotaUsedPercent) % used, $($Summary.QuotaFreeGB) GB available" } else { 'Tenant quota could not be read' }

    & $addRow 'Overview' 'Tenant SharePoint storage quota' $null $Summary.TenantQuotaGB $null $null $quotaNote
    & $addRow 'Overview' 'SharePoint sites volume' @($Summary.Sites | Where-Object { $_.Kind -eq 'SharePoint' }).Count $Summary.SharePointUsedGB $null $null 'Counts against the tenant quota.'
    & $addRow 'Overview' 'OneDrive volume' @($Summary.Sites | Where-Object { $_.Kind -eq 'OneDrive' }).Count $Summary.OneDriveUsedGB $null $null $(if ($null -eq $Summary.OneDriveUsedGB) { 'Not collected.' } else { 'Per-user storage: outside the tenant quota.' })

    foreach ($item in $Summary.Items) {
        & $addRow 'Reclaimable' $item.Item $item.Sites $item.VolumeGB $item.Treatment $item.EstimatedYearlyGain $item.Note
    }

    return $rows
}
