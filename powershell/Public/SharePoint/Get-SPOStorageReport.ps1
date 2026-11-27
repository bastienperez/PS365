function Get-SPOStorageReport {
    <#
    .SYNOPSIS
    Builds a FinOps view of the SharePoint Online storage: quota usage, volume per service, reclaimable
    volume and estimated yearly gain.

    .DESCRIPTION
    Get-SPOStorageReport answers "how full is the tenant storage quota, and how much could be freed":
    - the tenant SharePoint storage quota and the share of it used by the SharePoint sites;
    - the volume per service: SharePoint and OneDrive;
    - the reclaimable volume, without double counting: version history (upper bound) and dormant sites
      (no content change for DormantDays, current content only);
    - with PricePerGBMonth, the estimated yearly gain of freeing that volume.

    Everything comes from tenant-level data (Get-PnPTenant and Get-PnPTenantSite -Detailed): no connection
    to each site, so the report stays fast even on a large tenant.

    ONLY SHAREPOINT COUNTS AGAINST THE QUOTA
    The tenant storage quota (1 TB plus 10 GB per licensed user, plus purchased extra storage) is consumed by
    the SharePoint sites only. OneDrive has its own per-user storage: it is reported for the volume picture,
    never added to the quota percentage nor to the reclaimable volume. Freeing OneDrive space does not lower
    the SharePoint storage bill. Exchange Online mailboxes are out of scope here: see
    Get-ExMailboxStatisticsInfo and Get-ExMailboxQuotaInfo.

    ESTIMATES, NOT A QUOTE
    - Version history is an upper bound: the part actually freed depends on the version limit applied
      (count limit or automatic expiration). Microsoft's version expiration report job
      (New-SPOSiteFileVersionExpirationReportJob) simulates a given limit site by site when an exact
      figure is needed.
    - Dormant is decided per site, on LastContentModifiedDate. Old files inside an active site are not
      detected: that would mean walking every file of every library.
    - The gain is VolumeGB x PricePerGBMonth x 12. It is a real saving only while the tenant buys extra
      storage; below the quota, it is the purchase avoided.

    Get-SPOSiteReport -IncludeStorageSummary adds the same summary to its own Excel and HTML exports, and
    adds a 'without owner or active admin' line when its governance controls were collected.

    PERMISSIONS
    The same app-only connection as Get-SPOSiteReport: an Entra app registration with a certificate and the
    SharePoint Sites.FullControl.All application permission. No Microsoft Graph permission is needed.

    MULTI-GEO
    Only the geo of the admin center used is covered. Pass AdminUrl to target another geo.

    .PARAMETER ClientId
    (Mandatory) Application (client) ID of the Entra app registration used for the app-only connection.

    .PARAMETER Tenant
    (Mandatory) Tenant domain name for the PnP connection, for example contoso.onmicrosoft.com.

    .PARAMETER CertificateThumbprint
    Thumbprint of the certificate located in the current user's Windows certificate store.
    Provide exactly one of CertificateThumbprint, CertificatePath or CertificateBase64Encoded.

    .PARAMETER CertificatePath
    Path to a local .pfx certificate file. Use CertificatePassword when the file is protected.

    .PARAMETER CertificatePassword
    (Optional) SecureString password protecting the .pfx file passed to CertificatePath.

    .PARAMETER CertificateBase64Encoded
    Base64-encoded certificate (with its private key) passed directly to Connect-PnPOnline.

    .PARAMETER AdminUrl
    (Optional) URL of the SharePoint admin center. Derived from Tenant when omitted; provide it when Tenant
    is a vanity domain, or to target a specific Multi-Geo location.
    An AdminUrl given as the root site (https://contoso.sharepoint.com), the OneDrive host (-my) or with a path is
    turned into the admin center URL with a warning: the tenant-level data is only readable there.
    A renamed SharePoint domain is detected: when the admin URL answers 'not a tenant administration site', the
    real root site is read from Microsoft Graph (sites/root, application permission Sites.Read.All) and the
    report switches to its admin center, with a warning giving the -AdminUrl to pass next time.

    .PARAMETER DormantDays
    (Optional) Number of days without any content change after which a SharePoint site is considered
    dormant. Default is 730 (2 years).

    .PARAMETER PricePerGBMonth
    (Optional) Price of 1 GB of extra SharePoint storage per month, as in your Microsoft agreement: the
    'Office 365 Extra File Storage' line of the invoice, or its price in the Microsoft 365 admin center
    (Billing > Purchase services). It varies by country, agreement type (CSP, EA, MCA) and discounts, so no
    default is assumed. When omitted, volumes are reported without any gain estimate.

    .PARAMETER Currency
    (Optional) Currency label of PricePerGBMonth. Default is EUR.

    .PARAMETER ExportToExcel
    (Optional) Exports the results to an Excel file: a SharePoint-Storage worksheet (overview and
    reclaimable items) and a SharePoint-StorageSites worksheet (one row per site).
    Can be combined with ExportToHtml.

    .PARAMETER ExportPath
    (Optional) Output directory for the Excel or HTML export. Defaults to the user profile.

    .PARAMETER ExportToHtml
    (Optional) Exports the summary as a single self-contained HTML file (tiles, volume per service,
    reclaimable items, largest reclaimable sites) and opens it. Can be combined with ExportToExcel.

    .EXAMPLE
    Get-SPOStorageReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb

    Returns the storage summary object: quota, volume per service, reclaimable items (Items) and one row per
    site (Sites).

    .EXAMPLE
    Get-SPOStorageReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -PricePerGBMonth $pricePerGB -Currency EUR -ExportToHtml

    Generates the HTML summary with the yearly gain estimate, $pricePerGB holding the monthly price of 1 GB of
    extra storage taken from your Microsoft agreement.

    .EXAMPLE
    Get-SPOStorageReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -DormantDays 365 -ExportToExcel

    Exports the summary and the per-site detail to Excel, with sites unchanged for one year counted as dormant.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-SPOStorageReport
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Tenant,

        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory = $false)]
        [string]$CertificatePath,

        [Parameter(Mandatory = $false)]
        [securestring]$CertificatePassword,

        [Parameter(Mandatory = $false)]
        [string]$CertificateBase64Encoded,

        [Parameter(Mandatory = $false)]
        [string]$AdminUrl,

        [Parameter(Mandatory = $false)]
        [ValidateRange(30, 3650)]
        [int]$DormantDays = 730,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 1000)]
        [Nullable[double]]$PricePerGBMonth,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Currency = 'EUR',

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel or HTML export (defaults to the user profile).')]
        [string]$ExportPath,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToHtml
    )

    # Exactly one certificate method is required for the app-only connection
    $certMethods = @($CertificateThumbprint, $CertificatePath, $CertificateBase64Encoded | Where-Object { $_ })
    if ($certMethods.Count -ne 1) {
        Write-Warning 'Provide exactly one certificate method: CertificateThumbprint, CertificatePath or CertificateBase64Encoded.'
        return
    }

    if (-not (Get-Module PnP.PowerShell -ListAvailable)) {
        Write-Warning 'Please install the PnP.PowerShell module: Install-Module PnP.PowerShell'
        return
    }

    $pnpAuthParams = @{
        ClientId = $ClientId
        Tenant   = $Tenant
    }
    if ($CertificateThumbprint) {
        $pnpAuthParams.Add('Thumbprint', $CertificateThumbprint)
    }
    elseif ($CertificatePath) {
        $pnpAuthParams.Add('CertificatePath', $CertificatePath)
        if ($CertificatePassword) {
            $pnpAuthParams.Add('CertificatePassword', $CertificatePassword)
        }
    }
    else {
        $pnpAuthParams.Add('CertificateBase64Encoded', $CertificateBase64Encoded)
    }

    # Admin center URL: derived from the tenant name when not provided, normalized when given as the root site,
    # the OneDrive host or with a path (the tenant-level cmdlets fail against anything but the admin center)
    $AdminUrl = Resolve-SPOAdminUrl -AdminUrl $AdminUrl -Tenant $Tenant

    # Connection and tenant settings in one step: Connect-SPOAdminCenter also detects a renamed SharePoint
    # domain (the derived admin URL then answers as a redirect site) and switches to the real admin center
    try {
        $adminCenter = Connect-SPOAdminCenter -AdminUrl $AdminUrl -PnpAuthParams $pnpAuthParams
    }
    catch {
        Write-Warning "Unable to connect to the SharePoint admin center: $_"
        return
    }

    $adminConnection = $adminCenter.Connection
    $AdminUrl = $adminCenter.AdminUrl

    # StorageQuota is returned in MB. A failed read leaves the quota empty (percentage 'n/a') instead of
    # aborting: the volumes and the reclaimable items do not depend on it
    $tenantQuotaGB = $null
    if ($adminCenter.TenantSettings) {
        if ($adminCenter.TenantSettings.StorageQuota -gt 0) {
            $tenantQuotaGB = [double]$adminCenter.TenantSettings.StorageQuota / 1024
        }
    }
    else {
        Write-Warning "Unable to read the tenant storage quota: $($adminCenter.TenantError). The quota percentage will be empty."
    }

    Write-Host -ForegroundColor Cyan 'Retrieving SharePoint Online and OneDrive sites (one call, can take several minutes on a large tenant)...'

    try {
        $tenantSites = @(Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Connection $adminConnection -ErrorAction Stop)
    }
    catch {
        Write-Warning "Unable to list the sites of the tenant: $_"
        return
    }

    Write-Host -ForegroundColor Cyan "Sites retrieved: $($tenantSites.Count)"

    $normalizedSites = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($tenantSite in $tenantSites) {
        # StorageUsage is in MB, VersionSize in bytes. Get-SPOSite names the usage StorageUsageCurrent
        $storageUsedMB = if ($null -ne $tenantSite.StorageUsageCurrent) { $tenantSite.StorageUsageCurrent } else { $tenantSite.StorageUsage }

        $normalizedSite = [PSCustomObject]@{
            Title               = $tenantSite.Title
            Url                 = $tenantSite.Url
            Template            = $tenantSite.Template
            UsedGB              = if ($null -ne $storageUsedMB) { [double]$storageUsedMB / 1024 } else { $null }
            VersionGB           = if ($null -ne $tenantSite.VersionSize) { [double]$tenantSite.VersionSize / 1GB } else { $null }
            LastContentModified = $tenantSite.LastContentModifiedDate
            IsOrphan            = $null
        }
        $normalizedSites.Add($normalizedSite)
    }

    $adminConnection = $null

    $summaryParams = @{
        Sites             = @($normalizedSites)
        TenantQuotaGB     = $tenantQuotaGB
        OneDriveCollected = $true
        DormantDays       = $DormantDays
        PricePerGBMonth   = $PricePerGBMonth
        Currency          = $Currency
    }
    $summary = Get-SPOStorageSummary @summaryParams

    $quotaText = if ($null -ne $summary.QuotaUsedPercent) { "$($summary.QuotaUsedPercent) % of the quota" } else { 'quota unknown' }
    Write-Host -ForegroundColor Green "SharePoint: $($summary.SharePointUsedGB) GB ($quotaText). Reclaimable: $($summary.ReclaimableGB) GB."

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
        $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-SPOStorageReport.xlsx"
        Write-Host -ForegroundColor Cyan 'Exporting the storage report to Excel...'

        ConvertTo-SPOStorageSummaryRow -Summary $summary | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-Storage' -TableStyle Light9
        $summary.Sites | Sort-Object -Property ReclaimableGB -Descending | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-StorageSites' -TableStyle Light9

        Add-PS365ExcelAboutSheet -Path $excelFilePath
        Write-Host -ForegroundColor Green 'Excel report exported to:'
        Write-Host -ForegroundColor White "  $excelFilePath"
    }

    # Not an elseif: both exports can be produced by the same run
    if ($ExportToHtml.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
        $htmlFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-SPOStorageReport.html"

        # Single-quoted here-string: no PS variable expansion, placeholders replaced with String.Replace().
        # The CSS variables match the Get-SPOSiteReport page, which the storage section relies on.
        $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SharePoint Storage Report - PS365</title>
  PS365_BRAND_FAVICON
  <style>
PS365_BRAND_CSS
  </style>
</head>
<body>

<header class="ps365-header">
  PS365_BRAND_HEADER
  <div class="spacer"></div>
  <span class="chip">PS365_STATS_TEXT</span>
  <span class="chip muted">PS365_GENERATED_DATE</span>
</header>
<div class="ps365-page">
PS365_BRAND_PAGEHEAD
PS365_STORAGE_HTML
</div>
PS365_BRAND_FOOTER
PS365_BRAND_SCRIPT
</body>
</html>
'@

        $siteCount = $summary.Sites.Count
        $statsText = "$siteCount site$(if ($siteCount -ne 1) { 's' })"

        $generatedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $branding = Get-PS365HtmlBranding -Title 'SharePoint Storage Report' -Subtitle "Generated on $generatedDate - $statsText - tenant storage quota and reclaimable volume"
        $html = $htmlTemplate
        $html = $html.Replace('PS365_BRAND_CSS', $branding.Css)
        $html = $html.Replace('PS365_BRAND_FAVICON', $branding.HeadHtml)
        $html = $html.Replace('PS365_BRAND_HEADER', $branding.HeaderBrandHtml)
        $html = $html.Replace('PS365_BRAND_FOOTER', $branding.FooterHtml)
        $html = $html.Replace('PS365_BRAND_PAGEHEAD', $branding.PageHeadHtml)
        $html = $html.Replace('PS365_STORAGE_HTML', (ConvertTo-SPOStorageSummaryHtml -Summary $summary -CopyButtonHtml $branding.CopyButtonHtml))
        $html = $html.Replace('PS365_BRAND_SCRIPT', $branding.ScriptHtml)
        $html = $html.Replace('PS365_GENERATED_DATE', $generatedDate)
        $html = $html.Replace('PS365_STATS_TEXT', $statsText)
        $html | Out-File -FilePath $htmlFilePath -Encoding UTF8
        Write-Host -ForegroundColor Green 'HTML report exported to:'
        Write-Host -ForegroundColor White "  $htmlFilePath"
        Invoke-Item -Path $htmlFilePath
    }

    if (-not ($ExportToExcel.IsPresent -or $ExportToHtml.IsPresent)) {
        return $summary
    }
}
