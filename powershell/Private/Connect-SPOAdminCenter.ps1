function Connect-SPOAdminCenter {
    <#
    .SYNOPSIS
    Connects to the SharePoint admin center and reads the tenant settings, detecting a renamed SharePoint
    domain on the way.

    .DESCRIPTION
    Shared by Get-SPOSiteReport and Get-SPOStorageReport. Connects with PnP app-only to AdminUrl, then reads
    Get-PnPTenant.

    RENAMED SHAREPOINT DOMAIN
    The admin URL is derived from the onmicrosoft.com name (contoso.onmicrosoft.com gives
    contoso-admin.sharepoint.com). When the SharePoint domain of the tenant was renamed (Start-SPOTenantRename)
    or never matched that name, the derived URL is the old one: it still answers, as a redirect site, so the
    connection succeeds, but every tenant-level cmdlet then fails with 'Current site is not a tenant
    administration site'. On that exact error, the real root site is read from Microsoft Graph
    (GET /v1.0/sites/root, on the same app-only connection), the admin URL is rebuilt from its host, a
    warning gives the URL to pass as -AdminUrl next time, and the connection is redone there.

    That Graph call only happens on this failure, so a tenant whose name matches needs no Graph permission. It
    needs the Microsoft Graph application permission Sites.Read.All (or Sites.FullControl.All); without it, a
    warning explains how to find the right URL by hand.

    Returned object: Connection (PnP connection to use), AdminUrl (the URL actually used), TenantSettings
    (Get-PnPTenant result, $null when unreadable), TenantError (the error message when unreadable), Renamed
    ($true when the URL was corrected through Graph).

    .PARAMETER AdminUrl
    Admin center URL to try first.

    .PARAMETER PnpAuthParams
    Connect-PnPOnline parameters (ClientId, Tenant, certificate), without Url.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$AdminUrl,

        [Parameter(Mandatory = $true)]
        [hashtable]$PnpAuthParams
    )

    # The first PnP call of a session also loads the PnP.PowerShell module, which alone can take a while:
    # the elapsed time is printed so a slow start is not mistaken for a hang
    Write-Host -ForegroundColor Cyan "Connecting to SharePoint admin center: $AdminUrl (loads PnP.PowerShell on first use)..."
    $stepStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Connection errors are left to the caller (terminating), which reports them and stops
    $connection = Connect-PnPOnline -Url $AdminUrl @PnpAuthParams -ReturnConnection -ErrorAction Stop
    Write-Host -ForegroundColor Cyan "Connected in $([int]$stepStopwatch.Elapsed.TotalSeconds) s. Reading the tenant settings..."

    $result = [PSCustomObject]@{
        Connection     = $connection
        AdminUrl       = $AdminUrl
        TenantSettings = $null
        TenantError    = $null
        Renamed        = $false
    }

    try {
        $result.TenantSettings = Get-PnPTenant -Connection $connection -ErrorAction Stop
        return $result
    }
    catch {
        $result.TenantError = $_.Exception.Message.Trim()
    }

    if ($result.TenantError -notmatch 'not a tenant administration site') {
        return $result
    }

    # The URL answers but is not the admin center: most likely a renamed SharePoint domain
    $rootWebUrl = $null
    try {
        $rootSite = Invoke-PnPGraphMethod -Url 'v1.0/sites/root?$select=webUrl' -Connection $connection -ErrorAction Stop
        $rootWebUrl = $rootSite.webUrl
    }
    catch {
        Write-Verbose "Unable to read the root site from Microsoft Graph: $($_.Exception.Message)"
    }

    $hostMatch = $null
    if ($rootWebUrl) {
        $hostMatch = [regex]::Match(([System.Uri]$rootWebUrl).Host, '^(?<name>[^.]+)\.(?<domain>sharepoint\.[a-z.]+)$', 'IgnoreCase')
    }

    if (-not $hostMatch -or -not $hostMatch.Success) {
        Write-Warning "$AdminUrl is not the SharePoint admin center of this tenant ('$($result.TenantError)'). The SharePoint domain was probably renamed: pass -AdminUrl 'https://<sharepoint-name>-admin.sharepoint.com'. To find it: Microsoft 365 admin center > Admin centers > SharePoint (its URL), or Get-SPOTenantRenameStatus. Automatic detection needs the Microsoft Graph application permission Sites.Read.All."
        return $result
    }

    $detectedAdminUrl = "https://$($hostMatch.Groups['name'].Value)-admin.$($hostMatch.Groups['domain'].Value)"
    if ($detectedAdminUrl -eq $AdminUrl.TrimEnd('/')) {
        return $result
    }

    Write-Warning "The SharePoint domain of this tenant is not the one derived from its name (renamed SharePoint domain): the root site is $rootWebUrl. Using the admin center $detectedAdminUrl. Pass -AdminUrl '$detectedAdminUrl' next time to skip this detection."

    try {
        $connection = Connect-PnPOnline -Url $detectedAdminUrl @PnpAuthParams -ReturnConnection -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to connect to the detected admin center $($detectedAdminUrl): $($_.Exception.Message)"
        return $result
    }

    $result.Connection = $connection
    $result.AdminUrl = $detectedAdminUrl
    $result.Renamed = $true
    $result.TenantError = $null

    try {
        $result.TenantSettings = Get-PnPTenant -Connection $connection -ErrorAction Stop
    }
    catch {
        $result.TenantError = $_.Exception.Message.Trim()
    }

    return $result
}
