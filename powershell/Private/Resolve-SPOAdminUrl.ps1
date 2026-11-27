function Resolve-SPOAdminUrl {
    <#
    .SYNOPSIS
    Returns the SharePoint admin center URL to connect to, from an explicit AdminUrl or from the tenant name.

    .DESCRIPTION
    The tenant-level cmdlets (Get-PnPTenant, Get-PnPTenantSite) only work against the admin center
    (https://<tenant>-admin.sharepoint.com). Against any other SharePoint URL they fail with 'Current site is
    not a tenant administration site', which stops the whole report. This helper therefore normalizes an
    AdminUrl given as the root site (https://contoso.sharepoint.com), as the OneDrive host
    (https://contoso-my.sharepoint.com) or with a path (https://contoso.sharepoint.com/sites/x) into the
    admin center URL, and says so with a warning. The sovereign clouds keep their own domain
    (sharepoint.us, sharepoint.de, sharepoint.cn...).

    Without AdminUrl, the admin URL is derived from the first label of Tenant, which is right for an
    onmicrosoft.com domain. A vanity domain (contoso.com) cannot be mapped to its SharePoint name: a warning
    asks for AdminUrl.

    .PARAMETER AdminUrl
    AdminUrl as passed by the caller, possibly empty.

    .PARAMETER Tenant
    Tenant as passed by the caller, for example contoso.onmicrosoft.com.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$AdminUrl,

        [Parameter(Mandatory = $true)]
        [string]$Tenant
    )

    if (-not $AdminUrl) {
        if ($Tenant -notmatch '\.onmicrosoft\.(com|us|de)$' -and $Tenant -match '\.') {
            Write-Warning "Tenant '$Tenant' is not an onmicrosoft.com domain: the admin center URL is guessed as https://$($Tenant.Split('.')[0])-admin.sharepoint.com. If the connection or the tenant read fails, pass -AdminUrl 'https://<tenant>-admin.sharepoint.com'."
        }
        return "https://$($Tenant.Split('.')[0])-admin.sharepoint.com"
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($AdminUrl.Trim(), [System.UriKind]::Absolute, [ref]$uri)) {
        return $AdminUrl
    }

    $hostMatch = [regex]::Match($uri.Host, '^(?<name>[^.]+?)(?<suffix>-admin|-my)?\.(?<domain>sharepoint\.[a-z.]+)$', 'IgnoreCase')
    if (-not $hostMatch.Success) {
        return $AdminUrl
    }

    $normalizedUrl = "https://$($hostMatch.Groups['name'].Value)-admin.$($hostMatch.Groups['domain'].Value)"
    if ($hostMatch.Groups['suffix'].Value -ne '-admin' -or $uri.AbsolutePath -ne '/') {
        Write-Warning "AdminUrl '$AdminUrl' is not the SharePoint admin center: using $normalizedUrl instead (the tenant-level data is only readable there)."
    }

    return $normalizedUrl
}
