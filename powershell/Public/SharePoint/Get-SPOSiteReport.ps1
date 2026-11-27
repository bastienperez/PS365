<#
    .SYNOPSIS
    Builds an inventory report of every SharePoint Online site of the tenant using PnP.PowerShell app-only.

    .DESCRIPTION
    Get-SPOSiteReport collects the configuration of the SharePoint Online sites: storage quota and usage,
    site collection administrators, site Owners/Members/Visitors, external sharing settings, sensitivity
    label, default sharing links, conditional access policy and hub site membership.

    The function connects with PnP.PowerShell in app-only mode (Entra app registration plus a certificate).
    This is the only reliable way to read site-level data (site collection administrators, site groups,
    regional settings) across every site of the tenant: the SharePoint Administrator role alone grants
    access to the admin center, not to the content of each site, so a delegated connection returns
    Access Denied on Get-PnPSiteCollectionAdmin. An app registration holding the SharePoint
    Sites.FullControl.All application permission bypasses the site collection admin requirement.

    That permission covers the personal OneDrive sites as well: they are ordinary site collections of the
    -my.sharepoint.com host, so they are enumerated and read exactly like any other site (use ExcludeOneDrive
    or OnlyOneDrive to filter them). The one documented exception is the list of site collection
    administrators of a OneDrive site, which stays invisible even under this permission: see SECONDARY
    ADMINISTRATORS ON ONEDRIVE SITES ARE NOT CAPTURED under the IncludeSiteAdmins parameter.

    Client secrets are not supported by SharePoint for app-only: a certificate is mandatory. Provide it
    through CertificateThumbprint (Windows certificate store), CertificatePath (.pfx file) or
    CertificateBase64Encoded (base64 string, handy for Azure Automation or a pipeline variable).

    The site-level collection runs in parallel (ForEach-Object -Parallel) with one PnP connection per site.
    Tune the concurrency with ThrottleLimit. PowerShell 7 is required.

    UNIFORM MEMBERS/OWNERS MODEL
    On a site connected to a Microsoft 365 group, the SharePoint site collection administrators and the
    Owners/Members/Visitors groups do not list individual users: they list a single claim placeholder that
    represents the whole group, of the form 'federateddirectoryclaimprovider|<groupId>_o' (group Owners) or
    'federateddirectoryclaimprovider|<groupId>' (group Members). Get-SPOSiteReport resolves these claims in
    a second pass, so the *Resolved columns list actual users regardless of the site type (classic site,
    group-connected site, with or without a Microsoft Teams team). Two well-known non-user claims are also
    labelled for readability: the 'Everyone' claim and the 'Everyone except external users' claim. Any claim
    that cannot be resolved (deleted group, insufficient Graph permission) is reported as '<unresolved:guid>'
    and counted in the ClaimResolutionErrorCount summary at the end of the run.

    PROVENANCE OF EACH RESOLVED IDENTITY
    Every identity in a *Resolved column is annotated so you can tell how it got its rights:
    - No annotation: the user is directly assigned to that SharePoint group/role (individual claim).
    - "<user> (via M365 group '<name>' Owners)" or "... Members)": the user only has rights because they
      belong to that Microsoft 365 group, which itself was added to the SharePoint group/role as a whole.
    The same user can legitimately appear twice for the same site - once without annotation (direct) and
    once with it (also a member of a group that was added) - that is not a duplicate, it reflects two
    distinct sources of access. The group's display name is resolved once per distinct group for the whole
    run and falls back to its GUID if that lookup fails (deleted group, insufficient Graph permission).

    M365GROUPSDETAILS - APP-ONLY REQUIREMENT
    M365GroupsDetails is served entirely through the PnP app-only connection already opened by this
    function (Get-PnPMicrosoft365Group, Get-PnPTeamsTeam, Get-PnPTeamsChannel, Get-PnPTeamsUser,
    Get-PnPMicrosoft365GroupOwner, Get-PnPMicrosoft365GroupMember): no Exchange Online or Microsoft Teams
    session is needed. The Entra app registration must additionally hold the Microsoft Graph application
    permission Group.ReadWrite.All (or Group.Read.All for the read-only subset), on top of the SharePoint
    Sites.FullControl.All permission used for the rest of the report, both consented on the same certificate.

    The retrieval is scoped to what the run actually needs. Only a site carrying a GroupId can have a group
    behind it, so that list is computed first: when no site in scope is group-connected (a OneDrive-only run,
    or SiteUrl on a classic site) nothing is retrieved at all, and up to 25 group-connected sites are looked
    up one by one. The tenant-wide enumeration - which is expensive, since IncludeSiteUrl and IncludeOwners
    each add their own Graph work per group - only kicks in beyond that, where it becomes the cheaper option.

    MULTI-GEO TENANTS
    By default, every geo location of a Multi-Geo tenant is enumerated automatically (detected via
    Get-PnPTenantInstance): the function opens one additional PnP connection per satellite geo admin center
    purely to list its sites, merges every site into a single result set, and tags each one with a Geo
    column so its origin stays visible. The site-level detail collection and the Microsoft 365 group layer
    do not need any geo-specific handling: the former connects directly to each site's own URL (already
    geo-specific), and Microsoft 365 groups are Entra ID directory objects, not partitioned by geo. This
    only happens for the full tenant-wide report: SiteUrl (single site) and an explicitly supplied AdminUrl
    (one geo targeted on purpose) both stay single-geo, exactly as before. On a single-geo tenant this adds
    one harmless detection call and otherwise behaves exactly as before. Use ExcludeGeo to skip one or more
    geo locations entirely (for example one that is too noisy for the report at hand) - MultiGeoStats
    reflects the same exclusion automatically, since it aggregates the sites this enumeration already produced.

    ACTIVE SITES COLUMNS
    The columns of the CSV export of the Active sites page of the SharePoint admin center are in every report:
    SiteCreatedDate, CreatedBy, CreatedFrom (named with GetSPOSiteCreationSources: Microsoft Teams, SharePoint
    start page, API...), LastActivityDate, FilesCount, FilesViewedOrEdited, PageViews, PagesVisited,
    ChannelSitesCount and IBMode come from the hidden list that page reads
    (DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS, read once with the admin connection),
    HubSiteName from Get-PnPHubSite, StorageUsedPercent is StorageUsedGB against StorageLimitGB and
    CustomScripts (Allowed or Blocked) is DenyAddAndCustomizePages. The list covers the main geo only: on a
    Multi-Geo tenant those columns stay empty for the sites of a satellite geo.

    Optional switches enrich the report:
    - IncludeSiteAdmins adds the full list of site collection administrators of each site (raw and resolved).
    - IncludeSiteMembers adds the site Owners and Members, resolved through the model described above.
    - IncludeSiteVisitors adds the site Visitors on top of IncludeSiteMembers. Kept separate because the
      Visitors group is often large (for example 'Everyone except external users') and less relevant to a
      rights audit than Owners/Members.
    - IncludeSharingLinks adds a per-site count of the existing sharing links, broken down by link type.
    - IncludeSharingLinksDetails drills into those links file by file, through Get-SPOSharingLinkReport.
    - RegionalSettingsDetails adds the time zone, hour format and locale of each site.
    - M365GroupsDetails adds the Microsoft 365 group and Microsoft Teams layer of the group-connected sites.

    Each of these switches literally adds its columns to the output objects: without the switch, the columns
    are absent, not empty. This is deliberate - an empty SiteAdmins column would read as "this site has no
    administrator", where the truth is "this was not collected". A column that is present but empty therefore
    always means the data really is empty (or could not be read, see the Status counters at the end of a run).

    FINDINGS SUMMARY
    ExportToExcel and ExportToHtml add a findings summary computed from the collected columns only (no extra
    call): each control counts the sites or Microsoft 365 groups it flags and lists them with a short detail.
    - EID-GROUP-08 Public Microsoft 365 groups (M365GroupsDetails).
    - SPO-SHARE-04 / SPO-SHARE-05 Private Microsoft 365 groups whose site grants access to 'Everyone except external
      users' / 'Everyone' (M365GroupsDetails, IncludeEveryoneAccess).
    - SPO-SHARE-01 / SPO-SHARE-02 Sites granting access to 'Everyone except external users' / 'Everyone'
      (IncludeEveryoneAccess): on the site itself, directly or through any SharePoint group, and on the
      libraries/lists with unique permissions; folders and files are not inspected. Without
      IncludeEveryoneAccess, SPO-SHARE-01, 02, 04 and 05 are still evaluated from IncludeSiteMembers plus
      IncludeSiteVisitors, limited to the associated Owners/Members/Visitors groups, and flagged as partial.
    - EID-GROUP-11 Sites with guests ('#EXT#' identities in the site rights, plus the Microsoft 365 group guests
      when M365GroupsDetails is used) (IncludeSiteMembers).
    - SPO-SHARE-03 Sites holding anonymous 'Anyone with the link' links (IncludeSharingLinks). Flexible links are
      not counted: their scope is only known through Get-SPOSharingLinkReport -IncludeGraphDetails.
    - EID-GROUP-13 / EID-GROUP-14 / EID-GROUP-15 Microsoft 365 groups without owner, with external owners only, with external
      users only (M365GroupsDetails).
    - SPO-OWNER-01 Sites without an active administrator: every site collection administrator is disabled, deleted
      or has not signed in for InactiveDaysThreshold days (IncludeSiteAdminsActivity). OneDrive sites are not
      evaluated, and a site where one admin could not be looked up is not flagged.
    A control whose switches were not used is reported as NotCollected (with the switch it needs), never as
    zero affected. IncludeAllDetails turns on everything the summary needs.

    .PARAMETER ClientId
    (Mandatory) Application (client) ID of the Entra app registration used for the app-only connection.

    .PARAMETER Tenant
    (Mandatory) Tenant domain name for the PnP connection, for example contoso.onmicrosoft.com.
    Used by Connect-PnPOnline (it expects the domain, not the tenant GUID).

    .PARAMETER CertificateThumbprint
    Thumbprint of the certificate located in the current user's Windows certificate store.
    Provide exactly one of CertificateThumbprint, CertificatePath or CertificateBase64Encoded.

    .PARAMETER CertificatePath
    Path to a local .pfx certificate file. Use CertificatePassword when the file is protected.
    Provide exactly one of CertificateThumbprint, CertificatePath or CertificateBase64Encoded.

    .PARAMETER CertificatePassword
    (Optional) SecureString password protecting the .pfx file passed to CertificatePath.

    .PARAMETER CertificateBase64Encoded
    Base64-encoded certificate (with its private key) passed directly to Connect-PnPOnline.
    Provide exactly one of CertificateThumbprint, CertificatePath or CertificateBase64Encoded.

    .PARAMETER AdminUrl
    (Optional) URL of the SharePoint admin center, for example https://contoso-admin.sharepoint.com.
    When omitted, it is derived from the tenant name of the Tenant parameter. Provide it explicitly
    when Tenant is a vanity domain (for example contoso.com) rather than an onmicrosoft.com domain.
    On a Multi-Geo tenant, supplying AdminUrl explicitly restricts the whole report to that one geo: the
    automatic loop over every geo (see MULTI-GEO TENANTS above) only runs when AdminUrl is omitted and
    derived automatically from Tenant.
    An AdminUrl given as the root site (https://contoso.sharepoint.com), the OneDrive host (-my) or with a path is
    turned into the admin center URL with a warning: the tenant-level data is only readable there.
    A renamed SharePoint domain is detected: when the admin URL answers 'not a tenant administration site', the
    real root site is read from Microsoft Graph (sites/root, application permission Sites.Read.All) and the
    report switches to its admin center, with a warning giving the -AdminUrl to pass next time.

    .PARAMETER SiteUrl
    (Optional) Restricts the report to a single site collection URL. When omitted, every site of the tenant
    is processed. Always single-site/single-geo: the automatic Multi-Geo enumeration never applies when
    SiteUrl is used.

    .PARAMETER ExcludeOneDrive
    (Optional) Excludes the personal OneDrive sites from the report. Cannot be combined with OnlyOneDrive.

    .PARAMETER OnlyOneDrive
    (Optional) Restricts the report to the personal OneDrive sites. Cannot be combined with ExcludeOneDrive.

    .PARAMETER ExcludeGeo
    (Optional) On a Multi-Geo tenant, excludes one or more geo locations from the automatic enumeration (see
    MULTI-GEO TENANTS above), matched by their geo code (for example 'CHE', 'TWN' - case-insensitive). Handy
    when a specific geo has a lot of noise that is not relevant to the current report. Tab-completes against
    the official Microsoft 365 Geography (PreferredDataLocation) codes, but any value is accepted - the
    completer is a suggestion only, not a restriction, so a geo code not yet in that list still works. A geo
    name that does not match any detected location is ignored with a warning; excluding every detected geo
    returns nothing with a warning rather than silently producing an empty report. Has no effect when
    SiteUrl or an explicit AdminUrl already restrict the report to a single geo.

    .PARAMETER IncludeAllDetails
    (Optional) Turns on every collection at once: IncludeSiteAdmins, IncludeSiteMembers, IncludeSiteVisitors,
    IncludeSharingLinks, IncludeSharingLinksDetails, IncludeEveryoneAccess, IncludeSiteAdminsActivity,
    RegionalSettingsDetails, M365GroupsDetails and, with an export, IncludeStorageSummary.

    THIS IS THE SLOWEST MODE. It includes the expensive collections: the file-level drill-down of the sharing
    links (IncludeSharingLinksDetails, the slowest part), the Microsoft 365 group and Teams layer, the
    libraries with unique permissions of every site and one Microsoft Graph lookup per distinct admin. On a
    large tenant, expect a run of several hours; a notice listing the slow parts is printed at the start.
    For a faster run, pass only the individual switches you need.

    It also uses every Microsoft Graph application permission of the report (Group.Read.All, User.Read.All,
    AuditLog.Read.All or Reports.Read.All) on top of SharePoint Sites.FullControl.All. A missing permission
    does not stop the run: the part that needs it is skipped with a warning, and its findings are reported
    as not collected.

    Does NOT affect ExcludeOneDrive/OnlyOneDrive: those are mutually exclusive site filters, not detail
    enrichment, and are left untouched. Individual switches passed alongside IncludeAllDetails have no
    additional effect (everything is already on).

    .PARAMETER IncludeSiteAdmins
    (Optional) Adds the full list of site collection administrators of each site, resolved with
    Get-PnPSiteCollectionAdmin, plus their resolved identities (SiteAdminsResolved). Requires the app-only
    connection described above.

    SECONDARY ADMINISTRATORS ON ONEDRIVE SITES ARE NOT CAPTURED
    SharePoint Administrators or Global Administrators cannot view the site collection administrators of a
    OneDrive site by default: they must be explicitly added as a site collection administrator on that
    specific OneDrive site to gain visibility into its admin list. Because of this restriction,
    Get-PnPSiteCollectionAdmin - which this switch relies on - does not pick up secondary administrators on
    OneDrive sites, even under the app-only certificate used by this function (see
    https://github.com/pnp/powershell/discussions/4697). Expect SiteAdmins/SiteAdminsResolved to be
    incomplete or empty for OneDrive sites: this is not a bug in this function, it reflects what the
    connection is actually allowed to see. The 'My Site Secondary Admin' admin center page (reminded at the
    end of every run) remains the only reliable source for OneDrive secondary administrators.

    .PARAMETER IncludeSiteMembers
    (Optional) Adds the site Owners and Members groups of each site (Get-PnPGroup -AssociatedOwnerGroup /
    -AssociatedMemberGroup), raw and resolved through the claim resolution model described above. This is
    the reliable, site-type-agnostic replacement for auditing who effectively has rights on a site.

    .PARAMETER IncludeSiteVisitors
    (Optional) Adds the site Visitors group of each site (Get-PnPGroup -AssociatedVisitorGroup), raw and
    resolved. Kept separate from IncludeSiteMembers because this group is often large and less relevant
    to a rights audit.

    .PARAMETER IncludeSharingLinks
    (Optional) Adds five columns counting the sharing links that currently exist on each site:
    SharingLinksAnyoneCount (anonymous 'Anyone with the link'), SharingLinksCompanyCount (people in the organization),
    SharingLinksFlexibleCount, SharingLinksOtherCount and SharingLinksTotalCount. Meant as a triage indicator: it answers
    "which sites have an exposure worth looking at", not "which file is shared with whom".

    SharingLinksFlexibleCount deserves a word: Flexible is the modern SharePoint link type, whose scope is
    set per link. The group name says the link exists, not whether it is restricted to named people or open
    to anyone holding it, and this report will not guess - counting Flexible links as 'specific people'
    would make SharingLinksAnyoneCount read as zero on a site that may well be exposing an anonymous link.
    A high SharingLinksFlexibleCount therefore means "worth resolving", which is what
    Get-SPOSharingLinkReport -IncludeGraphDetails does.

    SharePoint stores every sharing link as a hidden site group named
    'SharingLinks.<itemGuid>.<linkType>.<linkGuid>', so this switch only costs one extra Get-PnPGroup call
    per site, and it reuses the app-only connection already opened for the site. The link type is derived
    from that group name (AnonymousEdit/AnonymousView, OrganizationEdit/OrganizationView, Flexible); since
    the naming is an internal SharePoint convention rather than a documented API, any unknown type is
    counted in SharingLinksOtherCount instead of being dropped. A non-zero SharingLinksOtherCount is worth investigating:
    it usually means a link type this function does not know about yet.

    What it deliberately does not provide: which file or folder is shared, and with whom. Use
    IncludeSharingLinksDetails, or Get-SPOSharingLinkReport directly, on the sites this report flags.

    .PARAMETER IncludeSharingLinksDetails
    (Optional) Drill-down of IncludeSharingLinks (which it turns on automatically): instead of stopping at
    the counts, each link is resolved down to the file or folder it points at and to the people it was
    shared with. The work is delegated to Get-SPOSharingLinkReport, which can also be called on its own.

    Only the sites whose counts came back non-zero are visited, so the pass is naturally limited to the
    sites that actually have something to show. It still opens one connection per site and resolves every
    shared item, which makes it the slowest part of IncludeAllDetails.

    No Microsoft Graph permission is involved: the drill-down runs in the Graph-free mode of
    Get-SPOSharingLinkReport, where the shared-with list comes from the members of the SharingLinks group
    itself. The columns Graph alone can provide - link URL, expiration date, password protection, download
    block, and the resolved scope of a Flexible link - are therefore absent here. Call
    Get-SPOSharingLinkReport -IncludeGraphDetails yourself when you need them, after reading the permission
    trade-off documented in that function.

    Shape of the output: every site object receives a SharingLinksDetails property holding the collection of
    its own links (an empty collection when it has none, never $null), so the report stays one object per
    site and the drill-down is $site.SharingLinksDetails. With ExportToExcel, that nested property is
    dropped from the main worksheet - a cell cannot hold a collection - and the links are written to their
    own SharePoint-SharingLinks worksheet, one link per row.

    .PARAMETER IncludeEveryoneAccess
    (Optional) Lists where the tenant-wide claims 'Everyone' and 'Everyone except external users' are granted
    a real permission (anything beyond 'Limited Access'): on the site itself, directly or through any
    SharePoint group holding permissions (not only the associated Owners/Members/Visitors ones), and on every
    non hidden library or list whose permission inheritance is broken. Adds two columns: EveryoneAccess (for
    example "Everyone except external users on library 'Finance' (Read, via group 'Finance readers')") and
    ListsWithUniquePermissionsCount. Costs two REST calls per site plus one per list with unique permissions,
    on the site connection already opened. Folders, files and subsites are not inspected: SharePoint cannot
    filter items on unique permissions, so it would mean walking every item. Sharing link groups are left
    out, they are counted by IncludeSharingLinks.

    .PARAMETER RegionalSettingsDetails
    (Optional) Adds the regional settings of each site (time zone, hour format, locale), read with Get-PnPWeb.
    This switch significantly increases the execution time on large tenants.

    .PARAMETER M365GroupsDetails
    (Optional) Adds the Microsoft 365 group and Microsoft Teams details of the group-connected sites (including
    M365GroupVisibility: Public, Private or HiddenMembership), served entirely through the PnP app-only connection (see the app-only requirement note above). No prior
    Exchange Online or Microsoft Teams session is needed.
    This switch significantly increases the execution time on large tenants.

    .PARAMETER IncludeSiteAdminsActivity
    (Optional) Checks whether the site collection administrators are still active, and turns on
    IncludeSiteAdmins. Every distinct admin account (group claims expanded to their users) is looked up once
    in Microsoft Graph through the PnP app-only connection, reading accountEnabled and signInActivity. An
    admin is active when the account is enabled and has signed in within InactiveDaysThreshold days; a
    disabled, deleted or never signed in account is not. Adds three columns: SiteAdminsActiveCount,
    SiteAdminsNotActive (each non active admin with the reason) and SiteAdminsActivityStatus ('OK',
    'PARTIAL: ...' when a lookup failed, 'NoAdminCollected', 'NoUserAdmin' or 'NotApplicable (OneDrive)').

    Requires the Microsoft Graph application permissions User.Read.All and AuditLog.Read.All on the app
    registration, and an Entra ID P1 license for signInActivity. Without AuditLog.Read.All or the license, the
    function warns once and falls back to the Microsoft 365 active users report (getOffice365ActiveUserDetail,
    last 180 days, one download for the whole run): the last activity date across Exchange, OneDrive,
    SharePoint, Teams... replaces the last sign-in. That fallback needs the Reports.Read.All application
    permission (no license), and the report concealment setting turned off ('Display concealed user, group,
    and site names in all reports'), otherwise user names are hashed and cannot be matched. An admin absent
    from the report (often an unlicensed account) gets no verdict. When neither source is usable, an enabled
    admin is considered active.

    .PARAMETER InactiveDaysThreshold
    (Optional) Number of days without sign-in after which an admin is considered inactive. Default is 90.
    Used by IncludeSiteAdminsActivity only.

    .PARAMETER IncludeStorageSummary
    (Optional) Adds the storage summary of Get-SPOStorageReport to the Excel export (SharePoint-Storage
    worksheet) and to the HTML export (section under the findings): tenant quota usage, volume per service,
    reclaimable volume (version history and dormant sites, without double counting) and, with
    PricePerGBMonth, the estimated yearly gain. Computed from the columns already collected, no extra call.
    When the no owner (EID-GROUP-13) or no active admin (SPO-OWNER-01) controls were evaluated, a 'dormant
    sites without owner or active admin' line is added. Requires ExportToExcel or ExportToHtml, and is
    skipped with SiteUrl or OnlyOneDrive (the quota usage needs every SharePoint site).

    .PARAMETER DormantDays
    (Optional) Number of days without any content change after which a SharePoint site is considered
    dormant by IncludeStorageSummary. Default is 730 (2 years).

    .PARAMETER PricePerGBMonth
    (Optional) Price of 1 GB of extra SharePoint storage per month, used by IncludeStorageSummary to estimate
    the yearly gain. When omitted, volumes are reported without any gain.

    .PARAMETER Currency
    (Optional) Currency label of PricePerGBMonth. Default is EUR.

    .PARAMETER ThrottleLimit
    (Optional) Maximum number of sites processed concurrently by the parallel site-level collection. Default is 8.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.
    Besides the SharePoint-SiteReport worksheet, the file holds a SharePoint-Findings worksheet (one row per
    control, see FINDINGS SUMMARY) and a SharePoint-FindingsSites worksheet (one row per flagged site).
    Can be combined with ExportToHtml: both files are then produced by the same run.

    .PARAMETER ExcelTemplatePath
    (Optional) Path to an existing .xlsx file used as a template for the Excel export. The template is
    copied to the destination path first, then the report data is written into a 'SharePoint-SiteReport'
    worksheet in that copy: any other worksheet, formatting or logo already present in the template is left
    untouched. Ignored when ExportToExcel is not specified.

    .PARAMETER ExportPath
    (Optional) Output directory for the Excel or HTML export. Defaults to the user profile.

    .PARAMETER ExportToHtml
    (Optional) If specified, exports the results as a single self-contained HTML file (no external
    dependency, works offline) instead of returning the objects. Every site is rendered as a collapsible
    tree: Site Collection Admins / Owners / Members / Visitors, each showing the individual users
    resolved through the claim model above, with identities inherited from a nested Microsoft 365 group
    shown under their own collapsible sub-node (name and role of that group). A search box filters sites
    by title or URL, useful when the tenant has many sites. A findings summary table (see FINDINGS SUMMARY)
    sits above the sites, each flagged site carries the codes of its findings as badges, and the search box
    also matches those codes (type SPO-SHARE-01 to list the sites it flags). The file opens automatically once written.
    Can be combined with ExportToExcel: both files are then produced by the same run.

    .PARAMETER MultiGeoStats
    (Optional) Skips every detailed collection (site collection admins, Owners/Members/Visitors, regional
    settings, Microsoft 365 group and Teams layer) and returns only a per-geo site count: GeoName,
    IsDefaultGeo, AdminUrl, TotalSites, SharePointSites and OneDriveSites. Since the full report already
    enumerates every geo of a Multi-Geo tenant by default (see MULTI-GEO TENANTS in the description), this
    switch performs no connection or enumeration of its own: it simply groups the sites already merged
    above by their Geo tag and computes the same aggregates. It remains the fast, zero-per-site-connection
    mode - it still skips the detailed per-site collection entirely, it just no longer duplicates the site
    enumeration work. Combine with SiteUrl or an explicit AdminUrl to restrict the aggregation to that one
    site/geo, exactly like the full report does. ExportToExcel is honored (worksheet
    'SharePoint-MultiGeoStats', ExcelTemplatePath applies too); ExportToHtml is ignored, since that export
    renders the detailed rights tree, which does not apply to a stats-only run.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -MultiGeoStats

    Returns, per Multi-Geo location (or a single row on a single-geo tenant), the total number of sites,
    the number of SharePoint sites and the number of OneDrive sites - without collecting any other detail.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -MultiGeoStats -ExportToExcel

    Exports the per-geo site counts to an Excel file (worksheet 'SharePoint-MultiGeoStats') instead of
    returning the objects.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb

    Returns every SharePoint Online site of the tenant, including the personal OneDrive sites, using an
    app-only connection based on a certificate from the Windows certificate store. On a Multi-Geo tenant,
    this automatically covers every geo location and tags each site with its Geo column; on a single-geo
    tenant it behaves exactly like a normal full report.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -AdminUrl 'https://contosodeu-admin.sharepoint.com' -CertificateThumbprint $thumb

    Restricts the report to a single Multi-Geo location by supplying its admin center URL explicitly -
    the automatic loop over every geo does not run in this case.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -ExcludeGeo 'CHE', 'TWN'

    Enumerates every Multi-Geo location except CHE and TWN, still looping automatically over the rest.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateBase64Encoded $certB64 -ExcludeOneDrive -IncludeSiteAdmins -IncludeSiteMembers

    Returns the SharePoint sites without the OneDrive ones and adds the site collection administrators plus
    the Owners and Members of each site, with individual users resolved even on group-connected sites.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificatePath 'C:\Certs\app.pfx' -CertificatePassword $pwd -SiteUrl 'https://contoso.sharepoint.com/sites/marketing' -RegionalSettingsDetails

    Returns a single site with its regional settings, using a .pfx certificate file.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -M365GroupsDetails -ExportToExcel

    Adds the Microsoft 365 group and Microsoft Teams layer (Owners, Members, Guests, Teams metadata) using
    only the PnP app-only connection, and exports the results to an Excel file in the user profile directory.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -ExportToExcel -ExcelTemplatePath 'C:\Templates\SPOSiteReport-Template.xlsx'

    Exports the results into a copy of the given template, keeping any other worksheet, formatting or
    logo already present in the template file.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -IncludeSharingLinks | Where-Object { $_.SharingLinksAnyoneCount -gt 0 }

    Lists the sites holding at least one anonymous 'Anyone with the link' sharing link. Pipe the result to
    Get-SPOSharingLinkReport to get the file-level detail of those links on the sites that came out.

    .EXAMPLE
    $report = Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -IncludeSharingLinksDetails
    $report[0].SharingLinksDetails | Format-Table ItemName, LinkType, SharedWith, ExpirationDateTime

    Same triage, in one command: the counts decide which sites deserve the file-level pass, and each site
    object carries its own links in SharingLinksDetails. Add -ExportToExcel to get them as a second
    worksheet instead.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -ExcludeOneDrive -IncludeAllDetails -ExportToExcel

    Turns on every detail switch (site admins, Owners/Members/Visitors, regional settings, Microsoft 365
    group and Teams layer) in one go, without having to remember each switch name individually.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -ExcludeOneDrive -IncludeAllDetails -ExportToHtml

    Generates a single self-contained HTML file with a collapsible rights tree per site (Admins, Owners,
    Members, Visitors, with identities inherited from a nested Microsoft 365 group shown separately), and
    opens it automatically.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -ExcludeOneDrive -IncludeAllDetails -InactiveDaysThreshold 180 -PricePerGBMonth $pricePerGB -ExportToExcel

    Slowest and most complete run: every collection, every control of the findings summary (an admin without
    sign-in for 180 days counts as inactive) and the storage summary, with the yearly gain computed from
    $pricePerGB, the monthly price of 1 GB of extra storage taken from your Microsoft agreement.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -ExcludeOneDrive -IncludeSiteMembers -IncludeSiteVisitors -M365GroupsDetails -ExportToHtml

    Faster run limited to the site rights and the Microsoft 365 groups: the findings that need other
    collections are reported as not collected.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-SPOSiteReport
#>

function Get-SPOSiteReport {
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
        [ValidateNotNullOrEmpty()]
        [string]$SiteUrl,

        [Parameter(Mandatory = $false)]
        [switch]$ExcludeOneDrive,

        [Parameter(Mandatory = $false)]
        [switch]$OnlyOneDrive,

        [Parameter(Mandatory = $false)]
        [ArgumentCompleter({
                param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)
                # The 5-parameter positional signature above is mandated by ArgumentCompleter itself -
                # only $WordToComplete is actually needed here, the rest must still be declared in order.
                $null = $CommandName, $ParameterName, $CommandAst, $FakeBoundParameters

                # Microsoft 365 Geography codes (PreferredDataLocation/PDL). Suggestions only - never
                # enforced, so a future/unlisted PDL code still works (the runtime already warns on values
                # that do not match any geo actually detected on the tenant).
                @(
                    'APC', 'AUS', 'AUT', 'BRA', 'CAN', 'CHL', 'DNK', 'EUR', 'FRA', 'DEU',
                    'IND', 'IDN', 'ISR', 'ITA', 'JPN', 'KOR', 'MYS', 'MEX', 'NZL', 'NOR',
                    'POL', 'QAT', 'ZAF', 'ESP', 'SWE', 'CHE', 'TWN', 'ARE', 'GBR', 'NAM'
                ) | Where-Object { $_ -like "$WordToComplete*" }
            })]
        [string[]]$ExcludeGeo,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeAllDetails,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSiteAdmins,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSiteMembers,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSiteVisitors,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSharingLinks,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeSharingLinksDetails,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeEveryoneAccess,

        [Parameter(Mandatory = $false)]
        [switch]$RegionalSettingsDetails,

        [Parameter(Mandatory = $false)]
        [switch]$M365GroupsDetails,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeStorageSummary,

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
        [switch]$IncludeSiteAdminsActivity,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 3650)]
        [int]$InactiveDaysThreshold = 90,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 32)]
        [int]$ThrottleLimit = 8,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional path to an existing .xlsx file used as a template for the Excel export.')]
        [string]$ExcelTemplatePath,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel or HTML export (defaults to the user profile).')]
        [string]$ExportPath,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToHtml,

        [Parameter(Mandatory = $false)]
        [switch]$MultiGeoStats
    )

    # https://diecknet.de/en/2021/07/09/Sharepoint-Online-Timezones-by-PowerShell/
    function Convert-SPOTimezoneToString {
        <#
        .SYNOPSIS
        Convert a SharePoint Online time zone ID to a human readable string.

        .NOTES
        By Andreas Dieckmann - https://diecknet.de
        Timezone IDs according to https://docs.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.spregionalsettings.timezones?view=sharepoint-server#Microsoft_SharePoint_SPRegionalSettings_TimeZones

        Licensed under MIT License
        Copyright 2021 Andreas Dieckmann

        Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

        The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

        THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

        .EXAMPLE
        Convert-SPOTimezoneToString -Id 14

        Returns '(UTC-09:00) Alaska'.

        .LINK
        https://diecknet.de/en/2021/07/09/Sharepoint-Online-Timezones-by-PowerShell/
        #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [int]$Id
        )

        $timezoneIDs = @{
            39  = '(UTC-12:00) International Date Line West'
            95  = '(UTC-11:00) Coordinated Universal Time-11'
            15  = '(UTC-10:00) Hawaii'
            14  = '(UTC-09:00) Alaska'
            78  = '(UTC-08:00) Baja California'
            13  = '(UTC-08:00) Pacific Time (US and Canada)'
            38  = '(UTC-07:00) Arizona'
            77  = '(UTC-07:00) Chihuahua, La Paz, Mazatlan'
            12  = '(UTC-07:00) Mountain Time (US and Canada)'
            55  = '(UTC-06:00) Central America'
            11  = '(UTC-06:00) Central Time (US and Canada)'
            37  = '(UTC-06:00) Guadalajara, Mexico City, Monterrey'
            36  = '(UTC-06:00) Saskatchewan'
            35  = '(UTC-05:00) Bogota, Lima, Quito'
            10  = '(UTC-05:00) Eastern Time (US and Canada)'
            34  = '(UTC-05:00) Indiana (East)'
            88  = '(UTC-04:30) Caracas'
            91  = '(UTC-04:00) Asuncion'
            9   = '(UTC-04:00) Atlantic Time (Canada)'
            81  = '(UTC-04:00) Cuiaba'
            33  = '(UTC-04:00) Georgetown, La Paz, Manaus, San Juan'
            28  = '(UTC-03:30) Newfoundland'
            8   = '(UTC-03:00) Brasilia'
            85  = '(UTC-03:00) Buenos Aires'
            32  = '(UTC-03:00) Cayenne, Fortaleza'
            60  = '(UTC-03:00) Greenland'
            90  = '(UTC-03:00) Montevideo'
            103 = '(UTC-03:00) Salvador'
            65  = '(UTC-03:00) Santiago'
            96  = '(UTC-02:00) Coordinated Universal Time-02'
            30  = '(UTC-02:00) Mid-Atlantic'
            29  = '(UTC-01:00) Azores'
            53  = '(UTC-01:00) Cabo Verde'
            86  = '(UTC) Casablanca'
            93  = '(UTC) Coordinated Universal Time'
            2   = '(UTC) Dublin, Edinburgh, Lisbon, London'
            31  = '(UTC) Monrovia, Reykjavik'
            4   = '(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna'
            6   = '(UTC+01:00) Belgrade, Bratislava, Budapest, Ljubljana, Prague'
            3   = '(UTC+01:00) Brussels, Copenhagen, Madrid, Paris'
            57  = '(UTC+01:00) Sarajevo, Skopje, Warsaw, Zagreb'
            69  = '(UTC+01:00) West Central Africa'
            83  = '(UTC+01:00) Windhoek'
            79  = '(UTC+02:00) Amman'
            5   = '(UTC+02:00) Athens, Bucharest, Istanbul'
            80  = '(UTC+02:00) Beirut'
            49  = '(UTC+02:00) Cairo'
            98  = '(UTC+02:00) Damascus'
            50  = '(UTC+02:00) Harare, Pretoria'
            59  = '(UTC+02:00) Helsinki, Kyiv, Riga, Sofia, Tallinn, Vilnius'
            101 = '(UTC+02:00) Istanbul'
            27  = '(UTC+02:00) Jerusalem'
            7   = '(UTC+02:00) Minsk (old)'
            104 = '(UTC+02:00) E. Europe'
            100 = '(UTC+02:00) Kaliningrad (RTZ 1)'
            26  = '(UTC+03:00) Baghdad'
            74  = '(UTC+03:00) Kuwait, Riyadh'
            109 = '(UTC+03:00) Minsk'
            51  = '(UTC+03:00) Moscow, St. Petersburg, Volgograd (RTZ 2)'
            56  = '(UTC+03:00) Nairobi'
            25  = '(UTC+03:30) Tehran'
            24  = '(UTC+04:00) Abu Dhabi, Muscat'
            54  = '(UTC+04:00) Baku'
            106 = '(UTC+04:00) Izhevsk, Samara (RTZ 3)'
            89  = '(UTC+04:00) Port Louis'
            82  = '(UTC+04:00) Tbilisi'
            84  = '(UTC+04:00) Yerevan'
            48  = '(UTC+04:30) Kabul'
            58  = '(UTC+05:00) Ekaterinburg (RTZ 4)'
            87  = '(UTC+05:00) Islamabad, Karachi'
            47  = '(UTC+05:00) Tashkent'
            23  = '(UTC+05:30) Chennai, Kolkata, Mumbai, New Delhi'
            66  = '(UTC+05:30) Sri Jayawardenepura'
            62  = '(UTC+05:45) Kathmandu'
            71  = '(UTC+06:00) Astana'
            102 = '(UTC+06:00) Dhaka'
            46  = '(UTC+06:00) Novosibirsk (RTZ 5)'
            61  = '(UTC+06:30) Yangon (Rangoon)'
            22  = '(UTC+07:00) Bangkok, Hanoi, Jakarta'
            64  = '(UTC+07:00) Krasnoyarsk (RTZ 6)'
            45  = '(UTC+08:00) Beijing, Chongqing, Hong Kong, Urumqi'
            63  = '(UTC+08:00) Irkutsk (RTZ 7)'
            21  = '(UTC+08:00) Kuala Lumpur, Singapore'
            73  = '(UTC+08:00) Perth'
            75  = '(UTC+08:00) Taipei'
            94  = '(UTC+08:00) Ulaanbaatar'
            20  = '(UTC+09:00) Osaka, Sapporo, Tokyo'
            72  = '(UTC+09:00) Seoul'
            70  = '(UTC+09:00) Yakutsk (RTZ 8)'
            19  = '(UTC+09:30) Adelaide'
            44  = '(UTC+09:30) Darwin'
            18  = '(UTC+10:00) Brisbane'
            76  = '(UTC+10:00) Canberra, Melbourne, Sydney'
            43  = '(UTC+10:00) Guam, Port Moresby'
            42  = '(UTC+10:00) Hobart'
            99  = '(UTC+10:00) Magadan'
            68  = '(UTC+10:00) Vladivostok, Magadan (RTZ 9)'
            107 = '(UTC+11:00) Chokurdakh (RTZ 10)'
            41  = '(UTC+11:00) Solomon Is., New Caledonia'
            108 = '(UTC+12:00) Anadyr, Petropavlovsk-Kamchatsky (RTZ 11)'
            17  = '(UTC+12:00) Auckland, Wellington'
            97  = '(UTC+12:00) Coordinated Universal Time+12'
            40  = '(UTC+12:00) Fiji'
            92  = '(UTC+12:00) Petropavlovsk-Kamchatsky - Old'
            67  = "(UTC+13:00) Nuku'alofa"
            16  = '(UTC+13:00) Samoa'
        }

        $timezoneString = $timezoneIDs[$Id]

        if ($null -ne $timezoneString) {
            return $timezoneString
        }

        return $Id
    }

    # Best-effort UPN/display value extraction from a Graph-backed user object, whatever the exact
    # property name used by the calling PnP cmdlet (Get-PnPTeamsUser, Get-PnPMicrosoft365GroupMember/Owner)
    function Get-SPOUserIdentifier {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            $UserObject
        )

        if ($null -eq $UserObject) {
            return $null
        }

        foreach ($propertyName in @('UserPrincipalName', 'User', 'Mail', 'Email', 'DisplayName', 'Id')) {
            $value = $UserObject.$propertyName
            if ($value) {
                return $value
            }
        }

        return $null
    }

    # Turns a version policy (auto-trim flag + manual limits) into one readable label, e.g. 'Automatic' or
    # 'Manual (max 500 versions, expiry: never)'. ExpireVersionsAfterDays = 0 means versions never expire.
    # Returns $null when the policy could not be read, so an unreadable value never looks like a real one.
    function Get-SPOVersionPolicyLabel {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            $AutoTrim,

            [Parameter(Mandatory = $false)]
            $MajorVersionLimit,

            [Parameter(Mandatory = $false)]
            $ExpireVersionsAfterDays
        )

        if ($null -eq $AutoTrim) {
            return $null
        }

        if ($AutoTrim) {
            return 'Automatic'
        }

        $expiry = if ($ExpireVersionsAfterDays -gt 0) { "expiry: $ExpireVersionsAfterDays days" } else { 'expiry: never' }

        return "Manual (max $MajorVersionLimit versions, $expiry)"
    }

    # Detects a Multi-Geo tenant via Get-PnPTenantInstance and normalizes its output into a small,
    # predictable shape (GeoName, AdminUrl, IsDefault). Get-PnPTenantInstance is a recently added PnP
    # cmdlet: its output schema was confirmed against a real Multi-Geo tenant (DataLocation, TenantAdminUrl,
    # IsDefaultDataLocation), but a couple of alternate property names are still tried defensively in case
    # this differs across PnP.PowerShell versions. Returns $null when the cmdlet fails/is unavailable, or
    # when it returns nothing - the caller then falls back to treating the tenant as single-geo.
    function Get-SPOGeoInstance {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            $Connection
        )

        $geoInstances = $null
        try {
            $geoInstances = Get-PnPTenantInstance -Connection $Connection -ErrorAction Stop
        }
        catch {
            Write-Verbose "Get-PnPTenantInstance failed or is unavailable (older PnP.PowerShell version?): $_"
            return $null
        }

        if (-not $geoInstances) {
            return $null
        }

        $cleaned = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($geoInstance in $geoInstances) {
            $geoAdminUrl = $null
            foreach ($propertyName in @('TenantAdminUrl', 'AdminUrl', 'Url', 'SPOAdminUrl')) {
                if ($geoInstance.$propertyName) {
                    $geoAdminUrl = $geoInstance.$propertyName
                    break
                }
            }

            $geoName = $null
            foreach ($propertyName in @('DataLocation', 'GeoLocation', 'Geo', 'Name', 'PreferredDataLocation')) {
                if ($geoInstance.$propertyName) {
                    $geoName = $geoInstance.$propertyName
                    break
                }
            }
            if (-not $geoName) {
                $geoName = $geoAdminUrl
            }

            if (-not $geoAdminUrl) {
                Write-Warning "Unable to determine the admin URL of a Multi-Geo instance from Get-PnPTenantInstance output. Run 'Get-PnPTenantInstance | Get-Member' to find the correct property name. Skipping this geo: $($geoInstance | Out-String)"
                continue
            }

            $cleaned.Add([PSCustomObject]@{
                    GeoName   = $geoName
                    AdminUrl  = $geoAdminUrl
                    IsDefault = [bool]$geoInstance.IsDefaultDataLocation
                })
        }

        return $cleaned
    }

    # Applies the same ExcludeOneDrive/IncludeOneDriveSites/OnlyOneDrive filtering logic against a given
    # connection, whichever geo it belongs to - factored out so the multi-geo enumeration loop below does
    # not repeat this three-way conditional once per geo.
    function Get-SPOGeoSite {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            $Connection,

            [Parameter(Mandatory = $true)]
            [bool]$ExcludeOneDriveSites,

            [Parameter(Mandatory = $true)]
            [bool]$OnlyOneDriveSites
        )

        if ($ExcludeOneDriveSites) {
            $sites = Get-PnPTenantSite -Detailed -Connection $Connection -ErrorAction Stop
        }
        else {
            $sites = Get-PnPTenantSite -IncludeOneDriveSites -Detailed -Connection $Connection -ErrorAction Stop
        }

        if ($OnlyOneDriveSites) {
            $sites = $sites | Where-Object { $_.Url -like '*-my.sharepoint.com/personal/*' }
        }

        return $sites
    }

    # Resolves an array of raw LoginNames (SiteAdmins, Owners, Members, Visitors) into a pipe-joined list
    # of actual identities. On a classic site every LoginName already is an individual user claim; on a
    # group-connected site, the site collection admins / Owners / Members groups instead carry a single
    # claim placeholder representing the whole Microsoft 365 group, of the form
    # 'federateddirectoryclaimprovider|<groupId>_o' (group Owners) or 'federateddirectoryclaimprovider|<groupId>'
    # (group Members). This function expands those placeholders through Graph, caching one Graph call per
    # distinct group GUID for the whole run (the cache hashtables are mutated in place, which works across
    # nested function calls without needing $script:/$using: since a hashtable is a reference type).
    #
    # Each identity coming from a group claim is annotated with "(via M365 group '<name>' Owners|Members)"
    # so the reader can tell a direct SharePoint assignment (no annotation) from an identity that only has
    # rights because it belongs to a nested Microsoft 365 group. The group's display name is resolved once
    # per distinct GUID (GroupNameCache), lazily through Get-PnPMicrosoft365Group when not already known
    # from -M365GroupsDetails; falls back to showing the raw GUID if that lookup also fails.
    #
    # IMPORTANT: the input MUST be an array of individual LoginName strings, never a single string with
    # LoginNames joined by '|'. A SharePoint claim uses '|' as ITS OWN internal delimiter
    # (c:0o.c|federateddirectoryclaimprovider|<guid>_o is ONE claim, not three), so splitting a pre-joined
    # string on '|' to recover individual LoginNames is ambiguous and silently corrupts every claim.
    function Resolve-SPOClaimLoginName {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            [string[]]$LoginNames,

            [Parameter(Mandatory = $true)]
            $Connection,

            [Parameter(Mandatory = $true)]
            [hashtable]$OwnersCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$MembersCache,

            [Parameter(Mandatory = $true)]
            [hashtable]$GroupNameCache
        )

        if (-not ($LoginNames -and $LoginNames.Count -gt 0)) {
            return $null
        }

        $resolvedValues = [System.Collections.Generic.List[string]]::new()

        foreach ($loginName in $LoginNames) {
            if ([string]::IsNullOrEmpty($loginName)) {
                continue
            }

            # Well-known non-user claims, labelled for readability instead of expanded through Graph
            if ($loginName -eq 'c:0(.s|true') {
                $resolvedValues.Add('Everyone')
                continue
            }
            if ($loginName -eq 'c:0!.s|windows') {
                $resolvedValues.Add('All authenticated Windows users')
                continue
            }
            if ($loginName -like '*rolemanager|spo-grid-all-users*') {
                $resolvedValues.Add('Everyone except external users')
                continue
            }

            $claimMatch = [regex]::Match($loginName, 'federateddirectoryclaimprovider\|([0-9a-fA-F-]{36})(_o)?', 'IgnoreCase')

            if (-not $claimMatch.Success) {
                # Individual user claim (i:0#.f|membership|user@contoso.com, i:05.t|adfs|..., ...):
                # the identity value is always the last pipe-delimited segment. No annotation: this user
                # is directly assigned, not inherited through a nested Microsoft 365 group.
                $segments = $loginName -split '\|'
                $resolvedValues.Add($segments[-1])
                continue
            }

            $groupId = $claimMatch.Groups[1].Value
            $isOwnerClaim = $claimMatch.Groups[2].Success
            $roleLabel = if ($isOwnerClaim) { 'Owners' } else { 'Members' }

            try {
                if (-not $GroupNameCache.ContainsKey($groupId)) {
                    try {
                        $groupInfo = Get-PnPMicrosoft365Group -Identity $groupId -Connection $Connection -ErrorAction Stop
                        $GroupNameCache[$groupId] = if ($groupInfo.DisplayName) { $groupInfo.DisplayName } else { $groupId }
                    }
                    catch {
                        $GroupNameCache[$groupId] = $groupId
                    }
                }
                $groupLabel = $GroupNameCache[$groupId]
                $provenanceSuffix = " (via M365 group '$groupLabel' $roleLabel)"

                if ($isOwnerClaim) {
                    if (-not $OwnersCache.ContainsKey($groupId)) {
                        $groupOwners = Get-PnPMicrosoft365GroupOwner -Identity $groupId -Connection $Connection -ErrorAction Stop
                        $OwnersCache[$groupId] = @($groupOwners | ForEach-Object { Get-SPOUserIdentifier -UserObject $_ } | Where-Object { $_ } | ForEach-Object { "$_$provenanceSuffix" })
                    }
                    foreach ($identifier in $OwnersCache[$groupId]) {
                        $resolvedValues.Add($identifier)
                    }
                }
                else {
                    if (-not $MembersCache.ContainsKey($groupId)) {
                        $groupMembers = Get-PnPMicrosoft365GroupMember -Identity $groupId -Connection $Connection -ErrorAction Stop
                        $MembersCache[$groupId] = @($groupMembers | ForEach-Object { Get-SPOUserIdentifier -UserObject $_ } | Where-Object { $_ } | ForEach-Object { "$_$provenanceSuffix" })
                    }
                    foreach ($identifier in $MembersCache[$groupId]) {
                        $resolvedValues.Add($identifier)
                    }
                }
            }
            catch {
                Write-Verbose "Unable to resolve claim group $groupId : $_"
                $resolvedValues.Add("<unresolved:$groupId>")
            }
        }

        return (($resolvedValues | Select-Object -Unique) -join '|')
    }

    # The three helpers below are used only by the -ExportToHtml report; they are no-ops otherwise.
    function ConvertTo-SPOHtmlEncoded {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            [string]$Text
        )

        if ([string]::IsNullOrEmpty($Text)) {
            return ''
        }

        return [System.Net.WebUtility]::HtmlEncode($Text)
    }

    # Splits a *Resolved column value (built by Resolve-SPOClaimLoginName) back into its structural parts:
    # identities assigned directly, identities inherited via a named Microsoft 365 group (grouped by
    # group name + role), and claims that failed to resolve. The flat pipe-joined string remains the
    # object's property value for console/Excel output; this split view is only used to render the HTML tree.
    function Split-SPOResolvedIdentity {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $false)]
            [string]$ResolvedString
        )

        $tree = [PSCustomObject]@{
            Direct     = [System.Collections.Generic.List[string]]::new()
            Groups     = [System.Collections.Generic.List[PSCustomObject]]::new()
            Unresolved = [System.Collections.Generic.List[string]]::new()
        }

        if ([string]::IsNullOrEmpty($ResolvedString)) {
            return $tree
        }

        $groupBuckets = [ordered]@{}

        foreach ($segment in ($ResolvedString -split '\|')) {
            if ([string]::IsNullOrEmpty($segment)) {
                continue
            }

            $unresolvedMatch = [regex]::Match($segment, '^<unresolved:([0-9a-fA-F-]{36})>$')
            if ($unresolvedMatch.Success) {
                $tree.Unresolved.Add($unresolvedMatch.Groups[1].Value)
                continue
            }

            $viaMatch = [regex]::Match($segment, "^(.*?) \(via M365 group '(.*)' (Owners|Members)\)$")
            if ($viaMatch.Success) {
                $bucketKey = "$($viaMatch.Groups[2].Value)|$($viaMatch.Groups[3].Value)"
                if (-not $groupBuckets.Contains($bucketKey)) {
                    $groupBuckets[$bucketKey] = [PSCustomObject]@{
                        GroupName = $viaMatch.Groups[2].Value
                        Role      = $viaMatch.Groups[3].Value
                        Users     = [System.Collections.Generic.List[string]]::new()
                    }
                }
                $groupBuckets[$bucketKey].Users.Add($viaMatch.Groups[1].Value)
            }
            else {
                $tree.Direct.Add($segment)
            }
        }

        foreach ($bucket in $groupBuckets.Values) {
            $tree.Groups.Add($bucket)
        }

        return $tree
    }

    # Renders one collapsible role section (Site Collection Admins / Owners / Members / Visitors) as HTML,
    # with a nested collapsible block per Microsoft 365 group the role inherits from. Returns an empty
    # string when the role was not collected (switch not used, or no data for this site).
    function ConvertTo-SPORoleHtml {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$RoleLabel,

            [Parameter(Mandatory = $true)]
            [string]$RoleCssClass,

            [Parameter(Mandatory = $false)]
            [string]$ResolvedString
        )

        if ([string]::IsNullOrEmpty($ResolvedString)) {
            return ''
        }

        $tree = Split-SPOResolvedIdentity -ResolvedString $ResolvedString
        $groupUsersCount = ($tree.Groups | ForEach-Object { $_.Users.Count } | Measure-Object -Sum).Sum
        $totalCount = $tree.Direct.Count + $groupUsersCount + $tree.Unresolved.Count

        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.Append("<details class='role-block' open><summary><span class='badge-role $RoleCssClass'>$(ConvertTo-SPOHtmlEncoded $RoleLabel)</span><span class='count'>($totalCount)</span></summary><ul class='identity-list'>")

        foreach ($user in $tree.Direct) {
            $null = $sb.Append("<li class='user-item'>$(ConvertTo-SPOHtmlEncoded $user)</li>")
        }

        foreach ($group in $tree.Groups) {
            $null = $sb.Append("<li><details class='group-block'><summary><span class='badge-group'>M365 group '$(ConvertTo-SPOHtmlEncoded $group.GroupName)' - $($group.Role)</span><span class='count'>($($group.Users.Count))</span></summary><ul class='identity-list'>")
            foreach ($user in $group.Users) {
                $null = $sb.Append("<li class='user-item'>$(ConvertTo-SPOHtmlEncoded $user)</li>")
            }
            $null = $sb.Append('</ul></details></li>')
        }

        foreach ($guid in $tree.Unresolved) {
            $null = $sb.Append("<li class='user-item unresolved'>Unresolved group claim: $(ConvertTo-SPOHtmlEncoded $guid)</li>")
        }

        $null = $sb.Append('</ul></details>')
        return $sb.ToString()
    }

    # Evaluates the findings summary of the Excel and HTML exports (see FINDINGS SUMMARY in the help). Every
    # control is computed from the columns already present on the site objects: no extra call is made. A
    # control whose input columns were not collected is reported as 'NotCollected' with the switch it needs,
    # never as zero affected - same reasoning as the optional columns being absent rather than empty.
    function Get-SPOSiteFinding {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [AllowEmptyCollection()]
            [object[]]$Sites,

            [Parameter(Mandatory = $true)]
            [hashtable]$CollectedSwitches
        )

        $controls = @(
            [PSCustomObject]@{ Category = 'Internal oversharing'; Code = 'EID-GROUP-08'; Control = 'Public Microsoft 365 groups'; Unit = 'group(s)'; Required = @('M365GroupsDetails'); Explanation = 'A public Microsoft 365 group is open to every employee, who can join it on their own and read its site content.' }
            [PSCustomObject]@{ Category = 'Internal oversharing'; Code = 'SPO-SHARE-04'; Control = 'Private Microsoft 365 groups with a site open to all internal users'; Unit = 'group(s)'; Required = @('M365GroupsDetails', 'IncludeEveryoneAccess'); Fallback = @('M365GroupsDetails', 'IncludeSiteMembers', 'IncludeSiteVisitors'); Explanation = "The group is private, but its site grants access to 'Everyone except external users'. The owners believe the space is closed; it is not." }
            [PSCustomObject]@{ Category = 'Internal oversharing'; Code = 'SPO-SHARE-05'; Control = 'Private Microsoft 365 groups with a site open to Everyone'; Unit = 'group(s)'; Required = @('M365GroupsDetails', 'IncludeEveryoneAccess'); Fallback = @('M365GroupsDetails', 'IncludeSiteMembers', 'IncludeSiteVisitors'); Explanation = "The group is private, but its site grants access to 'Everyone', guests included." }
            [PSCustomObject]@{ Category = 'Internal oversharing'; Code = 'SPO-SHARE-01'; Control = 'Sites open to all internal users'; Unit = 'site(s)'; Required = @('IncludeEveryoneAccess'); Fallback = @('IncludeSiteMembers', 'IncludeSiteVisitors'); Explanation = "The site grants access to 'Everyone except external users': every employee can read it, whatever their need." }
            [PSCustomObject]@{ Category = 'Internal oversharing'; Code = 'SPO-SHARE-02'; Control = 'Sites open to Everyone'; Unit = 'site(s)'; Required = @('IncludeEveryoneAccess'); Fallback = @('IncludeSiteMembers', 'IncludeSiteVisitors'); Explanation = "The site grants access to 'Everyone', guests included." }
            [PSCustomObject]@{ Category = 'External exposure'; Code = 'EID-GROUP-11'; Control = 'Sites and Microsoft 365 groups with guests'; Unit = 'site(s)'; Required = @('IncludeSiteMembers'); Explanation = 'Guest accounts have access to the site, through its SharePoint groups or its Microsoft 365 group.' }
            [PSCustomObject]@{ Category = 'External exposure'; Code = 'SPO-SHARE-03'; Control = 'Anonymous sharing links'; Unit = 'site(s)'; Required = @('IncludeSharingLinks'); Explanation = "An 'Anyone with the link' link requires no authentication: whoever holds the link reaches the content. Flexible links are not counted here, their scope is only known through Get-SPOSharingLinkReport -IncludeGraphDetails." }
            [PSCustomObject]@{ Category = 'Governance'; Code = 'EID-GROUP-13'; Control = 'Microsoft 365 groups without owner'; Unit = 'group(s)'; Required = @('M365GroupsDetails'); Explanation = 'Nobody is accountable for the group: its access cannot be reviewed or approved.' }
            [PSCustomObject]@{ Category = 'Governance'; Code = 'EID-GROUP-14'; Control = 'Microsoft 365 groups with external owners only'; Unit = 'group(s)'; Required = @('M365GroupsDetails'); Explanation = 'Every owner is a guest: the organization no longer controls the access to its own group.' }
            [PSCustomObject]@{ Category = 'Governance'; Code = 'EID-GROUP-15'; Control = 'Microsoft 365 groups with external users only'; Unit = 'group(s)'; Required = @('M365GroupsDetails'); Explanation = 'The group only contains guests: no employee follows what is shared in it.' }
            [PSCustomObject]@{ Category = 'Governance'; Code = 'SPO-OWNER-01'; Control = 'Sites without an active administrator'; Unit = 'site(s)'; Required = @('IncludeSiteAdminsActivity'); Explanation = 'No site collection administrator is still active (disabled, deleted or no recent sign-in): nobody manages the site access anymore. OneDrive sites are not evaluated.' }
        )

        $details = [System.Collections.Generic.List[PSCustomObject]]::new()
        $itemsByCode = @{ 'SPO-SHARE-03' = 0 }

        $roleLabels = [ordered]@{
            SiteAdminsResolved   = 'Admins'
            SiteOwnersResolved   = 'Owners'
            SiteMembersResolved  = 'Members'
            SiteVisitorsResolved = 'Visitors'
        }

        foreach ($site in $Sites) {
            $siteFindings = [ordered]@{}
            $isGroup = $site.Type -eq 'M365 Group'
            $visibility = $site.M365GroupVisibility

            # Where the 'everyone' claims are granted. With IncludeEveryoneAccess, the EveryoneAccess column
            # already covers every permission of the site (associated groups included) and of the libraries
            # with unique permissions, so only the site collection admins are added from the role columns;
            # without it, the associated groups of the role columns are the only source.
            # The same pass collects the guests (the module's '#EXT#' convention), deduplicated once their
            # provenance annotation is stripped.
            $internalWhere = [System.Collections.Generic.List[string]]::new()
            $everyoneWhere = [System.Collections.Generic.List[string]]::new()
            $siteGuests = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $everyoneAccessCollected = $null -ne $site.EveryoneAccess

            foreach ($column in $roleLabels.Keys) {
                $value = $site.$column
                if ([string]::IsNullOrEmpty($value)) {
                    continue
                }

                $useForEveryone = -not $everyoneAccessCollected -or $column -eq 'SiteAdminsResolved'
                $whereLabel = "$($roleLabels[$column]) group"
                if ($column -eq 'SiteAdminsResolved') {
                    $whereLabel = 'site collection admins'
                }

                foreach ($segment in ($value -split '\|')) {
                    if ($segment -eq 'Everyone except external users') {
                        if ($useForEveryone -and -not $internalWhere.Contains($whereLabel)) { $internalWhere.Add($whereLabel) }
                    }
                    elseif ($segment -eq 'Everyone') {
                        if ($useForEveryone -and -not $everyoneWhere.Contains($whereLabel)) { $everyoneWhere.Add($whereLabel) }
                    }
                    elseif ($segment -match '#EXT#') {
                        $null = $siteGuests.Add(($segment -replace " \(via M365 group '.*' (Owners|Members)\)$", ''))
                    }
                }
            }

            if ($everyoneAccessCollected) {
                foreach ($grant in ("$($site.EveryoneAccess)" -split '\|' | Where-Object { $_ })) {
                    if ($grant -like 'Everyone except external users on *') {
                        $internalWhere.Add($grant.Substring('Everyone except external users on '.Length))
                    }
                    elseif ($grant -like 'Everyone on *') {
                        $everyoneWhere.Add($grant.Substring('Everyone on '.Length))
                    }
                }
            }

            # $null means "could not be read" (the warning was already written), '' means "read, nobody".
            # Assigned outside an if expression: an empty array returned by an if statement collapses to $null
            $groupOwnersRead = $null -ne $site.M365GroupOwners
            $groupOwners = @("$($site.M365GroupOwners)" -split '\|' | Where-Object { $_ })
            $groupOwnersAllExternal = $groupOwners.Count -gt 0 -and @($groupOwners | Where-Object { $_ -notmatch '#EXT#' }).Count -eq 0

            if ($isGroup -and $visibility -eq 'Public') {
                $siteFindings['EID-GROUP-08'] = 'Visibility: Public'
            }

            if ($isGroup -and $visibility -and $visibility -ne 'Public') {
                if ($internalWhere.Count -gt 0) {
                    $siteFindings['SPO-SHARE-04'] = "Visibility: $visibility - 'Everyone except external users' on $($internalWhere -join '; ')"
                }
                if ($everyoneWhere.Count -gt 0) {
                    $siteFindings['SPO-SHARE-05'] = "Visibility: $visibility - 'Everyone' on $($everyoneWhere -join '; ')"
                }
            }

            if ($internalWhere.Count -gt 0) {
                $siteFindings['SPO-SHARE-01'] = "'Everyone except external users' on $($internalWhere -join '; ')"
            }

            if ($everyoneWhere.Count -gt 0) {
                $siteFindings['SPO-SHARE-02'] = "'Everyone' on $($everyoneWhere -join '; ')"
            }

            $groupGuestsCount = if ($null -ne $site.M365GroupGuestsCount) { [int]$site.M365GroupGuestsCount } else { 0 }
            if ($siteGuests.Count -gt 0 -or $groupGuestsCount -gt 0) {
                $guestParts = [System.Collections.Generic.List[string]]::new()
                if ($siteGuests.Count -gt 0) { $guestParts.Add("$($siteGuests.Count) guest(s) in the site rights") }
                if ($groupGuestsCount -gt 0) { $guestParts.Add("$groupGuestsCount guest(s) in the Microsoft 365 group") }
                $siteFindings['EID-GROUP-11'] = $guestParts -join ', '
            }

            if ($site.SharingLinksAnyoneCount -gt 0) {
                $siteFindings['SPO-SHARE-03'] = "$($site.SharingLinksAnyoneCount) anonymous link(s)"
                $itemsByCode['SPO-SHARE-03'] += [int]$site.SharingLinksAnyoneCount
            }

            if ($isGroup -and $groupOwnersRead -and $groupOwners.Count -eq 0) {
                $siteFindings['EID-GROUP-13'] = 'No owner'
            }

            if ($isGroup -and $groupOwnersAllExternal) {
                $siteFindings['EID-GROUP-14'] = "Owners: $($groupOwners -join ', ')"
            }

            if ($isGroup -and $groupGuestsCount -gt 0 -and $site.M365GroupMembersCount -eq 0 -and $groupOwnersRead -and ($groupOwners.Count -eq 0 -or $groupOwnersAllExternal)) {
                $siteFindings['EID-GROUP-15'] = "$groupGuestsCount guest(s), no internal member or owner"
            }

            # Only a complete check ('OK') gives a verdict: a site with a failed lookup is not flagged on a guess
            if ($site.SiteAdminsActivityStatus -eq 'OK' -and $site.SiteAdminsActiveCount -eq 0) {
                $siteFindings['SPO-OWNER-01'] = "$($site.SiteAdminsNotActive -replace '\|', ', ')"
            }

            foreach ($code in $siteFindings.Keys) {
                $detail = [PSCustomObject][ordered]@{
                    Code    = $code
                    Control = ($controls | Where-Object { $_.Code -eq $code }).Control
                    SPTitle = $site.SPTitle
                    Url     = $site.Url
                    Detail  = $siteFindings[$code]
                }
                $details.Add($detail)
            }
        }

        $summary = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($control in $controls) {
            # Required gives the full coverage. Some controls also accept a smaller Fallback set, evaluated
            # with a narrower scope: the status says so and names the switch that would complete it
            $missingSwitches = @($control.Required | Where-Object { -not $CollectedSwitches[$_] })
            # Not assigned from an if expression: an empty array returned by an if statement collapses to $null
            $hasFallback = $control.PSObject.Properties.Name -contains 'Fallback'
            $missingFallback = @($control.Fallback | Where-Object { $_ -and -not $CollectedSwitches[$_] })
            $isEvaluated = $missingSwitches.Count -eq 0 -or ($hasFallback -and $missingFallback.Count -eq 0)

            $status = if ($missingSwitches.Count -eq 0) {
                'Evaluated'
            }
            elseif ($isEvaluated) {
                "Evaluated (partial, associated groups only: add -$($missingSwitches -join ', -'))"
            }
            else {
                "NotCollected (requires -$($missingSwitches -join ', -'))"
            }

            $summaryRow = [PSCustomObject][ordered]@{
                Category      = $control.Category
                Code          = $control.Code
                Control       = $control.Control
                Status        = $status
                AffectedCount = if ($isEvaluated) { @($details | Where-Object { $_.Code -eq $control.Code }).Count } else { $null }
                AffectedUnit  = $control.Unit
                ItemsCount    = if ($isEvaluated -and $itemsByCode.ContainsKey($control.Code)) { $itemsByCode[$control.Code] } else { $null }
                ItemsUnit     = if ($itemsByCode.ContainsKey($control.Code)) { 'anonymous link(s)' } else { $null }
                Explanation   = $control.Explanation
            }
            $summary.Add($summaryRow)
        }

        # Details of a control that is not fully collected would be a partial, misleading list: dropped
        $evaluatedCodes = @($summary | Where-Object { $_.Status -like 'Evaluated*' }).Code
        $evaluatedDetails = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($detail in $details) {
            if ($evaluatedCodes -contains $detail.Code) {
                $evaluatedDetails.Add($detail)
            }
        }

        return [PSCustomObject]@{
            Summary = $summary
            Details = $evaluatedDetails
        }
    }

    # Invoke-PnPGraphMethod with a retry on transient errors (throttling, service unavailable, timeout), with
    # an exponential backoff (2, 4, 8 seconds). Any other error, or a transient one that persists, is thrown
    # back to the caller unchanged, so a permission or not found error is never retried.
    function Invoke-SPOGraphWithRetry {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$Url,

            [Parameter(Mandatory = $true)]
            $Connection,

            [Parameter(Mandatory = $false)]
            [switch]$Raw,

            [Parameter(Mandatory = $false)]
            [ValidateRange(1, 10)]
            [int]$MaxAttempts = 4
        )

        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            try {
                return (Invoke-PnPGraphMethod -Url $Url -Raw:$Raw -Connection $Connection -ErrorAction Stop)
            }
            catch {
                $isTransient = $_.Exception.Message -match '429|TooManyRequests|Too many requests|throttl|503|ServiceUnavailable|504|GatewayTimeout|timed out|timeout'
                if (-not $isTransient -or $attempt -eq $MaxAttempts) {
                    throw
                }

                $delaySeconds = [math]::Pow(2, $attempt)
                Write-Verbose "Transient Graph error on $Url (attempt $attempt/$MaxAttempts), retrying in $delaySeconds s: $($_.Exception.Message)"
                Start-Sleep -Seconds $delaySeconds
            }
        }
    }

    # Fallback of IncludeSiteAdminsActivity when signInActivity cannot be read (AuditLog.Read.All missing or no
    # Entra ID P1): the Microsoft 365 active users report gives, per user, the last activity date of each
    # service (Exchange, OneDrive, SharePoint, Teams...). Downloaded once for the whole run (one CSV over the
    # last 180 days, the longest period available) and cached in ReportState. Needs the Microsoft Graph
    # application permission Reports.Read.All, and no license. When that report cannot be used either, the
    # admin is judged on the enabled/disabled state alone.
    function Get-SPOUsageActivityState {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string]$UserPrincipalName,

            [Parameter(Mandatory = $true)]
            $Connection,

            [Parameter(Mandatory = $true)]
            [hashtable]$ReportState,

            [Parameter(Mandatory = $true)]
            [datetime]$InactiveSince,

            [Parameter(Mandatory = $true)]
            [int]$InactiveDaysThreshold
        )

        if ($ReportState.Status -eq 'NotLoaded') {
            # Marked first: a failed download is not retried for every admin
            $ReportState.Status = 'Unavailable'

            try {
                $csvContent = Invoke-SPOGraphWithRetry -Url "v1.0/reports/getOffice365ActiveUserDetail(period='D180')" -Raw -Connection $Connection
                if ($csvContent -is [byte[]]) {
                    $csvContent = [System.Text.Encoding]::UTF8.GetString($csvContent)
                }
                $reportRows = @("$csvContent".TrimStart([char]0xFEFF) | ConvertFrom-Csv)
            }
            catch {
                Write-Warning "The Microsoft 365 usage report cannot be read either (Microsoft Graph application permission Reports.Read.All missing?): an enabled admin is considered active. Detail: $($_.Exception.Message)"
                $reportRows = $null
            }

            if ($null -ne $reportRows) {
                $upnColumn = 'User Principal Name'

                if ($reportRows.Count -eq 0 -or -not ($reportRows[0].PSObject.Properties.Name -contains $upnColumn)) {
                    Write-Warning 'The Microsoft 365 usage report came back empty or in an unexpected format: an enabled admin is considered active.'
                }
                elseif (@($reportRows | Where-Object { $_.$upnColumn -match '@' }).Count -eq 0) {
                    # Report concealment (tenant default since 2021) replaces every UPN with a hash
                    Write-Warning "The Microsoft 365 usage report hides user names, so it cannot be matched to the admins: an enabled admin is considered active. Turn off 'Display concealed user, group, and site names in all reports' (Microsoft 365 admin center > Settings > Org settings > Reports) to use it."
                }
                else {
                    $dateColumns = @($reportRows[0].PSObject.Properties.Name | Where-Object { $_ -match 'Last Activity Date$' })
                    $lastActivityByUpn = @{}

                    foreach ($reportRow in $reportRows) {
                        $lastActivity = $null
                        foreach ($dateColumn in $dateColumns) {
                            $parsedDate = [datetime]::MinValue
                            if ($reportRow.$dateColumn -and [datetime]::TryParseExact("$($reportRow.$dateColumn)", 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]'AdjustToUniversal, AssumeUniversal', [ref]$parsedDate)) {
                                if ($null -eq $lastActivity -or $parsedDate -gt $lastActivity) {
                                    $lastActivity = $parsedDate
                                }
                            }
                        }
                        $lastActivityByUpn[$reportRow.$upnColumn.ToLowerInvariant()] = $lastActivity
                    }

                    $ReportState.LastActivityByUpn = $lastActivityByUpn
                    $ReportState.Status = 'Loaded'
                    Write-Host -ForegroundColor Cyan "Microsoft 365 usage report loaded: $($lastActivityByUpn.Count) user(s), last 180 days."

                    if ($InactiveDaysThreshold -gt 180) {
                        Write-Warning "The usage report only covers the last 180 days: an admin with no activity in that window is reported inactive, even though InactiveDaysThreshold is $InactiveDaysThreshold."
                    }
                }
            }
        }

        if ($ReportState.Status -ne 'Loaded') {
            return [PSCustomObject]@{ State = 'Active'; Label = 'Enabled, activity not available' }
        }

        $upnKey = $UserPrincipalName.ToLowerInvariant()
        if (-not $ReportState.LastActivityByUpn.ContainsKey($upnKey)) {
            # Typically an unlicensed account: no verdict rather than a guess
            return [PSCustomObject]@{ State = 'Unknown'; Label = 'Not in the usage report' }
        }

        $lastActivity = $ReportState.LastActivityByUpn[$upnKey]
        if ($null -eq $lastActivity) {
            return [PSCustomObject]@{ State = 'Inactive'; Label = 'No Microsoft 365 activity in the last 180 days' }
        }
        if ($lastActivity -lt $InactiveSince) {
            return [PSCustomObject]@{ State = 'Inactive'; Label = "Last Microsoft 365 activity $($lastActivity.ToString('yyyy-MM-dd'))" }
        }
        return [PSCustomObject]@{ State = 'Active'; Label = "Last Microsoft 365 activity $($lastActivity.ToString('yyyy-MM-dd'))" }
    }

    if ($ExcludeOneDrive.IsPresent -and $OnlyOneDrive.IsPresent) {
        Write-Warning 'ExcludeOneDrive and OnlyOneDrive cannot be used together. Choose one of them.'
        return
    }

    # IncludeAllDetails turns on every detail switch below, so callers don't have to remember each one.
    # It does NOT touch ExcludeOneDrive/OnlyOneDrive: those are mutually exclusive site filters, not
    # detail enrichment, and forcing either one on would silently change which sites are even returned.
    # Explicit request only: when it comes from IncludeAllDetails, a run without export must not warn about it
    $storageSummaryRequested = $IncludeStorageSummary.IsPresent

    if ($IncludeAllDetails.IsPresent) {
        $IncludeSiteAdmins = [switch]$true
        $IncludeSiteMembers = [switch]$true
        $IncludeSiteVisitors = [switch]$true
        $IncludeSharingLinks = [switch]$true
        $RegionalSettingsDetails = [switch]$true
        $M365GroupsDetails = [switch]$true
        $IncludeEveryoneAccess = [switch]$true
        $IncludeSharingLinksDetails = [switch]$true
        $IncludeSiteAdminsActivity = [switch]$true
        $IncludeStorageSummary = [switch]$true

        # Everything is on, the slow collections included: say it upfront, with the permissions they need,
        # so a long run or a warning about a missing permission later on is not a surprise
        Write-Host -ForegroundColor Yellow 'IncludeAllDetails turns on every collection, the slow ones included. Expect a long run on a large tenant:'
        Write-Host -ForegroundColor Yellow '  - per site: admins, Owners/Members/Visitors, Everyone access (libraries with unique permissions), sharing links, regional settings'
        Write-Host -ForegroundColor Yellow '  - IncludeSharingLinksDetails: every shared file of every site holding a link (the slowest part)'
        Write-Host -ForegroundColor Yellow '  - M365GroupsDetails: Microsoft 365 group and Teams layer of each group-connected site'
        Write-Host -ForegroundColor Yellow '  - IncludeSiteAdminsActivity: one Microsoft Graph lookup per distinct admin'
        Write-Host -ForegroundColor Yellow '  Microsoft Graph application permissions used on top of SharePoint Sites.FullControl.All: Group.Read.All, User.Read.All, AuditLog.Read.All (or Reports.Read.All). A missing one only skips its part, with a warning.'
        Write-Host -ForegroundColor Yellow '  For a faster run, use the individual switches instead.'
    }

    # IncludeSharingLinksDetails is the drill-down of IncludeSharingLinks and needs its counts to know which
    # sites are worth visiting, so it turns the cheap switch on. Its cost is driven by the number of shared
    # items, a different order of magnitude: the IncludeAllDetails notice above calls it out as the slowest part.
    if ($IncludeSharingLinksDetails.IsPresent) {
        $IncludeSharingLinks = [switch]$true
    }

    # IncludeSiteAdminsActivity checks the accounts listed by IncludeSiteAdmins, so it turns that switch on.
    # It needs extra Microsoft Graph application permissions (User.Read.All, AuditLog.Read.All or
    # Reports.Read.All); without them it only warns and leaves SPO-OWNER-01 not evaluated.
    if ($IncludeSiteAdminsActivity.IsPresent) {
        $IncludeSiteAdmins = [switch]$true
    }

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

    # Build the common Connect-PnPOnline parameters, reused for the admin connection and for every site
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

    # AdminUrl explicitly supplied by the caller targets one precise geo on purpose: capture this BEFORE
    # the derivation below overwrites $AdminUrl, so the site enumeration knows not to loop over every geo.
    $adminUrlWasExplicit = [bool]$AdminUrl

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

    # Tenant-wide version history policy, read once: it is the default applied to new sites and to every
    # site that inherits it, so it feeds the VersionPolicyEffective and VersionPolicyTenantDefault columns.
    # A failed read (limited app permissions) leaves them empty rather than aborting the report.
    $tenantVersionPolicyLabel = $null
    $tenantVersionPolicy = $adminCenter.TenantSettings

    if ($tenantVersionPolicy) {
        $tenantVersionPolicyLabel = Get-SPOVersionPolicyLabel -AutoTrim $tenantVersionPolicy.EnableAutoExpirationVersionTrim -MajorVersionLimit $tenantVersionPolicy.MajorVersionLimit -ExpireVersionsAfterDays $tenantVersionPolicy.ExpireVersionsAfterDays
        Write-Host -ForegroundColor Cyan "Tenant version history policy (default for new sites): $tenantVersionPolicyLabel"
    }
    else {
        Write-Warning "Unable to read the tenant version history policy: $($adminCenter.TenantError). The VersionPolicyEffective and VersionPolicyTenantDefault columns will be empty."
    }

    # Get-PnPTenantSite -Detailed is a single call that can last several minutes on a large tenant
    Write-Host -ForegroundColor Cyan 'Retrieving SharePoint Online sites (one call, can take several minutes on a large tenant)...'

    # Geo name -> admin center URL / is-default-geo, used later by MultiGeoStats to report each geo's admin
    # URL and default flag without tagging every single site with them (both are properties of the geo,
    # not of each site). Note: in the multi-geo loop below, the default geo keeps its real GeoName (e.g.
    # 'FRA'), never the literal string 'Default' - so IsDefaultGeo cannot be derived by comparing the geo
    # name to 'Default', it must be looked up here instead.
    $geoAdminUrlByName = @{}
    $geoIsDefaultByName = @{}
    $geoEnumerationErrorCount = 0

    $spoSites = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($ExcludeGeo -and @($ExcludeGeo).Count -gt 0 -and ($SiteUrl -or $adminUrlWasExplicit)) {
        Write-Warning 'ExcludeGeo has no effect here: SiteUrl and an explicit AdminUrl both already restrict the report to a single geo.'
    }

    try {
        if ($SiteUrl) {
            # Single site targeted explicitly: no Multi-Geo notion applies, unchanged from before.
            try {
                $singleSite = Get-PnPTenantSite -Identity $SiteUrl -Detailed -Connection $adminConnection -ErrorAction Stop
            }
            catch {
                # Tenant-level read denied ('Attempted to perform an unauthorized operation': app limited to
                # Sites.Selected, or identity without the SharePoint Administrator role). The site-level pass
                # below still works through a direct connection to the site itself, so degrade to a minimal
                # site object and keep going instead of aborting the whole report
                Write-Warning "Unable to read the tenant-level properties of $SiteUrl ($($_.Exception.Message.Trim())). Continuing with site-level data only: storage, owner, sharing and the other tenant-level columns will be empty for this site."

                $singleSite = [PSCustomObject]@{
                    Url                 = $SiteUrl
                    Title               = $null
                    GroupId             = $null
                    Owner               = $null
                    StorageQuota        = $null
                    StorageUsage        = $null
                    StorageUsageCurrent = $null
                    # Carried into the SiteDataStatus column so the failure is visible in the report itself
                    TenantDataError     = "$($_.Exception.Message.Trim())"
                }
            }

            # No fabricated tenant-level values when the tenant read failed: empty beats wrong
            $singleSiteGeo = if ($singleSite.PSObject.Properties.Name -contains 'TenantDataError') { $null } else { 'Default' }
            $singleSite | Add-Member -NotePropertyName 'Geo' -NotePropertyValue $singleSiteGeo -Force
            $spoSites.Add($singleSite)
            $geoAdminUrlByName['Default'] = $AdminUrl
            $geoIsDefaultByName['Default'] = $true
        }
        elseif ($adminUrlWasExplicit) {
            # An explicit AdminUrl means the caller deliberately targets that one geo: no Multi-Geo loop.
            Write-Verbose 'AdminUrl was explicitly provided: restricting enumeration to this geo only.'
            $geoSitesRaw = Get-SPOGeoSite -Connection $adminConnection -ExcludeOneDriveSites $ExcludeOneDrive.IsPresent -OnlyOneDriveSites $OnlyOneDrive.IsPresent
            foreach ($s in $geoSitesRaw) {
                $s | Add-Member -NotePropertyName 'Geo' -NotePropertyValue 'Default' -Force
            }
            foreach ($geoSite in @($geoSitesRaw)) { $spoSites.Add($geoSite) }
            $geoAdminUrlByName['Default'] = $AdminUrl
            $geoIsDefaultByName['Default'] = $true
        }
        else {
            # Full tenant-wide report with an auto-derived AdminUrl: cover every Multi-Geo location so
            # sites in satellite geos are never silently missed.
            $geoInstances = Get-SPOGeoInstance -Connection $adminConnection

            if ($geoInstances -and $ExcludeGeo -and @($ExcludeGeo).Count -gt 0) {
                $excludedInstances = @($geoInstances | Where-Object { $_.GeoName -in $ExcludeGeo })
                if ($excludedInstances.Count -gt 0) {
                    Write-Host -ForegroundColor Yellow "Excluding geo location(s) from enumeration: $(($excludedInstances.GeoName) -join ', ')"
                }

                $unmatchedExclusions = @($ExcludeGeo | Where-Object { $_ -notin $geoInstances.GeoName })
                if ($unmatchedExclusions.Count -gt 0) {
                    Write-Warning "ExcludeGeo value(s) not found among detected geo locations, ignored: $($unmatchedExclusions -join ', ')"
                }

                $geoInstances = @($geoInstances | Where-Object { $_.GeoName -notin $ExcludeGeo })
            }

            if (-not $geoInstances -or @($geoInstances).Count -eq 0) {
                if ($ExcludeGeo -and @($ExcludeGeo).Count -gt 0) {
                    Write-Warning 'All Multi-Geo locations were excluded by ExcludeGeo - nothing to enumerate.'
                    return
                }

                # No Multi-Geo instances detected at all: a genuinely single-geo tenant, or the detection
                # cmdlet is unavailable/failed. Enumerate the default geo directly, no per-instance loop needed.
                Write-Host -ForegroundColor Cyan 'Single-geo tenant (or Multi-Geo detection unavailable) - enumerating the default geo only.'
                $geoSitesRaw = Get-SPOGeoSite -Connection $adminConnection -ExcludeOneDriveSites $ExcludeOneDrive.IsPresent -OnlyOneDriveSites $OnlyOneDrive.IsPresent
                foreach ($s in $geoSitesRaw) {
                    $s | Add-Member -NotePropertyName 'Geo' -NotePropertyValue 'Default' -Force
                }
                foreach ($geoSite in @($geoSitesRaw)) { $spoSites.Add($geoSite) }
                $geoAdminUrlByName['Default'] = $AdminUrl
                $geoIsDefaultByName['Default'] = $true
            }
            else {
                # One or more Multi-Geo instances remain after any ExcludeGeo filtering. Looping here even
                # when only one instance is left (rather than special-casing it) matters: after exclusion,
                # that single remaining geo could be a satellite, not the default - the per-instance
                # IsDefault check below is what correctly decides whether to reuse $adminConnection or open
                # a fresh connection, instead of assuming "only one left" always means "the default one".
                if (@($geoInstances).Count -gt 1) {
                    Write-Host -ForegroundColor Cyan "Multi-Geo tenant detected: $(@($geoInstances).Count) geo location(s). Enumerating sites in each geo..."
                }

                foreach ($geoInstance in $geoInstances) {
                    try {
                        # The default geo is already the one $adminConnection points to: reuse it instead
                        # of opening a second, redundant connection to the same admin center.
                        $geoConnectionToUse = if ($geoInstance.IsDefault) {
                            $adminConnection
                        }
                        else {
                            Connect-PnPOnline -Url $geoInstance.AdminUrl @pnpAuthParams -ReturnConnection -ErrorAction Stop
                        }

                        $geoSitesRaw = Get-SPOGeoSite -Connection $geoConnectionToUse -ExcludeOneDriveSites $ExcludeOneDrive.IsPresent -OnlyOneDriveSites $OnlyOneDrive.IsPresent
                        $geoConnectionToUse = $null

                        foreach ($s in $geoSitesRaw) {
                            $s | Add-Member -NotePropertyName 'Geo' -NotePropertyValue $geoInstance.GeoName -Force
                        }
                        foreach ($geoSite in @($geoSitesRaw)) { $spoSites.Add($geoSite) }
                        $geoAdminUrlByName[$geoInstance.GeoName] = $geoInstance.AdminUrl
                        $geoIsDefaultByName[$geoInstance.GeoName] = $geoInstance.IsDefault

                        Write-Host -ForegroundColor Cyan "Geo '$($geoInstance.GeoName)': $(@($geoSitesRaw).Count) site(s)."
                    }
                    catch {
                        $geoEnumerationErrorCount++
                        Write-Warning "Unable to enumerate sites for geo '$($geoInstance.GeoName)' ($($geoInstance.AdminUrl)): $_"
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Unable to retrieve the SharePoint Online sites: $_"
        return
    }

    $sitesCount = @($spoSites).Count
    Write-Host -ForegroundColor Cyan "SharePoint Online sites count: $sitesCount"

    if ($MultiGeoStats.IsPresent) {
        # The sites are already enumerated across every geo above (see the site enumeration block), each
        # tagged with its Geo property: no separate connection or re-enumeration is needed here anymore,
        # just an aggregation of what has already been fetched.
        Write-Host -ForegroundColor Cyan 'Building per-geo site counts from the sites already enumerated above...'

        $statsArray = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($geoGroup in ($spoSites | Group-Object -Property Geo)) {
            $geoOneDriveSites = @($geoGroup.Group | Where-Object { $_.Url -like '*-my.sharepoint.com/personal/*' })

            $statsArray.Add([PSCustomObject][ordered]@{
                    GeoName         = $geoGroup.Name
                    IsDefaultGeo    = [bool]$geoIsDefaultByName[$geoGroup.Name]
                    AdminUrl        = $geoAdminUrlByName[$geoGroup.Name]
                    TotalSites      = $geoGroup.Count
                    SharePointSites = $geoGroup.Count - $geoOneDriveSites.Count
                    OneDriveSites   = $geoOneDriveSites.Count
                })
        }

        if ($ExportToHtml.IsPresent) {
            Write-Warning 'MultiGeoStats ignores ExportToHtml: the HTML export renders the detailed rights tree, which does not apply to a stats-only run. Use ExportToExcel instead, or pipe the returned rows to Export-Csv yourself.'
        }

        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
            $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-SPOSiteReportMultiGeoStats.xlsx"

            if ($ExcelTemplatePath) {
                if (-not (Test-Path -Path $ExcelTemplatePath -PathType Leaf)) {
                    Write-Warning "ExcelTemplatePath not found: $ExcelTemplatePath"
                    return
                }
                Copy-Item -Path $ExcelTemplatePath -Destination $excelFilePath -Force
                Write-Host -ForegroundColor Cyan "Using Excel template: $ExcelTemplatePath"
            }

            Write-Host -ForegroundColor Cyan 'Exporting Multi-Geo stats to Excel...'
            $statsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-MultiGeoStats' -TableStyle Light9
            Add-PS365ExcelAboutSheet -Path $excelFilePath
            Write-Host -ForegroundColor Green 'Excel report exported to:'
            Write-Host -ForegroundColor White "  $excelFilePath"
        }
        else {
            return $statsArray
        }
    }

    # Build the Microsoft 365 groups lookup table once, keyed on the SharePoint site URL. Served entirely
    # through the PnP app-only connection: no Connect-MgGraph session exists here, so the module's mandatory
    # Test-MgGraphPermission check does not apply (it inspects Get-MgContext, which PnP never populates).
    # The app registration certificate must instead hold the Graph application permission
    # Group.ReadWrite.All (or Group.Read.All for the read-only subset) alongside Sites.FullControl.All.
    $m365GroupsLookup = @{}

    # Group GUID -> DisplayName, used by Resolve-SPOClaimLoginName to annotate identities inherited through
    # a nested Microsoft 365 group. Pre-seeded here for free when M365GroupsDetails already fetched every
    # group of the tenant; otherwise Resolve-SPOClaimLoginName lazily fills it in, one Graph call per
    # distinct group GUID actually encountered in a claim, whether or not M365GroupsDetails is used.
    $groupNameCache = @{}

    # $false when the Microsoft 365 groups could not be read at all: the group findings are then reported as not
    # collected instead of 'clear', and the report goes on without the group columns
    $m365GroupsAvailable = $M365GroupsDetails.IsPresent

    if ($M365GroupsDetails.IsPresent) {
        # Only the sites carrying a GroupId can have a Microsoft 365 group behind them. Computing that list
        # first avoids two pointless situations: enumerating every group of the tenant to enrich a handful
        # of sites, and enumerating them for a scope that cannot contain a single group-connected site (a
        # OneDrive-only run, or -SiteUrl on a classic site). Get-PnPTenantSite already returned GroupId, so
        # this costs nothing.
        $groupConnectedSites = @($spoSites | Where-Object {
                $_.GroupId -and $_.GroupId.ToString() -ne '00000000-0000-0000-0000-000000000000'
            })

        # Groups are matched to their site through the GroupId of Get-PnPTenantSite, never through
        # -IncludeSiteUrl: that switch makes PnP resolve the site of every group, and one group whose site
        # cannot be resolved (provisioning failed, renamed SharePoint domain) throws 'Object reference not set
        # to an instance of an object' for the whole call.
        $groupsById = @{}

        # Owners attached as an Owners property, the way -IncludeOwners does, when the group came without them
        $addGroupOwners = {
            param($Group, [string]$GroupId)
            try {
                $groupOwners = @(Get-PnPMicrosoft365GroupOwner -Identity $GroupId -Connection $adminConnection -ErrorAction Stop)
                $Group | Add-Member -NotePropertyName 'Owners' -NotePropertyValue $groupOwners -Force
            }
            catch {
                # $null Owners: read as 'unknown' downstream, never as 'no owner'
                $Group | Add-Member -NotePropertyName 'Owners' -NotePropertyValue $null -Force
                $Group | Add-Member -NotePropertyName 'OwnersUnreadable' -NotePropertyValue $true -Force
                Write-Verbose "Unable to read the owners of the Microsoft 365 group $($GroupId): $($_.Exception.Message)"
            }
        }

        $isPermissionError = {
            param([string]$Message)
            return $Message -match 'Forbidden|403|Authorization_RequestDenied|Insufficient privileges|Unauthorized|401'
        }

        # Below this many group-connected sites, one targeted lookup per site beats enumerating the whole
        # tenant: -IncludeOwners makes the tenant-wide call expensive on its own, and a tenant holds orders of
        # magnitude more groups than a narrow run needs.
        $targetedLookupThreshold = 25

        if ($groupConnectedSites.Count -eq 0) {
            Write-Host -ForegroundColor Cyan 'No group-connected site in scope: skipping the Microsoft 365 groups retrieval.'
        }
        else {
            $tenantWideError = $null

            if ($groupConnectedSites.Count -gt $targetedLookupThreshold) {
                Write-Host -ForegroundColor Cyan 'Retrieving Microsoft 365 groups...'

                # With owners first; if that fails, without owners (fetched below per group in scope)
                $allM365Groups = $null
                $ownersIncluded = $true
                try {
                    $allM365Groups = @(Get-PnPMicrosoft365Group -IncludeOwners -Connection $adminConnection -ErrorAction Stop)
                }
                catch {
                    Write-Verbose "Tenant-wide group read with owners failed, retrying without owners: $($_.Exception.Message)"
                    $ownersIncluded = $false
                    try {
                        $allM365Groups = @(Get-PnPMicrosoft365Group -Connection $adminConnection -ErrorAction Stop)
                    }
                    catch {
                        $tenantWideError = $_.Exception.Message.Trim()
                        Write-Verbose "Tenant-wide group read failed, falling back to one lookup per site: $tenantWideError"
                    }
                }

                if ($null -ne $allM365Groups) {
                    foreach ($m365Group in $allM365Groups) {
                        # Property name defensively checked (Id vs GroupId) - not guaranteed identical across PnP versions
                        $m365GroupId = if ($m365Group.Id) { $m365Group.Id.ToString() } elseif ($m365Group.GroupId) { $m365Group.GroupId.ToString() } else { $null }
                        if ($m365GroupId) {
                            $groupsById[$m365GroupId.ToLowerInvariant()] = $m365Group
                            if ($m365Group.DisplayName) {
                                $groupNameCache[$m365GroupId] = $m365Group.DisplayName
                            }
                        }
                    }

                    if (-not $ownersIncluded) {
                        Write-Host -ForegroundColor Cyan "Reading the owners of $($groupConnectedSites.Count) Microsoft 365 group(s) one by one..."
                        foreach ($groupConnectedSite in $groupConnectedSites) {
                            $siteGroupKey = $groupConnectedSite.GroupId.ToString().ToLowerInvariant()
                            if ($groupsById.ContainsKey($siteGroupKey)) {
                                & $addGroupOwners $groupsById[$siteGroupKey] $siteGroupKey
                            }
                        }
                    }
                }
            }

            # Targeted lookups: the normal path for a narrow run, and the fallback when the tenant-wide read failed
            if ($groupConnectedSites.Count -le $targetedLookupThreshold -or $tenantWideError) {
                Write-Host -ForegroundColor Cyan "Retrieving the Microsoft 365 group of $($groupConnectedSites.Count) group-connected site(s)..."
                $targetedErrors = 0

                foreach ($groupConnectedSite in $groupConnectedSites) {
                    $siteGroupId = $groupConnectedSite.GroupId.ToString()
                    $m365Group = $null

                    try {
                        $m365Group = Get-PnPMicrosoft365Group -Identity $siteGroupId -IncludeOwners -Connection $adminConnection -ErrorAction Stop
                    }
                    catch {
                        try {
                            $m365Group = Get-PnPMicrosoft365Group -Identity $siteGroupId -Connection $adminConnection -ErrorAction Stop
                            & $addGroupOwners $m365Group $siteGroupId
                        }
                        catch {
                            $targetedErrors++
                            $lastTargetedError = $_.Exception.Message.Trim()
                            Write-Verbose "Unable to retrieve the Microsoft 365 group $siteGroupId of $($groupConnectedSite.Url): $lastTargetedError"
                            continue
                        }
                    }

                    $groupsById[$siteGroupId.ToLowerInvariant()] = $m365Group
                    if ($m365Group.DisplayName) {
                        $groupNameCache[$siteGroupId] = $m365Group.DisplayName
                    }
                }

                if ($targetedErrors -gt 0) {
                    Write-Warning "The Microsoft 365 group of $targetedErrors site(s) could not be read (last error: $lastTargetedError). Run with -Verbose for each site."
                    if (& $isPermissionError $lastTargetedError) {
                        Write-Warning 'M365GroupsDetails requires the Microsoft Graph application permission Group.Read.All (or Group.ReadWrite.All) on the certificate used for the PnP connection, with admin consent.'
                    }
                }
            }

            # Site URL -> group, through the GroupId each site already carries
            foreach ($groupConnectedSite in $groupConnectedSites) {
                $siteGroupKey = $groupConnectedSite.GroupId.ToString().ToLowerInvariant()
                if ($groupsById.ContainsKey($siteGroupKey)) {
                    $m365GroupsLookup[$groupConnectedSite.Url] = $groupsById[$siteGroupKey]
                }
            }

            Write-Host -ForegroundColor Cyan "Microsoft 365 groups matched to their site: $($m365GroupsLookup.Count) of $($groupConnectedSites.Count)"

            if ($m365GroupsLookup.Count -eq 0) {
                # Nothing could be read: the report goes on, the group findings say 'not collected'
                $m365GroupsAvailable = $false
                Write-Warning 'No Microsoft 365 group could be read: the Microsoft 365 group columns stay empty and the group findings are reported as not collected.'
            }
        }
    }

    # The Active sites page of the SharePoint admin center and its CSV export read a hidden list of the admin
    # site, DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS: one item per site with what
    # Get-PnPTenantSite does not return (creation date and creator, last activity, file and page counts,
    # channel sites, how the site was created, information barriers mode). Read once with the admin
    # connection and indexed by URL; a failed read leaves those columns empty rather than aborting the report.
    # The main geo only: the list of a satellite geo lives on that geo's admin site.
    $activeSitesLookup = @{}
    $activeSitesFields = @('SiteUrl', 'TimeCreated', 'CreatedBy', 'LastActivityOn', 'NumOfFiles', 'FileViewedOrEdited', 'PageViews', 'PagesVisited', 'ChannelSitesCount', 'SiteCreationSource', 'IBMode', 'TimeDeleted')
    try {
        $activeSitesItems = Get-PnPListItem -List 'DO_NOT_DELETE_SPLIST_TENANTADMIN_AGGREGATED_SITECOLLECTIONS' -Fields $activeSitesFields -PageSize 5000 -Connection $adminConnection -ErrorAction Stop
        foreach ($activeSitesItem in $activeSitesItems) {
            $activeSitesValues = $activeSitesItem.FieldValues
            if ($activeSitesValues['TimeDeleted']) {
                continue
            }
            $activeSitesUrl = "$($activeSitesValues['SiteUrl'])".TrimEnd('/').ToLowerInvariant()
            if ($activeSitesUrl) {
                $activeSitesLookup[$activeSitesUrl] = $activeSitesValues
            }
        }
        Write-Host -ForegroundColor Cyan "Active sites list of the admin center: $($activeSitesLookup.Count) site(s)"
    }
    catch {
        Write-Warning "Unable to read the Active sites list of the SharePoint admin center: $($_.Exception.Message.Trim()). The SiteCreatedDate, CreatedBy, LastActivityDate, FilesCount, FilesViewedOrEdited, PageViews, PagesVisited, ChannelSitesCount, CreatedFrom and IBMode columns will be empty."
    }

    # How a site was created: SiteCreationSource is an id that the admin center names with
    # GetSPOSiteCreationSources (Microsoft Teams, SharePoint start page, API...), read like the Active sites page
    $creationSourceById = @{}
    try {
        $creationSources = Invoke-PnPSPRestMethod -Method Get -Url '/_api/SPO.Tenant/GetSPOSiteCreationSources' -Connection $adminConnection -ErrorAction Stop
        foreach ($creationSource in @($creationSources.value)) {
            if ($creationSource.Id) {
                $creationSourceById["$($creationSource.Id)".ToLowerInvariant()] = $creationSource.DisplayName
            }
        }
    }
    catch {
        Write-Verbose "Unable to read the site creation sources: $($_.Exception.Message)"
    }

    # Hub sites, by id: the HubSiteName column names the hub a site is associated with (HubSiteId)
    $hubNameById = @{}
    try {
        foreach ($hubSite in @(Get-PnPHubSite -Connection $adminConnection -ErrorAction Stop)) {
            foreach ($hubKey in @($hubSite.ID, $hubSite.SiteId)) {
                if ($hubKey -and "$hubKey" -ne [guid]::Empty.ToString()) {
                    $hubNameById["$hubKey".ToLowerInvariant()] = $hubSite.Title
                }
            }
        }
    }
    catch {
        Write-Warning "Unable to read the hub sites: $($_.Exception.Message.Trim()). The HubSiteName column will be empty."
    }

    # Site-level data (admins, Owners/Members/Visitors groups, regional settings) requires one PnP connection
    # per site. Collect it in parallel and index it by URL, so the main loop below just merges the results.
    # Claim resolution (federateddirectoryclaimprovider) is intentionally NOT done here: parallel runspaces
    # do not share mutable state with each other, so a cache of Graph lookups written concurrently by several
    # runspaces is not trivial to get right. It is deferred to the sequential merge loop below instead.
    $perSiteLookup = @{}
    $needPerSite = $IncludeSiteAdmins.IsPresent -or $IncludeSiteMembers.IsPresent -or $IncludeSiteVisitors.IsPresent -or $IncludeSharingLinks.IsPresent -or $RegionalSettingsDetails.IsPresent -or $IncludeEveryoneAccess.IsPresent

    if ($needPerSite) {
        Write-Host -ForegroundColor Cyan "Collecting site-level data with $ThrottleLimit concurrent connection(s)..."

        # Progress of the parallel pass: results stream out of ForEach-Object -Parallel as each site
        # completes, so counting them downstream gives a live counter. One console line every 10 % only:
        # Write-Progress is not used, it slows a long loop down noticeably
        $perSiteDone = 0
        $perSiteTotal = [math]::Max(1, $spoSites.Count)
        $perSiteStep = [math]::Max(1, [math]::Ceiling($perSiteTotal / 10))
        $perSiteStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        $perSiteResults = $spoSites | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            $auth = $using:pnpAuthParams
            $doAdmins = $using:IncludeSiteAdmins
            $doMembers = $using:IncludeSiteMembers
            $doVisitors = $using:IncludeSiteVisitors
            $doSharingLinks = $using:IncludeSharingLinks
            $doRegional = $using:RegionalSettingsDetails
            $doEveryoneAccess = $using:IncludeEveryoneAccess
            $site = $_

            $result = [ordered]@{
                Url                  = $site.Url
                AdminLoginNames      = $null
                AdminsCount          = 0
                SiteOwnerLoginNames  = $null
                SiteMemberLoginNames = $null
                SiteVisitorLoginNames = $null
                SharingLinksAnyoneCount     = $null
                SharingLinksCompanyCount    = $null
                SharingLinksFlexibleCount = $null
                SharingLinksOtherCount      = $null
                SharingLinksTotalCount      = $null
                TimeZoneId           = $null
                TimeZoneString       = $null
                HourFormat           = $null
                RegionalLocaleId     = $null
                EveryoneAccessGrants = $null
                ListsWithUniquePermissionsCount = $null
                Status                = 'OK'
            }

            $siteConnection = $null
            try {
                $siteConnection = Connect-PnPOnline -Url $site.Url @auth -ReturnConnection -ErrorAction Stop
            }
            catch {
                $result.Status = "ERROR: $($_.Exception.Message)"
                [PSCustomObject]$result
                return
            }

            # Each block below is isolated in its own try/catch: a .NET-level failure specific to one
            # data point (seen in practice on a couple of atypical sites: "Object reference not set to an
            # instance of an object" / "Nullable object must have a value" inside PnP's own CSOM/Graph
            # plumbing) must not wipe out the other data points already collected for the same site.
            # $result.Status only reflects a connection failure; per-block failures are tracked via
            # $result.PartialErrors and reported, but do not blank the whole site.
            $partialErrors = [System.Collections.Generic.List[string]]::new()

            if ($doAdmins) {
                try {
                    $siteAdmins = Get-PnPSiteCollectionAdmin -Connection $siteConnection -ErrorAction Stop
                    # Kept as an array, NOT joined with '|': a SharePoint claim already uses '|' as its own
                    # internal delimiter, so joining multiple LoginNames with '|' then splitting them back
                    # is ambiguous and corrupts every claim. See Resolve-SPOClaimLoginName below.
                    $result.AdminLoginNames = @($siteAdmins.LoginName)
                    $result.AdminsCount = @($siteAdmins).Count
                }
                catch {
                    $partialErrors.Add("Admins: $($_.Exception.Message)")
                }
            }

            if ($doMembers) {
                try {
                    $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup -Includes Users -Connection $siteConnection -ErrorAction Stop
                    $memberGroup = Get-PnPGroup -AssociatedMemberGroup -Includes Users -Connection $siteConnection -ErrorAction Stop
                    $result.SiteOwnerLoginNames = @($ownerGroup.Users.LoginName)
                    $result.SiteMemberLoginNames = @($memberGroup.Users.LoginName)
                }
                catch {
                    $partialErrors.Add("Members: $($_.Exception.Message)")
                }
            }

            if ($doVisitors) {
                # Some group-connected sites have no distinct Visitors group; treat that as empty, not an error
                try {
                    $visitorGroup = Get-PnPGroup -AssociatedVisitorGroup -Includes Users -Connection $siteConnection -ErrorAction Stop
                    $result.SiteVisitorLoginNames = @($visitorGroup.Users.LoginName)
                }
                catch {
                    $result.SiteVisitorLoginNames = $null
                }
            }

            if ($doSharingLinks) {
                # SharePoint materialises every sharing link as a hidden site group named
                # 'SharingLinks.<itemGuid>.<linkType>.<linkGuid>', so counting those groups gives a per-site
                # picture of the sharing surface for the cost of a single call, instead of one Graph call per
                # file. The link type is read from the group name: AnonymousEdit/AnonymousView (Anyone links),
                # OrganizationEdit/OrganizationView (people in the organization) and Flexible (specific
                # people, and the newer link types). This naming is an implementation detail of SharePoint,
                # not a documented contract: anything that does not match a known type is counted as Other
                # rather than silently dropped, so a future rename shows up as a bucket that fills instead of
                # counts that quietly go to zero. These counts say how many links exist and of which kind,
                # nothing more: expiration dates, password protection and the link URLs themselves require
                # the per-item Graph calls of Get-SPOSharingLinkReport.
                try {
                    $siteGroups = Get-PnPGroup -Connection $siteConnection -ErrorAction Stop
                    $sharingLinkGroups = @($siteGroups | Where-Object { $_.Title -like 'SharingLinks.*' })

                    $anyoneLinks = 0
                    $companyLinks = 0
                    $flexibleLinks = 0
                    $otherLinks = 0

                    foreach ($sharingLinkGroup in $sharingLinkGroups) {
                        $groupTitle = $sharingLinkGroup.Title
                        if ($groupTitle -match '\.Anonymous[^.]*\.') { $anyoneLinks++ }
                        elseif ($groupTitle -match '\.Organization[^.]*\.') { $companyLinks++ }
                        elseif ($groupTitle -match '\.Flexible\.') { $flexibleLinks++ }
                        else { $otherLinks++ }
                    }

                    $result.SharingLinksAnyoneCount = $anyoneLinks
                    $result.SharingLinksCompanyCount = $companyLinks
                    $result.SharingLinksFlexibleCount = $flexibleLinks
                    $result.SharingLinksOtherCount = $otherLinks
                    $result.SharingLinksTotalCount = $sharingLinkGroups.Count
                }
                catch {
                    $partialErrors.Add("SharingLinks: $($_.Exception.Message)")
                }
            }

            if ($doRegional) {
                try {
                    $web = Get-PnPWeb -Includes RegionalSettings, RegionalSettings.TimeZone -Connection $siteConnection -ErrorAction Stop
                    $regional = $web.RegionalSettings
                    if ($null -ne $regional) {
                        $result.TimeZoneId = $regional.TimeZone.Id
                        $result.TimeZoneString = $regional.TimeZone.Description
                        $result.HourFormat = if ($regional.Time24) { '24' } else { '12' }
                        $result.RegionalLocaleId = $regional.LocaleId
                    }
                }
                catch {
                    $partialErrors.Add("Regional: $($_.Exception.Message)")
                }
            }

            if ($doEveryoneAccess) {
                # Where the tenant-wide claims ('Everyone', 'Everyone except external users') are granted:
                # on the site itself (directly or through ANY SharePoint group holding permissions, not only
                # the three associated ones) and on every library/list whose inheritance is broken. One REST
                # call for the site, one for the list inventory, one per list with unique permissions.
                # Folders and files are not inspected: SharePoint cannot filter items on unique permissions,
                # so it would mean walking every item of every library.
                try {
                    $everyoneGrants = [System.Collections.Generic.List[string]]::new()

                    $getRestRows = {
                        param($Response)
                        if ($null -ne $Response.value) { return @($Response.value) }
                        if ($null -ne $Response.d.results) { return @($Response.d.results) }
                        return @()
                    }

                    $addEveryoneGrants = {
                        param($RoleAssignments, [string]$ScopeLabel)

                        foreach ($roleAssignment in $RoleAssignments) {
                            # 'Limited Access' is granted automatically by SharePoint when something below
                            # is shared: it opens nothing by itself, so it is not reported
                            $roleNames = @($roleAssignment.RoleDefinitionBindings | Where-Object { $_.Name -notmatch 'Limited Access' } | ForEach-Object { $_.Name })
                            if ($roleNames.Count -eq 0) {
                                continue
                            }

                            $member = $roleAssignment.Member
                            # Sharing link groups are counted by IncludeSharingLinks, not here
                            if ("$($member.LoginName)" -like 'SharingLinks.*' -or "$($member.Title)" -like 'SharingLinks.*') {
                                continue
                            }

                            $candidates = [System.Collections.Generic.List[PSCustomObject]]::new()
                            $candidates.Add([PSCustomObject]@{ LoginName = "$($member.LoginName)"; ViaGroup = $null })
                            # Users of a SharePoint group: a plain array in the nometadata JSON format, wrapped in
                            # .results in the verbose one. Tested on the property itself, not with '$null -ne',
                            # which is true for any array (member enumeration of a missing property included)
                            $groupUsers = @($member.Users)
                            if ($null -ne $member.Users -and $member.Users.PSObject.Properties.Name -contains 'results') {
                                $groupUsers = @($member.Users.results)
                            }
                            foreach ($groupUser in $groupUsers) {
                                if ($groupUser.LoginName) {
                                    $candidates.Add([PSCustomObject]@{ LoginName = "$($groupUser.LoginName)"; ViaGroup = "$($member.Title)" })
                                }
                            }

                            foreach ($candidate in $candidates) {
                                $principalLabel = if ($candidate.LoginName -eq 'c:0(.s|true') { 'Everyone' }
                                elseif ($candidate.LoginName -like '*spo-grid-all-users*') { 'Everyone except external users' }
                                else { $null }

                                if ($principalLabel) {
                                    $grantDetail = $roleNames -join ', '
                                    if ($candidate.ViaGroup) {
                                        $grantDetail += ", via group '$($candidate.ViaGroup)'"
                                    }
                                    $everyoneGrants.Add("$principalLabel on $ScopeLabel ($grantDetail)")
                                }
                            }
                        }
                    }

                    $roleAssignmentsQuery = 'roleassignments?$expand=Member/Users,RoleDefinitionBindings'
                    $webRoleAssignments = Invoke-PnPSPRestMethod -Url "/_api/web/$roleAssignmentsQuery" -Connection $siteConnection -ErrorAction Stop
                    & $addEveryoneGrants (& $getRestRows $webRoleAssignments) 'site'

                    $listsResponse = Invoke-PnPSPRestMethod -Url '/_api/web/lists?$select=Id,Title,Hidden,BaseType,HasUniqueRoleAssignments&$top=5000' -Connection $siteConnection -ErrorAction Stop
                    # Hidden lists are system plumbing (catalogs, internal lists): not where content is shared
                    $uniqueLists = @(& $getRestRows $listsResponse | Where-Object { $_.HasUniqueRoleAssignments -and -not $_.Hidden })
                    $result.ListsWithUniquePermissionsCount = $uniqueLists.Count

                    foreach ($uniqueList in $uniqueLists) {
                        # BaseType 1 = document library, anything else is reported as a list
                        $listKind = if ("$($uniqueList.BaseType)" -in @('1', 'DocumentLibrary')) { 'library' } else { 'list' }
                        $listRoleAssignments = Invoke-PnPSPRestMethod -Url "/_api/web/lists(guid'$($uniqueList.Id)')/$roleAssignmentsQuery" -Connection $siteConnection -ErrorAction Stop
                        & $addEveryoneGrants (& $getRestRows $listRoleAssignments) "$listKind '$($uniqueList.Title)'"
                    }

                    $result.EveryoneAccessGrants = @($everyoneGrants | Select-Object -Unique)
                }
                catch {
                    $partialErrors.Add("EveryoneAccess: $($_.Exception.Message)")
                }
            }

            if ($partialErrors.Count -gt 0) {
                $result.Status = "PARTIAL: $($partialErrors -join ' | ')"
            }

            # Disconnect-PnPOnline cannot target a specific connection; releasing the variable disposes
            # it. This matters because the parallel runspace pool is reused across sites, so leaving
            # connection objects alive would accumulate them inside each runspace.
            $siteConnection = $null

            [PSCustomObject]$result
        } | ForEach-Object {
            $perSiteDone++
            if ($perSiteDone % $perSiteStep -eq 0 -and $perSiteDone -lt $perSiteTotal) {
                $remaining = [timespan]::FromSeconds(($perSiteStopwatch.Elapsed.TotalSeconds / $perSiteDone) * ($perSiteTotal - $perSiteDone))
                Write-Host -ForegroundColor Cyan "  Site-level data: $perSiteDone / $perSiteTotal site(s) ($([int]($perSiteDone / $perSiteTotal * 100)) %), about $($remaining.ToString('hh\:mm\:ss')) remaining"
            }
            $_
        }

        Write-Host -ForegroundColor Cyan "Site-level data collected for $perSiteDone site(s) in $($perSiteStopwatch.Elapsed.ToString('hh\:mm\:ss'))."

        foreach ($perSite in $perSiteResults) {
            $perSiteLookup[$perSite.Url] = $perSite
        }
    }

    $spoSitesInfosArray = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Access denied / claim resolution counters, reported once at the end instead of flooding the console
    $siteErrorCount = 0
    # Failing site URLs, listed in the final summary so identifying them does not require -Verbose
    $siteErrorUrls = [System.Collections.Generic.List[string]]::new()
    $sitePartialErrorCount = 0
    $claimResolutionErrorCount = 0

    # Sequential claim resolution caches: one Graph lookup per distinct group GUID for the whole run.
    # A plain hashtable is safe here because this loop is not parallelized.
    $groupClaimOwnersCache = @{}
    $groupClaimMembersCache = @{}

    # Progress of the merge pass: it resolves the group claims and reads the Microsoft 365 group and Teams
    # layer through Graph, site after site, which can take a while on a large tenant. One line every 10 %
    $mergeIndex = 0
    $mergeTotal = [math]::Max(1, $spoSites.Count)
    $mergeStep = [math]::Max(1, [math]::Ceiling($mergeTotal / 10))
    $mergeStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host -ForegroundColor Cyan 'Resolving rights and Microsoft 365 groups...'

    foreach ($spoSite in $spoSites) {
        Write-Verbose "Get details for SharePoint site $($spoSite.Url)"

        $mergeIndex++
        if ($mergeIndex % $mergeStep -eq 0 -and $mergeIndex -lt $mergeTotal) {
            $remaining = [timespan]::FromSeconds(($mergeStopwatch.Elapsed.TotalSeconds / $mergeIndex) * ($mergeTotal - $mergeIndex))
            Write-Host -ForegroundColor Cyan "  Rights and groups: $mergeIndex / $mergeTotal site(s) ($([int]($mergeIndex / $mergeTotal * 100)) %), about $($remaining.ToString('hh\:mm\:ss')) remaining"
        }

        # Init variables for each site to avoid leaking values from the previous iteration
        $siteType = $null
        $team = $channelCount = $teamOwners = $teamMembersCount = $teamGuestsCount = $null
        $m365GroupOwnersResolved = $m365GroupMembersCount = $m365GroupGuestsCount = $null
        $primarySmtpAddress = $whenCreatedUTC = $m365GroupVisibility = $null
        $siteAdmins = $siteAdminsRaw = $siteAdminsResolved = $null
        $siteOwnersLoginNames = $siteOwnersRaw = $siteOwnersResolved = $null
        $siteMembersLoginNames = $siteMembersRaw = $siteMembersResolved = $null
        $siteVisitorsLoginNames = $siteVisitorsRaw = $siteVisitorsResolved = $null
        $sharingLinksAnyoneCount = $sharingLinksCompanyCount = $sharingLinksFlexibleCount = $null
        $sharingLinksOtherCount = $sharingLinksTotalCount = $null
        $timezoneId = $timezoneString = $hourFormat = $localeId = $localeIdString = $null
        $everyoneAccess = $listsWithUniquePermissionsCount = $null

        # Get-PnPTenantSite exposes GroupId as a System.Guid (not a wrapper with a .Guid property)
        $groupId = if ($spoSite.GroupId) { $spoSite.GroupId.ToString() } else { $null }

        # Merge the parallel site-level data collected above, then resolve claims sequentially
        $perSite = $perSiteLookup[$spoSite.Url]

        # Carried into the report row: empty site-level columns must be distinguishable from a collection failure
        $siteDataStatus = 'NotCollected'

        if ($null -ne $perSite) {
            $siteDataStatus = 'OK'

            if ($perSite.Status -like 'ERROR:*') {
                # Connection to the site itself failed: no site-level data at all for this site
                $siteErrorCount++
                $siteErrorUrls.Add($spoSite.Url)
                $siteDataStatus = "$($perSite.Status)"
                Write-Verbose "Site-level collection failed for $($spoSite.Url): $($perSite.Status)"
            }
            elseif ($perSite.Status -like 'PARTIAL:*') {
                # Connection succeeded but one or more data points failed; the others are still usable
                $sitePartialErrorCount++
                $siteDataStatus = "$($perSite.Status)"
                Write-Verbose "Site-level collection partially failed for $($spoSite.Url): $($perSite.Status)"
            }

            # Kept as arrays here (source of truth for claim resolution below); joined to a '|' string
            # only for the raw display columns, never re-split - see Resolve-SPOClaimLoginName for why.
            $siteAdmins = $perSite.AdminLoginNames
            $siteOwnersLoginNames = $perSite.SiteOwnerLoginNames
            $siteMembersLoginNames = $perSite.SiteMemberLoginNames
            $siteVisitorsLoginNames = $perSite.SiteVisitorLoginNames
            $siteAdminsRaw = if ($siteAdmins) { $siteAdmins -join '|' } else { $null }
            $siteOwnersRaw = if ($siteOwnersLoginNames) { $siteOwnersLoginNames -join '|' } else { $null }
            $siteMembersRaw = if ($siteMembersLoginNames) { $siteMembersLoginNames -join '|' } else { $null }
            $siteVisitorsRaw = if ($siteVisitorsLoginNames) { $siteVisitorsLoginNames -join '|' } else { $null }
            $sharingLinksAnyoneCount = $perSite.SharingLinksAnyoneCount
            $sharingLinksCompanyCount = $perSite.SharingLinksCompanyCount
            $sharingLinksFlexibleCount = $perSite.SharingLinksFlexibleCount
            $sharingLinksOtherCount = $perSite.SharingLinksOtherCount
            $sharingLinksTotalCount = $perSite.SharingLinksTotalCount
            $timezoneId = $perSite.TimeZoneId
            $hourFormat = $perSite.HourFormat
            # Empty string (not $null) when collected and nothing found: keeps "collected, none" apart from
            # "could not be read" for the findings summary
            $everyoneAccess = if ($null -ne $perSite.EveryoneAccessGrants) { @($perSite.EveryoneAccessGrants) -join '|' } else { $null }
            $listsWithUniquePermissionsCount = $perSite.ListsWithUniquePermissionsCount
            $localeId = $perSite.RegionalLocaleId

            if ($perSite.TimeZoneString) {
                $timezoneString = $perSite.TimeZoneString
            }
            elseif ($null -ne $perSite.TimeZoneId) {
                $timezoneString = Convert-SPOTimezoneToString -Id $perSite.TimeZoneId
            }

            if ($localeId) {
                $lang = [System.Globalization.CultureInfo][int]$localeId
                $localeIdString = "$($lang.Name)|$($lang.DisplayName)"
            }

            $resolveParams = @{
                Connection    = $adminConnection
                OwnersCache   = $groupClaimOwnersCache
                MembersCache  = $groupClaimMembersCache
                GroupNameCache = $groupNameCache
            }

            if ($siteAdmins) {
                $siteAdminsResolved = Resolve-SPOClaimLoginName -LoginNames $siteAdmins @resolveParams
            }
            if ($siteOwnersLoginNames) {
                $siteOwnersResolved = Resolve-SPOClaimLoginName -LoginNames $siteOwnersLoginNames @resolveParams
            }
            if ($siteMembersLoginNames) {
                $siteMembersResolved = Resolve-SPOClaimLoginName -LoginNames $siteMembersLoginNames @resolveParams
            }
            if ($siteVisitorsLoginNames) {
                $siteVisitorsResolved = Resolve-SPOClaimLoginName -LoginNames $siteVisitorsLoginNames @resolveParams
            }

            foreach ($resolvedValue in @($siteAdminsResolved, $siteOwnersResolved, $siteMembersResolved, $siteVisitorsResolved)) {
                if ($resolvedValue -match '<unresolved:') {
                    $claimResolutionErrorCount += ([regex]::Matches($resolvedValue, '<unresolved:')).Count
                }
            }
        }

        if ($M365GroupsDetails.IsPresent) {
            if ([string]::IsNullOrEmpty($groupId) -or $groupId -eq '00000000-0000-0000-0000-000000000000') {
                $siteType = 'SharePoint Site'
            }
            elseif ($spoSite.Template -eq 'GROUP#0') {
                $m365Group = $m365GroupsLookup[$spoSite.Url]

                if ($null -ne $m365Group) {
                    $siteType = 'M365 Group'
                    $primarySmtpAddress = $m365Group.Mail
                    # 'Public', 'Private' or 'HiddenMembership': already part of the group object, no extra call
                    $m365GroupVisibility = "$($m365Group.Visibility)"
                    # Property name for the creation date is not guaranteed identical across PnP versions
                    $whenCreatedUTC = if ($m365Group.CreatedDateTime) { $m365Group.CreatedDateTime } else { $m365Group.WhenCreated }

                    if ($m365Group.HasTeam) {
                        try {
                            $team = Get-PnPTeamsTeam -Identity $groupId -Connection $adminConnection -ErrorAction Stop
                        }
                        catch {
                            Write-Warning "Unable to retrieve the Teams metadata for $($spoSite.Url): $_"
                        }

                        try {
                            $channels = Get-PnPTeamsChannel -Team $groupId -Connection $adminConnection -ErrorAction Stop
                            $channelCount = @($channels).Count

                            # Get-PnPTeamsUser -Role only accepts a single scalar value, hence three calls
                            $teamOwnersUsers = Get-PnPTeamsUser -Team $groupId -Role Owner -Connection $adminConnection -ErrorAction Stop
                            $teamMembersUsers = Get-PnPTeamsUser -Team $groupId -Role Member -Connection $adminConnection -ErrorAction Stop
                            $teamGuestsUsers = Get-PnPTeamsUser -Team $groupId -Role Guest -Connection $adminConnection -ErrorAction Stop

                            $teamOwners = (@($teamOwnersUsers) | ForEach-Object { Get-SPOUserIdentifier -UserObject $_ } | Where-Object { $_ }) -join '|'
                            $teamMembersCount = @($teamMembersUsers).Count
                            $teamGuestsCount = @($teamGuestsUsers).Count
                        }
                        catch {
                            Write-Warning "Unable to retrieve the Teams membership for $($spoSite.Url): $_"
                        }

                        $m365GroupOwnersResolved = $teamOwners
                        $m365GroupMembersCount = $teamMembersCount
                        $m365GroupGuestsCount = $teamGuestsCount
                    }
                    else {
                        # No Team: Owners already fetched via -IncludeOwners (no extra Graph call).
                        # Get-PnPMicrosoft365GroupMember has no -UserType parameter (confirmed against
                        # PnP.PowerShell 3.2.0 - a parameter cannot be found error), so members and guests
                        # are split client-side on the '#EXT#' UPN pattern, the module's own documented
                        # guest-detection convention (see CLAUDE.md "User Filtering Patterns").
                        # Owners that could not be read stay $null ('unknown'), never '' ('no owner'): an
                        # empty string would flag the group in EID-GROUP-13 on a read failure
                        $m365GroupOwnersResolved = if ($m365Group.PSObject.Properties.Name -contains 'OwnersUnreadable') {
                            $null
                        }
                        else {
                            (@($m365Group.Owners) | ForEach-Object { Get-SPOUserIdentifier -UserObject $_ } | Where-Object { $_ }) -join '|'
                        }

                        try {
                            $allGroupMembers = Get-PnPMicrosoft365GroupMember -Identity $groupId -Connection $adminConnection -ErrorAction Stop
                            $memberIdentifiers = @($allGroupMembers | ForEach-Object { Get-SPOUserIdentifier -UserObject $_ } | Where-Object { $_ })
                            $guestIdentifiers = @($memberIdentifiers | Where-Object { $_ -match '#EXT#' })

                            $m365GroupGuestsCount = $guestIdentifiers.Count
                            $m365GroupMembersCount = $memberIdentifiers.Count - $guestIdentifiers.Count
                        }
                        catch {
                            Write-Warning "Unable to retrieve the Microsoft 365 group membership for $($spoSite.Url): $_"
                        }
                    }
                }
                else {
                    $siteType = 'M365 group but not connected (?)'
                }
            }
        }

        # Get-PnPTenantSite exposes StorageUsage (MB) instead of Get-SPOSite's StorageUsageCurrent
        $storageUsed = if ($null -ne $spoSite.StorageUsageCurrent) { $spoSite.StorageUsageCurrent } else { $spoSite.StorageUsage }

        # The properties are accumulated in an ordered dictionary rather than declared in a single literal:
        # the columns coming from an optional switch are added only when that switch was used (an absent
        # column reads as "not collected", where an empty one would wrongly read as "collected, and there is
        # nobody/nothing"), while still landing at their historical position in the column order.
        $properties = [ordered]@{
            SPTitle        = $spoSite.Title
            GroupID        = $groupId
            Url            = $spoSite.Url
            Geo            = $spoSite.Geo
            # StorageQuota/StorageUsage are returned in MB by SPO/PnP: convert to GB.
            # Left empty (not 0, which would read as real data) when the tenant-level read failed
            StorageLimitGB = if ($spoSite.TenantDataError) { $null } else { [math]::Round(($spoSite.StorageQuota) / 1024, 2) }
            StorageUsedGB  = if ($spoSite.TenantDataError) { $null } else { [math]::Round(($storageUsed) / 1024, 2) }
            Owner          = $spoSite.Owner
            # 'OK', 'ERROR: ...' (no site-level data at all), 'PARTIAL: ...' (some data points missing)
            # or 'NotCollected': tells whether the empty site-level columns of this row mean
            # 'nothing there' or 'could not read'. When the tenant-level read failed too (single-site
            # fallback), its error is prepended so the report carries the full picture on its own
            SiteDataStatus = if ($spoSite.TenantDataError) { "TenantError: $($spoSite.TenantDataError) | SiteLevel: $siteDataStatus" } else { $siteDataStatus }
        }

        if ($IncludeSiteAdmins.IsPresent) {
            $properties['SiteAdmins'] = $siteAdminsRaw
            $properties['SiteAdminsResolved'] = $siteAdminsResolved
        }

        if ($IncludeSiteMembers.IsPresent) {
            $properties['SiteOwnersLoginNames'] = $siteOwnersRaw
            $properties['SiteOwnersResolved'] = $siteOwnersResolved
            $properties['SiteMembersLoginNames'] = $siteMembersRaw
            $properties['SiteMembersResolved'] = $siteMembersResolved
        }

        if ($IncludeSiteVisitors.IsPresent) {
            $properties['SiteVisitorsLoginNames'] = $siteVisitorsRaw
            $properties['SiteVisitorsResolved'] = $siteVisitorsResolved
        }

        if ($IncludeEveryoneAccess.IsPresent) {
            $properties['EveryoneAccess'] = $everyoneAccess
            $properties['ListsWithUniquePermissionsCount'] = $listsWithUniquePermissionsCount
        }

        $properties['SharingCapability'] = $spoSite.SharingCapability
        $properties['SharingAllowedDomain'] = $spoSite.SharingAllowedDomainList -join '|'
        $properties['SharingBlockedDomain'] = $spoSite.SharingBlockedDomainList -join '|'
        $properties['SiteDefinedSharingCapability'] = $spoSite.SiteDefinedSharingCapability

        if ($IncludeSharingLinks.IsPresent) {
            $properties['SharingLinksAnyoneCount'] = $sharingLinksAnyoneCount
            $properties['SharingLinksCompanyCount'] = $sharingLinksCompanyCount
            $properties['SharingLinksFlexibleCount'] = $sharingLinksFlexibleCount
            $properties['SharingLinksOtherCount'] = $sharingLinksOtherCount
            $properties['SharingLinksTotalCount'] = $sharingLinksTotalCount
        }

        $properties['LockState'] = $spoSite.LockState

        if ($RegionalSettingsDetails.IsPresent) {
            $properties['LocaleID'] = $localeId
            $properties['LocaleIDString'] = $localeIdString
            $properties['Timezone'] = $timezoneId
            $properties['TimezoneString'] = $timezoneString
            $properties['HourFormat'] = $hourFormat
        }

        $properties['Template'] = $spoSite.Template
        $properties['ConditionalAccessPolicy'] = $spoSite.ConditionalAccessPolicy
        $properties['LastContentModifiedDate'] = $spoSite.LastContentModifiedDate
        $properties['IsTeamsConnected'] = $spoSite.IsTeamsConnected
        $properties['IsTeamsChannelConnected'] = $spoSite.IsTeamsChannelConnected
        $properties['SensitivityLabel'] = $spoSite.SensitivityLabel
        $properties['DefaultLinkPermission'] = $spoSite.DefaultLinkPermission
        $properties['DefaultSharingLinkType'] = $spoSite.DefaultSharingLinkType
        $properties['DefaultLinkToExistingAccess'] = $spoSite.DefaultLinkToExistingAccess
        $properties['AnonymousLinkExpirationInDays'] = $spoSite.AnonymousLinkExpirationInDays
        $properties['OverrideTenantAnonymousLinkExpirationPolicy'] = $spoSite.OverrideTenantAnonymousLinkExpirationPolicy
        $properties['ExternalUserExpirationInDays'] = $spoSite.ExternalUserExpirationInDays
        $properties['OverrideTenantExternalUserExpirationPolicy'] = $spoSite.OverrideTenantExternalUserExpirationPolicy
        $properties['IsHubSite'] = $spoSite.IsHubSite

        # The columns of the admin center's Active sites export (his ask of 2026-11-27), from the list read
        # above and Get-PnPTenantSite. Empty when the list was not read or does not hold the site.
        $activeSite = $activeSitesLookup["$($spoSite.Url)".TrimEnd('/').ToLowerInvariant()]
        $hubSiteKey = if ($spoSite.HubSiteId -and "$($spoSite.HubSiteId)" -ne [guid]::Empty.ToString()) { "$($spoSite.HubSiteId)".ToLowerInvariant() } else { $null }
        $properties['HubSiteName'] = if ($hubSiteKey -and -not $spoSite.IsHubSite) { $hubNameById[$hubSiteKey] } else { $null }
        $properties['StorageUsedPercent'] = if (-not $spoSite.TenantDataError -and $spoSite.StorageQuota -gt 0) { [math]::Round(($storageUsed / $spoSite.StorageQuota) * 100, 2) } else { $null }
        $properties['SiteCreatedDate'] = if ($activeSite) { $activeSite['TimeCreated'] } else { $null }
        $properties['CreatedBy'] = if ($activeSite) { $activeSite['CreatedBy'] } else { $null }
        $creationSourceId = if ($activeSite -and $activeSite['SiteCreationSource']) { "$($activeSite['SiteCreationSource'])".Trim('{', '}').ToLowerInvariant() } else { $null }
        $properties['CreatedFrom'] = if ($creationSourceId) { if ($creationSourceById[$creationSourceId]) { $creationSourceById[$creationSourceId] } else { $creationSourceId } } else { $null }
        $properties['LastActivityDate'] = if ($activeSite) { $activeSite['LastActivityOn'] } else { $null }
        $properties['FilesCount'] = if ($activeSite) { $activeSite['NumOfFiles'] } else { $null }
        $properties['FilesViewedOrEdited'] = if ($activeSite) { $activeSite['FileViewedOrEdited'] } else { $null }
        $properties['PageViews'] = if ($activeSite) { $activeSite['PageViews'] } else { $null }
        $properties['PagesVisited'] = if ($activeSite) { $activeSite['PagesVisited'] } else { $null }
        $properties['ChannelSitesCount'] = if ($activeSite) { $activeSite['ChannelSitesCount'] } else { $null }
        $properties['IBMode'] = if ($activeSite) { $activeSite['IBMode'] } else { $null }
        # DenyAddAndCustomizePages Enabled means custom scripts are blocked (the export's IsCustomScriptsEnabled)
        $properties['CustomScripts'] = switch ("$($spoSite.DenyAddAndCustomizePages)") {
            'Enabled' { 'Blocked' }
            'Disabled' { 'Allowed' }
            default { $null }
        }

        # Version history: the site's own settings, the tenant default, and what really applies. A site that
        # inherits the tenant policy follows the tenant value, otherwise its own values win. Empty (not
        # 'Automatic') when the tenant-level read failed, so a missing value never reads as a real one
        $siteVersionPolicyLabel = Get-SPOVersionPolicyLabel -AutoTrim $spoSite.EnableAutoExpirationVersionTrim -MajorVersionLimit $spoSite.MajorVersionLimit -ExpireVersionsAfterDays $spoSite.ExpireVersionsAfterDays
        $versionPolicyEffective = if ($spoSite.TenantDataError) {
            $null
        }
        elseif ($spoSite.InheritVersionPolicyFromTenant -and $tenantVersionPolicyLabel) {
            "Tenant: $tenantVersionPolicyLabel"
        }
        elseif ($siteVersionPolicyLabel) {
            "Site: $siteVersionPolicyLabel"
        }
        else {
            $null
        }

        $properties['VersionPolicyEffective'] = $versionPolicyEffective
        $properties['VersionPolicyInheritsTenant'] = $spoSite.InheritVersionPolicyFromTenant
        $properties['VersionPolicyTenantDefault'] = $tenantVersionPolicyLabel
        $properties['VersionPolicySite'] = $siteVersionPolicyLabel
        $properties['VersionCount'] = $spoSite.VersionCount
        # VersionSize is returned in bytes by SPO/PnP
        $properties['VersionSizeGB'] = if ($null -ne $spoSite.VersionSize) { [math]::Round(($spoSite.VersionSize) / 1GB, 2) } else { $null }

        # If the team has been renamed, its DisplayName differs from the SharePoint site Title
        if ($M365GroupsDetails.IsPresent) {
            $properties['PrimarySmtpAddress'] = $primarySmtpAddress
            $properties['Type'] = $siteType
            $properties['M365WhenCreatedUTC'] = $whenCreatedUTC
            $properties['M365GroupVisibility'] = $m365GroupVisibility
            $properties['M365GroupOwners'] = $m365GroupOwnersResolved
            $properties['M365GroupMembersCount'] = $m365GroupMembersCount
            $properties['M365GroupGuestsCount'] = $m365GroupGuestsCount
            $properties['TeamDisplayName'] = $team.DisplayName
            $properties['TeamDescription'] = $team.Description
            $properties['TeamVisibility'] = $team.Visibility
            $properties['TeamIsArchived'] = $team.IsArchived
            $properties['TeamWebUrl'] = $team.WebUrl
            $properties['TeamChannelCount'] = $channelCount
        }

        $spoSitesInfosArray.Add([PSCustomObject]$properties)
    }

    # Site admins activity: every distinct admin account is looked up once for the whole run (cached) through
    # Microsoft Graph, over the PnP app-only connection already open - no Microsoft Graph PowerShell session.
    # Done after the merge loop because it reads the resolved admins (group claims expanded to their users).
    if ($IncludeSiteAdminsActivity.IsPresent) {
        Write-Host -ForegroundColor Cyan "Checking the activity of the site collection administrators (inactive = no sign-in for $InactiveDaysThreshold days)..."

        $adminActivityCache = @{}
        $signInActivityAvailable = $true
        # Fallback source, downloaded only if signInActivity turns out to be unreadable
        $usageReportState = @{ Status = 'NotLoaded'; LastActivityByUpn = $null }
        $activityLookupErrorCount = 0
        $inactiveSince = (Get-Date).ToUniversalTime().AddDays(-$InactiveDaysThreshold)
        $activityIndex = 0
        $activityTotal = [math]::Max(1, $spoSitesInfosArray.Count)
        $activityStep = [math]::Max(1, [math]::Ceiling($activityTotal / 10))

        foreach ($siteInfo in $spoSitesInfosArray) {
            $activityIndex++
            if ($activityIndex % $activityStep -eq 0 -and $activityIndex -lt $activityTotal) {
                Write-Host -ForegroundColor Cyan "  Admin activity: $activityIndex / $activityTotal site(s) ($([int]($activityIndex / $activityTotal * 100)) %), $($adminActivityCache.Count) distinct admin(s) checked"
            }

            $activeCount = $null
            $notActiveList = [System.Collections.Generic.List[string]]::new()
            $siteLookupErrors = 0

            # Only accounts with a UPN can be looked up: well-known claims ('Everyone...'), unresolved group
            # claims and app principals (<appId>@<tenantId>) are not users and are left out
            $adminIdentities = @(
                "$($siteInfo.SiteAdminsResolved)" -split '\|' |
                    ForEach-Object { $_ -replace " \(via M365 group '.*' (Owners|Members)\)$", '' } |
                    Where-Object { $_ -match '@' -and $_ -notmatch '^[0-9a-fA-F-]{36}@[0-9a-fA-F-]{36}$' } |
                    Select-Object -Unique
            )

            if ($siteInfo.Template -like 'SPSPERS*') {
                # OneDrive admins are not reliably readable (see IncludeSiteAdmins): no verdict on them
                $activityStatus = 'NotApplicable (OneDrive)'
            }
            elseif ([string]::IsNullOrEmpty($siteInfo.SiteAdminsResolved)) {
                $activityStatus = 'NoAdminCollected'
            }
            elseif ($adminIdentities.Count -eq 0) {
                $activityStatus = 'NoUserAdmin'
            }
            else {
                $activeCount = 0

                foreach ($adminIdentity in $adminIdentities) {
                    if (-not $adminActivityCache.ContainsKey($adminIdentity)) {
                        $accountState = $null
                        $selectSets = if ($signInActivityAvailable) { @('accountEnabled,signInActivity', 'accountEnabled') } else { @('accountEnabled') }

                        foreach ($selectSet in $selectSets) {
                            try {
                                $graphUser = Invoke-SPOGraphWithRetry -Url "v1.0/users/$([System.Uri]::EscapeDataString($adminIdentity))?`$select=$selectSet" -Connection $adminConnection

                                if ($selectSet -eq 'accountEnabled' -and $signInActivityAvailable) {
                                    # signInActivity refused but the user itself is readable: AuditLog.Read.All is
                                    # missing or the tenant has no Entra ID P1. Switch the rest of the run to the
                                    # enabled/disabled check only, and say so once
                                    $signInActivityAvailable = $false
                                    Write-Warning 'The sign-in activity of the site admins cannot be read (Microsoft Graph application permission AuditLog.Read.All missing, or no Entra ID P1 license): falling back to the last activity dates of the Microsoft 365 usage report (Reports.Read.All).'
                                }

                                $lastSignIn = $null
                                if ($graphUser.signInActivity) {
                                    foreach ($dateProperty in 'lastSuccessfulSignInDateTime', 'lastSignInDateTime', 'lastNonInteractiveSignInDateTime') {
                                        $rawDate = $graphUser.signInActivity.$dateProperty
                                        if (-not $rawDate) {
                                            continue
                                        }
                                        $parsedDate = if ($rawDate -is [datetime]) { $rawDate.ToUniversalTime() } else { [datetime]::Parse("$rawDate", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]'AdjustToUniversal, AssumeUniversal') }
                                        if ($null -eq $lastSignIn -or $parsedDate -gt $lastSignIn) {
                                            $lastSignIn = $parsedDate
                                        }
                                    }
                                }

                                $accountState = if ($graphUser.accountEnabled -eq $false) {
                                    [PSCustomObject]@{ State = 'Disabled'; Label = 'Disabled' }
                                }
                                elseif ($selectSet -eq 'accountEnabled') {
                                    Get-SPOUsageActivityState -UserPrincipalName $adminIdentity -Connection $adminConnection -ReportState $usageReportState -InactiveSince $inactiveSince -InactiveDaysThreshold $InactiveDaysThreshold
                                }
                                elseif ($null -eq $lastSignIn) {
                                    [PSCustomObject]@{ State = 'Inactive'; Label = 'Never signed in' }
                                }
                                elseif ($lastSignIn -lt $inactiveSince) {
                                    [PSCustomObject]@{ State = 'Inactive'; Label = "Last sign-in $($lastSignIn.ToString('yyyy-MM-dd'))" }
                                }
                                else {
                                    [PSCustomObject]@{ State = 'Active'; Label = "Last sign-in $($lastSignIn.ToString('yyyy-MM-dd'))" }
                                }
                                break
                            }
                            catch {
                                $errorMessage = $_.Exception.Message
                                if ($errorMessage -match 'Request_ResourceNotFound|ResourceNotFound|NotFound|404|does not exist') {
                                    # Deleted or renamed account: it cannot administer anything anymore
                                    $accountState = [PSCustomObject]@{ State = 'Inactive'; Label = 'Account not found' }
                                    break
                                }
                                Write-Verbose "Activity lookup failed for $adminIdentity ($selectSet): $errorMessage"

                                # Only a permission/licence refusal justifies the fallback without signInActivity:
                                # a transient error (throttling, timeout) must not downgrade the rest of the run
                                if ($selectSet -ne 'accountEnabled' -and $errorMessage -notmatch 'Forbidden|403|Authorization|Premium|AuditLog|signInActivity') {
                                    break
                                }
                            }
                        }

                        if ($null -eq $accountState) {
                            $accountState = [PSCustomObject]@{ State = 'Unknown'; Label = 'Lookup failed' }
                        }
                        $adminActivityCache[$adminIdentity] = $accountState
                    }

                    $cachedState = $adminActivityCache[$adminIdentity]
                    switch ($cachedState.State) {
                        'Active' { $activeCount++ }
                        'Unknown' { $siteLookupErrors++; $notActiveList.Add("$adminIdentity ($($cachedState.Label))") }
                        default { $notActiveList.Add("$adminIdentity ($($cachedState.Label))") }
                    }
                }

                $activityStatus = if ($siteLookupErrors -gt 0) { "PARTIAL: $siteLookupErrors lookup(s) failed" } else { 'OK' }
                if ($siteLookupErrors -gt 0) {
                    $activityLookupErrorCount++
                }
            }

            $siteInfo | Add-Member -NotePropertyName 'SiteAdminsActiveCount' -NotePropertyValue $activeCount -Force
            $siteInfo | Add-Member -NotePropertyName 'SiteAdminsNotActive' -NotePropertyValue ($(if ($notActiveList.Count -gt 0) { $notActiveList -join '|' } else { $null })) -Force
            $siteInfo | Add-Member -NotePropertyName 'SiteAdminsActivityStatus' -NotePropertyValue $activityStatus -Force
        }

        Write-Host -ForegroundColor Cyan "Site admins checked: $($adminActivityCache.Count) distinct account(s)."

        if ($activityLookupErrorCount -gt 0) {
            Write-Host -ForegroundColor Yellow "The activity of some admins could not be determined on $activityLookupErrorCount site(s) (SiteAdminsActivityStatus = PARTIAL, see SiteAdminsNotActive): 'Lookup failed' usually means the app registration lacks the Microsoft Graph application permission User.Read.All, 'Not in the usage report' an account absent from the Microsoft 365 usage report (often unlicensed). Run with -Verbose for details."
        }
    }

    # Disconnect-PnPOnline cannot target a specific connection; releasing the variable disposes it
    $adminConnection = $null

    Write-Host -ForegroundColor Green "Processed $($spoSitesInfosArray.Count) SharePoint Online site(s)."

    if ($geoEnumerationErrorCount -gt 0) {
        Write-Host -ForegroundColor Yellow "Site enumeration failed for $geoEnumerationErrorCount geo location(s) - their sites are entirely missing from this report. Run with -Verbose to see each failing geo."
    }

    if ($siteErrorCount -gt 0) {
        Write-Host -ForegroundColor Yellow "The PnP connection could not be established on $siteErrorCount site(s) - no site-level data at all for those. Run with -Verbose to see the exact error per site:"

        foreach ($siteErrorUrl in $siteErrorUrls) {
            Write-Host -ForegroundColor Yellow "  $siteErrorUrl"
        }

        Write-Host -ForegroundColor Yellow 'A recurring error usually means the app registration lacks the SharePoint Sites.FullControl.All application permission, or the OneDrive personal sites block app-only access.'
    }

    if ($sitePartialErrorCount -gt 0) {
        Write-Host -ForegroundColor Yellow "One or more site-level data points (Admins, Members, Regional settings) could not be read on $sitePartialErrorCount site(s) - the other data points collected for those sites are still usable. Run with -Verbose to see the exact failure per site."
    }

    if ($claimResolutionErrorCount -gt 0) {
        Write-Host -ForegroundColor Yellow "$claimResolutionErrorCount group claim(s) could not be resolved to individual users - look for '<unresolved:guid>' in the *Resolved columns. Run with -Verbose for details."
        Write-Host -ForegroundColor Yellow 'A recurring error usually means the app registration lacks the Microsoft Graph application permission Group.ReadWrite.All (or Group.Read.All), or the group was deleted.'
    }

    # Drill-down: the per-site counts collected above already say which sites carry at least one link, so the
    # expensive file-level pass only visits those. Delegated to Get-SPOSharingLinkReport rather than
    # duplicated here - it owns the item resolution and the Graph calls, and stays usable on its own.
    $sharingLinksDetailsArray = [System.Collections.Generic.List[PSCustomObject]]::new()

    if ($IncludeSharingLinksDetails.IsPresent) {
        $sitesWithLinks = @($spoSitesInfosArray | Where-Object { $_.SharingLinksTotalCount -gt 0 })

        if ($sitesWithLinks.Count -eq 0) {
            Write-Host -ForegroundColor Cyan 'No site holds a sharing link: skipping the file-level sharing link collection.'
        }
        else {
            Write-Host -ForegroundColor Cyan "Collecting the file-level sharing links of $($sitesWithLinks.Count) site(s) holding at least one link..."

            $sharingLinkParams = $pnpAuthParams.Clone()
            # Get-SPOSharingLinkReport names the tenant parameter Tenant and the certificate the same way,
            # except for the PnP-specific Thumbprint alias used by Connect-PnPOnline.
            if ($sharingLinkParams.ContainsKey('Thumbprint')) {
                $sharingLinkParams.Remove('Thumbprint')
                $sharingLinkParams.Add('CertificateThumbprint', $CertificateThumbprint)
            }
            $sharingLinkParams.Add('ThrottleLimit', $ThrottleLimit)

            try {
                $collectedLinks = @($sitesWithLinks.Url | Get-SPOSharingLinkReport @sharingLinkParams -ErrorAction Stop)
                foreach ($collectedLink in $collectedLinks) {
                    $sharingLinksDetailsArray.Add($collectedLink)
                }
            }
            catch {
                Write-Warning "Unable to collect the file-level sharing links: $_"
            }

            # Attached per site so the caller keeps a single object per site and drills down with
            # $site.SharingLinksDetails. Sites with no link get an empty collection, never $null: iterating
            # the property must be safe without a null check.
            $linksBySite = $sharingLinksDetailsArray | Group-Object -Property SiteUrl -AsHashTable -AsString
            foreach ($siteInfo in $spoSitesInfosArray) {
                $siteLinks = if ($linksBySite -and $linksBySite.ContainsKey($siteInfo.Url)) { @($linksBySite[$siteInfo.Url]) } else { @() }
                $siteInfo | Add-Member -NotePropertyName 'SharingLinksDetails' -NotePropertyValue $siteLinks -Force
            }
        }
    }

    Write-Host -ForegroundColor Yellow "Reminder: check 'My Site Secondary Admin' too - https://<tenant>-admin.sharepoint.com/_layouts/15/Online/PersonalSites.aspx?PersonalSitesOverridden=1"
    Write-Host -ForegroundColor Yellow "  Why: that tenant setting silently makes an account a secondary site collection administrator on every OneDrive, but it is not exposed by any supported cmdlet or API (Get-SPOTenant / Graph / PnP), so this report cannot read it."
    Write-Host -ForegroundColor Yellow '  On top of that, an app-only connection is not granted access to OneDrive sites it is not explicitly admin of, so SiteAdmins/SiteAdminsResolved are expected to be incomplete or empty for those sites. The admin center page above is the only reliable source.'

    $findings = $null
    if ($ExportToExcel.IsPresent -or $ExportToHtml.IsPresent) {
        $collectedSwitches = @{
            M365GroupsDetails         = $m365GroupsAvailable
            IncludeSiteMembers        = $IncludeSiteMembers.IsPresent
            IncludeSiteVisitors       = $IncludeSiteVisitors.IsPresent
            IncludeSharingLinks       = $IncludeSharingLinks.IsPresent
            IncludeSiteAdminsActivity = $IncludeSiteAdminsActivity.IsPresent
            IncludeEveryoneAccess     = $IncludeEveryoneAccess.IsPresent
        }
        $findings = Get-SPOSiteFinding -Sites @($spoSitesInfosArray) -CollectedSwitches $collectedSwitches
    }

    # Storage summary (same computation as Get-SPOStorageReport), from the columns already collected. Its
    # quota percentage only makes sense over every SharePoint site of the tenant, hence the scope checks.
    $storageSummary = $null
    if ($IncludeStorageSummary.IsPresent) {
        if (-not ($ExportToExcel.IsPresent -or $ExportToHtml.IsPresent)) {
            if ($storageSummaryRequested) {
                Write-Warning 'IncludeStorageSummary only adds a section to the Excel and HTML exports: use it with ExportToExcel or ExportToHtml, or call Get-SPOStorageReport to get the summary as an object.'
            }
        }
        elseif ($SiteUrl -or $OnlyOneDrive.IsPresent) {
            if ($storageSummaryRequested) {
                Write-Warning 'IncludeStorageSummary is skipped with SiteUrl or OnlyOneDrive: the quota usage needs every SharePoint site of the tenant.'
            }
        }
        else {
            # A site is orphaned when it is flagged by 'no owner' or 'no active admin', provided at least one
            # of those two controls was evaluated; otherwise the information is unknown ($null)
            $orphanCodes = @('EID-GROUP-13', 'SPO-OWNER-01')
            $orphanEvaluated = @($findings.Summary | Where-Object { $orphanCodes -contains $_.Code -and $_.Status -like 'Evaluated*' }).Count -gt 0
            $orphanUrls = @($findings.Details | Where-Object { $orphanCodes -contains $_.Code } | ForEach-Object { $_.Url })

            $storageSites = foreach ($siteInfo in $spoSitesInfosArray) {
                [PSCustomObject]@{
                    Title               = $siteInfo.SPTitle
                    Url                 = $siteInfo.Url
                    Template            = $siteInfo.Template
                    UsedGB              = $siteInfo.StorageUsedGB
                    VersionGB           = $siteInfo.VersionSizeGB
                    LastContentModified = $siteInfo.LastContentModifiedDate
                    IsOrphan            = if ($orphanEvaluated) { $orphanUrls -contains $siteInfo.Url } else { $null }
                }
            }

            # $tenantVersionPolicy is the Get-PnPTenant object read at the start; StorageQuota is in MB
            $tenantQuotaGB = if ($tenantVersionPolicy -and $tenantVersionPolicy.StorageQuota -gt 0) { [double]$tenantVersionPolicy.StorageQuota / 1024 } else { $null }

            $storageSummaryParams = @{
                Sites             = @($storageSites)
                TenantQuotaGB     = $tenantQuotaGB
                OneDriveCollected = -not $ExcludeOneDrive.IsPresent
                DormantDays       = $DormantDays
                PricePerGBMonth   = $PricePerGBMonth
                Currency          = $Currency
            }
            $storageSummary = Get-SPOStorageSummary @storageSummaryParams
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
        $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-SPOSiteReport.xlsx"

        if ($ExcelTemplatePath) {
            if (-not (Test-Path -Path $ExcelTemplatePath -PathType Leaf)) {
                Write-Warning "ExcelTemplatePath not found: $ExcelTemplatePath"
                return
            }

            # Export-Excel accepts an existing .xlsx file and only touches the named worksheet, leaving
            # every other sheet untouched - copying the template first is the officially documented way
            # to use it: the copy becomes the report, with the template's other sheets/formatting intact.
            Copy-Item -Path $ExcelTemplatePath -Destination $excelFilePath -Force
            Write-Host -ForegroundColor Cyan "Using Excel template: $ExcelTemplatePath"
        }

        Write-Host -ForegroundColor Cyan 'Exporting SharePoint sites to Excel...'

        # SharingLinksDetails holds a collection, which a spreadsheet cell cannot represent (it would render
        # as the type name). It is dropped from the main worksheet and written to its own sheet below, where
        # one link per row is the natural shape anyway.
        $sitesToExport = if ($IncludeSharingLinksDetails.IsPresent) {
            $spoSitesInfosArray | Select-Object -Property * -ExcludeProperty 'SharingLinksDetails'
        }
        else {
            $spoSitesInfosArray
        }

        $sitesToExport | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-SiteReport' -TableStyle Light9

        if ($IncludeSharingLinksDetails.IsPresent -and $sharingLinksDetailsArray.Count -gt 0) {
            Write-Host -ForegroundColor Cyan "Exporting $($sharingLinksDetailsArray.Count) sharing link(s) to the SharePoint-SharingLinks worksheet..."
            $sharingLinksDetailsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-SharingLinks' -TableStyle Light9
        }

        $findings.Summary | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-Findings' -TableStyle Light9

        if ($findings.Details.Count -gt 0) {
            $findings.Details | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-FindingsSites' -TableStyle Light9
        }

        if ($null -ne $storageSummary) {
            ConvertTo-SPOStorageSummaryRow -Summary $storageSummary | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SharePoint-Storage' -TableStyle Light9
        }

        Add-PS365ExcelAboutSheet -Path $excelFilePath
        Write-Host -ForegroundColor Green 'Excel report exported to:'
        Write-Host -ForegroundColor White "  $excelFilePath"
    }

    # Not an elseif: both exports can be produced by the same run, which avoids collecting everything twice
    if ($ExportToHtml.IsPresent) {
        Write-Verbose 'Preparing HTML report export...'
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
        $htmlFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-SPOSiteReport.html"
        $generatedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $siteCount = $spoSitesInfosArray.Count
        $statsText = "$siteCount site$(if ($siteCount -ne 1) { 's' })"

        # Branding first: the copy buttons of the sections below come from it
        $branding = Get-PS365HtmlBranding -Title 'SharePoint Site Report' -Subtitle "Generated on $generatedDate - $statsText" -ActionsHtml '<button class="btn" type="button" onclick="expandAll()">Expand all</button><button class="btn" type="button" onclick="collapseAll()">Collapse all</button>'
        $copyButtonHtml = $branding.CopyButtonHtml

        # Every site is rendered server-side into a static, collapsible <details> tree: no client-side
        # data model or JSON payload is needed (unlike a force-directed graph), so the JS below is only
        # a thin layer for expand/collapse-all and the search filter.
        $sitesHtmlBuilder = [System.Text.StringBuilder]::new()

        # Findings summary table, grouped by category, with the affected sites of each control in a
        # collapsible list. Not collected controls stay visible (greyed) so the reader knows they were
        # not evaluated rather than reading them as clean.
        $findingCodesBySite = @{}
        foreach ($detail in $findings.Details) {
            if (-not $findingCodesBySite.ContainsKey($detail.Url)) {
                $findingCodesBySite[$detail.Url] = [System.Collections.Generic.List[string]]::new()
            }
            $findingCodesBySite[$detail.Url].Add($detail.Code)
        }

        # 365Ops layout: a summary card (how many controls need attention, stacked bar, legend), then the
        # controls in a table card. Flagged controls come first, then the clear ones, then the controls not
        # collected, which stay visible so the reader knows they were not evaluated rather than clean.
        $evaluatedRows = @($findings.Summary | Where-Object { $_.Status -like 'Evaluated*' })
        $flaggedRows = @($evaluatedRows | Where-Object { $_.AffectedCount -gt 0 })
        $clearRows = @($evaluatedRows | Where-Object { $_.AffectedCount -eq 0 })
        $notCollectedRows = @($findings.Summary | Where-Object { $_.Status -notlike 'Evaluated*' })

        $findingsHtmlBuilder = [System.Text.StringBuilder]::new()
        $headline = if ($flaggedRows.Count -eq 1) { '1 control needs attention' } else { "$($flaggedRows.Count) controls need attention" }
        $null = $findingsHtmlBuilder.Append("<section class='findings' data-copy-root><div class='section-head'><h2>Findings</h2>$copyButtonHtml</div><div class='card summary-card'>")
        $null = $findingsHtmlBuilder.Append("<div class='summary-head'><div class='summary-title'>$headline</div><div class='summary-of'>of $($evaluatedRows.Count) evaluated</div></div>")
        $null = $findingsHtmlBuilder.Append("<div class='stack-bar no-copy' role='img' aria-label='$($flaggedRows.Count) flagged, $($clearRows.Count) clear, $($notCollectedRows.Count) not collected'>")
        foreach ($segment in @(@{ Count = $flaggedRows.Count; Color = 'var(--flag)' }, @{ Count = $clearRows.Count; Color = 'var(--ok)' }, @{ Count = $notCollectedRows.Count; Color = 'var(--skip)' })) {
            if ($segment.Count -gt 0) {
                $null = $findingsHtmlBuilder.Append("<div style='flex:$($segment.Count);background:$($segment.Color)'></div>")
            }
        }
        $null = $findingsHtmlBuilder.Append('</div>')
        $null = $findingsHtmlBuilder.Append("<div class='legend'><span class='st'><span class='dot' style='background:var(--flag)'></span>Flagged $($flaggedRows.Count)</span><span class='st'><span class='dot' style='background:var(--ok)'></span>Clear $($clearRows.Count)</span><span class='st'><span class='dot' style='background:var(--skip)'></span>Not collected $($notCollectedRows.Count)</span></div>")
        $null = $findingsHtmlBuilder.Append('</div>')

        $null = $findingsHtmlBuilder.Append("<div class='card table-card'><div class='table-scroll'><table class='ps365-table'><thead><tr><th style='width:52%'>Control</th><th>Affected</th><th>Category</th><th>Status</th></tr></thead><tbody>")

        foreach ($row in @($flaggedRows) + @($clearRows) + @($notCollectedRows)) {
            $isEvaluated = $row.Status -like 'Evaluated*'
            $controlCell = "<div class='ctl-name'>$(ConvertTo-SPOHtmlEncoded $row.Control)</div><div class='id'>$($row.Code) - $(ConvertTo-SPOHtmlEncoded $row.Explanation)</div>"

            if (-not $isEvaluated) {
                $affectedCell = "<span class='pill pill-low'>-</span>"
                $statusCell = "<span class='st'><span class='dot' style='background:var(--skip)'></span>Not collected</span><div class='id'>$(ConvertTo-SPOHtmlEncoded ($row.Status -replace '^NotCollected ', ''))</div>"
            }
            else {
                $pillClass = if ($row.AffectedCount -gt 0) { 'pill pill-high' } else { 'pill pill-low' }
                $affectedCell = "<span class='$pillClass'>$($row.AffectedCount) $(ConvertTo-SPOHtmlEncoded $row.AffectedUnit)</span>"
                if ($null -ne $row.ItemsCount) {
                    $affectedCell += "<div class='id'>$($row.ItemsCount) $(ConvertTo-SPOHtmlEncoded $row.ItemsUnit)</div>"
                }

                $statusCell = if ($row.AffectedCount -gt 0) { "<span class='st' style='color:var(--flag)'><span class='dot' style='background:var(--flag)'></span>Flagged</span>" } else { "<span class='st' style='color:var(--ok)'><span class='dot' style='background:var(--ok)'></span>Clear</span>" }
                if ($row.Status -ne 'Evaluated') {
                    $statusCell += "<div class='id'>$(ConvertTo-SPOHtmlEncoded ($row.Status -replace '^Evaluated ', ''))</div>"
                }

                $rowDetails = @($findings.Details | Where-Object { $_.Code -eq $row.Code })
                if ($rowDetails.Count -gt 0) {
                    # 'site(s)' / 'group(s)' turned into a proper singular or plural for the toggle label
                    $unitWord = if ($rowDetails.Count -eq 1) { $row.AffectedUnit -replace '\(s\)$', '' } else { $row.AffectedUnit -replace '\(s\)$', 's' }
                    $controlCell += "<details class='affected-list'><summary>View $($rowDetails.Count) $(ConvertTo-SPOHtmlEncoded $unitWord)</summary><table class='aff-table'><tbody>"
                    foreach ($rowDetail in $rowDetails) {
                        # A site without title (a OneDrive root, for example) is named by its URL
                        $affectedName = if ([string]::IsNullOrWhiteSpace($rowDetail.SPTitle)) { $rowDetail.Url } else { $rowDetail.SPTitle }
                        $affectedUrlLine = if ([string]::IsNullOrWhiteSpace($rowDetail.SPTitle)) { '' } else { "<div class='aff-url'>$(ConvertTo-SPOHtmlEncoded $rowDetail.Url)</div>" }
                        $controlCell += "<tr><td><div class='aff-title'>$(ConvertTo-SPOHtmlEncoded $affectedName)</div>$affectedUrlLine<div class='aff-detail'>$(ConvertTo-SPOHtmlEncoded $rowDetail.Detail)</div></td></tr>"
                    }
                    $controlCell += '</tbody></table></details>'
                }
            }

            $rowClass = if ($isEvaluated) { '' } else { " class='muted-row'" }
            $null = $findingsHtmlBuilder.Append("<tr$rowClass><td>$controlCell</td><td>$affectedCell</td><td>$(ConvertTo-SPOHtmlEncoded $row.Category)</td><td>$statusCell</td></tr>")
        }

        $null = $findingsHtmlBuilder.Append('</tbody></table></div></div></section>')

        foreach ($site in $spoSitesInfosArray) {
            $siteFindingCodes = if ($findingCodesBySite.ContainsKey($site.Url)) { @($findingCodesBySite[$site.Url]) } else { @() }
            # Finding codes are part of the search text, so typing 'SPO-SHARE-01' lists the sites it flags
            $searchText = (ConvertTo-SPOHtmlEncoded ("$($site.SPTitle) $($site.Url) $($siteFindingCodes -join ' ')").ToLowerInvariant())
            $typeBadge = if ($M365GroupsDetails.IsPresent -and $site.Type) { "<span class='pill pill-low'>$(ConvertTo-SPOHtmlEncoded $site.Type)</span>" } else { '' }
            $findingBadges = ($siteFindingCodes | ForEach-Object { "<span class='pill pill-high'>$_</span>" }) -join ''

            $null = $sitesHtmlBuilder.Append("<details class='site-block' data-copy-root data-search='$searchText'>")
            $siteName = if ([string]::IsNullOrWhiteSpace($site.SPTitle)) { '(no title)' } else { $site.SPTitle }
            $null = $sitesHtmlBuilder.Append("<summary><span class='site-title'>$(ConvertTo-SPOHtmlEncoded $siteName)</span> <span class='site-url'>$(ConvertTo-SPOHtmlEncoded $site.Url)</span><span class='site-badges'>$findingBadges$typeBadge</span>$copyButtonHtml</summary>")
            $null = $sitesHtmlBuilder.Append("<div class='site-body'>")
            $null = $sitesHtmlBuilder.Append("<div class='site-meta'>Template: $(ConvertTo-SPOHtmlEncoded $site.Template) - Storage: $($site.StorageUsedGB) / $($site.StorageLimitGB) GB - Sharing: $(ConvertTo-SPOHtmlEncoded $site.SharingCapability)</div>")

            $null = $sitesHtmlBuilder.Append((ConvertTo-SPORoleHtml -RoleLabel 'Site Collection Admins' -RoleCssClass 'admins' -ResolvedString $site.SiteAdminsResolved))
            $null = $sitesHtmlBuilder.Append((ConvertTo-SPORoleHtml -RoleLabel 'Owners' -RoleCssClass 'owners' -ResolvedString $site.SiteOwnersResolved))
            $null = $sitesHtmlBuilder.Append((ConvertTo-SPORoleHtml -RoleLabel 'Members' -RoleCssClass 'members' -ResolvedString $site.SiteMembersResolved))
            $null = $sitesHtmlBuilder.Append((ConvertTo-SPORoleHtml -RoleLabel 'Visitors' -RoleCssClass 'visitors' -ResolvedString $site.SiteVisitorsResolved))

            if ($M365GroupsDetails.IsPresent -and $site.Type -eq 'M365 Group') {
                $null = $sitesHtmlBuilder.Append("<div class='m365-info'>")
                $null = $sitesHtmlBuilder.Append("<div class='m365-row'><span class='m365-lbl'>Microsoft 365 group</span><span class='m365-val'>$(ConvertTo-SPOHtmlEncoded $site.PrimarySmtpAddress)</span></div>")
                if ($site.M365GroupOwners) {
                    $null = $sitesHtmlBuilder.Append("<div class='m365-row'><span class='m365-lbl'>Group owners</span><span class='m365-val'>$(ConvertTo-SPOHtmlEncoded $site.M365GroupOwners)</span></div>")
                }
                if ($null -ne $site.M365GroupMembersCount) {
                    $null = $sitesHtmlBuilder.Append("<div class='m365-row'><span class='m365-lbl'>Members / Guests</span><span class='m365-val'>$($site.M365GroupMembersCount) / $($site.M365GroupGuestsCount)</span></div>")
                }
                if ($site.TeamDisplayName) {
                    $null = $sitesHtmlBuilder.Append("<div class='m365-row'><span class='m365-lbl'>Microsoft Teams team</span><span class='m365-val'>$(ConvertTo-SPOHtmlEncoded $site.TeamDisplayName) ($($site.TeamChannelCount) channel(s))</span></div>")
                }
                $null = $sitesHtmlBuilder.Append('</div>')
            }

            $null = $sitesHtmlBuilder.Append('</div></details>')
        }


        # Single-quoted here-string: no PS variable expansion inside HTML/JS.
        # Placeholders are replaced via .NET String.Replace() (literal, not regex), same pattern as
        # the HTML graph export in Get-NestedGroup.
        $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SharePoint Site Report - PS365</title>
  PS365_BRAND_FAVICON
  <style>
PS365_BRAND_CSS
    /* Shared PS365 design (365Ops look: tokens, top bar, page heading, cards, buttons, pills, tables,
       footer) comes from Get-PS365HtmlBranding above. Below: what is specific to this report */

    /* Findings: summary card then the controls table card */
    .summary-card { padding: 20px; margin-bottom: 16px; }
    .summary-head { display: flex; flex-wrap: wrap; align-items: baseline; gap: 8px 24px; margin-bottom: 14px; }
    .summary-title { font-size: 22px; font-weight: 700; letter-spacing: -0.01em; }
    .summary-of { color: var(--muted); }
    .stack-bar { display: flex; height: 10px; border-radius: 5px; overflow: hidden; gap: 2px; background: var(--divider); }
    .legend { display: flex; flex-wrap: wrap; gap: 8px 24px; margin-top: 12px; font-size: 13px; color: #3D4350; }
    .legend .st { font-weight: 500; }
    .table-card { overflow: hidden; }
    .table-scroll { overflow-x: auto; }
    .ctl-name { font-weight: 600; }
    .ps365-table tr.muted-row .ctl-name { color: var(--muted); }
    /* Affected items of a control: a discreet toggle, then a compact table in a light panel */
    details.affected-list { margin-top: 8px; }
    details.affected-list > summary { display: inline-flex; align-items: center; gap: 6px; cursor: pointer; list-style: none; color: var(--accent); font-size: 12.5px; font-weight: 600; padding: 3px 10px 3px 8px; border: 1px solid var(--border); border-radius: 14px; background: var(--surface); }
    details.affected-list > summary::-webkit-details-marker { display: none; }
    details.affected-list > summary::before { content: ""; width: 6px; height: 6px; border-right: 2px solid currentColor; border-bottom: 2px solid currentColor; transform: rotate(-45deg); transition: transform .15s; margin-right: 2px; }
    details.affected-list[open] > summary::before { transform: rotate(45deg); }
    details.affected-list > summary:hover { background: var(--surface2); }
    table.aff-table { width: 100%; margin-top: 8px; border-collapse: separate; border-spacing: 0; background: var(--surface2); border: 1px solid var(--divider); border-radius: 8px; overflow: hidden; }
    .aff-table td { padding: 8px 12px; border-bottom: 1px solid var(--divider); vertical-align: top; font-size: 13px; }
    .aff-table tr:last-child td { border-bottom: none; }
    .aff-title { font-weight: 600; overflow-wrap: anywhere; }
    .aff-url { color: var(--sub); font-size: 12px; font-family: var(--mono); overflow-wrap: anywhere; }
    .aff-detail { color: var(--muted); margin-top: 2px; }

    /* Sites: one card, a toolbar with the search, then one collapsible row per site */
    .sites-card { overflow: hidden; }
    .sites-card .card-toolbar .search { flex: 1 1 240px; max-width: 360px; }
    .sites-card .card-toolbar .count { color: var(--sub); font-size: 13px; margin-left: auto; }
    details.site-block { border-bottom: 1px solid var(--divider); }
    details.site-block:last-of-type { border-bottom: none; }
    details.site-block > summary { padding: 12px 16px; cursor: pointer; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; list-style: none; }
    details.site-block > summary:hover { background: var(--surface2); }
    details.site-block > summary::-webkit-details-marker { display: none; }
    details.site-block > summary::before { content: "\25B8"; color: var(--sub); font-size: 11px; transition: transform .15s; flex-shrink: 0; }
    details.site-block[open] > summary::before { transform: rotate(90deg); }
    .site-title { font-weight: 600; }
    .site-url { color: var(--sub); font-size: 12px; font-family: var(--mono); }
    .site-badges { display: flex; gap: 6px; flex-wrap: wrap; margin-left: auto; }
    details.site-block > summary .btn-copy { margin-left: 4px; }
    .findings .section-head, section .section-head { margin-top: 4px; }
    .site-body { padding: 4px 16px 16px 38px; background: var(--surface2); border-top: 1px solid var(--divider); }
    .site-meta { font-size: 12.5px; color: var(--muted); padding: 10px 0; }

    details.role-block { margin: 6px 0; }
    details.role-block > summary { cursor: pointer; list-style: none; display: flex; align-items: center; gap: 8px; padding: 4px 0; }
    details.role-block > summary::-webkit-details-marker { display: none; }
    details.role-block > summary::before { content: "\25B8"; color: var(--sub); font-size: 10px; transition: transform .15s; }
    details.role-block[open] > summary::before { transform: rotate(90deg); }
    .badge-role { border-radius: 5px; padding: 2px 8px; font-size: 12px; font-weight: 600; color: #fff; }
    .badge-role.admins { background: var(--admins); }
    .badge-role.owners { background: var(--owners); }
    .badge-role.members { background: var(--members); }
    .badge-role.visitors { background: var(--visitors); }
    .count { color: var(--sub); font-size: 12px; }

    ul.identity-list { list-style: none; margin: 4px 0 4px 20px; padding-left: 14px; border-left: 2px solid var(--divider); }
    li.user-item { font-size: 13px; padding: 2px 0; color: var(--text); }
    li.user-item.unresolved { color: var(--admins); font-style: italic; }

    details.group-block { margin: 2px 0; }
    details.group-block > summary { cursor: pointer; list-style: none; display: flex; align-items: center; gap: 8px; padding: 2px 0; }
    details.group-block > summary::-webkit-details-marker { display: none; }
    details.group-block > summary::before { content: "\25B8"; color: var(--sub); font-size: 9px; transition: transform .15s; }
    details.group-block[open] > summary::before { transform: rotate(90deg); }
    .badge-group { background: var(--accent-light); color: #92400E; border-radius: 5px; padding: 2px 8px; font-size: 12px; font-weight: 600; }

    .m365-info { margin-top: 10px; padding-top: 10px; border-top: 1px solid var(--divider); }
    .m365-row { display: flex; gap: 10px; font-size: 13px; padding: 3px 0; }
    .m365-lbl { color: var(--sub); min-width: 160px; flex-shrink: 0; }
    .m365-val { color: var(--text); }

    .empty-state { text-align: center; color: var(--sub); padding: 40px 0; }
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

PS365_FINDINGS_HTML

<section data-copy-root>
  <div class="section-head"><h2>Sites</h2>PS365_COPY_BUTTON</div>
  <div class="card sites-card">
    <div class="card-toolbar">
      <input id="search" class="search" type="search" placeholder="Find a site by title, URL or finding code" autocomplete="off" aria-label="Find a site" />
      <span class="count no-copy" id="site-count">PS365_STATS_TEXT</span>
    </div>
    <div id="sites">
PS365_SITES_HTML
    </div>
    <div class="empty-state" id="empty-state" style="display:none">No site matches your search.</div>
  </div>
</section>
</div>

PS365_BRAND_FOOTER

PS365_BRAND_SCRIPT
<script>
  function expandAll() {
    document.querySelectorAll('#sites details').forEach(function (d) { d.open = true; });
  }
  function collapseAll() {
    document.querySelectorAll('#sites > details.site-block').forEach(function (d) { d.open = false; });
  }

  document.getElementById('search').addEventListener('input', function () {
    var term = this.value.trim().toLowerCase();
    var blocks = document.querySelectorAll('#sites > details.site-block');
    var visibleCount = 0;
    blocks.forEach(function (block) {
      var matches = !term || block.dataset.search.indexOf(term) !== -1;
      block.style.display = matches ? '' : 'none';
      if (matches) {
        visibleCount++;
        if (term) { block.open = true; }
      }
    });
    document.getElementById('site-count').textContent = term ? visibleCount + ' of ' + blocks.length + ' sites' : blocks.length + (blocks.length === 1 ? ' site' : ' sites');
    document.getElementById('empty-state').style.display = visibleCount === 0 ? 'block' : 'none';
  });
</script>
</body>
</html>
'@

        $html = $htmlTemplate
        $html = $html.Replace('PS365_BRAND_CSS', $branding.Css)
        $html = $html.Replace('PS365_BRAND_FAVICON', $branding.HeadHtml)
        $html = $html.Replace('PS365_BRAND_HEADER', $branding.HeaderBrandHtml)
        $html = $html.Replace('PS365_BRAND_FOOTER', $branding.FooterHtml)
        $html = $html.Replace('PS365_BRAND_PAGEHEAD', $branding.PageHeadHtml)
        $html = $html.Replace('PS365_BRAND_SCRIPT', $branding.ScriptHtml)
        $html = $html.Replace('PS365_COPY_BUTTON', $copyButtonHtml)
        $storageHtml = if ($null -ne $storageSummary) { ConvertTo-SPOStorageSummaryHtml -Summary $storageSummary -CopyButtonHtml $copyButtonHtml } else { '' }
        $html = $html.Replace('PS365_FINDINGS_HTML', $findingsHtmlBuilder.ToString() + $storageHtml)
        $html = $html.Replace('PS365_SITES_HTML', $sitesHtmlBuilder.ToString())
        $html = $html.Replace('PS365_GENERATED_DATE', $generatedDate)
        $html = $html.Replace('PS365_STATS_TEXT', $statsText)

        Write-Verbose "HTML report file path: $htmlFilePath"
        $html | Out-File -FilePath $htmlFilePath -Encoding UTF8
        Write-Host -ForegroundColor Green 'HTML report exported to:'
        Write-Host -ForegroundColor White "  $htmlFilePath"
        Invoke-Item -Path $htmlFilePath
    }

    if (-not ($ExportToExcel.IsPresent -or $ExportToHtml.IsPresent)) {
        return $spoSitesInfosArray
    }
}
