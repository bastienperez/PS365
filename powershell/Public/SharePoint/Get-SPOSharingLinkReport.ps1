function Get-SPOSharingLinkReport {
    <#
    .SYNOPSIS
    Lists every sharing link of one or more SharePoint Online sites, one row per link, using PnP.PowerShell app-only.

    .DESCRIPTION
    Get-SPOSharingLinkReport is the file-level companion of Get-SPOSiteReport -IncludeSharingLinks: where the
    latter answers "which sites have an exposure worth looking at" with a per-site count, this function
    answers "which file is shared, with whom, until when, and through which URL" for the sites you point it
    at. It is deliberately not tenant-wide by default: the detail costs Microsoft Graph calls per shared
    item, so the intended workflow is to triage with Get-SPOSiteReport, then drill into the sites it flags.

    HOW THE SHARED ITEMS ARE FOUND
    A naive implementation walks every document library, loads HasUniqueRoleAssignments on every single item,
    and asks Graph for the links of the ones that have unique permissions. That is one round trip per file,
    which does not survive a real tenant. This function inverts the problem: SharePoint already materialises
    every sharing link as a hidden site group named 'SharingLinks.<itemUniqueId>.<linkType>.<linkGuid>', so a
    single Get-PnPGroup call per site yields the exact list of items that carry a link. Only those items are
    then resolved and queried. The cost is proportional to the number of shared items, not to the number of
    files in the site, which is usually several orders of magnitude smaller.

    Each candidate item is resolved from its unique ID through the REST endpoints /_api/web/GetFileById and,
    if that returns nothing, /_api/web/GetFolderById - an item is one or the other, and the group name does
    not say which. The links themselves are then read with Get-PnPFileSharingLink / Get-PnPFolderSharingLink,
    which is what provides the authoritative link type (the Graph 'scope' property), the expiration date, the
    password protection, the download block and the identities the link was shared with. The group name is
    used only as an index of what to look at, never as the source of truth for the link itself.

    TWO MODES, AND WHY THE DEFAULT AVOIDS GRAPH
    By default the function never calls Microsoft Graph. Everything it reports comes from the SharingLinks
    groups themselves: the shared item (resolved through SharePoint REST), the link type read from the group
    name, and - the part that matters most for an access audit - the people the link was shared with, who
    are literally the members of that group. This runs with the SharePoint Sites.FullControl.All permission
    the rest of the module already needs, and nothing more.

    IncludeGraphDetails adds what only Graph knows: the link URL, the expiration date, the password
    protection, the download block, the role granted, and the authoritative scope of a Flexible link. That
    costs one Graph call per shared item AND the Microsoft Graph APPLICATION permissions Files.Read.All and
    Sites.Read.All - which mean "read every file of the tenant, with no user behind the call". On a
    production tenant that is a decision worth making deliberately, not a checkbox: consider Sites.Selected
    instead, which scopes the app to the sites you explicitly grant it.

    WHAT THE DEFAULT MODE CANNOT TELL YOU
    A link type of 'Flexible' is reported as Flexible, never as 'SpecificPeople'. Flexible is the modern
    SharePoint link model where the scope is set per link, so the group name genuinely does not say whether
    such a link is restricted to named people or open to anyone holding it. Guessing here would be actively
    harmful: reporting a Flexible link as 'SpecificPeople' makes SharingLinksAnyoneCount read as zero on a
    site that may well have an anonymous link. When that distinction matters, IncludeGraphDetails is the
    only way to resolve it - the shared-with list of the default mode is often enough to conclude on its own.

    PERMISSIONS
    - Default mode: SharePoint (Office 365 SharePoint Online) Sites.FullControl.All application permission,
      the same one Get-SPOSiteReport uses. No Graph permission at all.
    - IncludeGraphDetails: the above, plus the Microsoft Graph APPLICATION permissions Files.Read.All and
      Sites.Read.All. These are a different Entra resource - consenting to the SharePoint one does not cover
      Graph. Without them every item fails at the last step with 'accessDenied 403' (Graph token present but
      insufficient) or 'Either scp or roles claim need to be present in the token' (no Graph permission on
      the certificate at all). After granting consent, allow a few minutes for propagation and reconnect:
      PnP caches the acquired token.

    Client secrets are not supported by SharePoint for app-only: a certificate is mandatory. Provide it
    through CertificateThumbprint (Windows certificate store), CertificatePath (.pfx file) or
    CertificateBase64Encoded (base64 string, handy for Azure Automation or a pipeline variable).

    Sites are processed in parallel (ForEach-Object -Parallel) with one PnP connection per site. Tune the
    concurrency with ThrottleLimit. PowerShell 7 is required.

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

    .PARAMETER SiteUrl
    (Mandatory) One or more site collection URLs to scan. Accepts pipeline input, by value or by the Url
    property, so the output of Get-SPOSiteReport can be piped straight in (its Url column binds to this
    parameter through its alias).

    .PARAMETER IncludeGraphDetails
    (Optional) Re-reads every link through Microsoft Graph to add the link URL, the expiration date, the
    password protection, the download block, the granted role, and the real scope of a Flexible link.
    Requires the Microsoft Graph Files.Read.All and Sites.Read.All application permissions - see the
    PERMISSIONS section above before turning this on against a production tenant.

    Without it, the report carries LinkType/LinkTypeRaw, SharedWith, the item, and the SharingLinks group it
    came from. With it, those Graph columns are added and the type is resolved. Columns that were not
    collected are absent rather than empty, so an empty cell always means the data really is empty.

    .PARAMETER LinkType
    (Optional) Restricts the output to one or more link types: Anyone (anonymous 'Anyone with the link'),
    Company (people in the organization), Flexible (modern link whose scope is only resolvable through
    Graph), SpecificPeople (requires IncludeGraphDetails), or Other for a type this function does not know
    about yet. Filtering happens after collection, so it does not make the run faster - it only makes the
    report shorter.

    .PARAMETER ExpiredOnly
    (Optional) Keeps only the links whose expiration date is in the past. Requires IncludeGraphDetails:
    expiration dates only exist in the Graph payload. Mutually exclusive with NeverExpiringOnly and
    ExpiringInDays.

    .PARAMETER NeverExpiringOnly
    (Optional) Keeps only the links that carry no expiration date at all. Requires IncludeGraphDetails.
    Mutually exclusive with ExpiredOnly and ExpiringInDays.

    .PARAMETER ExpiringInDays
    (Optional) Keeps only the links expiring within that many days from now, excluding those already
    expired. Requires IncludeGraphDetails. Mutually exclusive with ExpiredOnly and NeverExpiringOnly.

    .PARAMETER ThrottleLimit
    (Optional) Number of sites processed concurrently. Defaults to 8. Lower it if SharePoint starts
    answering 429: past a certain rate, more concurrency means more throttling and a longer total run.

    .PARAMETER ExportToExcel
    (Optional) Exports the result to an .xlsx file instead of returning the objects to the pipeline.

    .PARAMETER ExportPath
    (Optional) Output directory for the Excel export. Defaults to the user profile directory.

    .EXAMPLE
    Get-SPOSharingLinkReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -SiteUrl 'https://contoso.sharepoint.com/sites/marketing'

    Lists every sharing link of a single site.

    .EXAMPLE
    Get-SPOSiteReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -IncludeSharingLinks |
        Where-Object { $_.SharingLinksAnyoneCount -gt 0 } |
        Get-SPOSharingLinkReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -LinkType Anyone

    The intended workflow: triage the whole tenant with the cheap per-site counts, then get the file-level
    detail of the anonymous links only on the sites that actually have one.

    .EXAMPLE
    Get-SPOSharingLinkReport -ClientId $clientId -Tenant 'contoso.onmicrosoft.com' -CertificateThumbprint $thumb -SiteUrl $urls -NeverExpiringOnly -ExportToExcel

    Exports every link that will never expire on the given sites to an Excel file in the user profile directory.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-SPOSharingLinkReport
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

        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias('Url')]
        [string[]]$SiteUrl,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeGraphDetails,

        [Parameter(Mandatory = $false)]
        [ValidateSet('Anyone', 'Company', 'SpecificPeople', 'Flexible', 'Other')]
        [string[]]$LinkType,

        [Parameter(Mandatory = $false)]
        [switch]$ExpiredOnly,

        [Parameter(Mandatory = $false)]
        [switch]$NeverExpiringOnly,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 3650)]
        [int]$ExpiringInDays,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 32)]
        [int]$ThrottleLimit = 8,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath
    )

    begin {
        [System.Collections.Generic.List[string]]$siteUrlsList = @()

        # Exactly one certificate method is required for the app-only connection
        $certMethods = @($CertificateThumbprint, $CertificatePath, $CertificateBase64Encoded | Where-Object { $_ })
        if ($certMethods.Count -ne 1) {
            Write-Warning 'Provide exactly one certificate method: CertificateThumbprint, CertificatePath or CertificateBase64Encoded.'
            $inputIsValid = $false
        }
        else {
            $inputIsValid = $true
        }

        # The three expiration filters answer mutually exclusive questions: combining them can only ever
        # return an empty report, so refuse instead of silently producing nothing.
        $expirationFilters = @($ExpiredOnly.IsPresent, $NeverExpiringOnly.IsPresent, ($ExpiringInDays -gt 0)) | Where-Object { $_ }
        if ($expirationFilters.Count -gt 1) {
            Write-Warning 'ExpiredOnly, NeverExpiringOnly and ExpiringInDays are mutually exclusive. Choose one of them.'
            $inputIsValid = $false
        }

        # Expiration dates only exist in the Graph payload. Filtering on them without IncludeGraphDetails
        # would silently return an empty report (nothing has an expiration date to compare), so it is
        # refused rather than answered with a misleading result.
        if ($expirationFilters.Count -gt 0 -and -not $IncludeGraphDetails.IsPresent) {
            Write-Warning 'ExpiredOnly, NeverExpiringOnly and ExpiringInDays require IncludeGraphDetails: expiration dates are only exposed by Microsoft Graph, not by the SharePoint sharing link groups.'
            $inputIsValid = $false
        }

        # Same reason: without Graph the link type comes from the group name, which reports Flexible links
        # as such because their real scope is per-link. Asking for SpecificPeople in that mode would return
        # nothing while Flexible links possibly are exactly that.
        if ($LinkType -contains 'SpecificPeople' -and -not $IncludeGraphDetails.IsPresent) {
            Write-Warning "LinkType 'SpecificPeople' requires IncludeGraphDetails. Without it, use 'Flexible': SharePoint only exposes the link type through the group name, which does not resolve the scope of a Flexible link."
            $inputIsValid = $false
        }

        if (-not (Get-Module PnP.PowerShell -ListAvailable)) {
            Write-Warning 'Please install the PnP.PowerShell module: Install-Module PnP.PowerShell'
            $inputIsValid = $false
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
        elseif ($CertificateBase64Encoded) {
            $pnpAuthParams.Add('CertificateBase64Encoded', $CertificateBase64Encoded)
        }
    }

    process {
        if (-not $inputIsValid) {
            return
        }

        foreach ($url in $SiteUrl) {
            if ($url -and -not $siteUrlsList.Contains($url)) {
                $siteUrlsList.Add($url)
            }
        }
    }

    end {
        if (-not $inputIsValid) {
            return
        }

        if ($siteUrlsList.Count -eq 0) {
            Write-Warning 'No site URL to process.'
            return
        }

        Write-Host -ForegroundColor Cyan "Scanning $($siteUrlsList.Count) site(s) for sharing links with $ThrottleLimit concurrent connection(s)..."

        $rawResults = $siteUrlsList | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            $auth = $using:pnpAuthParams
            $useGraph = $using:IncludeGraphDetails
            $currentSiteUrl = $_

            $siteConnection = $null
            try {
                $siteConnection = Connect-PnPOnline -Url $currentSiteUrl @auth -ReturnConnection -ErrorAction Stop
            }
            catch {
                [PSCustomObject]@{
                    SiteUrl = $currentSiteUrl
                    Links   = @()
                    Status  = "ERROR: $($_.Exception.Message)"
                }
                return
            }

            # One call per site: the hidden 'SharingLinks.<itemUniqueId>.<linkType>.<linkGuid>' groups are the
            # index of every item carrying a link. Users are expanded in that same call because the members of
            # such a group ARE the people the link was shared with - that is the whole 'shared with' answer,
            # obtained without touching Microsoft Graph.
            try {
                $siteGroups = Get-PnPGroup -Includes Users -Connection $siteConnection -ErrorAction Stop
            }
            catch {
                [PSCustomObject]@{
                    SiteUrl = $currentSiteUrl
                    Links   = @()
                    Status  = "ERROR: unable to list the site groups - $($_.Exception.Message)"
                }
                return
            }

            # One entry per sharing link group, kept whole: the item ID says what to resolve, the link type
            # and the members are already the answer in the default (Graph-free) mode.
            [System.Collections.Generic.List[PSCustomObject]]$linkGroups = @()
            $itemUniqueIds = [System.Collections.Generic.HashSet[string]]::new()

            foreach ($siteGroup in $siteGroups) {
                if ($siteGroup.Title -notlike 'SharingLinks.*') {
                    continue
                }

                # 'SharingLinks' . <itemUniqueId> . <linkType> . <linkGuid>
                $titleParts = $siteGroup.Title.Split('.')
                if ($titleParts.Count -lt 2 -or -not $titleParts[1]) {
                    continue
                }

                $null = $itemUniqueIds.Add($titleParts[1])

                # The type token is read as-is, never normalised into something it may not be: Flexible in
                # particular carries its real scope per link, which only Graph can tell.
                $linkGroups.Add([PSCustomObject]@{
                        ItemUniqueId = $titleParts[1]
                        TypeToken    = if ($titleParts.Count -ge 3) { $titleParts[2] } else { $null }
                        LinkGuid     = if ($titleParts.Count -ge 4) { $titleParts[3] } else { $null }
                        GroupTitle   = $siteGroup.Title
                        GroupId      = $siteGroup.Id
                        SharedWith   = (@($siteGroup.Users | ForEach-Object {
                                    if ($_.Email) { $_.Email } else { $_.LoginName }
                                }) -join '|')
                    })
            }

            [System.Collections.Generic.List[PSCustomObject]]$siteLinks = @()
            [System.Collections.Generic.List[string]]$itemErrors = @()
            $resolvedItems = @{}

            foreach ($itemUniqueId in $itemUniqueIds) {
                # The group name does not say whether the item is a file or a folder, and the two have
                # distinct REST endpoints and distinct PnP cmdlets. Try the file first (by far the common
                # case), fall back to the folder.
                $itemUrl = $null
                $itemName = $null
                $itemType = $null

                try {
                    $fileResponse = Invoke-PnPSPRestMethod -Url "/_api/web/GetFileById(guid'$itemUniqueId')" -Connection $siteConnection -ErrorAction Stop
                    if ($fileResponse -and $fileResponse.ServerRelativeUrl) {
                        $itemUrl = $fileResponse.ServerRelativeUrl
                        $itemName = $fileResponse.Name
                        $itemType = 'File'
                    }
                }
                catch {
                    # Not a file (or no longer exists): the folder endpoint below decides
                    Write-Verbose "GetFileById failed for $itemUniqueId on $currentSiteUrl - $($_.Exception.Message)"
                }

                if (-not $itemUrl) {
                    try {
                        $folderResponse = Invoke-PnPSPRestMethod -Url "/_api/web/GetFolderById(guid'$itemUniqueId')" -Connection $siteConnection -ErrorAction Stop
                        if ($folderResponse -and $folderResponse.ServerRelativeUrl) {
                            $itemUrl = $folderResponse.ServerRelativeUrl
                            $itemName = $folderResponse.Name
                            $itemType = 'Folder'
                        }
                    }
                    catch {
                        Write-Verbose "GetFolderById failed for $itemUniqueId on $currentSiteUrl - $($_.Exception.Message)"
                    }
                }

                if (-not $itemUrl) {
                    # A SharingLinks group whose item resolves to neither a file nor a folder is an orphan
                    # left behind by a deleted item. Reported rather than dropped: those groups keep granting
                    # nothing, but their presence skews the per-site counts of Get-SPOSiteReport.
                    $itemErrors.Add("Item $itemUniqueId could not be resolved (deleted item, or orphaned SharingLinks group)")
                    continue
                }

                $resolvedItems[$itemUniqueId] = [PSCustomObject]@{
                    ItemType = $itemType
                    ItemName = $itemName
                    ItemUrl  = $itemUrl
                }
            }

            if (-not $useGraph) {
                # Default mode: everything comes from the site groups already read above. One row per group,
                # which is one row per link, and no Graph call anywhere in the path.
                foreach ($linkGroup in $linkGroups) {
                    $resolvedItem = $resolvedItems[$linkGroup.ItemUniqueId]
                    if ($null -eq $resolvedItem) {
                        continue
                    }

                    $siteLinks.Add([PSCustomObject]@{
                            ItemType   = $resolvedItem.ItemType
                            ItemName   = $resolvedItem.ItemName
                            ItemUrl    = $resolvedItem.ItemUrl
                            TypeToken  = $linkGroup.TypeToken
                            SharedWith = $linkGroup.SharedWith
                            GroupTitle = $linkGroup.GroupTitle
                            LinkGuid   = $linkGroup.LinkGuid
                        })
                }
            }
            else {
                # Graph mode: the links are re-read from Graph, which is the only source for the expiration
                # date, the password protection, the download block and the authoritative scope of a
                # Flexible link. One call per shared item, and the Graph application permissions that go
                # with it - see the PERMISSIONS section of the help.
                foreach ($itemUniqueId in $resolvedItems.Keys) {
                    $resolvedItem = $resolvedItems[$itemUniqueId]

                    try {
                        if ($resolvedItem.ItemType -eq 'File') {
                            $sharingLinks = Get-PnPFileSharingLink -Identity $resolvedItem.ItemUrl -Connection $siteConnection -ErrorAction Stop
                        }
                        else {
                            $sharingLinks = Get-PnPFolderSharingLink -Folder $resolvedItem.ItemUrl -Connection $siteConnection -ErrorAction Stop
                        }
                    }
                    catch {
                        $itemErrors.Add("Links of $($resolvedItem.ItemUrl) - $($_.Exception.Message)")
                        continue
                    }

                    foreach ($sharingLink in $sharingLinks) {
                        $link = $sharingLink.Link
                        if ($null -eq $link) {
                            continue
                        }

                        $siteLinks.Add([PSCustomObject]@{
                                ItemType           = $resolvedItem.ItemType
                                ItemName           = $resolvedItem.ItemName
                                ItemUrl            = $resolvedItem.ItemUrl
                                Scope              = $link.Scope
                                Permission         = $link.Type
                                Roles              = ($sharingLink.Roles -join '|')
                                SharedWith         = ($sharingLink.GrantedToIdentitiesV2.User.Email -join '|')
                                HasPassword        = $sharingLink.HasPassword
                                PreventsDownload   = $link.PreventsDownload
                                ExpirationDateTime = $sharingLink.ExpirationDateTime
                                LinkUrl            = $link.WebUrl
                                LinkId             = $sharingLink.Id
                            })
                    }
                }
            }

            $status = if ($itemErrors.Count -gt 0) { "PARTIAL: $($itemErrors -join ' | ')" } else { 'OK' }

            # Releasing the variable disposes the connection: the parallel runspace pool is reused across
            # sites, so live connection objects would otherwise accumulate inside each runspace.
            $siteConnection = $null

            [PSCustomObject]@{
                SiteUrl = $currentSiteUrl
                Links   = $siteLinks
                Status  = $status
            }
        }

        [System.Collections.Generic.List[PSCustomObject]]$sharingLinksArray = @()
        $siteErrorCount = 0
        $sitePartialErrorCount = 0
        $currentDate = (Get-Date).Date

        foreach ($rawResult in $rawResults) {
            if ($rawResult.Status -like 'ERROR:*') {
                $siteErrorCount++
                Write-Verbose "Sharing link collection failed for $($rawResult.SiteUrl): $($rawResult.Status)"
                continue
            }

            if ($rawResult.Status -like 'PARTIAL:*') {
                $sitePartialErrorCount++
                Write-Verbose "Sharing link collection partially failed for $($rawResult.SiteUrl): $($rawResult.Status)"
            }

            foreach ($link in $rawResult.Links) {
                if ($IncludeGraphDetails.IsPresent) {
                    # Graph mode: the scope is authoritative. An unknown value is surfaced as 'Other' with
                    # its raw form kept in LinkScope, so a new SharePoint link type shows up in the report
                    # instead of being silently misfiled.
                    $friendlyLinkType = switch ($link.Scope) {
                        'anonymous' { 'Anyone' }
                        'organization' { 'Company' }
                        'users' { 'SpecificPeople' }
                        default { 'Other' }
                    }
                    $rawLinkType = $link.Scope
                }
                else {
                    # Default mode: the type comes from the group name, which is all SharePoint exposes
                    # without Graph. Flexible is reported as such and NOT mapped to 'SpecificPeople': that
                    # link type carries its scope per link, so calling it anything else would be a guess -
                    # and a guess that reads as 'no anonymous link here' is exactly the wrong way to be
                    # wrong in an audit. Use IncludeGraphDetails to resolve those.
                    $friendlyLinkType = switch -Wildcard ($link.TypeToken) {
                        'Anonymous*' { 'Anyone' }
                        'Organization*' { 'Company' }
                        'Flexible' { 'Flexible' }
                        default { 'Other' }
                    }
                    $rawLinkType = $link.TypeToken
                }

                if ($LinkType -and $friendlyLinkType -notin $LinkType) {
                    continue
                }

                $expirationDate = $null
                $daysToExpiry = $null
                $linkStatus = 'Active'

                if ($link.ExpirationDateTime) {
                    $expirationDate = ([datetime]$link.ExpirationDateTime).ToLocalTime()
                    $daysToExpiry = (New-TimeSpan -Start $currentDate -End $expirationDate).Days
                    if ($expirationDate -lt $currentDate) {
                        $linkStatus = 'Expired'
                    }
                }

                if ($ExpiredOnly.IsPresent -and $linkStatus -ne 'Expired') {
                    continue
                }

                if ($NeverExpiringOnly.IsPresent -and $null -ne $expirationDate) {
                    continue
                }

                if ($ExpiringInDays -gt 0) {
                    if (($null -eq $expirationDate) -or ($linkStatus -eq 'Expired') -or ($daysToExpiry -gt $ExpiringInDays)) {
                        continue
                    }
                }

                # Same rule as Get-SPOSiteReport: a column that was not collected is absent, never present
                # and empty. An empty ExpirationDateTime would read as "this link never expires", which is
                # the opposite of "nobody asked Graph".
                $linkProperties = [ordered]@{
                    SiteUrl      = $rawResult.SiteUrl
                    ItemType     = $link.ItemType
                    ItemName     = $link.ItemName
                    ItemUrl      = $link.ItemUrl
                    LinkType     = $friendlyLinkType
                    LinkTypeRaw  = $rawLinkType
                    SharedWith   = $link.SharedWith
                }

                if ($IncludeGraphDetails.IsPresent) {
                    $linkProperties['Permission'] = $link.Permission
                    $linkProperties['Roles'] = $link.Roles
                    $linkProperties['LinkStatus'] = $linkStatus
                    $linkProperties['ExpirationDateTime'] = $expirationDate
                    $linkProperties['DaysToExpiry'] = $daysToExpiry
                    $linkProperties['HasPassword'] = $link.HasPassword
                    $linkProperties['PreventsDownload'] = $link.PreventsDownload
                    $linkProperties['LinkUrl'] = $link.LinkUrl
                    $linkProperties['LinkId'] = $link.LinkId
                }
                else {
                    $linkProperties['SharePointGroup'] = $link.GroupTitle
                    $linkProperties['LinkGuid'] = $link.LinkGuid
                }

                $sharingLinksArray.Add([PSCustomObject]$linkProperties)
            }
        }

        if ($siteErrorCount -gt 0) {
            Write-Host -ForegroundColor Yellow "$siteErrorCount site(s) could not be scanned at all. Run with -Verbose for the reason of each one."
            Write-Host -ForegroundColor Yellow 'A recurring error usually means the app registration lacks the SharePoint Sites.FullControl.All application permission.'
        }

        if ($sitePartialErrorCount -gt 0) {
            Write-Host -ForegroundColor Yellow "$sitePartialErrorCount site(s) returned partial results (unresolvable items, or links that could not be read). Run with -Verbose for details."

            if ($IncludeGraphDetails.IsPresent) {
                Write-Host -ForegroundColor Yellow 'A Graph error on every item points at the app registration: IncludeGraphDetails reads the links through Microsoft Graph, a DIFFERENT Entra resource from the SharePoint Sites.FullControl.All permission used everywhere else.'
                Write-Host -ForegroundColor Yellow "  'accessDenied 403' : the Graph token is valid but lacks the permission - add the Microsoft Graph APPLICATION permissions Files.Read.All and Sites.Read.All, then grant admin consent."
                Write-Host -ForegroundColor Yellow "  'Either scp or roles claim need to be present in the token' : the certificate carries no Graph permission at all - same fix."
                Write-Host -ForegroundColor Yellow '  Allow a few minutes for the consent to propagate and reconnect: PnP caches the acquired token. Dropping IncludeGraphDetails also works, at the cost of the expiration/URL/password columns.'
            }
            else {
                Write-Host -ForegroundColor Yellow '  Unresolvable items are usually orphaned SharingLinks groups left behind by deleted files - they grant nothing, but they do inflate the per-site counts.'
            }
        }

        Write-Host -ForegroundColor Green "Found $($sharingLinksArray.Count) sharing link(s) across $($siteUrlsList.Count) site(s)."

        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
            $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-SPOSharingLinkReport.xlsx"

            Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
            $sharingLinksArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'SPO-SharingLinks' -TableStyle Light9
            Write-Host -ForegroundColor Green 'Export completed successfully!'
        }
        else {
            return $sharingLinksArray
        }
    }
}
