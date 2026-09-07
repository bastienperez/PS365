<#
    .SYNOPSIS
    Inventories the custom attributes usable in Entra ID dynamic group rules: the 15 built-in extension attributes and the directory extensions.

    .DESCRIPTION
    Returns one row per custom attribute available in the tenant, whether it carries a value or not, so the
    inventory is exhaustive.

    Two mechanisms are covered:
    - the 15 built-in extension attributes exposed through onPremisesExtensionAttributes (extensionAttribute1..15);
    - the directory extensions declared by an application, named extension_<AppId without dashes>_<Name>.

    Custom security attributes are deliberately out of scope: they cannot be used in dynamic membership rules.
    Use Get-MgCustomSecurityAttributeInfo for those.

    For every attribute the function reports whether it is populated, how many objects carry a value, and which
    dynamic groups reference it. It also flags the two situations that are hard to see from the portal:

    - Orphaned: the application that declared the directory extension no longer exists. The values remain in the
      directory and keep driving dynamic group membership, but the attribute can no longer be listed or corrected,
      and recreating the application does not help since the new AppId differs.
    - Referenced but empty: a dynamic group rule points at an attribute no object carries, so the group stays empty.

    .PARAMETER TargetObject
    Object types to count values on. Valid values: User, Group, Device. Default is User only. Adding Device also
    counts the built-in extension attributes on devices, which carry them under a different property name and can
    be addressed by a device membership rule.

    .PARAMETER ExcludeBuiltIn
    Leaves the 15 built-in extension attributes out of the inventory and reports directory extensions only.

    .PARAMETER SkipUsageCount
    Skips counting the objects carrying a value. The count is one Graph query per attribute and per object type,
    which is the slow part of the scan on a large tenant.

    .PARAMETER SkipDynamicGroups
    Skips reading the dynamic group membership rules, so the DynamicGroups columns are left empty.

    .PARAMETER ForceNewToken
    Forces a new token to be requested from Microsoft Graph.

    .PARAMETER ExportToExcel
    Exports the result to an Excel file in the user's profile directory instead of returning it.

    .PARAMETER ExportPath
    Optional output directory for the Excel export. Defaults to the user profile.

    .EXAMPLE
    Get-MgExtensionAttributeInfo

    Returns every built-in extension attribute and every directory extension, with their usage and their status.

    .EXAMPLE
    Get-MgExtensionAttributeInfo -TargetObject User, Group, Device

    Counts the objects carrying a value on users, groups and devices instead of users only.

    .EXAMPLE
    Get-MgExtensionAttributeInfo -SkipUsageCount

    Lists the attributes without counting values. Fast inventory on a large tenant.

    .EXAMPLE
    Get-MgExtensionAttributeInfo -ExportToExcel

    Exports the inventory to an Excel file in the user's profile directory.

    .NOTES
    Required Microsoft Graph permissions:
        - Application.Read.All
        - Directory.Read.All
        - Group.Read.All
        - User.Read.All

    The value counts use advanced queries ($count with ConsistencyLevel eventual). When Graph refuses a filter on a
    given attribute, the count is left null and the reason is reported in the CountError column rather than failing
    the whole scan. A count that failed on one object type is never reported as a partial total: an attribute
    counted at zero on users and unreadable on devices is left Unknown rather than offered for cleanup.

    Definitions come from getAvailableExtensionProperties rather than from a walk of the applications collection.
    That is the only supported way to see an extension whose declaring application has been deleted, since Graph
    exposes no extensionProperties navigation on the recycle bin.

    Dynamic group rules are matched on both the user. and device. prefixes: extension attributes and directory
    extensions can be addressed either way depending on the type of group.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgExtensionAttributeInfo
#>

function Get-MgExtensionAttributeInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [ValidateSet('User', 'Group', 'Device')]
        [string[]]$TargetObject = @('User'),

        [Parameter(Mandatory = $false)]
        [switch]$ExcludeBuiltIn,

        [Parameter(Mandatory = $false)]
        [switch]$SkipUsageCount,

        [Parameter(Mandatory = $false)]
        [switch]$SkipDynamicGroups,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath
    )

    $modules = @(
        'Microsoft.Graph.Authentication'
    )

    foreach ($module in $modules) {
        try {
            $null = Import-Module $module -ErrorAction Stop
        }
        catch {
            Write-Warning "Please install $module first"
            return
        }
    }

    $permissionsNeeded = @(
        'Application.Read.All'
        'Directory.Read.All'
        'Group.Read.All'
        'User.Read.All'
    )

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
    if ($ForceNewToken.IsPresent) {
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        $isConnected = $false
    }
    if (-not $isConnected) {
        Write-Host -ForegroundColor Cyan 'Connecting to Microsoft Graph'
        $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
    }

    if (-not (Test-MgGraphPermission -RequiredScopes $permissionsNeeded -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    $collectionByObject = @{
        User   = 'users'
        Group  = 'groups'
        Device = 'devices'
    }

    function Get-GraphCollection {
        param(
            [Parameter(Mandatory = $true)] [string]$Uri
        )

        $items = [System.Collections.Generic.List[PSCustomObject]]@()
        $next = $Uri

        do {
            # Through the retry wrapper: a throttled page would land in a catch block that carries
            # on with a partial list, and a directory extension missing from an inventory reads as
            # an extension that does not exist.
            $response = Invoke-MgGraphRequestWithRetry -Method GET -Uri $next -OutputType PSObject
            foreach ($item in $response.value) { $items.Add($item) }
            $next = $response.'@odata.nextLink'
        } while ($next)

        return $items
    }

    # Directory extension definitions are read with getAvailableExtensionProperties, which returns
    # every definition registered in the tenant, including those declared by an application that has
    # since been deleted. Walking the applications collection cannot reach those: the application is
    # gone from that collection, and Graph exposes no extensionProperties navigation on the recycle
    # bin. The declaring application is identified by the AppId embedded in the attribute name.
    Write-Host -ForegroundColor Cyan 'Retrieving directory extension definitions'
    $availableExtensions = @()
    try {
        $response = Invoke-MgGraphRequestWithRetry -Method POST -Uri 'https://graph.microsoft.com/v1.0/directoryObjects/getAvailableExtensionProperties' -Body '{}' -ContentType 'application/json' -OutputType PSObject
        $availableExtensions = @($response.value)
    }
    catch {
        Write-Warning "Unable to retrieve the directory extension definitions: $_"
        return
    }

    # Applications are read only to name the owner of each extension and to tell a live application
    # from a missing one. The AppId is normalized the way it appears in an attribute name: no dashes.
    $applicationsByAppId = @{}
    try {
        $applications = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName&$top=999'
        foreach ($application in $applications) { $applicationsByAppId[($application.appId -replace '-', '').ToLowerInvariant()] = $application }
    }
    catch {
        Write-Warning "Unable to retrieve applications, every extension will be reported as orphaned: $_"
    }

    $deletedApplicationsByAppId = @{}
    try {
        $deletedApplications = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.application?$select=id,appId,displayName,deletedDateTime&$top=999'
        foreach ($application in $deletedApplications) { $deletedApplicationsByAppId[($application.appId -replace '-', '').ToLowerInvariant()] = $application }
    }
    catch {
        # Reading the recycle bin can be denied without breaking the rest: an extension whose
        # application is missing is still reported as orphaned, only the deletion date is lost.
        Write-Warning "Unable to read deleted applications, deletion dates will be missing: $_"
    }

    [System.Collections.Generic.List[PSCustomObject]]$extensionDefinitions = @()

    foreach ($property in $availableExtensions) {
        # extension_<AppId without dashes>_<Name>. An on-premises synced extension follows the same
        # shape; anything else is left with an unknown owner rather than guessed at.
        $ownerAppKey = if ($property.name -match '^extension_([0-9a-fA-F]{32})_') { $Matches[1].ToLowerInvariant() } else { $null }

        $ownerApplication = $null
        $isDeleted = $false
        if ($ownerAppKey) {
            if ($applicationsByAppId.ContainsKey($ownerAppKey)) {
                $ownerApplication = $applicationsByAppId[$ownerAppKey]
            }
            elseif ($deletedApplicationsByAppId.ContainsKey($ownerAppKey)) {
                $ownerApplication = $deletedApplicationsByAppId[$ownerAppKey]
                $isDeleted = $true
            }
            else {
                # Neither live nor in the recycle bin: the application is gone for good, which is the
                # worst case since the extension can no longer be maintained at all.
                $isDeleted = $true
            }
        }

        $extensionDefinitions.Add([PSCustomObject]@{
                Name           = $property.name
                DataType       = $property.dataType
                IsMultiValued  = $property.isMultiValued
                TargetObjects  = ($property.targetObjects -join ', ')
                AppDisplayName = if ($ownerApplication) { $ownerApplication.displayName } else { $property.appDisplayName }
                AppId          = if ($ownerApplication) { $ownerApplication.appId } else { $ownerAppKey }
                AppDeleted     = $isDeleted
                AppDeletedDate = if ($ownerApplication) { $ownerApplication.deletedDateTime } else { $null }
            })
    }

    Write-Host -ForegroundColor Cyan "Found $($extensionDefinitions.Count) directory extension(s)"

    # Membership rules are read once and searched per attribute afterwards: one pass over the
    # groups instead of one query per attribute.
    $dynamicGroups = @()
    if (-not $SkipDynamicGroups.IsPresent) {
        Write-Host -ForegroundColor Cyan 'Retrieving dynamic group membership rules'
        try {
            $dynamicGroups = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'DynamicMembership')&`$select=id,displayName,membershipRule&`$top=999"
            Write-Host -ForegroundColor Cyan "Found $($dynamicGroups.Count) dynamic group(s)"
        }
        catch {
            Write-Warning "Unable to retrieve dynamic groups, the DynamicGroups columns will be empty: $_"
        }
    }

    function Get-GroupsUsingAttribute {
        param(
            [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [object[]]$Groups,
            [Parameter(Mandatory = $true)] [string]$AttributeName
        )

        # A rule addresses the attribute as 'user.<name>' or, on a device group, as
        # 'device.<name>': both prefixes are supported for extension attributes and for
        # directory extensions. The trailing word boundary matters: a plain substring match
        # on 'user.extensionAttribute1' also hits a rule that only uses extensionAttribute15.
        $pattern = "(?:user|device)\.$([regex]::Escape($AttributeName))\b"
        return @($Groups | Where-Object { $_.membershipRule -and $_.membershipRule -match $pattern })
    }

    function Get-ValueCount {
        param(
            [Parameter(Mandatory = $true)] [string]$Collection,
            [Parameter(Mandatory = $true)] [string]$Filter
        )

        $uri = "https://graph.microsoft.com/v1.0/$Collection/`$count?`$filter=$([uri]::EscapeDataString($Filter))"
        # One count per attribute and per collection is where throttling actually bites on a large
        # tenant, so this is the call that most needs the backoff.
        $raw = Invoke-MgGraphRequestWithRetry -Method GET -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' } -OutputType Text
        return [int]$raw
    }

    [System.Collections.Generic.List[PSCustomObject]]$resultsArray = @()

    # Built-in extension attributes: they always exist, so the useful information is which of
    # the fifteen actually carry a value.
    if (-not $ExcludeBuiltIn.IsPresent) {
        Write-Host -ForegroundColor Cyan 'Inspecting the 15 built-in extension attributes'

        foreach ($index in 1..15) {
            $attributeName = "extensionAttribute$index"
            $usedByGroups = Get-GroupsUsingAttribute -Groups $dynamicGroups -AttributeName $attributeName

            $count = $null
            $countError = $null
            if (-not $SkipUsageCount.IsPresent) {
                $countIsComplete = $true

                # Users and devices hold the fifteen attributes under different property names, and
                # a membership rule can address either. Devices are counted only when asked for.
                $countSources = [ordered]@{ users = "onPremisesExtensionAttributes/$attributeName ne null" }
                if ($TargetObject -contains 'Device') {
                    $countSources['devices'] = "extensionAttributes/$attributeName ne null"
                }

                foreach ($collection in $countSources.Keys) {
                    try {
                        $count = [int]$count + (Get-ValueCount -Collection $collection -Filter $countSources[$collection])
                    }
                    catch {
                        $countError = $_.Exception.Message
                        $countIsComplete = $false
                    }
                }

                if (-not $countIsComplete) { $count = $null }
            }

            $resultsArray.Add([PSCustomObject][ordered]@{
                    Kind              = 'BuiltIn'
                    AttributeName     = $attributeName
                    FriendlyName      = $attributeName
                    TargetObjects     = if ($TargetObject -contains 'Device') { 'User, Device' } else { 'User' }
                    DataType          = 'String'
                    IsMultiValued     = $false
                    OwnerApp          = $null
                    OwnerAppId        = $null
                    OwnerAppDeleted   = $null
                    OwnerAppDeletedOn = $null
                    ObjectsWithValue  = $count
                    CountError        = $countError
                    DynamicGroupCount = $usedByGroups.Count
                    DynamicGroupNames = ($usedByGroups.displayName -join ' | ')
                    Status            = if ($usedByGroups.Count -gt 0 -and $count -eq 0) { 'ReferencedButEmpty' }
                    elseif ($count -gt 0) { 'InUse' }
                    elseif ($null -eq $count) { 'Unknown' }
                    else { 'Empty' }
                })
        }
    }

    # Directory extensions: here the attribute itself may be dead weight, or worse, orphaned.
    if ($extensionDefinitions.Count -gt 0) {
        Write-Host -ForegroundColor Cyan 'Inspecting directory extensions'
    }

    foreach ($definition in $extensionDefinitions) {
        $usedByGroups = Get-GroupsUsingAttribute -Groups $dynamicGroups -AttributeName $definition.Name

        $totalCount = $null
        $countError = $null
        if (-not $SkipUsageCount.IsPresent) {
            $countIsComplete = $true
            foreach ($object in $TargetObject) {
                # An extension declared for users only cannot be counted on devices.
                if ($definition.TargetObjects -and $definition.TargetObjects -notmatch $object) { continue }

                try {
                    $count = Get-ValueCount -Collection $collectionByObject[$object] -Filter "$($definition.Name) ne null"
                    $totalCount = [int]$totalCount + $count
                }
                catch {
                    $countError = $_.Exception.Message
                    $countIsComplete = $false
                }
            }

            # A partial total is worse than no total: an attribute counted at zero on users and
            # unreadable on devices would otherwise be reported as empty and offered for cleanup.
            if (-not $countIsComplete) { $totalCount = $null }
        }

        $resultsArray.Add([PSCustomObject][ordered]@{
                Kind              = 'DirectoryExtension'
                AttributeName     = $definition.Name
                FriendlyName      = ($definition.Name -split '_', 3)[-1]
                TargetObjects     = $definition.TargetObjects
                DataType          = $definition.DataType
                IsMultiValued     = $definition.IsMultiValued
                OwnerApp          = $definition.AppDisplayName
                OwnerAppId        = $definition.AppId
                OwnerAppDeleted   = $definition.AppDeleted
                OwnerAppDeletedOn = $definition.AppDeletedDate
                ObjectsWithValue  = $totalCount
                CountError        = $countError
                DynamicGroupCount = $usedByGroups.Count
                DynamicGroupNames = ($usedByGroups.displayName -join ' | ')
                Status            = if ($definition.AppDeleted) { 'Orphaned' }
                elseif ($usedByGroups.Count -gt 0 -and $totalCount -eq 0) { 'ReferencedButEmpty' }
                elseif ($totalCount -gt 0) { 'InUse' }
                elseif ($null -eq $totalCount) { 'Unknown' }
                else { 'Empty' }
            })
    }

    if ($resultsArray.Count -eq 0) {
        Write-Host -ForegroundColor Yellow 'No custom attribute found in this tenant.'
        return
    }

    $orphaned = @($resultsArray | Where-Object { $_.Status -eq 'Orphaned' })
    if ($orphaned.Count -gt 0) {
        Write-Host -ForegroundColor Yellow "$($orphaned.Count) directory extension(s) belong to a deleted application. Their values still apply but the attributes can no longer be maintained."
    }

    $referencedButEmpty = @($resultsArray | Where-Object { $_.Status -eq 'ReferencedButEmpty' })
    if ($referencedButEmpty.Count -gt 0) {
        Write-Host -ForegroundColor Yellow "$($referencedButEmpty.Count) attribute(s) are referenced by a dynamic group rule but carried by no object."
    }

    Write-Host -ForegroundColor Green "Found $($resultsArray.Count) custom attribute(s)."

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$(if ($ExportPath) { $ExportPath } else { $env:userprofile })\$now-MgExtensionAttributeInfo.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting extension attribute inventory to Excel file: $excelFilePath"

        $resultsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-ExtensionAttributes' -TableStyle Light9

        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $resultsArray
    }
}
