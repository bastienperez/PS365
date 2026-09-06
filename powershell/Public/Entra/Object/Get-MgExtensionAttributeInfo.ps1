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
    Object types to count values on. Valid values: User, Group, Device. Default is User only, which is what the
    built-in extension attributes support.

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
    the whole scan.

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
            $response = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
            foreach ($item in $response.value) { $items.Add($item) }
            $next = $response.'@odata.nextLink'
        } while ($next)

        return $items
    }

    # Directory extensions are declared per application, so the applications have to be
    # enumerated to reach them. Deleted applications are read separately: their extensions
    # are the orphaned ones, and they are invisible from the applications collection.
    Write-Host -ForegroundColor Cyan 'Retrieving applications and their directory extensions'
    $applicationsById = @{}
    try {
        $applications = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/applications?$select=id,appId,displayName&$top=999'
        foreach ($application in $applications) { $applicationsById[$application.id] = $application }
    }
    catch {
        Write-Warning "Unable to retrieve applications: $_"
        return
    }

    $deletedApplicationsById = @{}
    try {
        $deletedApplications = Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.application?$select=id,appId,displayName,deletedDateTime&$top=999'
        foreach ($application in $deletedApplications) { $deletedApplicationsById[$application.id] = $application }
    }
    catch {
        # Reading the recycle bin can be denied without breaking the rest: an extension whose
        # application is missing is still reported as orphaned, only the deletion date is lost.
        Write-Warning "Unable to read deleted applications, deletion dates will be missing: $_"
    }

    [System.Collections.Generic.List[PSCustomObject]]$extensionDefinitions = @()

    foreach ($applicationId in ($applicationsById.Keys + $deletedApplicationsById.Keys)) {
        $isDeleted = $deletedApplicationsById.ContainsKey($applicationId)
        $application = if ($isDeleted) { $deletedApplicationsById[$applicationId] } else { $applicationsById[$applicationId] }

        try {
            $properties = Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/applications/$applicationId/extensionProperties"
        }
        catch {
            if (-not $isDeleted) {
                Write-Warning "Unable to read the extension properties of '$($application.displayName)': $_"
            }
            continue
        }

        foreach ($property in $properties) {
            $extensionDefinitions.Add([PSCustomObject]@{
                    Name           = $property.name
                    DataType       = $property.dataType
                    IsMultiValued  = $property.isMultiValued
                    TargetObjects  = ($property.targetObjects -join ', ')
                    AppDisplayName = $application.displayName
                    AppId          = $application.appId
                    AppDeleted     = $isDeleted
                    AppDeletedDate = $application.deletedDateTime
                })
        }
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

        # A rule references the attribute as 'user.<name>'. The trailing word boundary
        # matters: a plain substring match on 'user.extensionAttribute1' also hits a rule
        # that only uses extensionAttribute15.
        $pattern = "user\.$([regex]::Escape($AttributeName))\b"
        return @($Groups | Where-Object { $_.membershipRule -and $_.membershipRule -match $pattern })
    }

    function Get-ValueCount {
        param(
            [Parameter(Mandatory = $true)] [string]$Collection,
            [Parameter(Mandatory = $true)] [string]$Filter
        )

        $uri = "https://graph.microsoft.com/v1.0/$Collection/`$count?`$filter=$([uri]::EscapeDataString($Filter))"
        $raw = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' } -OutputType Text -ErrorAction Stop
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
                try {
                    $count = Get-ValueCount -Collection 'users' -Filter "onPremisesExtensionAttributes/$attributeName ne null"
                }
                catch {
                    $countError = $_.Exception.Message
                }
            }

            $resultsArray.Add([PSCustomObject][ordered]@{
                    Kind              = 'BuiltIn'
                    AttributeName     = $attributeName
                    FriendlyName      = $attributeName
                    TargetObjects     = 'User'
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
            foreach ($object in $TargetObject) {
                # An extension declared for users only cannot be counted on devices.
                if ($definition.TargetObjects -and $definition.TargetObjects -notmatch $object) { continue }

                try {
                    $count = Get-ValueCount -Collection $collectionByObject[$object] -Filter "$($definition.Name) ne null"
                    $totalCount = [int]$totalCount + $count
                }
                catch {
                    $countError = $_.Exception.Message
                }
            }
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
