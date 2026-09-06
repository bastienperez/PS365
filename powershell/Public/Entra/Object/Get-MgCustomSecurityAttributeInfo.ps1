<#
    .SYNOPSIS
    Reports custom security attributes assigned to users, devices, and service principals (enterprise apps) in Microsoft Entra ID.

    .DESCRIPTION
    Queries Microsoft Graph to enumerate custom security attribute assignments across users, devices, and service principals.
    Auto-discovers all attribute sets in the tenant, or restricts the scope to a single set when -AttributeSet is provided.
    The output is one row per (entity, attribute set, attribute name, value) so it can be filtered/pivoted easily.

    .PARAMETER AttributeSet
    Restricts the report to a single attribute set name. If omitted, all attribute sets discovered in the tenant are reported.

    .PARAMETER EntityType
    Limits the entity types scanned. Valid values: User, Device, ServicePrincipal. Default is all three.

    .PARAMETER OnlyAssigned
    If specified, only entities that actually have at least one custom security attribute assignment are returned.
    This is the default behavior; the switch is kept for explicit/discoverable usage.

    .PARAMETER UnusedOnly
    Returns the attribute definitions that exist in the tenant but are assigned to no entity, instead of the
    assignments themselves. These are the cleanup candidates: a definition nobody uses still shows up in every
    attribute picker and still has to be governed.
    The scan is the same either way, so the count is always reported, whether or not this switch is present.
    The count is only trustworthy when every entity type was scanned without error. With a restricted -EntityType,
    or after an entity read failure, the result is reported as scoped: an attribute assigned only on a skipped
    entity type would otherwise look unused. A failure to read the definitions themselves is different: there is
    nothing to compare the assignments against, so the unused report is unavailable rather than scoped.

    .PARAMETER ForceNewToken
    Switch parameter to force getting a new token from Microsoft Graph.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-MgCustomSecurityAttributeInfo

    Auto-discovers all attribute sets and returns assignments across users, devices, and service principals.

    .EXAMPLE
    Get-MgCustomSecurityAttributeInfo -AttributeSet 'ComplianceData'

    Returns assignments only for the 'ComplianceData' attribute set.

    .EXAMPLE
    Get-MgCustomSecurityAttributeInfo -EntityType User, ServicePrincipal

    Returns assignments only for users and service principals (skips devices).

    .EXAMPLE
    Get-MgCustomSecurityAttributeInfo -UnusedOnly

    Returns only the attribute definitions assigned to nobody, with their set and their status.

    .EXAMPLE
    Get-MgCustomSecurityAttributeInfo -ExportToExcel

    Exports results to an Excel file in the user's profile directory, with one worksheet per entity type,
    plus an 'Unused' worksheet listing the definitions nobody uses.

    .NOTES
    Required Microsoft Graph permissions:
        - CustomSecAttributeDefinition.Read.All
        - CustomSecAttributeAssignment.Read.All
        - User.Read.All
        - Device.Read.All
        - Application.Read.All

    The custom security attribute on devices is in preview at the time of writing and uses the Graph beta endpoint.
    Reading customSecurityAttributes requires the caller to be granted the 'Attribute Assignment Reader' (or higher) directory role
    in addition to the application/delegated permissions above.


    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgCustomSecurityAttributeInfo
#>

function Get-MgCustomSecurityAttributeInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$AttributeSet,

        [Parameter(Mandatory = $false)]
        [ValidateSet('User', 'Device', 'ServicePrincipal')]
        [string[]]$EntityType = @('User', 'Device', 'ServicePrincipal'),

        [Parameter(Mandatory = $false)]
        [switch]$OnlyAssigned,

        [Parameter(Mandatory = $false)]
        [switch]$UnusedOnly,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath
    )

    # Import required modules
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
        'CustomSecAttributeDefinition.Read.All'
        'CustomSecAttributeAssignment.Read.All'
        'User.Read.All'
        'Device.Read.All'
        'Application.Read.All'
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

    # Discover attribute sets (used for filtering and to surface empty sets)
    Write-Host -ForegroundColor Cyan 'Retrieving attribute sets'
    try {
        $attributeSetsResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/directory/attributeSets' -OutputType PSObject
        $attributeSetsList = $attributeSetsResponse.value
    }
    catch {
        Write-Warning "Unable to retrieve attribute sets: $_"
        return
    }

    if ($AttributeSet) {
        $attributeSetsList = $attributeSetsList | Where-Object { $_.id -eq $AttributeSet }
        if (-not $attributeSetsList) {
            Write-Warning "Attribute set '$AttributeSet' not found in this tenant."
            return
        }
    }

    $attributeSetsAllowed = @{}
    foreach ($set in $attributeSetsList) {
        $attributeSetsAllowed[$set.id] = $true
    }

    Write-Host -ForegroundColor Cyan "Found $($attributeSetsList.Count) attribute set(s) to inspect"

    # Definitions are needed to report the attributes that carry no assignment.
    Write-Host -ForegroundColor Cyan 'Retrieving attribute definitions'
    [System.Collections.Generic.List[PSCustomObject]]$definitionsArray = @()
    $definitionsRetrieved = $true
    try {
        $definitionsUri = 'https://graph.microsoft.com/v1.0/directory/customSecurityAttributeDefinitions'
        do {
            $definitionsResponse = Invoke-MgGraphRequest -Method GET -Uri $definitionsUri -OutputType PSObject
            foreach ($definition in $definitionsResponse.value) {
                if (-not $attributeSetsAllowed.ContainsKey($definition.attributeSet)) { continue }
                $definitionsArray.Add($definition)
            }
            $definitionsUri = $definitionsResponse.'@odata.nextLink'
        } while ($definitionsUri)
    }
    catch {
        # Not fatal: the assignment report stays valid without it.
        $definitionsRetrieved = $false
        Write-Warning "Unable to retrieve attribute definitions, the unused report will be skipped: $_"
    }

    # An attribute is only unused if every entity type was scanned without error.
    # A restricted -EntityType, or a failed read, makes the answer scoped rather
    # than tenant-wide, and reporting it as unused would invite a wrong deletion.
    $scanIsComplete = -not @('User', 'Device', 'ServicePrincipal').Where({ $EntityType -notcontains $_ })

    [System.Collections.Generic.List[PSCustomObject]]$assignmentsArray = @()

    function Convert-CustomSecurityAttributesToRows {
        param(
            [Parameter(Mandatory = $true)] [string]$EntityType,
            [Parameter(Mandatory = $true)] [PSObject]$Entity,
            [Parameter(Mandatory = $true)] [hashtable]$AllowedSets
        )

        $rows = [System.Collections.Generic.List[PSCustomObject]]@()
        $csa = $Entity.customSecurityAttributes

        if ($null -eq $csa) {
            return $rows
        }

        foreach ($setProperty in $csa.PSObject.Properties) {
            $setName = $setProperty.Name
            if (-not $AllowedSets.ContainsKey($setName)) { continue }

            $setValues = $setProperty.Value
            if ($null -eq $setValues) { continue }

            foreach ($attrProperty in $setValues.PSObject.Properties) {
                # Skip OData annotations like '@odata.type' or '<attr>@odata.type'
                if ($attrProperty.Name -match '@odata') { continue }

                $value = $attrProperty.Value
                if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                    $valueText = ($value | ForEach-Object { "$_" }) -join '; '
                }
                else {
                    $valueText = "$value"
                }

                $row = [PSCustomObject][ordered]@{
                    EntityType        = $EntityType
                    DisplayName       = $Entity.displayName
                    Identifier        = if ($EntityType -eq 'User') { $Entity.userPrincipalName } elseif ($EntityType -eq 'ServicePrincipal') { $Entity.appId } else { $Entity.id }
                    ObjectId          = $Entity.id
                    OperatingSystem   = if ($EntityType -eq 'Device') { $Entity.operatingSystem } else { $null }
                    AttributeSet      = $setName
                    AttributeName     = $attrProperty.Name
                    AttributeValue    = $valueText
                }

                $rows.Add($row)
            }
        }

        return $rows
    }

    # Users
    if ($EntityType -contains 'User') {
        Write-Host -ForegroundColor Cyan 'Retrieving users with custom security attributes'
        $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,customSecurityAttributes&$count=true&$top=999'
        $headers = @{ ConsistencyLevel = 'eventual' }

        try {
            do {
                $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers -OutputType PSObject
                foreach ($user in $response.value) {
                    $rows = Convert-CustomSecurityAttributesToRows -EntityType 'User' -Entity $user -AllowedSets $attributeSetsAllowed
                    foreach ($row in $rows) { $assignmentsArray.Add($row) }
                }
                $uri = $response.'@odata.nextLink'
            } while ($uri)
        }
        catch {
            $scanIsComplete = $false
            Write-Warning "Unable to retrieve users: $_"
        }
    }

    # Service principals (enterprise apps)
    if ($EntityType -contains 'ServicePrincipal') {
        Write-Host -ForegroundColor Cyan 'Retrieving service principals with custom security attributes'
        $uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,displayName,appId,customSecurityAttributes&$count=true&$top=999'
        $headers = @{ ConsistencyLevel = 'eventual' }

        try {
            do {
                $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers -OutputType PSObject
                foreach ($sp in $response.value) {
                    $rows = Convert-CustomSecurityAttributesToRows -EntityType 'ServicePrincipal' -Entity $sp -AllowedSets $attributeSetsAllowed
                    foreach ($row in $rows) { $assignmentsArray.Add($row) }
                }
                $uri = $response.'@odata.nextLink'
            } while ($uri)
        }
        catch {
            $scanIsComplete = $false
            Write-Warning "Unable to retrieve service principals: $_"
        }
    }

    # Devices (beta endpoint - preview)
    if ($EntityType -contains 'Device') {
        Write-Host -ForegroundColor Cyan 'Retrieving devices with custom security attributes (beta endpoint)'
        $uri = 'https://graph.microsoft.com/beta/devices?$select=id,displayName,operatingSystem,customSecurityAttributes&$count=true&$top=999'
        $headers = @{ ConsistencyLevel = 'eventual' }

        try {
            do {
                $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers -OutputType PSObject
                foreach ($device in $response.value) {
                    $rows = Convert-CustomSecurityAttributesToRows -EntityType 'Device' -Entity $device -AllowedSets $attributeSetsAllowed
                    foreach ($row in $rows) { $assignmentsArray.Add($row) }
                }
                $uri = $response.'@odata.nextLink'
            } while ($uri)
        }
        catch {
            $scanIsComplete = $false
            Write-Warning "Unable to retrieve devices: $_"
        }
    }

    # Computed before the early return on an empty scan: zero assignment means
    # every definition is unused.
    [System.Collections.Generic.List[PSCustomObject]]$unusedArray = @()
    if ($definitionsArray.Count -gt 0) {
        $assignedKeys = @{}
        foreach ($assignment in $assignmentsArray) {
            $assignedKeys["$($assignment.AttributeSet)/$($assignment.AttributeName)"] = $true
        }

        foreach ($definition in $definitionsArray) {
            if ($assignedKeys.ContainsKey("$($definition.attributeSet)/$($definition.name)")) { continue }

            $unusedArray.Add([PSCustomObject][ordered]@{
                    AttributeSet  = $definition.attributeSet
                    AttributeName = $definition.name
                    Status        = $definition.status
                    Type          = $definition.type
                    IsCollection  = $definition.isCollection
                    Description   = $definition.description
                })
        }
    }

    if (-not $definitionsRetrieved) {
        # Without the definitions there is nothing to subtract from, so silence is the
        # only honest answer: an empty list here would read as "nothing unused".
        Write-Warning 'The unused report is unavailable because the attribute definitions could not be read.'
        if ($UnusedOnly.IsPresent) {
            return
        }
    }
    elseif ($scanIsComplete) {
        Write-Host -ForegroundColor $(if ($unusedArray.Count -gt 0) { 'Yellow' } else { 'Green' }) "$($unusedArray.Count) attribute definition(s) out of $($definitionsArray.Count) are assigned to nobody.$(if ($unusedArray.Count -gt 0 -and -not $UnusedOnly.IsPresent) { ' Use -UnusedOnly to list them.' })"
    }
    else {
        Write-Warning "$($unusedArray.Count) attribute definition(s) out of $($definitionsArray.Count) carry no assignment within the scanned scope ($($EntityType -join ', ')). The scan is partial, so these are candidates and not confirmed unused: an attribute assigned only on an entity type that was skipped or failed would appear here."
    }

    if ($UnusedOnly.IsPresent) {
        if ($unusedArray.Count -eq 0) {
            Write-Host -ForegroundColor Green 'Every attribute definition is assigned to at least one entity.'
            return
        }
    }
    elseif ($assignmentsArray.Count -eq 0) {
        Write-Host -ForegroundColor Yellow 'No entities found with custom security attributes for the requested scope.'
        if (-not ($ExportToExcel.IsPresent -and $unusedArray.Count -gt 0)) {
            return
        }
    }
    else {
        Write-Host -ForegroundColor Green "Found $($assignmentsArray.Count) attribute assignment(s)."
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$(if ($ExportPath) { $ExportPath } else { $env:userprofile })\$now-MgCustomSecurityAttributeInfo.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting custom security attribute report to Excel file: $excelFilePath"

        # One worksheet per entity type, plus a consolidated 'All' sheet. Skipped in
        # -UnusedOnly mode, where the assignments are not what the caller asked for.
        if (-not $UnusedOnly.IsPresent -and $assignmentsArray.Count -gt 0) {
            $assignmentsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-CustomSecAttr-All' -TableStyle Light9

            foreach ($type in ($assignmentsArray.EntityType | Sort-Object -Unique)) {
                $sheetName = "Entra-CustomSecAttr-$type"
                $assignmentsArray | Where-Object { $_.EntityType -eq $type } | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName $sheetName -TableStyle Light9
            }
        }

        if ($unusedArray.Count -gt 0) {
            $unusedArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-CustomSecAttr-Unused' -TableStyle Light9
        }

        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif ($UnusedOnly.IsPresent) {
        return $unusedArray
    }
    else {
        return $assignmentsArray
    }
}
