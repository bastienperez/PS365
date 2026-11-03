<#
    .SYNOPSIS
    Retrieves the dynamic administrative units from Entra ID with membership rule analysis.

    .DESCRIPTION
    The Get-MgDynamicAdminUnit function retrieves the Entra ID administrative units whose membership
    is dynamic, with their membership rule, the user and device attributes the rule reads and the
    member count. It flags the administrative units whose rule uses the deprecated MemberOf rule
    operator and lists the source groups that rule reads.

    By default, only the member count is returned (MembersCount). Use -IncludeMembers to also get
    the concatenated lists of member names and IDs, or -MemberReport to get one row per
    (administrative unit, member).

    Dynamic groups are covered by Get-DynamicGroup.

    .PARAMETER AdminUnitId
    When specified, retrieves only the dynamic administrative unit with this object ID.
    If not specified, retrieves all dynamic administrative units.

    .PARAMETER DeprecatedMemberOfRuleOnly
    When specified, returns only the administrative units whose membership rule uses the deprecated
    MemberOf rule operator (e.g. user.memberof -any (group.objectId -in [...])). Microsoft ends this
    preview operator: after November 3, 2026, affected memberships stop updating and stay in their
    last known state.
    Without this switch, the DeprecatedMemberOfRule property still flags affected administrative units.

    .PARAMETER IncludeMembers
    When specified, populates the MembersName (pipe-separated display names) and MembersId
    (pipe-separated object IDs) properties in addition to MembersCount.
    By default, only MembersCount is populated to avoid slow enumeration on large administrative units.
    Ignored when -MemberReport is used.

    .PARAMETER MemberReport
    When specified, the function returns one row per (administrative unit, member) pair instead of
    the per-administrative unit objects.
    Output columns: AdminUnitId, AdminUnitName, MembershipRule, MemberId, MemberDisplayName,
    MemberPrincipal, MemberType.

    .PARAMETER ExportToExcel
    When specified, exports the result to an Excel file in the user's profile directory instead of
    returning it. When combined with -MemberReport, the member rows are written to a second worksheet
    named "Members" in the same workbook.

    .PARAMETER ExportPath
    Optional output directory for the Excel export (defaults to the user profile).

    .EXAMPLE
    Get-MgDynamicAdminUnit

    Retrieves all dynamic administrative units (members count only).

    .EXAMPLE
    Get-MgDynamicAdminUnit -DeprecatedMemberOfRuleOnly

    Returns only the dynamic administrative units using the deprecated MemberOf rule operator, whose
    memberships stop updating after November 3, 2026.

    .EXAMPLE
    Get-MgDynamicAdminUnit -DeprecatedMemberOfRuleOnly -ExportToExcel

    Exports the dynamic administrative units using the deprecated MemberOf rule operator to an Excel file.

    .EXAMPLE
    Get-MgDynamicAdminUnit -AdminUnitId '00000000-0000-0000-0000-000000000000' -IncludeMembers

    Retrieves one dynamic administrative unit with MembersName and MembersId populated (pipe-separated).

    .EXAMPLE
    Get-MgDynamicAdminUnit -MemberReport -ExportToExcel

    Exports both the per-administrative unit sheet ("DynamicAdminUnits") and the per-member sheet
    ("Members") in the same Excel workbook.

    .OUTPUTS
    System.Collections.Generic.List[PSCustomObject]

    .NOTES
    OUTPUT PROPERTIES (default / -IncludeMembers mode)
    - AdminUnitId: Unique identifier of the administrative unit
    - Name: Display name of the administrative unit
    - Type: Dynamic User or Dynamic Device, from the objects the membership rule reads
    - MembershipRule: The membership rule, on one line
    - MembershipRuleProcessingState: On or Paused
    - UserAttributes: Pipe-separated list of user attributes referenced in the membership rule
    - DeviceAttributes: Pipe-separated list of device attributes referenced in the membership rule
    - SourceGroupIds: Pipe-separated list of the group IDs a MemberOf rule reads (empty otherwise)
    - MembersCount: Number of current members of the administrative unit
    - MembersName: Pipe-separated list of member display names (only with -IncludeMembers)
    - MembersId: Pipe-separated list of member IDs (only with -IncludeMembers)
    - Description, Visibility, IsMemberManagementRestricted
    - DeprecatedMemberOfRule: True if the membership rule uses the deprecated MemberOf rule operator
      (preview ending: memberships stop updating after November 3, 2026)
    - Warning: Deprecation warning if the administrative unit uses the MemberOf rule operator

    OUTPUT PROPERTIES (-MemberReport mode)
    - AdminUnitId, AdminUnitName, MembershipRule
    - MemberId, MemberDisplayName, MemberPrincipal, MemberType

    The members of an administrative unit with hidden membership (Visibility HiddenMembership) can
    only be listed with the Member.Read.Hidden scope.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgDynamicAdminUnit
#>

function Get-MgDynamicAdminUnit {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$AdminUnitId,

        [Parameter(Mandatory = $false)]
        [switch]$DeprecatedMemberOfRuleOnly,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeMembers,

        [Parameter(Mandatory = $false)]
        [switch]$MemberReport,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath
    )

    # The members of a dynamic administrative unit are users or devices: without a scope that
    # reads them, Graph returns only their id and type
    $requiredScopes = @('AdministrativeUnit.Read.All|Directory.Read.All')
    if ($IncludeMembers.IsPresent -or $MemberReport.IsPresent) {
        $requiredScopes += 'User.ReadBasic.All|Directory.Read.All'
        $requiredScopes += 'Device.Read.All|Directory.Read.All'
    }
    if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    if ($MemberReport.IsPresent) {
        Write-Verbose 'MemberReport mode: returning one row per (administrative unit, member).'
        [System.Collections.Generic.List[PSCustomObject]]$memberReportArray = @()
    }
    elseif (-not $IncludeMembers.IsPresent) {
        Write-Warning 'Members enumeration is disabled by default (only MembersCount is populated). Use -IncludeMembers to also retrieve MembersName and MembersId, or -MemberReport to get one row per member. This can be slow on large administrative units.'
    }

    # Deprecated MemberOf rule operator (preview ending after November 3, 2026), same pattern as Get-DynamicGroup
    $deprecatedMemberOfRulePattern = '(?i)\b(user|device)\.memberof\b'
    $groupIdPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
    $selectProperties = 'id,displayName,description,membershipType,membershipRule,membershipRuleProcessingState,visibility,isMemberManagementRestricted'

    # Distinct attribute names after 'user.' or 'device.' in a membership rule
    $getRuleAttributes = {
        param([string]$Rule, [string]$Prefix)
        $attributes = foreach ($match in [regex]::Matches($Rule, "(?i)\b$Prefix\.([a-zA-Z][a-zA-Z0-9_]*)")) {
            $match.Groups[1].Value
        }
        ($attributes | Sort-Object -Unique) -join '|'
    }

    if (-not [string]::IsNullOrWhiteSpace($AdminUnitId)) {
        Write-Verbose "Retrieving the administrative unit $AdminUnitId..."
        $escapedId = [System.Uri]::EscapeDataString($AdminUnitId)
        try {
            $adminUnits = @(Invoke-MgGraphPagedRequest -Uri "/v1.0/directory/administrativeUnits/$($escapedId)?`$select=$selectProperties" -ErrorAction Stop)
        }
        catch {
            Write-Error "Error retrieving the administrative unit $AdminUnitId. $($_.Exception.Message)"
            return
        }
        if ($adminUnits[0].membershipType -ne 'Dynamic') {
            Write-Error "The administrative unit $AdminUnitId does not have a dynamic membership."
            return
        }
    }
    else {
        Write-Verbose 'Retrieving the administrative units...'
        try {
            $adminUnits = @(Invoke-MgGraphPagedRequest -Uri "/v1.0/directory/administrativeUnits?`$select=$selectProperties" -ErrorAction Stop)
        }
        catch {
            Write-Error "Error retrieving the administrative units. $($_.Exception.Message)"
            return
        }
        # membershipType has no documented $filter support: filter on the client side
        $adminUnits = @($adminUnits | Where-Object { $_.membershipType -eq 'Dynamic' })
        Write-Verbose "Found $($adminUnits.Count) dynamic administrative units"
    }

    # Filter early on the membership rule so the member calls only run on affected administrative units
    if ($DeprecatedMemberOfRuleOnly.IsPresent) {
        $adminUnits = @($adminUnits | Where-Object { $_.membershipRule -match $deprecatedMemberOfRulePattern })
        Write-Verbose "DeprecatedMemberOfRuleOnly: $($adminUnits.Count) dynamic administrative units use the deprecated MemberOf rule"
    }

    [System.Collections.Generic.List[PSCustomObject]]$adminUnitsArray = @()
    $counter = 0
    $total = $adminUnits.Count
    foreach ($adminUnit in $adminUnits) {
        $counter++
        Write-Verbose "[$counter/$total] Processing administrative unit: $($adminUnit.displayName)"

        $rule = [string]$adminUnit.membershipRule
        $usesMemberOf = $rule -match $deprecatedMemberOfRulePattern

        $type = $null
        if ($rule -match '(?i)\bdevice\.') {
            $type = 'Dynamic Device'
        }
        elseif ($rule -match '(?i)\buser\.') {
            $type = 'Dynamic User'
        }

        $sourceGroupIds = $null
        if ($usesMemberOf) {
            $groupIds = foreach ($match in [regex]::Matches($rule, $groupIdPattern)) {
                $match.Value.ToLower()
            }
            $sourceGroupIds = ($groupIds | Sort-Object -Unique) -join '|'
        }

        $warning = $null
        if ($usesMemberOf) {
            $warning = 'This administrative unit uses the deprecated MemberOf rule operator (preview ending). After November 3, 2026, membership stops updating and stays in its last known state. Replace the rule or convert the administrative unit to assigned membership. See https://learn.microsoft.com/entra/identity/users/groups-dynamic-rule-member-of'
        }

        $object = [PSCustomObject][ordered]@{
            AdminUnitId                   = $adminUnit.id
            Name                          = $adminUnit.displayName
            Type                          = $type
            # The rule can be multiline: one line, single spaces
            MembershipRule                = (($rule -replace '[\r\n]+', ' ' -replace '\s+', ' ') -replace '(\))(\w)', '$1 $2' -replace '(\w)(\()', '$1 $2' -replace '\(\s+', '(').Trim()
            MembershipRuleProcessingState = $adminUnit.membershipRuleProcessingState
            UserAttributes                = & $getRuleAttributes $rule 'user'
            DeviceAttributes              = & $getRuleAttributes $rule 'device'
            SourceGroupIds                = $sourceGroupIds
            MembersCount                  = $null
            MembersName                   = $null
            MembersId                     = $null
            Description                   = $adminUnit.description
            Visibility                    = $adminUnit.visibility
            IsMemberManagementRestricted  = [bool]$adminUnit.isMemberManagementRestricted
            DeprecatedMemberOfRule        = $usesMemberOf
            Warning                       = $warning
        }

        $membersUri = "/v1.0/directory/administrativeUnits/$($adminUnit.id)/members"
        try {
            if ($MemberReport.IsPresent -or $IncludeMembers.IsPresent) {
                $members = @(Invoke-MgGraphPagedRequest -Uri $membersUri -ErrorAction Stop)
                $object.MembersCount = $members.Count
                if ($MemberReport.IsPresent) {
                    foreach ($member in $members) {
                        $memberReportArray.Add([PSCustomObject][ordered]@{
                                AdminUnitId       = $adminUnit.id
                                AdminUnitName     = $adminUnit.displayName
                                MembershipRule    = $object.MembershipRule
                                MemberId          = $member.id
                                MemberDisplayName = $member.displayName
                                MemberPrincipal   = $member.userPrincipalName
                                MemberType        = ($member.'@odata.type' -replace '#microsoft\.graph\.', '')
                            })
                    }
                }
                else {
                    $object.MembersName = ($members | ForEach-Object { $_.displayName }) -join '|'
                    $object.MembersId = ($members | ForEach-Object { $_.id }) -join '|'
                }
            }
            else {
                try {
                    # Single lightweight $count call instead of paging through every member
                    $rawCount = Invoke-MgGraphRequest -Method GET -Uri "$membersUri/`$count" -Headers @{ ConsistencyLevel = 'eventual' } -OutputType Json -ErrorAction Stop
                    $object.MembersCount = [int]($rawCount -replace '[^\d]', '')
                }
                catch {
                    Write-Verbose "The member count call failed for $($adminUnit.displayName), counting the pages instead: $($_.Exception.Message)"
                    $object.MembersCount = @(Invoke-MgGraphPagedRequest -Uri $membersUri -ErrorAction Stop).Count
                }
            }
        }
        catch {
            Write-Warning "Could not read the members of $($adminUnit.displayName): $($_.Exception.Message)"
        }

        $adminUnitsArray.Add($object)
    }

    if ($ExportToExcel.IsPresent) {
        Write-Verbose 'Preparing Excel export...'
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $exportDirectory = if ($ExportPath) { $ExportPath } elseif ($env:USERPROFILE) { $env:USERPROFILE } else { $HOME }
        $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-DynamicAdminUnits.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting dynamic administrative units to Excel file: $excelFilePath"

        # Excel hard limit: 32,767 characters per cell. Also strip illegal XML control chars.
        $maxCellLength = 32700
        $illegalXmlChars = '[\x00-\x08\x0B\x0C\x0E-\x1F]'
        $sanitizeForExcel = {
            param($collection)
            foreach ($item in $collection) {
                foreach ($prop in $item.PSObject.Properties) {
                    if ($prop.Value -is [string]) {
                        $value = $prop.Value -replace $illegalXmlChars, ''
                        if ($value.Length -gt $maxCellLength) {
                            $value = $value.Substring(0, $maxCellLength) + '...[TRUNCATED]'
                        }
                        $prop.Value = $value
                    }
                }
            }
        }

        & $sanitizeForExcel $adminUnitsArray
        $adminUnitsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'DynamicAdminUnits' -TableStyle Light9
        if ($MemberReport.IsPresent) {
            & $sanitizeForExcel $memberReportArray
            $memberReportArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Members' -TableStyle Light9
        }
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif ($MemberReport.IsPresent) {
        Write-Verbose "Returning $($memberReportArray.Count) member rows"
        return $memberReportArray
    }
    else {
        Write-Verbose "Returning $($adminUnitsArray.Count) dynamic administrative units"
        return $adminUnitsArray
    }
}
