<#
    .SYNOPSIS
    Retrieves the owners of all or a specific Entra ID group.

    .DESCRIPTION
    The Get-MgGroupOwnerInfo function queries Microsoft Entra ID via the Microsoft Graph API
    to return the owners of each group. Results can be filtered by a specific group ID
    or display name, and are exported to Excel by default.

    Groups with no owners are included in the output with an empty owner list, which is
    useful for identifying groups that lack an accountable owner.

    Note: This function is named Get-MgGroupOwnerInfo to avoid conflict with the built-in
    Microsoft Graph PowerShell cmdlet Get-MgGroupOwner.

    .PARAMETER GroupId
    Optional. The unique object ID of a specific Entra ID group to retrieve owners for.
    If omitted, all groups are processed.

    .PARAMETER DisplayName
    Optional. Filter groups by display name (exact match). Use with GroupId to narrow the scope.
    If omitted, all groups are processed.

    .PARAMETER ExportToExcel
    When specified, exports the results to an Excel file in the user's profile directory.
    Requires the ImportExcel module.

    .EXAMPLE
    Get-MgGroupOwnerInfo

    Retrieves the owners for all Entra ID groups and outputs them to the console.

    .EXAMPLE
    Get-MgGroupOwnerInfo -ExportToExcel

    Retrieves the owners of all groups and exports the results to an Excel file.

    .EXAMPLE
    Get-MgGroupOwnerInfo -GroupId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'

    Retrieves owners for a specific group by its object ID and exports to Excel.

    .EXAMPLE
    Get-MgGroupOwnerInfo -DisplayName 'Marketing Team'

    Retrieves owners for the group named 'Marketing Team' and outputs to the console only.

    .NOTES
    Requires Connect-MgGraph with scopes: 'Group.Read.All', 'User.Read.All'

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgGroupOwnerInfo
#>

function Get-MgGroupOwnerInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$GroupId,

        [Parameter(Mandatory = $false)]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )


    [System.Collections.Generic.List[PSCustomObject]]$results = @()
    $headers = @{ ConsistencyLevel = 'eventual' }

    # Fetch groups: all, by ID, or by display name
    [System.Collections.Generic.List[Object]]$groups = @()

    if ($GroupId) {
        Write-Host -ForegroundColor Cyan "Fetching group with ID: $GroupId"
        try {
            $group = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=id,displayName,groupTypes,securityEnabled,mailEnabled,mail" -ErrorAction Stop
            $groups.Add($group)
        }
        catch {
            Write-Error "Group with ID '$GroupId' not found."
            return
        }
    }
    else {
        Write-Host -ForegroundColor Cyan 'Fetching all groups...'

        $selectFields = 'id,displayName,groupTypes,securityEnabled,mailEnabled,mail'
        $filter = ''

        if ($DisplayName) {
            $filter = "&`$filter=displayName eq '$DisplayName'"
        }

        $uri = "https://graph.microsoft.com/v1.0/groups?`$select=$selectFields&`$top=999&`$count=true$filter"

        do {
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
            foreach ($g in $response.value) {
                $groups.Add($g)
            }
            $uri = $response.'@odata.nextLink'
        } while ($uri)
    }

    $totalCount = $groups.Count
    Write-Host -ForegroundColor Cyan "Found $totalCount group(s). Retrieving owners..."

    $processed = 0

    foreach ($group in $groups) {
        $processed++
        Write-Progress -Activity 'Retrieving group owners' -Status "$processed / $totalCount" -PercentComplete (($processed / $totalCount) * 100)

        # Determine group type
        $groupType = switch ($true) {
            { $group.groupTypes -contains 'Unified' -and $group.groupTypes -contains 'DynamicMembership' } { 'M365 Dynamic Group' }
            { $group.groupTypes -contains 'Unified' } { 'Microsoft 365 Group' }
            { $group.groupTypes -contains 'DynamicMembership' } { 'Dynamic Security Group' }
            { $group.mailEnabled -and -not $group.securityEnabled } { 'Distribution Group' }
            { $group.mailEnabled -and $group.securityEnabled } { 'Mail-enabled Security Group' }
            { $group.securityEnabled } { 'Security Group' }
            default { 'Other' }
        }

        # Retrieve owners for the group
        try {
            $ownersUri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/owners?`$select=id,displayName,userPrincipalName,mail"
            $ownerResponse = Invoke-MgGraphRequest -Method GET -Uri $ownersUri -ErrorAction Stop

            $owners = $ownerResponse.value

            if ($owners.Count -eq 0) {
                $object = [PSCustomObject][ordered]@{
                    GroupDisplayName      = $group.displayName
                    GroupId               = $group.id
                    GroupType             = $groupType
                    GroupMail             = $group.mail
                    OwnerDisplayName      = $null
                    OwnerUserPrincipalName = $null
                    OwnerId               = $null
                    OwnerMail             = $null
                    TotalOwners           = 0
                }
                $results.Add($object)
            }
            else {
                foreach ($owner in $owners) {
                    $object = [PSCustomObject][ordered]@{
                        GroupDisplayName      = $group.displayName
                        GroupId               = $group.id
                        GroupType             = $groupType
                        GroupMail             = $group.mail
                        OwnerDisplayName      = $owner.displayName
                        OwnerUserPrincipalName = $owner.userPrincipalName
                        OwnerId               = $owner.id
                        OwnerMail             = $owner.mail
                        TotalOwners           = $owners.Count
                    }
                    $results.Add($object)
                }
            }
        }
        catch {
            Write-Warning "Failed to retrieve owners for group '$($group.displayName)' ($($group.id)): $_"
        }
    }

    Write-Progress -Activity 'Retrieving group owners' -Completed
    Write-Host -ForegroundColor Green "Done. $($results.Count) row(s) returned for $totalCount group(s)."

    # Export to Excel if requested
    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgGroupOwnerInfo_Report.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $results | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-GroupOwners'
    }

    return $results
}
