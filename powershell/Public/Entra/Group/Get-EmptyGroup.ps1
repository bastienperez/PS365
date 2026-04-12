<#
    .SYNOPSIS
    Retrieves all empty groups (zero members) in Microsoft Entra ID.

    .DESCRIPTION
    The Get-EmptyGroup function retrieves all groups from Microsoft Entra ID via the
    Microsoft Graph API and identifies those with no members. It supports optional
    export to Excel or CSV for further analysis and cleanup.

    .PARAMETER ExportToExcel
    When specified, exports the results to an Excel file in the user's profile directory.
    Requires the ImportExcel module.

    .PARAMETER NoPermissionCheck
    (Optional) Skip the Microsoft Graph scope verification performed against the current Get-MgContext token.

    .EXAMPLE
    Get-EmptyGroup

    Retrieves all empty groups and outputs them to the console.

    .EXAMPLE
    Get-EmptyGroup -ExportToExcel

    Retrieves all empty groups and exports the results to an Excel file.

    .EXAMPLE
    Get-EmptyGroup | Export-Csv 'C:\temp\empty-groups.csv' -NoTypeInformation

    Retrieves all empty groups and pipes the results to Export-Csv.

    .OUTPUTS
    System.Collections.Generic.List[Object]

    .NOTES
    OUTPUT PROPERTIES
    Returns a collection of custom objects with the following properties:
    - DisplayName: Display name of the group
    - ObjectId: Unique identifier of the group
    - Type: Type of group (Microsoft 365, Dynamic, Mail-enabled Security, Security, Distribution, Other)
    - Mail: Primary email address of the group
    - Created: Creation date of the group (yyyy-MM-dd)
    - Description: Description of the group

    Requires Microsoft.Graph module: Connect-MgGraph -Scopes 'Group.Read.All'

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-EmptyGroup
#>

function Get-EmptyGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [switch]$NoPermissionCheck
    )

    if (-not $NoPermissionCheck.IsPresent) {
        $requiredScopes = @('Group.Read.All')
        if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }
    }

    Write-Verbose 'Fetching all groups...'

    # Fetch all groups with pagination
    [System.Collections.Generic.List[Object]]$allGroups = @()
    $uri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,groupTypes,securityEnabled,mailEnabled,mail,createdDateTime,description&`$top=999&`$count=true"
    $headers = @{ ConsistencyLevel = 'eventual' }

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $headers
        foreach ($group in $response.value) {
            $allGroups.Add($group)
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)

    $totalCount = $allGroups.Count
    Write-Verbose "Found $totalCount groups. Scanning for empty groups..."
    Write-Host -ForegroundColor Cyan "Found $totalCount groups. Scanning for empty groups..."

    # Check member count for each group
    [System.Collections.Generic.List[Object]]$emptyGroups = @()
    $processed = 0

    foreach ($group in $allGroups) {
        $processed++
        if ($processed % 50 -eq 0) {
            Write-Progress -Activity 'Scanning group members' -Status "$processed / $totalCount" -PercentComplete (($processed / $totalCount) * 100)
        }

        try {
            $membersUri = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members/`$count"
            $memberCount = Invoke-MgGraphRequest -Method GET -Uri $membersUri -Headers @{ ConsistencyLevel = 'eventual' }

            if ($memberCount -eq 0) {
                # Determine group type
                $type = if ($group.groupTypes -contains 'Unified') { 'Microsoft 365' }
                        elseif ($group.groupTypes -contains 'DynamicMembership') { 'Dynamic' }
                        elseif ($group.securityEnabled -and $group.mailEnabled) { 'Mail-enabled Security' }
                        elseif ($group.securityEnabled) { 'Security' }
                        elseif ($group.mailEnabled) { 'Distribution' }
                        else { 'Other' }

                $object = [PSCustomObject][ordered]@{
                    DisplayName = $group.displayName
                    ObjectId    = $group.id
                    Type        = $type
                    Mail        = $group.mail
                    Created     = if ($group.createdDateTime) { ([datetime]$group.createdDateTime).ToString('yyyy-MM-dd') } else { '' }
                    Description = $group.description
                }

                $emptyGroups.Add($object)
            }
        }
        catch {
            Write-Warning "Failed to check members for group '$($group.displayName)': $_"
        }
    }

    Write-Progress -Activity 'Scanning group members' -Completed

    Write-Host -ForegroundColor Yellow "`nEmpty Groups: $($emptyGroups.Count) / $totalCount total`n"

    if ($emptyGroups.Count -eq 0) {
        Write-Host -ForegroundColor Green 'No empty groups found.'
        return
    }

    if ($ExportToExcel.IsPresent) {
        Write-Verbose 'Preparing Excel export...'
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:USERPROFILE)\$now-EmptyGroups.xlsx"
        Write-Verbose "Excel file path: $excelFilePath"
        Write-Host -ForegroundColor Cyan "Exporting empty groups to Excel file: $excelFilePath"
        $emptyGroups | Sort-Object DisplayName | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EmptyGroups'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        Write-Verbose "Returning $($emptyGroups.Count) empty groups"
        return $emptyGroups
    }
}
