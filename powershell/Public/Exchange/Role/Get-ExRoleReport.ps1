<#
    .SYNOPSIS
    Reports on Exchange RBAC roles and their effective membership, including groups expanded recursively.

    .DESCRIPTION
    Produces a report of the membership of Exchange RBAC role groups.
    By default, the report contains only the roles that have at least one member.

    When a role member is itself a group (distribution group, mail-enabled security group,
    dynamic distribution group, or a nested role group), its members are resolved recursively
    and included in the report with DirectMember set to $false and MemberViaGroup set to the
    name of the group that is a direct member of the role. Circular group references are
    detected and skipped automatically.

    .PARAMETER OnPrem
    When specified, queries an on-premises Exchange server instead of Exchange Online.
    Group expansion is also performed for on-premises role groups.

    .OUTPUTS
    [System.Collections.Generic.List[Object]] containing PSCustomObject rows with the following properties:
    Role, MemberName, MemberDisplayName, MemberPrimarySMTPAddres, MemberIsDirSynced,
    MemberObjectID, MemberRecipientTypeDetails, RoleDescription, DirectMember, MemberViaGroup.

    .EXAMPLE
    Get-ExRoleReport

    Retrieves the Exchange RBAC role report for Exchange Online, including recursive group expansion.

    .EXAMPLE
    Get-ExRoleReport | Where-Object { $_.DirectMember -eq $false } | Format-Table Role, MemberName, MemberViaGroup

    Lists all users/objects resolved through group membership, showing which group is a direct member of the role.

    .EXAMPLE
    Get-ExRoleReport | Export-Csv -NoTypeInformation "$(Get-Date -Format yyyyMMdd)_adminRoles.csv" -Encoding UTF8

    Exports the full report (including group-expanded members) to a CSV file.

    .NOTES
    Requires ExchangeOnlineManagement module and an active Connect-ExchangeOnline session for Exchange Online.
    For on-premises Exchange, requires the Exchange Management Shell or the Exchange snap-in loaded.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExRoleReport
#>
function Get-ExRoleReport {
    [CmdletBinding()]
    param (
        [switch]$OnPrem
    )

    if (-not $OnPrem) {
        try {
            Import-Module ExchangeOnlineManagement -ErrorAction stop
        }
        catch {
            Write-Warning 'First, install the official Microsoft Import-Module ExchangeOnlineManagement module : Install-Module ExchangeOnlineManagement'
            return
        }

        if (-not (Get-ConnectionInformation | Where-Object { $_.ConnectionUri -eq 'https://outlook.office365.com' })) {
            Connect-ExchangeOnline
        }
    }

    # Recursively expands a group member and adds its individual members to $ResultList.
    # $VisitedGroups prevents infinite loops when circular group references exist.
    function Expand-ExGroupMember {
        param (
            [Parameter(Mandatory = $true)]
            $Member,
            [Parameter(Mandatory = $true)]
            [string]$ParentGroupName,
            [Parameter(Mandatory = $true)]
            [string]$RoleName,
            [Parameter(Mandatory = $true)]
            [string]$RoleDescription,
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.List[Object]]$ResultList,
            [Parameter(Mandatory = $true)]
            [System.Collections.Generic.HashSet[string]]$VisitedGroups,
            [Parameter(Mandatory = $true)]
            [bool]$IsOnPrem
        )

        # Use ExchangeObjectId when available, fall back to PrimarySmtpAddress as deduplication key
        $groupKey = if ($Member.ExchangeObjectId) { $Member.ExchangeObjectId.ToString() } else { $Member.PrimarySmtpAddress }
        if (-not $VisitedGroups.Add($groupKey)) {
            return
        }

        try {
            # Role groups are expanded via Get-RoleGroupMember; all other group types via Get-DistributionGroupMember
            if ($Member.RecipientTypeDetails -eq 'RoleGroup') {
                $subMembers = @(Get-RoleGroupMember -Identity $Member.ExchangeObjectId -ResultSize Unlimited -ErrorAction Stop)
            }
            else {
                $subMembers = @(Get-DistributionGroupMember -Identity $Member.PrimarySmtpAddress -ResultSize Unlimited -ErrorAction Stop)
            }
        }
        catch {
            Write-Warning "Could not expand group '$($Member.Name)': $($_.Exception.Message)"
            return
        }

        foreach ($subMember in $subMembers) {
            $object = [PSCustomObject][ordered]@{
                'Role'                       = $RoleName
                'MemberName'                 = $subMember.Name
                'MemberDisplayName'          = $subMember.DisplayName
                'MemberPrimarySMTPAddres'    = $subMember.PrimarySmtpAddress
                'MemberIsDirSynced'          = $subMember.IsDirSynced
                'MemberObjectID'             = $subMember.ExternalDirectoryObjectId
                'MemberRecipientTypeDetails' = $subMember.RecipientTypeDetails
                'RoleDescription'            = $RoleDescription
                'DirectMember'               = $false
                'MemberViaGroup'             = $ParentGroupName
            }

            $ResultList.Add($object)

            # Recurse into any nested group
            if ($subMember.RecipientTypeDetails -like '*Group*') {
                Expand-ExGroupMember -Member $subMember `
                    -ParentGroupName $ParentGroupName `
                    -RoleName $RoleName `
                    -RoleDescription $RoleDescription `
                    -ResultList $ResultList `
                    -VisitedGroups $VisitedGroups `
                    -IsOnPrem $IsOnPrem
            }
        }
    }

    try {
        if ($OnPrem) {
            $exchangeRoles = Get-RoleGroup -ErrorAction Stop
        }
        else {
            # -ShowPartnerLinked (only for Exchange Online)
            # This `-ShowPartnerLinked` switch specifies whether to return built-in role groups that are of type PartnerRoleGroup. You don't need to specify a value with this switch.
            # This type of role group is used in the cloud-based service to allow partner service providers to manage their customer organizations.
            # These types of role groups can't be edited and are not shown by default.
            $exchangeRoles = Get-RoleGroup -ShowPartnerLinked -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Unable to retrieve Exchange RBAC roles. $($_.Exception.Message)"
    }

    [System.Collections.Generic.List[Object]]$exchangeRolesMembership = @()
    foreach ($exchangeRole in $exchangeRoles) {        
        try {
            $roleMembers = @(Get-RoleGroupMember -Identity $exchangeRole.ExchangeObjectId -ResultSize Unlimited)

            # Add green color if member found into the role
            if ($roleMembers.count -gt 0) {
                Write-Host -ForegroundColor Green "Role $($exchangeRole.Name) - Member(s) found: $($roleMembers.count)"
            }
            else {
                Write-Host -ForegroundColor Cyan "Role $($exchangeRole.Name) - Member found: $($roleMembers.count)"
            }

            if ($exchangeRole.Description -eq '' -and $exchangeRole.Name -like 'ISVMailboxUsers_*') {
                $roleDescription = 'Third-party application developer mailbox role'
            }
            else {
                $roleDescription = $exchangeRole.Description
            }

            if ($roleMembers.count -eq 0) {
                $object = [PSCustomObject][ordered]@{
                    'Role'                       = $exchangeRole.Name
                    'MemberName'                 = '-'
                    'MemberDisplayName'          = '-'
                    'MemberPrimarySMTPAddres'    = '-'
                    'MemberIsDirSynced'          = '-'
                    'MemberObjectID'             = '-'
                    'MemberRecipientTypeDetails' = '-'
                    'RoleDescription'            = $roleDescription
                    'DirectMember'               = '-'
                    'MemberViaGroup'             = '-'
                }
                
                $exchangeRolesMembership.Add($object)

            }
            else {         
                # Fresh set per role to allow the same group to appear in multiple roles
                $visitedGroups = [System.Collections.Generic.HashSet[string]]::new()

                foreach ($roleMember in $roleMembers) {
                    $object = [PSCustomObject][ordered]@{
                        'Role'                       = $exchangeRole.Name
                        'MemberName'                 = $roleMember.Name
                        'MemberDisplayName'          = $roleMember.DisplayName
                        'MemberPrimarySMTPAddres'    = $roleMember.PrimarySmtpAddress
                        'MemberIsDirSynced'          = $roleMember.IsDirSynced
                        'MemberObjectID'             = $roleMember.ExternalDirectoryObjectId
                        'MemberRecipientTypeDetails' = $roleMember.RecipientTypeDetails
                        'RoleDescription'            = $roleDescription
                        'DirectMember'               = $true
                        'MemberViaGroup'             = '-'
                    }

                    $exchangeRolesMembership.Add($object)

                    # Expand group members recursively
                    if ($roleMember.RecipientTypeDetails -like '*Group*') {
                        Expand-ExGroupMember -Member $roleMember `
                            -ParentGroupName $roleMember.Name `
                            -RoleName $exchangeRole.Name `
                            -RoleDescription $roleDescription `
                            -ResultList $exchangeRolesMembership `
                            -VisitedGroups $visitedGroups `
                            -IsOnPrem ([bool]$OnPrem)
                    }
                }
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }

    return $exchangeRolesMembership
}