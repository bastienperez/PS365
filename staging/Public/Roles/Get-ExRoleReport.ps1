<#
# TODO: add MFA state
.SYNOPSIS
Get-ExchangeRoleReport - Reports on Exchange RBAC roles and permissions.

.DESCRIPTION 
This script produces a report of the membership of Exchange RBAC role groups.
By default, the report contains only the groups with members.

.OUTPUTS
The report is output to an array contained all the audit logs found.
To export in a csv, do Get-ExchangeRoleReport | Export-CSV -NoTypeInformation "$(Get-Date -Format yyyyMMdd)_adminRoles.csv" -Encoding UTF8

.EXAMPLE
Get-ExchangeRoleReport

.LINK

.NOTES
Written by Bastien Perez (ITPro-Tips.com)

Version history:
V1.0, 14 april 2022 - Initial version
v2.0 22 november 2024 - Add option to see permission graph - not work by now

Version History:
## [2.1] - 2025-07-03
### Added
- Add OnPrem switch to use Get-RoleGroup instead

### Changed
- Use another method to detect if we are connected to Exchange Online or not

## [2.0] - 2024-11-22
### Added
- Add option to see permissions graph - not work by now

## [1.0] - 2024-11-22
### Initial Release

#>
function Get-ExchangeRoleReport {
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
                }
                
                $exchangeRolesMembership.Add($object)

            }
            else {         

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
                    }

                    $exchangeRolesMembership.Add($object)
                }
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }

    # TODO: Add Get-ManagementRoleAssignment to obtain permissions, as this is sometimes not done via a group
    #Get-ManagementRoleAssignment -RoleAssigneeType User
    
    return $exchangeRolesMembership
}