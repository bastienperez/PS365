<#

# TODO: Add Get-ManagementRoleAssignment to obtain permissions, as this is sometimes not done via a group
#Get-ManagementRoleAssignment -RoleAssigneeType User
    
    .SYNOPSIS
    Get-PurviewRoleReport - Reports on Purview RBAC roles and permissions.

    .DESCRIPTION 
    This script produces a report of the membership of Purview RBAC role groups.
    By default, the report contains only the groups with members.

    .OUTPUTS
    The report is output to an array contained all the audit logs found.
    To export in a csv, do Get-PurviewRoleReport | Export-CSV -NoTypeInformation "$(Get-Date -Format yyyyMMdd)_adminRoles.csv" -Encoding UTF8

    .EXAMPLE
    Get-PurviewRoleReport

    .LINK

    .NOTES
    Written by Bastien Perez (Clidsys.com - ITPro-Tips.com)
    For more Office 365/Microsoft 365 tips and news, check out ITPro-Tips.com.

    Version History:
    ## [1.0] - 2025-07-03
    ### Initial Release
#>
function Get-PurviewRoleReport {
    [CmdletBinding()]
    param (
        [switch]$ShowGraph
    )

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction stop
    }
    catch {
        Write-Warning 'First, install the official Microsoft Import-Module ExchangeOnlineManagement module : Install-Module ExchangeOnlineManagement'
        return
    }

    # if at least one result, we are connected
    if (-not (Get-ConnectionInformation | Where-Object { $_.ConnectionUri -like '*.compliance.protection.outlook.com' })) {
        Connect-IPPSSession
    }

    try {
        # -ShowPartnerLinked : 
        # This ShowPartnerLinked switch specifies whether to return built-in role groups that are of type PartnerRoleGroup. You don't need to specify a value with this switch.
        # This type of role group is used in the cloud-based service to allow partner service providers to manage their customer organizations.
        # These types of role groups can't be edited and are not shown by default.
        $purviewRoles = Get-RoleGroup -ShowPartnerLinked -ErrorAction Stop
    }
    catch {
        Write-Warning 'You are not connected to the Purview service. Please connect using Connect-IPPSSession.'
        return
    }

    [System.Collections.Generic.List[Object]]$purviewRolesMembership = @()

    foreach ($purviewRole in $purviewRoles) {
        try {
            # we need to use ExchangeObjectId instead of `Identity` or `DistinguishedName` otherwise we get the following error:
            # Get-RoleGroupMember: Ex9E65A2|Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException|The operation couldn't be performed because object:
            # 'FFO.extest.microsoft.com/Microsoft Exchange Hosted Organizations/xxx.onmicrosoft.com/Configuration/OrganizationManagement' matches multiple entries.
            $roleMembers = @(Get-RoleGroupMember -Identity $purviewRole.ExchangeObjectId -ResultSize Unlimited)

            # Add green color if member found into the role
            if ($roleMembers.count -gt 0) {
                Write-Host -ForegroundColor Green "Role $($purviewRole.Name) - Member(s) found: $($roleMembers.count)"
            }
            else {
                Write-Host -ForegroundColor Cyan "Role $($purviewRole.Name) - Member found: $($roleMembers.count)"
            }

            if ($purviewRole.Description -eq '' -and $purviewRole.Name -like 'ISVMailboxUsers_*') {
                $roleDescription = 'Third-party application developer mailbox role'
            }
            else {
                $roleDescription = $purviewRole.Description
            }

            if ($roleMembers.count -eq 0) {
                $object = [PSCustomObject][ordered]@{
                    'Role'                       = $purviewRole.Name
                    'MemberName'                 = '-'
                    'MemberDisplayName'          = '-'
                    'MemberPrimarySMTPAddres'    = '-'
                    'MemberIsDirSynced'          = '-'
                    'MemberObjectID'             = '-'
                    'MemberRecipientTypeDetails' = '-'
                    'RoleDescription'            = $roleDescription
                }

                $purviewRolesMembership.Add($object)

            }
            else {         

                foreach ($roleMember in $roleMembers) {                
                   
                    $object = [PSCustomObject][ordered]@{
                        'Role'                       = $purviewRole.Name
                        'MemberName'                 = $roleMember.Name
                        'MemberDisplayName'          = $roleMember.DisplayName
                        'MemberPrimarySMTPAddres'    = $roleMember.PrimarySmtpAddress
                        'MemberIsDirSynced'          = $roleMember.IsDirSynced
                        'MemberObjectID'             = $roleMember.ExternalDirectoryObjectId
                        'MemberRecipientTypeDetails' = $roleMember.RecipientTypeDetails
                        'RoleDescription'            = $roleDescription
                    }

                    $purviewRolesMembership.Add($object)
                }

            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }
    }

    return $purviewRolesMembership
}