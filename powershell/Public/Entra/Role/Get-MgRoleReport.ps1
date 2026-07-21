<#
    .SYNOPSIS
    Get-MgRoleReport.ps1 - Reports on Microsoft Entra ID (Azure AD) roles

    .DESCRIPTION 
    By default, the report contains only the roles with members.
    To get all the role, included empty roles, add -IncludeEmptyRoles $true

    .OUTPUTS
    The report is output to an array contained all the audit logs found.
    To export in a csv, do Get-MgRoleReport | Export-CSV -NoTypeInformation "$(Get-Date -Format yyyyMMdd)_adminRoles.csv" -Encoding UTF8

    .PARAMETER Identity
    Filter the report on a specific role. Accepts the role display name (e.g. 'Global Administrator') or the role definition Id (GUID).

    .PARAMETER PrincipalID
    Filter the report on a specific principal. Accepts the UPN (user), AppId (service principal) or ObjectId.

    .PARAMETER PrincipalDisplayName
    Filter the report on a specific principal display name (exact match, case-insensitive).

    .PARAMETER Scope
    Filter the report on the assignment scope (AssignedRoleScope / directoryScopeId).
    Examples: '/' (tenant-wide), '/administrativeUnits/<id>' (AU-scoped), or any resource scope.

    .PARAMETER TierLevel
    Filter the report on a privileged role tier: '0' (control plane), '1' (service/workload admins) or '2' (lower-privilege / read-mostly).
    Tiering is based on Sean Metcalf's (PyroTek3) classification. Regardless of this filter, every row is always annotated with a RoleTier property (Tier0/Tier1/Tier2/Untiered).

    .PARAMETER IncludeEmptyRoles
    Switch parameter to include empty roles in the report

    .PARAMETER ExcludePIMEligibleAssignments
    Switch parameter to exclude PIM eligible assignments from the report. Default is $false (includes them)

    .PARAMETER ForceNewToken
    Switch parameter to force getting a new token from Microsoft Graph

    .PARAMETER MaesterMode
    Switch parameter to use with the Maester framework (internal process not presented here)

    .PARAMETER ExportToExcel
    Switch parameter to export the report to an Excel file in the user's profile directory

    .EXAMPLE
    Get-MgRoleReport

    Get all the roles with members, including PIM eligible assignments but without empty roles

    .EXAMPLE
    Get-MgRoleReport -Identity 'Global Administrator'

    Returns only the assignments of the Global Administrator role (filter accepts both role name and roleDefinitionId).

    .EXAMPLE
    Get-MgRoleReport -PrincipalID 'alice@contoso.com'

    Returns only the role assignments of alice@contoso.com (direct or via group membership).

    .EXAMPLE
    Get-MgRoleReport -PrincipalID '11111111-2222-3333-4444-555555555555'

    Returns only the role assignments for the principal matching this ObjectId or AppId.

    .EXAMPLE
    Get-MgRoleReport -PrincipalDisplayName 'Alice Doe'

    Returns only the role assignments for the principal whose DisplayName is 'Alice Doe'.

    .EXAMPLE
    Get-MgRoleReport -Scope '/'

    Returns only the role assignments at tenant scope.

    .EXAMPLE
    Get-MgRoleReport -TierLevel 0

    Returns only the assignments of Tier 0 (control plane) roles. Each row carries a RoleTier property.

    .EXAMPLE
    Get-MgRoleReport -IncludeEmptyRoles

    Get all the roles, including the ones without members

    .EXAMPLE
    Get-MgRoleReport -ExcludePIMEligibleAssignments

    Get all the roles with members (without empty roles), but without PIM eligible assignments

    .EXAMPLE
    Get-MgRoleReport | Export-CSV -NoTypeInformation "$(Get-Date -Format yyyyMMdd)_adminRoles.csv" -Encoding UTF8

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgRoleReport

    .NOTES
    https://itpro-tips.com/get-the-office-365-admin-roles-and-track-the-changes/
    
    Written by Bastien Perez (Clidsys.com - ITPro-Tips.com)
    For more Office 365/Microsoft 365 tips and news, check out ITPro-Tips.com.
#>

function Get-MgRoleReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [string]$PrincipalID,

        [Parameter(Mandatory = $false)]
        [string]$PrincipalDisplayName,

        [Parameter(Mandatory = $false)]
        [string]$Scope,

        [Parameter(Mandatory = $false)]
        [ValidateSet('0', '1', '2')]
        [string]$TierLevel,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeEmptyRoles,

        [Parameter(Mandatory = $false)]
        [switch]$ExcludePIMEligibleAssignments,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        # using with the Maester framework
        [Parameter(Mandatory = $false)]
        [switch]$MaesterMode,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    [System.Collections.Generic.List[PSObject]]$rolesMembersArray = @()
    [System.Collections.Generic.List[Object]]$objectsCacheArray = @()
    [System.Collections.Generic.List[Object]]$mgRolesArrayAssignment = @()
    $scopeTypeCache = @{}

    # Privileged role tiering, inspired by Sean Metcalf's (PyroTek3) Get-EntraIDAdmins.ps1
    # https://github.com/PyroTek3/EntraID/blob/main/Get-EntraIDAdmins.ps1
    # Tier 0 = control plane (can compromise the whole tenant), Tier 1 = service/workload admins,
    # Tier 2 = lower-privilege / read-mostly roles. Roles not listed are tagged 'Untiered'.
    $tier0Roles = @(
        'Application Administrator'
        'Cloud Application Administrator'
        'Conditional Access Administrator'
        'Global Administrator'
        'Hybrid Identity Administrator'
        'Partner Tier2 Support'
        'Privileged Authentication Administrator'
        'Privileged Role Administrator'
        'Security Administrator'
    )

    $tier1Roles = @(
        'AI Administrator'
        'Attribute Provisioning Administrator'
        'Authentication Administrator'
        'Authentication Extensibility Administrator'
        'Authentication Policy Administrator'
        'B2C IEF Keyset Administrator'
        'Cloud App Security Administrator'
        'Compliance Administrator'
        'Directory Synchronization Accounts'
        'Directory Writers'
        'Domain Name Administrator'
        'Dynamics 365 Administrator'
        'Exchange Administrator'
        'External ID User Flow Administrator'
        'External Identity Provider Administrator'
        'Global Secure Access Administrator'
        'Groups Administrator'
        'Helpdesk Administrator'
        'Identity Governance Administrator'
        'Intune Administrator'
        'Knowledge Administrator'
        'Lifecycle Workflows Administrator'
        'Microsoft 365 Backup Administrator'
        'Microsoft 365 Migration Administrator'
        'On Premises Directory Sync Account'
        'Partner Tier1 Support'
        'Password Administrator'
        'Power Platform Administrator'
        'Security Operator'
        'SharePoint Administrator'
        'Skype for Business Administrator'
        'Teams Administrator'
        'Teams Telephony Administrator'
        'User Administrator'
        'Windows 365 Administrator'
        'Yammer Administrator'
    )

    $tier2Roles = @(
        'Application Developer'
        'Azure DevOps Administrator'
        'Azure Information Protection Administrator'
        'B2C IEF Policy Administrator'
        'Billing Administrator'
        'Cloud Device Administrator'
        'Customer Lockbox Access Approver'
        'Exchange Recipient Administrator'
        'External ID User Flow Attribute Administrator'
        'Global Reader'
        'License Administrator'
        'Microsoft Entra Joined Device Local Administrator'
        'Security Reader'
        'Teams Communications Administrator'
        'Teams Communications Support Engineer'
    )

    # Build a single role -> tier lookup for O(1) classification
    $roleTierLookup = @{}
    foreach ($role in $tier0Roles) { $roleTierLookup[$role] = 'Tier0' }
    foreach ($role in $tier1Roles) { $roleTierLookup[$role] = 'Tier1' }
    foreach ($role in $tier2Roles) { $roleTierLookup[$role] = 'Tier2' }

    $modules = @(
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Identity.Governance'
        'Microsoft.Graph.Users'
        'Microsoft.Graph.Groups'
        'Microsoft.Graph.Beta.Reports'
    )
    
    foreach ($module in $modules) {
        try {
            Import-Module $module -ErrorAction Stop 
        }
        catch {
            Write-Warning "First, install module $module"
            return
        }
    }

    $isConnected = $false

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
    
    if ($ForceNewToken.IsPresent) {
        Write-Verbose 'Disconnecting from Microsoft Graph'
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        $isConnected = $false
    }
    
    # Audit.Log.Read.All for sign-in activity
    # RoleManagement.Read.All for role assignment (PIM eligible and permanent)
    # Directory.Read.All for user and group and service principal information
    $permissionsNeeded = 'Directory.Read.All', 'RoleManagement.Read.All', 'AuditLog.Read.All'

    if (-not $isConnected) {
        Write-Verbose "Connecting to Microsoft Graph. Scopes: $permissionsNeeded"
        $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
    }

    if (-not (Test-MgGraphPermission -RequiredScopes $permissionsNeeded -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    Write-Verbose 'Collecting  roles with assignments...'

    try {
        #$mgRolesArrayAssignment = Get-MgRoleManagementDirectoryRoleDefinition -ErrorAction Stop
        
        Get-MgRoleManagementDirectoryRoleAssignment -All -ExpandProperty Principal | ForEach-Object {
            $mgRolesArrayAssignment.Add($_)
        }

        #$mgRolesArrayAssignment = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' -OutputType PSObject).Value

        $mgRolesDefinition = Get-MgRoleManagementDirectoryRoleAssignment -All -ExpandProperty roleDefinition

    }
    catch {
        Write-Warning $($_.Exception.Message)
    }

    # In *Assignment, we don't have the role definition, so we need to get it and add it to the object
    foreach ($assignment in $mgRolesArrayAssignment) {
        # Add the role definition to the object
        Add-Member -InputObject $assignment -MemberType NoteProperty -Name RoleDefinitionExtended -Value ($mgRolesDefinition | Where-Object { $_.id -eq $assignment.id }).roleDefinition 
        # Add-Member -InputObject $assignment -MemberType NoteProperty -Name RoleDefinitionExtended -Value ($mgRolesDefinition | Where-Object { $_.id -eq $assignment.id }).roleDefinition.description 
    } 
    

    if (-not $ExcludePIMEligibleAssignments) {
        Write-Verbose 'Collecting PIM eligible role assignments...'
        try {
            (Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty * -ErrorAction Stop | Select-Object id, principalId, directoryScopeId, roleDefinitionId, status, principal, @{Name = 'RoleDefinitionExtended'; Expression = { $_.roleDefinition } }) | ForEach-Object {
                $mgRolesArrayAssignment.Add($_)
            }
            #$mgRoles += (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedule' -OutputType PSObject).Value
        }
        catch {
            Write-Warning "Unable to get PIM eligible role assignments: $($_.Exception.Message)"
        }
    }

    foreach ($assignment in $mgRolesArrayAssignment) {
        $principal = switch ($assignment.principal.AdditionalProperties.'@odata.type') {
            '#microsoft.graph.user' { $assignment.principal.AdditionalProperties.userPrincipalName; break }
            '#microsoft.graph.servicePrincipal' { $assignment.principal.AdditionalProperties.appId; break }
            '#microsoft.graph.group' { $assignment.principalid; break }
            'default' { '-' }
        }

        $directoryScopeId = $assignment.directoryScopeId
        if ($scopeTypeCache.ContainsKey($directoryScopeId)) {
            $scopeType = $scopeTypeCache[$directoryScopeId].Type
            $scopeName = $scopeTypeCache[$directoryScopeId].Name
        }
        else {
            $scopeType = 'Unknown'
            $scopeName = $null

            switch -Wildcard ($directoryScopeId) {
                '/' {
                    $scopeType = 'Tenant'
                    $scopeName = 'Tenant'
                    break
                }
                '/administrativeUnits/*' {
                    $scopeType = 'AdministrativeUnit'
                    $auId = $directoryScopeId -replace '^/administrativeUnits/', ''
                    try {
                        $auObject = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directory/administrativeUnits/$auId" -OutputType PSObject -ErrorAction Stop
                        $scopeName = $auObject.displayName
                    }
                    catch {
                        Write-Verbose "Unable to resolve AU displayName for '$directoryScopeId': $($_.Exception.Message)"
                    }
                    break
                }
                default {
                    # Resolve directory object via Graph (e.g. application, servicePrincipal, group)
                    try {
                        $scopeObjectId = $directoryScopeId.TrimStart('/')
                        $scopeObject = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/directoryObjects/$scopeObjectId" -OutputType PSObject -ErrorAction Stop
                        if ($scopeObject.'@odata.type') {
                            $scopeType = $scopeObject.'@odata.type' -replace '^#microsoft\.graph\.', ''
                        }
                        if ($scopeObject.displayName) {
                            $scopeName = $scopeObject.displayName
                        }
                    }
                    catch {
                        Write-Verbose "Unable to resolve scope for '$directoryScopeId': $($_.Exception.Message)"
                    }
                }
            }

            $scopeTypeCache[$directoryScopeId] = @{ Type = $scopeType; Name = $scopeName }
        }

        $object = [PSCustomObject][ordered]@{
            Principal                = $principal
            PrincipalDisplayName     = $assignment.principal.AdditionalProperties.displayName
            PrincipalType            = $assignment.principal.AdditionalProperties.'@odata.type'.Split('.')[-1]
            PrincipalObjectID        = $assignment.principal.id
            AssignedRole             = $assignment.RoleDefinitionExtended.displayName
            AssignedRoleDefinitionId = $assignment.RoleDefinitionId
            AssignedRoleScope        = $assignment.directoryScopeId
            AssignedRoleScopeType    = $scopeType
            AssignedRoleScopeName    = $scopeName
            AssignmentType           = if ($assignment.status -eq 'Provisioned') { 'Eligible' } else { 'Permanent' }
            RoleIsBuiltIn            = $assignment.RoleDefinitionExtended.isBuiltIn
            RoleType                 = if ($assignment.RoleDefinitionExtended.isBuiltIn) { 'Built-in' } else { 'Custom' }
            RoleTemplate             = $assignment.RoleDefinitionExtended.templateId
            DirectMember             = $true
            Recommendations          = 'Check if the user has alternate email or alternate phone number on Microsoft Entra ID'
        }

        if ($object.PrincipalType -eq 'servicePrincipal') {
            $object.Recommendations = 'Verify who is the owner of this resource'
        }

        $rolesMembersArray.Add($object)

        if ($object.PrincipalType -eq 'group') {
            # need to get ID for Get-MgGroupMember
            $group = Get-MgGroup -GroupId $object.Principal -Property Id, onPremisesSyncEnabled
            $object | Add-Member -MemberType NoteProperty -Name 'OnPremisesSyncEnabled' -Value $([bool]($group.onPremisesSyncEnabled -eq $true))

            #$group = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($object.Principal)" -OutputType PSObject)

            $groupMembers = Get-MgGroupMember -GroupId $group.Id -Property displayName, userPrincipalName
            #$groupMembers = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$($group.Id)/members" -OutputType PSObject).Value

            foreach ($member in $groupMembers) {
                $typeMapping = @{
                    '#microsoft.graph.user'             = 'user'
                    '#microsoft.graph.group'            = 'group'
                    '#microsoft.graph.servicePrincipal' = 'servicePrincipal'
                    '#microsoft.graph.device'           = 'device'
                    '#microsoft.graph.orgContact'       = 'contact'
                    '#microsoft.graph.application'      = 'application'
                }

                $memberType = if ($typeMapping[$member.AdditionalProperties.'@odata.type']) {
                    $typeMapping[$member.AdditionalProperties.'@odata.type']
                }
                else {
                    'Unknown'
                }

                $object = [PSCustomObject][ordered]@{
                    Principal                = $member.AdditionalProperties.userPrincipalName
                    PrincipalDisplayName     = $member.AdditionalProperties.displayName
                    PrincipalType            = $memberType
                    PrincipalObjectID        = $member.Id
                    AssignedRole             = $assignment.RoleDefinitionExtended.displayName
                    AssignedRoleDefinitionId = $assignment.RoleDefinitionId
                    AssignedRoleScope        = $assignment.directoryScopeId
                    AssignedRoleScopeType    = $scopeType
                    AssignedRoleScopeName    = $scopeName
                    AssignmentType           = if ($assignment.status -eq 'Provisioned') { 'Eligible' } else { 'Permanent' }
                    RoleIsBuiltIn            = $assignment.RoleDefinitionExtended.isBuiltIn
                    RoleType                 = if ($assignment.RoleDefinitionExtended.isBuiltIn) { 'Built-in' } else { 'Custom' }
                    RoleTemplate             = $assignment.RoleDefinitionExtended.templateId
                    DirectMember             = $false
                    Recommendations      = 'Check if the user has alternate email or alternate phone number on Microsoft Entra ID'
                }

                if ($object.PrincipalType -eq 'servicePrincipal') {
                    $object.Recommendations = 'Verify who is the owner of this resource'
                }

                $rolesMembersArray.Add($object)
            }
        }
    }

    $object = [PSCustomObject] [ordered]@{
        Principal             = 'Partners'
        PrincipalDisplayName  = 'Partners'
        PrincipalType         = 'Partners'
        AssignedRole          = 'Partners'
        AssignedRoleScope     = 'Partners'
        AssignedRoleScopeType = 'Partners'
        AssignedRoleScopeName = 'Partners'
        AssignmentType        = 'Partners'
        RoleIsBuiltIn         = 'Not applicable'
        RoleType              = 'Not applicable'
        RoleTemplate          = 'Not applicable'
        DirectMember          = 'Not applicable'
        Recommendations       = 'Please check this URL to identify if you have partner with admin roles https: / / admin.microsoft.com / AdminPortal / Home#/partners. More information on https://practical365.com/identifying-potential-unwanted-access-by-your-msp-csp-reseller/'
    }
    
    $rolesMembersArray.Add($object)

    #foreach user, we check if the user is global administrator. If global administrator, we add a new parameter to the object recommandationRole to tell the other role is not useful
    $globalAdminsHash = @{}
    $rolesMembersArray | Where-Object { $_.AssignedRole -eq 'Global Administrator' } | ForEach-Object {
        $globalAdminsHash[$_.Principal] = $true
    }

    $rolesMembersArray | ForEach-Object {
        if ($globalAdminsHash.ContainsKey($_.Principal) -and $_.AssignedRole -ne 'Global Administrator') {
            $_.Recommendations += ' | This user is Global Administrator. The other role(s) is/are not useful.'
        }
    }

    # Roles that have been deprecated by Microsoft and should no longer be used
    $deprecatedRoles = @(
        'AdHoc License Administrator'
        'Device Join'
        'Device Managers'
        'Device Users'
        'Email Verified User Creator'
        'Mailbox Administrator'
        'Workplace Device Join'
    )

    $rolesMembersArray | ForEach-Object {
        if ($deprecatedRoles -contains $_.AssignedRole) {
            $_.Recommendations = 'This role is deprecated by Microsoft and should no longer be used. Remove this assignment.'
        }
    }

    foreach ($member in $rolesMembersArray) {
        Write-Verbose "Processing $($member.AssignedRole) - $($member.Principal)"

        $lastSignInDateTime = $null
        $accountEnabled = $null
        $onPremisesSyncEnabled = $null
        
        if ($objectsCacheArray.Principal -contains $member.Principal) {
            $accountEnabled = ($objectsCacheArray | Where-Object { $_.Principal -eq $member.Principal }).AccountEnabled
            $lastSignInDateTime = ($objectsCacheArray | Where-Object { $_.Principal -eq $member.Principal }).LastSignInDateTime
            $lastNonInteractiveSignInDateTime = ($objectsCacheArray | Where-Object { $_.Principal -eq $member.Principal }).LastNonInteractiveSignInDateTime
            $onPremisesSyncEnabled = ($objectsCacheArray | Where-Object { $_.Principal -eq $member.Principal }).onPremisesSyncEnabled
        }
        else {
            $lastSignInActivity = $null

            switch ($member.PrincipalType) {
                'user' {
                    # If we use Get-MgUser -UserId $member.Principal -Property AccountEnabled, SignInActivity, onPremisesSyncEnabled, 
                    # we encounter the error 'Get-MgUser_Get: Get By Key only supports UserId, and the key must be a valid GUID'.
                    # This is because the sign-in data comes from a different source that requires a GUID to retrieve the account's sign-in activity. 
                    # Therefore, we must provide the account's object identifier for the command to function correctly.
                    # To overcome this issue, we use the -Filter parameter to search for the user by their UserPrincipalName.
                    $mgUser = Get-MgUser -Filter "UserPrincipalName eq '$(ConvertTo-ODataEscapedString -Value $member.Principal)'" -Property AccountEnabled, SignInActivity, onPremisesSyncEnabled
                    $accountEnabled = $mgUser.AccountEnabled
                    $lastSignInDateTime = $mgUser.signInActivity.LastSignInDateTime
                    $lastNonInteractiveSignInDateTime = $mgUser.signInActivity.LastNonInteractiveSignInDateTime
                    $onPremisesSyncEnabled = [bool]($mgUser.onPremisesSyncEnabled -eq $true)
                    
                    break
                }

                'group' {
                    $accountEnabled = 'Not applicable'
                    $lastSignInDateTime = 'Not applicable'
                    $lastNonInteractiveSignInDateTime = 'Not applicable'
                    # onpremisesSyncEnabled already get from Get-MgGroup in the previous loop
                    $onPremisesSyncEnabled = $member.OnPremisesSyncEnabled

                    break
                }

                'servicePrincipal' {
                    $lastSignInActivity = (Get-MgBetaReportServicePrincipalSignInActivity -Filter "appId eq '$(ConvertTo-ODataEscapedString -Value $member.Principal)'").LastSignInActivity
                    $accountEnabled = 'Not applicable'
                    $lastSignInDateTime = $lastSignInActivity.LastSignInDateTime
                    $lastNonInteractiveSignInDateTime = $lastSignInActivity.LastNonInteractiveSignInDateTime
                    $onPremisesSyncEnabled = $false
                    
                    break
                }
                
                'Partners' {
                    $accountEnabled = 'Not applicable'
                    $lastSignInDateTime = 'Not applicable'
                    $lastNonInteractiveSignInDateTime = 'Not applicable'
                    $onPremisesSyncEnabled = 'Not applicable'

                    break
                }

                'default' {
                    $accountEnabled = 'Not applicable'
                    $lastSignInDateTime = 'Not applicable'
                    $lastNonInteractiveSignInDateTime = 'Not applicable'
                    $onPremisesSyncEnabled = 'Not applicable'

                }
            }
        }

        $member | Add-Member -MemberType NoteProperty -Name 'LastSignInDateTime' -Value $lastSignInDateTime
        $member | Add-Member -MemberType NoteProperty -Name 'LastNonInteractiveSignInDateTime' -Value $lastNonInteractiveSignInDateTime
        $member | Add-Member -MemberType NoteProperty -Name 'AccountEnabled' -Value $accountEnabled
        $member | Add-Member -MemberType NoteProperty -Name 'OnPremisesSyncEnabled' -Value $onPremisesSyncEnabled -Force

        if ($onPremisesSyncEnabled) {
            $member.Recommendations += ' | Privileged accounts should be cloud-only.'
        }

        # only add if not already in the cache
        if (-not $objectsCacheArray.Principal -contains $member.Principal) {
            $objectsCacheArray.Add($member)
        }
    }
    
    if ($IncludeEmptyRoles.IsPresent) {

        Write-Verbose 'Collecting all roles...'
        try {
            #$mgRolesArrayAssignment = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions' -OutputType PSObject).Value
            $mgRolesDefinition = Get-MgRoleManagementDirectoryRoleDefinition -All -ErrorAction Stop

            $emptyRoles = $mgRolesDefinition | Where-Object { $mgRolesArrayAssignment.RoleDefinitionId -notcontains $_.id }

            foreach ($emptyRole in $emptyRoles) {
                $object = [PSCustomObject][ordered]@{
                    Principal                        = 'Role has no members'
                    PrincipalDisplayName             = $null
                    PrincipalType                    = $null
                    PrincipalObjectID                = $null
                    AssignedRole                     = $emptyRole.displayName
                    AssignedRoleDefinitionId         = $emptyRole.id
                    AssignedRoleScope                = $null
                    AssignedRoleScopeType            = $null
                    AssignedRoleScopeName            = $null
                    AssignmentType                   = $null
                    RoleIsBuiltIn                    = $emptyRole.isBuiltIn
                    RoleType                         = if ($emptyRole.isBuiltIn) { 'Built-in' } else { 'Custom' }
                    RoleTemplate                     = $emptyRole.templateId
                    DirectMember                     = $null
                    Recommendations                  = $null
                    LastSignInDateTime               = $null
                    LastNonInteractiveSignInDateTime = $null
                    AccountEnabled                   = $null
                    OnPremisesSyncEnabled            = $null
                }

                $rolesMembersArray.Add($object)
            }
        }
        catch {
            Write-Warning $($_.Exception.Message)   
        }   

    }

    # Classify each assignment into its privileged tier (Tier0/Tier1/Tier2/Untiered)
    $rolesMembersArray | ForEach-Object {
        if ($_.PrincipalType -eq 'Partners') {
            $roleTier = 'Not applicable'
        }
        elseif ($roleTierLookup.ContainsKey($_.AssignedRole)) {
            $roleTier = $roleTierLookup[$_.AssignedRole]
        }
        else {
            $roleTier = 'Untiered'
        }

        $_ | Add-Member -MemberType NoteProperty -Name 'RoleTier' -Value $roleTier -Force
    }

    # Apply optional filters on the final result set
    if ($Identity -or $PrincipalID -or $PrincipalDisplayName -or $Scope -or $TierLevel) {
        $tierFilter = if ($TierLevel) { "Tier$TierLevel" } else { $null }

        $rolesMembersArray = $rolesMembersArray | Where-Object {
            (-not $Identity             -or $_.AssignedRole -eq $Identity -or $_.AssignedRoleDefinitionId -eq $Identity) -and
            (-not $PrincipalID          -or $_.Principal -eq $PrincipalID -or $_.PrincipalObjectID -eq $PrincipalID) -and
            (-not $PrincipalDisplayName -or $_.PrincipalDisplayName -eq $PrincipalDisplayName) -and
            (-not $Scope                -or $_.AssignedRoleScope -eq $Scope) -and
            (-not $tierFilter           -or $_.RoleTier -eq $tierFilter)
        }

        if (-not $rolesMembersArray) {
            Write-Warning 'No role assignment found matching the specified filter(s).'
            return
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFileName = "$($env:userprofile)\$now-MgRoleReport.xlsx"
        Write-Verbose "Exporting report to Excel file: $excelFileName"

        $rolesMembersArray | Export-Excel -Path $excelFileName -AutoSize -AutoFilter -Title 'Microsoft Entra ID Role Report' -WorksheetName 'Role Report' -TableName 'MgRoleReport' -FreezeTopRow
    }
    else {
        return $rolesMembersArray
    }
}