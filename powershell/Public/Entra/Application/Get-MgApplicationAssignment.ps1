<#
    .SYNOPSIS
    Retrieves all Entra ID applications and their assignment types.
    
    .DESCRIPTION
    This function returns a list of all Entra ID applications with their assignment information,
    identifying whether they are assigned to all users or have specific assignments.
    If no assignments exist, it indicates whether the application is available to "all users".
    Give information about assigned users, groups, or service principals and if the group is protected/static/dynamic.
    
    .PARAMETER ApplicationId
    (Optional) One or more Application IDs to filter the results. If not provided, all
    applications will be processed.

    .PARAMETER AllApplications
    (Optional) If specified, retrieves all service principals regardless of type.
    By default, only Enterprise Applications (tagged 'WindowsAzureActiveDirectoryIntegratedApp') are returned.

    .PARAMETER AssignmentNotEnforced
    (Optional) If specified, only applications where AppRoleAssignmentRequired is $false (open to all users, no assignment needed)
    will be returned.

    .PARAMETER AssignmentEmpty
    (Optional) If specified, only applications with no specific user/group/service principal assignments will be returned.
    Note: these apps may still be accessible to all users if AppRoleAssignmentRequired is $false.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER NoPermissionCheck
    (Optional) Skip the Microsoft Graph scope verification performed against the current Get-MgContext token.

    .EXAMPLE
    Get-MgApplicationAssignment

    Retrieves all applications and their assignment types.

    .EXAMPLE
    Get-MgApplicationAssignment -ApplicationId "xxx", "yyy"

    Retrieves assignment types for the specified application IDs.
    
    .EXAMPLE
    Get-MgApplicationAssignment -AssignmentEmpty

    Retrieves only applications with no specific user/group/service principal assignments.

    .EXAMPLE
    Get-MgApplicationAssignment -ExportToExcel
    
    Gets all applications and exports them to an Excel file

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationAssignment
#>
    
function Get-MgApplicationAssignment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [String[]]$ApplicationId,

        [Parameter(Mandatory = $false)]
        [switch]$AllApplications,

        [Parameter(Mandatory = $false)]
        [switch]$AssignmentNotEnforced,

        [Parameter(Mandatory = $false)]
        [switch]$AssignmentEmpty,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [switch]$NoPermissionCheck
    )

    if (-not $NoPermissionCheck.IsPresent) {
        $requiredScopes = @(
            'Application.Read.All',
            'User.Read.All',
            'Group.Read.All'
        )
        if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }
    }

    $spProperty = 'Id,AppId,DisplayName,AppRoles,AppRoleAssignmentRequired,PublisherName,Tags,ServicePrincipalType,CreatedDateTime'

    # Get all service principals (enterprise applications)
    if ($ApplicationId) {
        $servicePrincipals = @()
        foreach ($appId in $ApplicationId) {
            try {
                $sp = Get-MgServicePrincipal -Filter "AppId eq '$appId'" -Property $spProperty -ErrorAction Stop
                if ($sp) {
                    $servicePrincipals += $sp
                }
            }
            catch {
                Write-Warning "Could not find application with ID: $appId"
            }
        }
    }
    elseif ($AllApplications) {
        $servicePrincipals = Get-MgServicePrincipal -All -Property $spProperty
    }
    else {
        # Default: Enterprise Applications only (tagged 'WindowsAzureActiveDirectoryIntegratedApp')
        $servicePrincipals = Get-MgServicePrincipal -All -Property $spProperty `
            -Filter "tags/Any(x: x eq 'WindowsAzureActiveDirectoryIntegratedApp')"
    }
    
    # Initialize results array
    [System.Collections.Generic.List[PSCustomObject]]$applicationAssignmentsArray = @()
    
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting analysis of $($servicePrincipals.Count) applications..." -ForegroundColor Cyan
    $counter = 0
    
    foreach ($sp in $servicePrincipals) {
        $counter++
        Write-Verbose "[$counter/$($servicePrincipals.Count)] Analyzing application: $($sp.DisplayName)"
        
        # Get app role assignments for this service principal
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id
        
        if ($appRoleAssignments.Count -gt 0) {
            # Process each assignment individually
            foreach ($assignment in $appRoleAssignments) {
                # Initialize assignment properties in logical order for readability
                $assignmentProps = [ordered]@{
                    # Application info (most important first)
                    ApplicationName            = $sp.DisplayName
                    ApplicationType            = if ($sp.Tags -contains 'WindowsAzureActiveDirectoryIntegratedApp') { 'Enterprise Application' } else { $sp.ServicePrincipalType }
                    AssignmentType             = ''
                    PrincipalType              = ''
                    IsRoleAssignableGroup      = $null
                    
                    # Principal details (User, Group, or Service Principal)
                    UserName                   = $null
                    UserPrincipalName          = $null
                    GroupName                  = $null
                    GroupType                  = $null
                    GroupMembershipType        = $null
                    ServicePrincipalName       = $null
                    ServicePrincipalType       = $null
                    PrincipalDisplayName       = $null
                    
                    # Role and permission details
                    AppRoleValue               = ($sp.AppRoles | Where-Object { $_.Id -eq $assignment.AppRoleId }).Value
                    AppRoleId                  = $assignment.AppRoleId
                    AppRoleAssignmentRequired  = $sp.AppRoleAssignmentRequired
                    
                    # Technical IDs (less important, at the end)
                    ApplicationId              = $sp.AppId
                    ServicePrincipalId         = $sp.Id
                    UserId                     = $null
                    GroupId                    = $null
                    AssignedServicePrincipalId = $null
                    PrincipalId                = $null
                    IsAssignableToRole         = $null
                    ServicePrincipalAppId      = $null
                    
                    # Metadata (at the end)
                    ApplicationPublisher       = $sp.PublisherName
                    CreatedDate                = $assignment.CreatedDateTime
                }
                
                if ($assignment.PrincipalType -eq 'User') {
                    $assignmentProps.AssignmentType = 'User Assignment'
                    $assignmentProps.PrincipalType = 'User'
                    
                    try {
                        $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($user) {
                            Write-Verbose "  -> User found: $($user.DisplayName)"
                            
                            $assignmentProps.UserName = $user.DisplayName
                            $assignmentProps.UserPrincipalName = $user.UserPrincipalName
                            $assignmentProps.UserId = $user.Id
                        }
                        else {
                            $assignmentProps.AssignmentType = 'User Assignment (Not Found)'
                            $assignmentProps.UserName = "User ID: $($assignment.PrincipalId)"
                            $assignmentProps.UserPrincipalName = 'Unknown'
                            $assignmentProps.UserId = $assignment.PrincipalId
                        }
                    }
                    catch {
                        Write-Warning "  -> Error retrieving user $($assignment.PrincipalId): $($_.Exception.Message)"
                        
                        $assignmentProps.AssignmentType = 'User Assignment (Error)'
                        $assignmentProps.UserName = "User ID: $($assignment.PrincipalId)"
                        $assignmentProps.UserPrincipalName = 'Unknown'
                        $assignmentProps.UserId = $assignment.PrincipalId
                    }
                }
                elseif ($assignment.PrincipalType -eq 'Group') {
                    $assignmentProps.AssignmentType = 'Group Assignment'
                    $assignmentProps.PrincipalType = 'Group'
                    
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($group) {
                            $protectedStatus = if ($null -ne $group.IsAssignableToRole) { $group.IsAssignableToRole } else { 'N/A' }
                            Write-Verbose "  -> Group found: $($group.DisplayName) (Protected: $protectedStatus)"
                            
                            $assignmentProps.GroupName = $group.DisplayName
                            $assignmentProps.GroupId = $group.Id
                            $assignmentProps.GroupType = $group.GroupTypes -join ','
                            $assignmentProps.GroupMembershipType = if ($group.MembershipRule -and $group.MembershipRuleProcessingState -eq 'On') { 'Dynamic' } else { 'Static' }
                            $assignmentProps.IsAssignableToRole = $group.IsAssignableToRole
                            $assignmentProps.IsRoleAssignableGroup = $group.IsAssignableToRole
                        }
                        else {
                            $assignmentProps.AssignmentType = 'Group Assignment (Not Found)'
                            $assignmentProps.GroupName = "Group ID: $($assignment.PrincipalId)"
                            $assignmentProps.GroupId = $assignment.PrincipalId
                            $assignmentProps.GroupType = 'Unknown'
                            $assignmentProps.GroupMembershipType = 'Unknown'
                            $assignmentProps.IsAssignableToRole = $null
                            $assignmentProps.IsRoleAssignableGroup = $null
                        }
                    }
                    catch {
                        Write-Warning "  -> Error retrieving group $($assignment.PrincipalId): $($_.Exception.Message)"
                        
                        $assignmentProps.AssignmentType = 'Group Assignment (Error)'
                        $assignmentProps.GroupName = "Group ID: $($assignment.PrincipalId)"
                        $assignmentProps.GroupId = $assignment.PrincipalId
                        $assignmentProps.GroupType = 'Unknown'
                        $assignmentProps.GroupMembershipType = 'Unknown'
                        $assignmentProps.IsAssignableToRole = $null
                        $assignmentProps.IsRoleAssignableGroup = $null
                    }
                }
                elseif ($assignment.PrincipalType -eq 'ServicePrincipal') {
                    $assignmentProps.AssignmentType = 'Service Principal Assignment'
                    $assignmentProps.PrincipalType = 'ServicePrincipal'
                    
                    try {
                        $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($servicePrincipal) {
                            Write-Verbose "  -> Service Principal found: $($servicePrincipal.DisplayName)"
                            
                            $assignmentProps.ServicePrincipalName = $servicePrincipal.DisplayName
                            $assignmentProps.ServicePrincipalAppId = $servicePrincipal.AppId
                            $assignmentProps.ServicePrincipalType = $servicePrincipal.ServicePrincipalType
                            $assignmentProps.AssignedServicePrincipalId = $servicePrincipal.Id
                        }
                        else {
                            $assignmentProps.AssignmentType = 'Service Principal Assignment (Not Found)'
                            $assignmentProps.ServicePrincipalName = "Service Principal ID: $($assignment.PrincipalId)"
                            $assignmentProps.AssignedServicePrincipalId = $assignment.PrincipalId
                            $assignmentProps.ServicePrincipalAppId = 'Unknown'
                            $assignmentProps.ServicePrincipalType = 'Unknown'
                        }
                    }
                    catch {
                        Write-Warning "  -> Error retrieving Service Principal $($assignment.PrincipalId): $($_.Exception.Message)"
                        
                        $assignmentProps.AssignmentType = 'Service Principal Assignment (Error)'
                        $assignmentProps.ServicePrincipalName = "Service Principal ID: $($assignment.PrincipalId)"
                        $assignmentProps.AssignedServicePrincipalId = $assignment.PrincipalId
                        $assignmentProps.ServicePrincipalAppId = 'Unknown'
                        $assignmentProps.ServicePrincipalType = 'Unknown'
                    }
                }
                else {
                    # Handle unknown/unsupported principal types
                    $assignmentProps.AssignmentType = "$($assignment.PrincipalType) Assignment"
                    $assignmentProps.PrincipalType = $assignment.PrincipalType
                    $assignmentProps.PrincipalId = $assignment.PrincipalId
                    $assignmentProps.PrincipalDisplayName = "Unknown $($assignment.PrincipalType)"
                    
                    Write-Warning "  -> Unhandled principal type: $($assignment.PrincipalType) (ID: $($assignment.PrincipalId))"
                }
                
                # Create and add the assignment object
                $object = [PSCustomObject]$assignmentProps
                $applicationAssignmentsArray.Add($object)
            }
        }
        else {
            # No assignments found - create one entry using the same complete structure
            $assignmentProps = [ordered]@{
                # Application info (most important first)
                ApplicationName            = $sp.DisplayName
                ApplicationType            = if ($sp.Tags -contains 'WindowsAzureActiveDirectoryIntegratedApp') { 'Enterprise Application' } else { $sp.ServicePrincipalType }
                AssignmentType             = ''
                PrincipalType              = ''
                IsRoleAssignableGroup      = $null
                
                # Principal details (User, Group, or Service Principal)
                UserName                   = $null
                UserPrincipalName          = $null
                GroupName                  = $null
                GroupType                  = $null
                GroupMembershipType        = $null
                ServicePrincipalName       = $null
                ServicePrincipalType       = $null
                PrincipalDisplayName       = $null
                
                # Role and permission details
                AppRoleValue               = $null
                AppRoleId                  = $null
                AppRoleAssignmentRequired  = $sp.AppRoleAssignmentRequired
                
                # Technical IDs (less important, at the end)
                ApplicationId              = $sp.AppId
                ServicePrincipalId         = $sp.Id
                UserId                     = $null
                GroupId                    = $null
                AssignedServicePrincipalId = $null
                PrincipalId                = $null
                IsAssignableToRole         = $null
                ServicePrincipalAppId      = $null
                
                # Metadata (at the end)
                ApplicationPublisher       = $sp.PublisherName
                CreatedDate                = $sp.CreatedDateTime
            }
            
            if (-not $sp.AppRoleAssignmentRequired) {
                $assignmentProps.AssignmentType = 'All Users (No Assignment Required)'
                $assignmentProps.PrincipalType = 'All Users'
                Write-Verbose '  -> Application available to all users (no assignment required)'
            }
            else {
                $assignmentProps.AssignmentType = 'Not Assigned'
                $assignmentProps.PrincipalType = 'None'
                Write-Verbose '  -> No assignments found'
            }
            
            # Create and add the no-assignment object
            $object = [PSCustomObject]$assignmentProps
            $applicationAssignmentsArray.Add($object)
        }
    }
    
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Analysis completed. Total items found: $($applicationAssignmentsArray.Count)" -ForegroundColor Cyan
    
    # Apply filtering if requested
    if ($AssignmentNotEnforced.IsPresent) {
        $beforeCount = $applicationAssignmentsArray.Count
        $applicationAssignmentsArray = $applicationAssignmentsArray | Where-Object {
            $_.AppRoleAssignmentRequired -eq $false
        }
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Filtering applied (AssignmentNotEnforced): $($applicationAssignmentsArray.Count)/$beforeCount items retained" -ForegroundColor Cyan
    }

    if ($AssignmentEmpty.IsPresent) {
        $beforeCount = $applicationAssignmentsArray.Count
        $applicationAssignmentsArray = $applicationAssignmentsArray | Where-Object {
            $_.AssignmentType -in @('Not Assigned', 'All Users (No Assignment Required)')
        }
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Filtering applied (AssignmentEmpty): $($applicationAssignmentsArray.Count)/$beforeCount items retained" -ForegroundColor Cyan
    }
    
    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationAssignment.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting application assignments to Excel file: $excelFilePath"
        $applicationAssignmentsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraApplicationAssignments'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $applicationAssignmentsArray
    }
}