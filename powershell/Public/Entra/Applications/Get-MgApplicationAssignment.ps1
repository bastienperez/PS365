<#
    .SYNOPSIS
    Retrieves all Entra ID applications and their assignment types.
    
    .DESCRIPTION
    This function returns a list of all Entra ID applications with their assignment information,
    identifying whether they are assigned to all users or have specific assignments.
    
    .EXAMPLE
    $apps = Get-MgApplicationAssignment

    .EXAMPLE
    $apps = Get-MgApplicationAssignment -ApplicationId "xxx", "yyy"
    
    .EXAMPLE
    $apps = Get-MgApplicationAssignment -OnlyAssignedToAllUsers

    .EXAMPLE
    Get-MgApplicationAssignment -ExportToExcel
    Gets all applications and exports them to an Excel file

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationAssignment
    #>
    
function Get-MgApplicationAssignment {
    param(
        [Parameter(Mandatory = $false)]
        [String[]]$ApplicationId,

        [Parameter(Mandatory = $false)]
        [switch]$OnlyAssignedToAllUsers,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )
    
    # Get all service principals (enterprise applications)
    if ($ApplicationId) {
        $servicePrincipals = @()
        foreach ($appId in $ApplicationId) {
            try {
                $sp = Get-MgServicePrincipal -Filter "AppId eq '$appId'" -ErrorAction Stop
                if ($sp) {
                    $servicePrincipals += $sp
                }
            }
            catch {
                Write-Warning "Could not find application with ID: $appId"
            }
        }
    }
    else {
        $servicePrincipals = Get-MgServicePrincipal -All
    }
    
    # Initialize results array
    [System.Collections.Generic.List[PSCustomObject]]$applicationAssignmentsArray = @()
    
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting analysis of $($servicePrincipals.Count) applications..." -ForegroundColor Cyan
    $counter = 0
    
    foreach ($sp in $servicePrincipals) {
        $counter++
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$counter/$($servicePrincipals.Count)] Analyzing application: $($sp.DisplayName)" -ForegroundColor Cyan
        
        # Get app role assignments for this service principal
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id
        
        if ($appRoleAssignments.Count -gt 0) {
            # Process each assignment individually
            foreach ($assignment in $appRoleAssignments) {
                # Initialize assignment properties in logical order for readability
                $assignmentProps = [ordered]@{
                    # Application info (most important first)
                    ApplicationName            = $sp.DisplayName
                    AssignmentType             = ''
                    PrincipalType              = ''
                    IsRoleAssignableGroup      = $null
                    
                    # Principal details (User, Group, or Service Principal)
                    UserName                   = $null
                    UserPrincipalName          = $null
                    GroupName                  = $null
                    GroupType                  = $null
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
                            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> User found: $($user.DisplayName)" -ForegroundColor Green
                            
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
                        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Error retrieving user $($assignment.PrincipalId): $($_.Exception.Message)" -ForegroundColor Red
                        
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
                            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Group found: $($group.DisplayName) (Protected: $protectedStatus)" -ForegroundColor Green
                            
                            $assignmentProps.GroupName = $group.DisplayName
                            $assignmentProps.GroupId = $group.Id
                            $assignmentProps.GroupType = $group.GroupTypes -join ','
                            $assignmentProps.IsAssignableToRole = $group.IsAssignableToRole
                            $assignmentProps.IsRoleAssignableGroup = $group.IsAssignableToRole
                        }
                        else {
                            $assignmentProps.AssignmentType = 'Group Assignment (Not Found)'
                            $assignmentProps.GroupName = "Group ID: $($assignment.PrincipalId)"
                            $assignmentProps.GroupId = $assignment.PrincipalId
                            $assignmentProps.GroupType = 'Unknown'
                            $assignmentProps.IsAssignableToRole = $null
                            $assignmentProps.IsRoleAssignableGroup = $null
                        }
                    }
                    catch {
                        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Error retrieving group $($assignment.PrincipalId): $($_.Exception.Message)" -ForegroundColor Red
                        
                        $assignmentProps.AssignmentType = 'Group Assignment (Error)'
                        $assignmentProps.GroupName = "Group ID: $($assignment.PrincipalId)"
                        $assignmentProps.GroupId = $assignment.PrincipalId
                        $assignmentProps.GroupType = 'Unknown'
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
                            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Service Principal found: $($servicePrincipal.DisplayName)" -ForegroundColor Magenta
                            
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
                        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Error retrieving Service Principal $($assignment.PrincipalId): $($_.Exception.Message)" -ForegroundColor Red
                        
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
                    
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Unhandled principal type: $($assignment.PrincipalType) (ID: $($assignment.PrincipalId))" -ForegroundColor DarkYellow
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
                AssignmentType             = ''
                PrincipalType              = ''
                IsRoleAssignableGroup      = $null
                
                # Principal details (User, Group, or Service Principal)
                UserName                   = $null
                UserPrincipalName          = $null
                GroupName                  = $null
                GroupType                  = $null
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
            
            if ($sp.AppRoleAssignmentRequired -eq $false) {
                $assignmentProps.AssignmentType = 'All Users (No Assignment Required)'
                $assignmentProps.PrincipalType = 'All Users'
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Application available to all users (no assignment required)" -ForegroundColor Yellow
            }
            else {
                $assignmentProps.AssignmentType = 'Not Assigned'
                $assignmentProps.PrincipalType = 'None'
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> No assignments found" -ForegroundColor Gray
            }
            
            # Create and add the no-assignment object
            $object = [PSCustomObject]$assignmentProps
            $applicationAssignmentsArray.Add($object)
        }
    }
    
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Analysis completed. Total items found: $($applicationAssignmentsArray.Count)" -ForegroundColor Cyan
    
    # Apply filtering if requested
    if ($OnlyAssignedToAllUsers.IsPresent) {
        $beforeCount = $applicationAssignmentsArray.Count
        $applicationAssignmentsArray = $applicationAssignmentsArray | Where-Object { 
            $_.AssignmentType -eq 'All Users (No Assignment Required)' 
        }
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Filtering applied (OnlyAssignedToAllUsers): $($applicationAssignmentsArray.Count)/$beforeCount items retained" -ForegroundColor Cyan
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