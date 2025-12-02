<#
    .SYNOPSIS
    Retrieves all Entra ID applications and their assignment types.
    
    .DESCRIPTION
    This function returns a list of all Entra ID applications with their assignment information,
    identifying whether they are assigned to all users or have specific assignments.
    
    .EXAMPLE
    $apps = Get-MgApplicationAssignment

    .EXAMPLE
    $apps = Get-MgApplicationAssignment -ApplicationIds "xxx", "yyy"
    
    .EXAMPLE
    $apps = Get-MgApplicationAssignment -OnlyAssignedToAllUsers

    .EXAMPLE
    Get-MgApplicationAssignment -ExportToExcel
    Gets all applications and exports them to an Excel file
    #>
    
function Get-MgApplicationAssignment {
    param(
        [Parameter(Mandatory = $false)]
        [String[]]$ApplicationIds,
        [Parameter(Mandatory = $false)]
        [switch]$OnlyAssignedToAllUsers,
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )
    
    # Get all service principals (enterprise applications)
    $servicePrincipals = Get-MgServicePrincipal -All
    
    # Initialize results array
    [System.Collections.Generic.List[PSCustomObject]]$applicationAssignmentsArray = @()
    
    foreach ($sp in $servicePrincipals) {
        # Get app role assignments for this service principal
        $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id
        
        # Determine assignment type
        $assignmentType = 'Not Assigned'
        $assignedToCount = 0
        
        [System.Collections.Generic.List[Object]]$assignedUsers = @()
        [System.Collections.Generic.List[Object]]$assignedGroups = @()
        
        if ($appRoleAssignments.Count -gt 0) {
            $assignedToCount = $appRoleAssignments.Count
            
            foreach ($assignment in $appRoleAssignments) {
                # Check if assigned to users or groups
                if ($assignment.PrincipalType -eq 'User') {
                    try {
                        $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($user) {
                            $assignedUsers.Add($user.DisplayName)
                        }
                    }
                    catch {
                        $assignedUsers.Add("User ID: $($assignment.PrincipalId)")
                    }
                }
                elseif ($assignment.PrincipalType -eq 'Group') {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($group) {
                            $assignedGroups.Add($group.DisplayName)
                        }
                    }
                    catch {
                        $assignedGroups.Add("Group ID: $($assignment.PrincipalId)")
                    }
                }
            }
            
            # Determine if it's assigned to "All Users" or specific assignments
            if ($assignedGroups -contains 'All Users' -or $sp.AppRoleAssignmentRequired -eq $false) {
                $assignmentType = 'All Users'
            }
            else {
                $assignmentType = 'Specific Assignment'
            }
        }
        else {
            # Check if user assignment is required
            if ($sp.AppRoleAssignmentRequired -eq $false) {
                $assignmentType = 'All Users (No Assignment Required)'
            }
            else {
                $assignmentType = 'Not Assigned'
            }
        }
        
        # Create result object
        $object = [PSCustomObject][ordered]@{
            ApplicationName           = $sp.DisplayName
            ApplicationId             = $sp.AppId
            ServicePrincipalId        = $sp.Id
            AssignmentType            = $assignmentType
            AppRoleAssignmentRequired = $sp.AppRoleAssignmentRequired
            AssignedToCount           = $assignedToCount
            AssignedUsers             = ($assignedUsers -join '; ')
            AssignedGroups            = ($assignedGroups -join '; ')
            ApplicationPublisher      = $sp.PublisherName
            ApplicationCategory       = $sp.Tags -join '; '
            CreatedDate               = $sp.CreatedDateTime
        }
        
        $applicationAssignmentsArray.Add($object)
    }
    
    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationAssignment.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting application assignments to Excel file: $excelFilePath"
        $applicationAssignmentsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraApplicationAssignments'
        Write-Host -ForegroundColor Green "Export completed successfully!"
    }
    else {
        return $applicationAssignmentsArray
    }
}