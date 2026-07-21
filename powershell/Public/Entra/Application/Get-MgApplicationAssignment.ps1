<#
    .SYNOPSIS
    Retrieves all Entra ID applications and their assignment types.
    
    .DESCRIPTION
    This function returns a list of all Entra ID applications with their assignment information,
    identifying whether they are assigned to all users or have specific assignments.
    If no assignments exist, it indicates whether the application is available to "all users".
    Give information about assigned users, groups, or service principals and if the group is protected/static/dynamic.
    
    .PARAMETER ApplicationId
    (Optional) One or more Application IDs (AppId) to filter the results. If not provided, all
    applications will be processed.

    .PARAMETER ObjectID
    (Optional) ObjectID (GUID) of a single service principal to target. Cannot be combined with -ApplicationId or -DisplayName.

    .PARAMETER DisplayName
    (Optional) Display name of a single service principal to target (exact match, with fallback on trimmed comparison). Cannot be combined with -ApplicationId or -ObjectID.

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

    .PARAMETER DisableParallel
    (Optional) Forces sequential processing. By default, on PowerShell 7+ the function analyzes applications in
    parallel (ForEach-Object -Parallel) to speed up processing; on PowerShell 5.1 it always runs sequentially.

    .PARAMETER ThrottleLimit
    (Optional) Maximum number of concurrent runspaces when running in parallel. Default is 5.
    Keep this value moderate to avoid Microsoft Graph throttling (HTTP 429).

    .EXAMPLE
    Get-MgApplicationAssignment

    Retrieves all applications and their assignment types.

    .EXAMPLE
    Get-MgApplicationAssignment -DisableParallel

    Forces sequential processing even on PowerShell 7+ (useful for debugging or to avoid concurrent Graph calls).

    .EXAMPLE
    Get-MgApplicationAssignment -ApplicationId "xxx", "yyy"

    Retrieves assignment types for the specified application IDs.

    .EXAMPLE
    Get-MgApplicationAssignment -ObjectID '11111111-2222-3333-4444-555555555555'

    Retrieves assignment types for the service principal matching this ObjectID.

    .EXAMPLE
    Get-MgApplicationAssignment -DisplayName 'My SAML App'

    Retrieves assignment types for the service principal matching this DisplayName.
    
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
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param(
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'ByApplicationId')]
        [String[]]$ApplicationId,

        [Parameter(Mandatory = $false, ParameterSetName = 'ByObjectId')]
        [Alias('Identity')]
        [string]$ObjectID,

        [Parameter(Mandatory = $false, ParameterSetName = 'ByDisplayName')]
        [string]$DisplayName,

        [Parameter(Mandatory = $false, ParameterSetName = 'All')]
        [switch]$AllApplications,

        [Parameter(Mandatory = $false)]
        [switch]$AssignmentNotEnforced,

        [Parameter(Mandatory = $false)]
        [switch]$AssignmentEmpty,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath,

        [Parameter(Mandatory = $false)]
        [switch]$DisableParallel,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 20)]
        [int]$ThrottleLimit = 5
    )

    $requiredScopes = @('Application.Read.All', 'User.Read.All', 'Group.Read.All')
    if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
        return
    }

    # Run in parallel by default on PowerShell 7+ (ForEach-Object -Parallel); always sequential on PowerShell 5.1.
    $useParallel = ($PSVersionTable.PSVersion.Major -ge 7) -and -not $DisableParallel

    $spProperty = 'Id,AppId,DisplayName,AppRoles,AppRoleAssignmentRequired,PublisherName,Tags,ServicePrincipalType,CreatedDateTime'

    # Get all service principals (enterprise applications)
    switch ($PSCmdlet.ParameterSetName) {
        'ByApplicationId' {
            # Resolve all AppIds in a single filtered query (was N+1: one
            # Get-MgServicePrincipal call per AppId).
            $servicePrincipals = @()
            if ($ApplicationId.Count -gt 0) {
                $filterParts = $ApplicationId | ForEach-Object { "AppId eq '$(ConvertTo-ODataEscapedString -Value $_)'" }
                $filter = $filterParts -join ' or '
                try {
                    $servicePrincipals = @(Get-MgServicePrincipal -Filter $filter -Property $spProperty -All -ErrorAction Stop)
                }
                catch {
                    Write-Warning "Could not retrieve service principals: $($_.Exception.Message)"
                }

                $foundAppIds = $servicePrincipals.AppId
                foreach ($appId in $ApplicationId) {
                    if ($appId -notin $foundAppIds) {
                        Write-Warning "Could not find application with ID: $appId"
                    }
                }
            }
        }
        'ByObjectId' {
            try {
                $servicePrincipals = @(Get-MgServicePrincipal -ServicePrincipalId $ObjectID -Property $spProperty -ErrorAction Stop)
            }
            catch {
                Write-Warning "Could not find service principal with ObjectID: $ObjectID"
                return
            }
        }
        'ByDisplayName' {
            $escaped = $DisplayName -replace "'", "''"
            $filter = "DisplayName eq '$escaped'"
            Write-Verbose "Filtering service principals with: $filter"
            $servicePrincipals = @(Get-MgServicePrincipal -Filter $filter -All -Property $spProperty)

            # Fallback: trimmed display name comparison if no exact match
            if (-not $servicePrincipals) {
                Write-Verbose "No exact match for '$DisplayName'. Searching with startswith and trimmed client-side comparison."
                $startsWithFilter = "startswith(DisplayName, '$escaped')"
                $candidates = Get-MgServicePrincipal -Filter $startsWithFilter -All -Property $spProperty
                $servicePrincipals = @($candidates | Where-Object { $_.DisplayName.Trim() -eq $DisplayName.Trim() })
            }

            if (-not $servicePrincipals) {
                Write-Warning "No service principal found matching DisplayName '$DisplayName'."
                return
            }
        }
        default {
            if ($AllApplications) {
                $servicePrincipals = Get-MgServicePrincipal -All -Property $spProperty
            }
            else {
                # Default: Enterprise Applications only (tagged 'WindowsAzureActiveDirectoryIntegratedApp')
                $servicePrincipals = Get-MgServicePrincipal -All -Property $spProperty `
                    -Filter "tags/Any(x: x eq 'WindowsAzureActiveDirectoryIntegratedApp')"
            }
        }
    }
    
    # Initialize results array
    [System.Collections.Generic.List[PSCustomObject]]$applicationAssignmentsArray = @()
    
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting analysis of $($servicePrincipals.Count) applications..." -ForegroundColor Cyan

    # Per-service-principal processing. Emits one object per assignment (or a fallback object).
    # Shared by the sequential and parallel paths.
    $processSp = {
        param($sp, $Prefix = '')

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ${Prefix}Analyzing application: $($sp.DisplayName)" -ForegroundColor Cyan

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
                        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Error retrieving group $($assignment.PrincipalId): $($_.Exception.Message)" -ForegroundColor Red
                        
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
                
                # Emit the assignment object
                [PSCustomObject]$assignmentProps
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
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> Application available to all users (no assignment required)" -ForegroundColor Yellow
            }
            else {
                $assignmentProps.AssignmentType = 'Not Assigned'
                $assignmentProps.PrincipalType = 'None'
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]   -> No assignments found" -ForegroundColor Gray
            }
            
            # Emit the no-assignment object
            [PSCustomObject]$assignmentProps
        }
    }

    if ($useParallel) {
        Write-Verbose "Analyzing applications in parallel (ThrottleLimit: $ThrottleLimit)..."
        $processText = $processSp.ToString()
        $parallelResults = $servicePrincipals | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
            Import-Module Microsoft.Graph.Applications -ErrorAction SilentlyContinue
            Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
            Import-Module Microsoft.Graph.Groups -ErrorAction SilentlyContinue
            $sb = [scriptblock]::Create($using:processText)
            & $sb $_
        }
        foreach ($result in $parallelResults) {
            if ($result) { $applicationAssignmentsArray.Add($result) }
        }
    }
    else {
        $counter = 0
        foreach ($sp in $servicePrincipals) {
            $counter++
            $results = & $processSp $sp "[$counter/$($servicePrincipals.Count)] "
            foreach ($result in $results) {
                if ($result) { $applicationAssignmentsArray.Add($result) }
            }
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
            $_.AssignmentType -eq 'All Users (No Assignment Required)' 
        }
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Filtering applied (AssignmentEmpty): $($applicationAssignmentsArray.Count)/$beforeCount items retained" -ForegroundColor Cyan
    }
    
    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$(if ($ExportPath) { $ExportPath } else { $env:userprofile })\$now-MgApplicationAssignment.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting application assignments to Excel file: $excelFilePath"
        $applicationAssignmentsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraApplicationAssignments'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $applicationAssignmentsArray
    }
}