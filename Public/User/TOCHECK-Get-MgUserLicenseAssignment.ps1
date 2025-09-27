<#
.SYNOPSIS
    Retrieves detailed license assignment information for users in Microsoft 365 tenant.

.DESCRIPTION
    This function analyzes license assignments for all users in the tenant, identifying whether licenses 
    were assigned directly to users or through group-based licensing. It provides comprehensive reporting 
    on license distribution and assignment methods.

.PARAMETER UserId
    Optional. Specific user ID to filter results. If not provided, retrieves data for all users.

.PARAMETER ExcludeDisabledUsers
    Optional. Exclude disabled users from the results. Default is $false.

.PARAMETER ShowErrorsOnly
    Optional. Only return users with license assignment errors or issues. Default is $false.

.EXAMPLE
    Get-MgUserLicenseAssignment
    Retrieves license assignments for all users in the tenant.

.EXAMPLE
    Get-MgUserLicenseAssignment -UserId "user@domain.com"
    Retrieves license assignments for a specific user.

.EXAMPLE
    Get-MgUserLicenseAssignment -ExcludeDisabledUsers
    Retrieves license assignments for enabled users only.

.EXAMPLE
    Get-MgUserLicenseAssignment -ShowErrorsOnly
    Retrieves only users with license assignment errors or issues.

.NOTES
    Requires Microsoft Graph PowerShell SDK with appropriate permissions:
    - User.Read.All
    - Group.Read.All
    - Organization.Read.All
#>
function Get-MgUserLicenseAssignment {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [string]$UserId,
        
        [Parameter()]
        [switch]$ExcludeDisabledUsers,
        
        [Parameter()]
        [switch]$ShowErrorsOnly
    )
    
    begin {
        Write-Verbose 'Starting license assignment analysis...'
        
        try {
            # Pre-fetch all SKUs for performance optimization
            Write-Verbose 'Retrieving tenant SKU information...'
            $skuLookup = @{}
            $skus = Get-MgSubscribedSku -All -Property SkuId, SkuPartNumber
            foreach ($sku in $skus) {
                $skuLookup[$sku.SkuId] = $sku.SkuPartNumber
            }
            Write-Verbose "Found $($skus.Count) SKUs in tenant"
            
            # Pre-fetch groups for better performance when resolving group assignments
            Write-Verbose 'Building group lookup cache...'
            $groupLookup = @{}
            $groups = Get-MgGroup -All -Property Id, DisplayName
            foreach ($group in $groups) {
                $groupLookup[$group.Id] = $group.DisplayName
            }
            Write-Verbose "Cached $($groups.Count) groups"
            
            $results = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        catch {
            Write-Error "Failed to initialize: $($_.Exception.Message)"
            return
        }
    }
    
    process {
        try {
            # Build user filter parameters
            $userParams = @{
                All      = $true
                Property = 'AssignedLicenses', 'LicenseAssignmentStates', 'DisplayName', 'Id', 'UserPrincipalName', 'AccountEnabled'
            }
            
            if ($UserId) {
                $userParams.Remove('All')
                $userParams.UserId = $UserId
                $users = @(Get-MgUser @userParams)
            }
            else {
                $users = Get-MgUser @userParams
                if ($ExcludeDisabledUsers) {
                    $users = $users | Where-Object { $_.AccountEnabled -eq $true }
                }
            }
            
            Write-Verbose "Processing $($users.Count) users..."
            
            foreach ($user in $users) {
                if (-not $user.LicenseAssignmentStates) {
                    continue # Skip users without licenses
                }
                
                $licenseAssignments = @{}
                $hasErrors = $false
                
                foreach ($assignment in $user.LicenseAssignmentStates) {
                    $skuId = $assignment.SkuId
                    
                    # Check for license assignment errors
                    if ($assignment.Error -or $assignment.State -eq 'Error') {
                        $hasErrors = $true
                    }
                    
                    $assignmentMethod = if ($assignment.AssignedByGroup) {
                        $groupName = $groupLookup[$assignment.AssignedByGroup]
                        if ($groupName) { $groupName } else { 'Unknown Group' }
                    }
                    else {
                        'Direct Assignment'
                    }
                    
                    if (-not $licenseAssignments.ContainsKey($skuId)) {
                        $licenseAssignments[$skuId] = @{
                            Methods = [System.Collections.Generic.List[string]]::new()
                            HasErrors = $false
                            ErrorMessage = $null
                        }
                    }
                    $licenseAssignments[$skuId].Methods.Add($assignmentMethod)
                    
                    if ($assignment.Error -or $assignment.State -eq 'Error') {
                        $licenseAssignments[$skuId].HasErrors = $true
                        $licenseAssignments[$skuId].ErrorMessage = if ($assignment.Error) { $assignment.Error } else { 'License assignment in error state' }
                    }
                }
                
                # If ShowErrorsOnly is true, skip users without errors
                if ($ShowErrorsOnly -and -not $hasErrors) {
                    continue
                }
                
                foreach ($skuId in $licenseAssignments.Keys) {
                    $skuPartNumber = $skuLookup[$skuId]
                    if (-not $skuPartNumber) {
                        $skuPartNumber = 'Unknown SKU'
                    }
                    
                    $assignmentMethods = ($licenseAssignments[$skuId].Methods | Sort-Object -Unique) -join ', '
                    
                    $userLicenseInfo = [PSCustomObject][ordered]@{
                        PSTypeName        = 'MgUserLicenseAssignment'
                        UserId            = $user.Id
                        UserDisplayName   = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                        AccountEnabled    = $user.AccountEnabled
                        SkuId             = $skuId
                        SkuPartNumber     = $skuPartNumber
                        AssignedBy        = $assignmentMethods
                        HasErrors         = $licenseAssignments[$skuId].HasErrors
                        ErrorMessage      = $licenseAssignments[$skuId].ErrorMessage
                    }
                    
                    $results.Add($userLicenseInfo)
                }
            }
        }
        catch {
            Write-Error "Error processing users: $($_.Exception.Message)"
        }
    }
    
    end {
        Write-Verbose "Analysis complete. Found $($results.Count) license assignments"
        
        return $results
    }
}