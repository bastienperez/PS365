
<#
.SYNOPSIS
    Generates a comprehensive report on group-based licensing in Microsoft 365 tenant.

.DESCRIPTION
    This function analyzes all groups with assigned licenses and provides detailed information about
    license assignments, user counts, and licensing errors. It can filter to show only groups
    with licensing errors when needed.

.PARAMETER GroupId
    Optional. Specific group ID to analyze. If not provided, analyzes all groups with licenses.

.PARAMETER ShowErrorsOnly
    Optional. Only return groups that have users with license assignment errors. Default is $false.

.EXAMPLE
    Get-MgGroupLicensingReport
    Generates a complete report for all groups with licenses.

.EXAMPLE
    Get-MgGroupLicensingReport -ShowErrorsOnly
    Shows only groups that have users with license assignment errors.

.EXAMPLE
    Get-MgGroupLicensingReport -GroupId "12345678-1234-1234-1234-123456789012"
    Generates a report for a specific group.

.NOTES
    Requires Microsoft Graph PowerShell SDK with appropriate permissions:
    - Group.Read.All
    - User.Read.All
    - Organization.Read.All
#>
function Get-MgGroupLicensingReport {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$GroupId,
        
        [Parameter()]
        [switch]$ShowErrorsOnly
    )
    
    begin {
        Write-Verbose 'Starting group licensing report analysis...'
        
        try {
            # Pre-fetch all SKUs for performance optimization
            Write-Verbose 'Retrieving tenant SKU information...'
            $skuLookup = @{}
            $skus = Get-MgSubscribedSku -All -Property SkuId, SkuPartNumber
            foreach ($sku in $skus) {
                $skuLookup[$sku.SkuId] = $sku.SkuPartNumber
            }
            Write-Verbose "Found $($skus.Count) SKUs in tenant"
            
            $results = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        catch {
            Write-Error "Failed to initialize: $($_.Exception.Message)"
            return
        }
    }
    
    process {
        try {
            # Get groups with licenses
            if ($GroupId) {
                Write-Verbose "Analyzing specific group: $GroupId"
                $groups = @(Get-MgGroup -GroupId $GroupId -Property DisplayName, Id, AssignedLicenses, LicenseProcessingState)
            }
            else {
                Write-Verbose 'Retrieving all groups with assigned licenses...'
                $groups = Get-MgGroup -All -Property DisplayName, Id, AssignedLicenses, LicenseProcessingState | 
                    Where-Object { $_.AssignedLicenses -and $_.AssignedLicenses.Count -gt 0 }
            }
            
            Write-Verbose "Processing $($groups.Count) groups with licenses..."
            
            foreach ($group in $groups) {
                Write-Verbose "Analyzing group: $($group.DisplayName)"
                
                # Get group members
                $groupMembers = Get-MgGroupMember -GroupId $group.Id -All -Property Id
                $totalMemberCount = $groupMembers.Count
                
                # Initialize counters
                $licensedUserCount = 0
                $errorUserCount = 0
                $errorDetails = [System.Collections.Generic.List[PSCustomObject]]::new()
                
                # Analyze each member's license status
                foreach ($member in $groupMembers) {
                    try {
                        $user = Get-MgUser -UserId $member.Id -Property DisplayName, Id, LicenseAssignmentStates -ErrorAction SilentlyContinue
                        
                        if ($user -and $user.LicenseAssignmentStates) {
                            $userHasGroupLicense = $false
                            $userHasErrors = $false
                            
                            foreach ($licenseState in $user.LicenseAssignmentStates) {
                                # Check if this license was assigned by current group
                                if ($licenseState.AssignedByGroup -eq $group.Id) {
                                    $userHasGroupLicense = $true
                                    
                                    # Check for errors
                                    if ($licenseState.Error -or $licenseState.State -eq 'Error') {
                                        $userHasErrors = $true
                                        
                                        $errorDetail = [PSCustomObject][ordered]@{
                                            UserDisplayName = $user.DisplayName
                                            UserId = $user.Id
                                            SkuId = $licenseState.SkuId
                                            SkuPartNumber = $skuLookup[$licenseState.SkuId]
                                            Error = if ($licenseState.Error) { $licenseState.Error } else { 'License assignment in error state' }
                                            State = $licenseState.State
                                        }
                                        $errorDetails.Add($errorDetail)
                                    }
                                }
                            }
                            
                            if ($userHasGroupLicense) {
                                $licensedUserCount++
                                if ($userHasErrors) {
                                    $errorUserCount++
                                }
                            }
                        }
                    }
                    catch {
                        Write-Warning "Could not analyze user $($member.Id): $($_.Exception.Message)"
                    }
                }
                
                # If ShowErrorsOnly is true, skip groups without errors
                if ($ShowErrorsOnly -and $errorUserCount -eq 0) {
                    continue
                }
                
                # Get license information
                $assignedLicenses = foreach ($license in $group.AssignedLicenses) {
                    $skuPartNumber = $skuLookup[$license.SkuId]
                    if (-not $skuPartNumber) {
                        $skuPartNumber = 'Unknown SKU'
                    }
                    "$skuPartNumber ($($license.SkuId))"
                }
                
                $groupReport = [PSCustomObject][ordered]@{
                    PSTypeName = 'MgGroupLicenseReport'
                    GroupId = $group.Id
                    GroupDisplayName = $group.DisplayName
                    AssignedLicenses = $assignedLicenses -join ', '
                    LicenseCount = $group.AssignedLicenses.Count
                    TotalMemberCount = $totalMemberCount
                    LicensedUserCount = $licensedUserCount
                    ErrorUserCount = $errorUserCount
                    LicenseProcessingState = $group.LicenseProcessingState.State
                    HasErrors = ($errorUserCount -gt 0)
                    ErrorDetails = if ($errorDetails.Count -gt 0) { $errorDetails } else { $null }
                }
                
                $results.Add($groupReport)
            }
        }
        catch {
            Write-Error "Error processing groups: $($_.Exception.Message)"
        }
    }
    
    end {
        Write-Verbose "Analysis complete. Found $($results.Count) groups"
        
        if ($ShowErrorsOnly) {
            Write-Verbose "Filtered to show only groups with licensing errors"
        }
        
        return $results
    }
}