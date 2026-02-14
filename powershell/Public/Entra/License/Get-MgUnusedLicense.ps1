<#
    .SYNOPSIS
    Detects users with assigned licenses who have not signed in for a specified number of days.

    .DESCRIPTION
    Retrieves all licensed users from Microsoft Graph and identifies those who have never signed in or have been inactive
    beyond a configurable threshold. The function calculates estimated wasted costs based on common license SKUs and
    provides recommendations for license optimization.

    This function helps administrators identify:
    - Users who have never signed in (potential orphaned accounts)
    - Users inactive beyond the specified threshold (potential license waste)
    - Estimated monthly and annual cost savings from license reclamation

    .PARAMETER InactiveDays
    The number of days of inactivity after which a user is considered inactive. Default is 90 days.

    .PARAMETER FilterByDomain
    Filters users by a specific domain. Only users from the specified domain will be analyzed (excluding guest users).

    .PARAMETER IncludeDisabledAccounts
    If specified, includes disabled accounts in the analysis. By default, only enabled accounts are analyzed.

    .PARAMETER IncludeGuestUsers
    If specified, includes guest users (#EXT#) in the analysis. By default, guest users are excluded.

    .PARAMETER ExportToExcel
    If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-MgUnusedLicense

    Retrieves all users with licenses who have been inactive for more than 90 days (default) and displays the results.

    .EXAMPLE
    Get-MgUnusedLicense -InactiveDays 60 -ExportToExcel

    Retrieves users inactive for more than 60 days and exports the results to an Excel report.

    .EXAMPLE
    Get-MgUnusedLicense -FilterByDomain 'contoso.com' -InactiveDays 30

    Retrieves users from the contoso.com domain who have been inactive for more than 30 days.

    .EXAMPLE
    Get-MgUnusedLicense -IncludeDisabledAccounts -IncludeGuestUsers

    Retrieves all users including disabled accounts and guest users who have unused licenses.

    .NOTES
    Requires Connect-MgGraph with scopes: 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All'.
    
    License cost estimates are approximate and based on common retail prices:
    - ENTERPRISEPREMIUM (E5): $57/month
    - ENTERPRISEPACK (E3): $23/month
    - SPE_E5 (Microsoft 365 E5): $57/month
    - SPE_E3 (Microsoft 365 E3): $36/month
    - POWERAPPS_PER_USER: $20/month
    - FLOW_PER_USER: $15/month
    
    For accurate pricing, consult your Microsoft licensing agreement.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgUnusedLicense
#>

function Get-MgUnusedLicense {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 365)]
        [int]$InactiveDays = 90,

        [Parameter(Mandatory = $false)]
        [string]$FilterByDomain,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeDisabledAccounts,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeGuestUsers,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    # Import required modules
    $modules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )
    
    foreach ($module in $modules) {
        try {
            $null = Import-Module $module -ErrorAction Stop
        }
        catch {
            Write-Warning "Please install $module first"
            return
        }
    }

    # Check connection status
    if (-not (Get-MgContext)) {
        Write-Host -ForegroundColor Cyan 'Connecting to Microsoft Graph'
        Connect-MgGraph -Scopes 'User.Read.All', 'Directory.Read.All', 'AuditLog.Read.All' -NoWelcome
    }

    # Build license SKU mapping with estimated prices
    Write-Host -ForegroundColor Cyan 'Retrieving license SKU information'
    $subscribedSkus = Get-MgSubscribedSku -All
    $skuHashTable = @{}

    foreach ($sku in $subscribedSkus) {
        $unitPrice = switch ($sku.SkuPartNumber) {
            'ENTERPRISEPREMIUM' { 57; break }      # E5
            'ENTERPRISEPACK' { 23; break }          # E3
            'SPE_E5' { 57; break }                  # Microsoft 365 E5
            'SPE_E3' { 36; break }                  # Microsoft 365 E3
            'POWERAPPS_PER_USER' { 20; break }      # Power Apps per user
            'FLOW_PER_USER' { 15; break }           # Power Automate per user
            'POWER_BI_PRO' { 10; break }            # Power BI Pro
            'PROJECTPREMIUM' { 55; break }          # Project Plan 5
            'VISIOONLINE_PLAN2' { 15; break }       # Visio Plan 2
            'EXCHANGEENTERPRISE' { 8; break }       # Exchange Online Plan 2
            'SHAREPOINTENTERPRISE' { 10; break }    # SharePoint Online Plan 2
            default { 0; break }
        }

        $skuHashTable[$sku.SkuId] = [PSCustomObject][ordered]@{
            SkuPartNumber = $sku.SkuPartNumber
            SkuId         = $sku.SkuId
            ConsumedUnits = $sku.ConsumedUnits
            PrepaidUnits  = $sku.PrepaidUnits.Enabled
            UnitPrice     = $unitPrice
        }
    }

    # Retrieve users with licenses
    $userParams = 'Id,DisplayName,UserPrincipalName,AccountEnabled,CreatedDateTime,SignInActivity,AssignedLicenses,UserType'

    if ($FilterByDomain) {
        Write-Host -ForegroundColor Cyan "Filtering by domain: $FilterByDomain"
        $mgUsersList = Get-MgUser -Filter "endswith(userPrincipalName,'$FilterByDomain') and not endswith(userPrincipalName,'#EXT#@$FilterByDomain')" -All -Property $userParams -ConsistencyLevel eventual
    }
    else {
        $mgUsersList = Get-MgUser -Filter "assignedLicenses/`$count ne 0" -All -Property $userParams -ConsistencyLevel eventual -CountVariable userCount
    }

    # Filter users with licenses (for domain filter case)
    $mgUsersList = $mgUsersList | Where-Object { $_.AssignedLicenses.Count -gt 0 }

    # Filter guest users unless explicitly included
    if (-not $IncludeGuestUsers) {
        $mgUsersList = $mgUsersList | Where-Object { $_.UserPrincipalName -notmatch '#EXT#' }
    }

    # Filter disabled accounts unless explicitly included
    if (-not $IncludeDisabledAccounts) {
        $mgUsersList = $mgUsersList | Where-Object { $_.AccountEnabled -eq $true }
    }

    # Analyze users
    
    [System.Collections.Generic.List[PSCustomObject]]$unusedLicensesList = @()
    $neverSignedInCount = 0
    $inactiveCount = 0
    $totalWastedCost = 0

    foreach ($mgUser in $mgUsersList) {
        $lastSignIn = $null
        $daysSinceLastSignIn = $null
        $status = $null
        $wastedCost = 0
        $recommendation = $null

        # Check sign-in activity
        if ($mgUser.SignInActivity -and $mgUser.SignInActivity.LastSignInDateTime) {
            $lastSignIn = $mgUser.SignInActivity.LastSignInDateTime
            
            # Check for Microsoft Graph null date
            if ($lastSignIn -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
                $daysSinceLastSignIn = ((Get-Date) - $lastSignIn).Days

                if ($daysSinceLastSignIn -gt $InactiveDays) {
                    $status = 'Inactive'
                    $inactiveCount++
                    $recommendation = "Audit - Inactive for $daysSinceLastSignIn days"
                }
                else {
                    # User is active, skip
                    continue
                }
            }
            else {
                $lastSignIn = $null
                $status = 'NeverSignedIn'
                $neverSignedInCount++
                $recommendation = 'Remove - Never used'
            }
        }
        else {
            $status = 'NeverSignedIn'
            $neverSignedInCount++
            $recommendation = 'Remove - Never used'
        }

        # Get license details and calculate wasted cost
        [System.Collections.Generic.List[string]]$licenseNamesList = @()
        
        foreach ($license in $mgUser.AssignedLicenses) {
            $skuInfo = $skuHashTable[$license.SkuId]
            if ($skuInfo) {
                $licenseNamesList.Add($skuInfo.SkuPartNumber)
                $wastedCost += $skuInfo.UnitPrice
            }
        }

        $totalWastedCost += $wastedCost

        # Calculate days since creation
        $daysSinceCreation = $null
        if ($mgUser.CreatedDateTime -and $mgUser.CreatedDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
            $daysSinceCreation = ((Get-Date) - $mgUser.CreatedDateTime).Days
        }

        $object = [PSCustomObject][ordered]@{
            UserPrincipalName       = $mgUser.UserPrincipalName
            DisplayName             = $mgUser.DisplayName
            Id                      = $mgUser.Id
            AccountEnabled          = $mgUser.AccountEnabled
            UserType                = $mgUser.UserType
            Status                  = $status
            LastSignInDateUTC       = $lastSignIn
            DaysSinceLastSignIn     = $daysSinceLastSignIn
            CreatedDateUTC          = if ($mgUser.CreatedDateTime -and $mgUser.CreatedDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) { $mgUser.CreatedDateTime } else { $null }
            DaysSinceCreation       = $daysSinceCreation
            AssignedLicenses        = ($licenseNamesList -join ', ')
            LicenseCount            = $mgUser.AssignedLicenses.Count
            EstimatedMonthlyCostUSD = $wastedCost
            EstimatedAnnualCostUSD  = $wastedCost * 12
            Recommendation          = $recommendation
        }

        $unusedLicensesList.Add($object)
    }

    # Sort results: never signed in first, then by days since last sign-in
    $unusedLicensesList = $unusedLicensesList | Sort-Object @{Expression = { $_.Status -eq 'NeverSignedIn' }; Descending = $true }, DaysSinceLastSignIn -Descending

    # Export or return results
    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgUnusedLicense_Report.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $unusedLicensesList | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-UnusedLicenses'
        Write-Host -ForegroundColor Green 'Export completed'
    }
    else {
        return $unusedLicensesList
    }
}
