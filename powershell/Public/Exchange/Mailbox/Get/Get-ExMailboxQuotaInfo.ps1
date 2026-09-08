<#
    .SYNOPSIS
    Reports the effective mailbox quota of each user against the storage their licenses entitle them to,
    including the new 100 GB entitlement of the Microsoft 365 Business suites.

    .DESCRIPTION
    Since the 2026 Microsoft 365 packaging updates (rolled out June-September 2026), Business Basic,
    Business Standard and Business Premium include an additional 50 GB of primary mailbox storage,
    bringing the entitlement from 50 GB to 100 GB. The increase is delivered through a dedicated
    service plan: the Business Exchange plan (BPOS_S_STANDARD) grants 50 GB, and the add-on plan
    EXCHANGE_STORAGE_50GB adds the other 50 GB. Both must be enabled for the mailbox to reach 100 GB.

    This function crosses two sources, one row per mailbox:
    - Exchange Online: the effective quotas (IssueWarningQuota, ProhibitSendQuota, ProhibitSendReceiveQuota)
    - Microsoft Graph: the assigned SKUs and the enabled Exchange service plans of the user

    It computes the expected maximum quota of EVERY mailbox from the licenses (100 GB for
    Exchange Online Plan 2 or a Business suite with the storage add-on, 50 GB for Plan 1 alone or a
    license-free shared/room/equipment mailbox, 2 GB for Kiosk) and flags the gaps in the Issue
    column. A mailbox that cannot be evaluated is flagged too, so an empty Issue always means
    'evaluated and consistent':
    - StorageAddOnNotProvisioned: Business suite license but the EXCHANGE_STORAGE_50GB plan is not
      enabled (disabled in the license options, or provisioning not completed yet)
    - QuotaBelowEntitlement: the add-on is there but ProhibitSendReceiveQuota is still below the
      entitlement (custom quota set by an administrator, preserved by design, or provisioning pending)
    - QuotaAboveEntitlement: ProhibitSendReceiveQuota exceeds what the current licenses entitle the
      user to (backfill applied on top of an already raised quota, or a 100 GB license since removed).
      Informational: the mailbox has more than its entitlement, not less
    - EntitlementUnknown (<plans>): the user holds an Exchange service plan whose quota is not in the
      internal map: the row is NOT evaluated (report it so the map can be extended)
    - NoExchangeLicense: user mailbox without any Exchange service plan

    Quotas are never additive across products: a user holding both a Business suite and an Enterprise
    plan is entitled to 100 GB, not 150. Custom quotas set by an administrator are preserved by the
    rollout, which is exactly what this report surfaces. The increase concerns the primary mailbox
    only (no archive change).

    .PARAMETER Identity
    (Optional) One or more user principal names. When omitted, every user, shared, room and
    equipment mailbox of the tenant is evaluated.

    .PARAMETER IncludeUsage
    (Optional) Also retrieves the current mailbox usage (Get-EXOMailboxStatistics): columns
    TotalItemSize, TotalItemSizeGB and UsagePercent (against ProhibitSendReceiveQuota).
    One extra call per mailbox: noticeably slower on large tenants.

    .PARAMETER OnlyIssues
    (Optional) Returns only the mailboxes with a non-empty Issue column.

    .PARAMETER ForceNewToken
    Switch parameter to force getting a new token from Microsoft Graph.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER ExportPath
    (Optional) Output directory for the Excel export. Defaults to the user profile.

    .EXAMPLE
    Get-ExMailboxQuotaInfo

    Quotas versus license entitlement for every user, shared, room and equipment mailbox of the tenant.

    .EXAMPLE
    Get-ExMailboxQuotaInfo -OnlyIssues

    Only the mailboxes whose quota does not match what the licenses entitle them to
    (Business suite without the +50 GB add-on, or quota still below the entitlement).

    .EXAMPLE
    Get-ExMailboxQuotaInfo -Identity 'user@contoso.com'

    The report for a single user.

    .EXAMPLE
    Get-ExMailboxQuotaInfo -IncludeUsage

    Same report with the current size of each mailbox and its usage percentage.

    .EXAMPLE
    Get-ExMailboxQuotaInfo -ExportToExcel

    Exports the report to an Excel file in the user's profile directory.

    .OUTPUTS
    System.Collections.Generic.List[PSCustomObject]

    .NOTES
    OUTPUT PROPERTIES
    - UserPrincipalName        : the mailbox owner
    - DisplayName              : display name of the mailbox
    - RecipientTypeDetails     : UserMailbox, SharedMailbox...
    - ProhibitSendQuota        : effective send quota reported by Exchange
    - ProhibitSendReceiveQuota : effective maximum size reported by Exchange
    - IssueWarningQuota        : warning threshold reported by Exchange
    - UseDatabaseQuotaDefaults : whether the database defaults apply
    - TotalItemSize            : current mailbox size (with -IncludeUsage, empty otherwise)
    - TotalItemSizeGB          : current mailbox size in GB (with -IncludeUsage)
    - UsagePercent             : TotalItemSizeGB / ProhibitSendReceiveQuota (with -IncludeUsage)
    - Licenses                 : SKU part numbers assigned to the user
    - ExchangeServicePlans     : enabled Exchange service plans (BPOS_S_STANDARD, EXCHANGE_STORAGE_50GB...)
    - HasBusinessSuite         : the user holds Business Basic, Business Standard or Business Premium
    - HasStorageAddOn          : the EXCHANGE_STORAGE_50GB service plan is enabled
    - ExpectedMaxQuotaGB       : maximum quota the licenses entitle the user to (empty when unknown)
    - Issue                    : StorageAddOnNotProvisioned, QuotaBelowEntitlement, QuotaAboveEntitlement,
                                 EntitlementUnknown, NoExchangeLicense, or empty (evaluated and consistent)

    Required Microsoft Graph permissions:
        - User.Read.All
        - Organization.Read.All

    Requires an Exchange Online connection (Connect-ExchangeOnline is called when none exists).
    Mailboxes are retrieved with Get-EXOMailbox (REST) for performance on large tenants.

    Reference: https://techcommunity.microsoft.com/blog/exchange/understanding-the-new-100-gb-mailbox-entitlement-for-microsoft-365-business-suit/4548243

    Version history:
        1.0 - Creation. Crosses Exchange quotas with Graph service plans to check the 100 GB
              Business suites entitlement (BPOS_S_STANDARD + EXCHANGE_STORAGE_50GB), flags the
              missing add-on and the quotas below entitlement (custom quotas are preserved by
              the rollout by design).

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExMailboxQuotaInfo
#>
function Get-ExMailboxQuotaInfo {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[PSCustomObject]])]
    param (
        [Parameter(Mandatory = $false, Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Identity,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeUsage,

        [Parameter(Mandatory = $false)]
        [switch]$OnlyIssues,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath
    )

    begin {
        $identityList = [System.Collections.Generic.List[string]]::new()

        # Business suites eligible for the +50 GB storage add-on (standalone Exchange Online Plan 1
        # uses the same Exchange plan but stays at 50 GB, hence the SKU-level check)
        $businessSuiteSkus = @('O365_BUSINESS_ESSENTIALS', 'O365_BUSINESS_PREMIUM', 'SPB')

        # Documented primary mailbox entitlement per Exchange service plan (GB). Graph exposes the
        # EXCHANGE_S_* names; the BPOS_S_* provisioning names used by the Exchange blog are kept as aliases
        $exchangePlanQuotaGB = @{
            'EXCHANGE_S_ENTERPRISE' = 100
            'BPOS_S_ENTERPRISE'     = 100
            'EXCHANGE_S_STANDARD'   = 50
            'BPOS_S_STANDARD'       = 50
            'EXCHANGE_S_DESKLESS'   = 2
            'BPOS_S_DESKLESS'       = 2
        }

        # Exchange plans that carry no primary mailbox quota of their own
        $quotaNeutralPlans = @('EXCHANGE_S_FOUNDATION', 'EXCHANGE_STORAGE_50GB')

        # Converts an Exchange quota string like '99 GB (106,300,440,576 bytes)' to GB
        function ConvertTo-QuotaGB {
            param ([string]$Quota)

            if ([string]::IsNullOrWhiteSpace($Quota) -or $Quota -eq 'Unlimited') {
                return $null
            }

            if ($Quota -match '\(([\d,]+)\s*bytes\)') {
                return [math]::Round(([long]($Matches[1] -replace ',', '')) / 1GB, 2)
            }

            return $null
        }
    }

    process {
        foreach ($userPrincipalName in $Identity) {
            $identityList.Add($userPrincipalName)
        }
    }

    end {
        try {
            $null = Import-Module 'Microsoft.Graph.Authentication' -ErrorAction Stop
        }
        catch {
            Write-Warning 'Please install Microsoft.Graph.Authentication first'
            return
        }

        $permissionsNeeded = @('User.Read.All', 'Organization.Read.All')

        $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
        if ($ForceNewToken.IsPresent) {
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
            $isConnected = $false
        }
        if (-not $isConnected) {
            Write-Host -ForegroundColor Cyan 'Connecting to Microsoft Graph'
            $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
        }

        if (-not (Test-MgGraphPermission -RequiredScopes $permissionsNeeded -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }

        if (-not (Get-ConnectionInformation | Where-Object { $_.ConnectionUri -eq 'https://outlook.office365.com' })) {
            Connect-ExchangeOnline
        }

        # subscribedSkus gives both maps: skuId -> SkuPartNumber and servicePlanId -> ServicePlanName
        # (users expose assignedPlans by servicePlanId only)
        Write-Host -ForegroundColor Cyan 'Getting the tenant subscribed SKUs'
        $skuPartNumberById = @{}
        $servicePlanNameById = @{}

        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -ErrorAction Stop

            foreach ($subscribedSku in @($response.value)) {
                $skuPartNumberById["$($subscribedSku.skuId)"] = $subscribedSku.skuPartNumber

                foreach ($servicePlan in @($subscribedSku.servicePlans)) {
                    $servicePlanNameById["$($servicePlan.servicePlanId)"] = $servicePlan.servicePlanName
                }
            }
        }
        catch {
            Write-Warning "Unable to get the subscribed SKUs. $($_.Exception.Message)"
            return
        }

        # Licenses and enabled plans of the users, indexed by UPN for the mailbox loop
        Write-Host -ForegroundColor Cyan 'Getting the users licenses and service plans'
        $licensesByUpn = @{}
        $selectClause = 'id,userPrincipalName,assignedLicenses,assignedPlans'

        try {
            if ($identityList.Count -gt 0) {
                $graphUsers = [System.Collections.Generic.List[PSObject]]::new()
                foreach ($userPrincipalName in $identityList) {
                    $singleUserUri = "https://graph.microsoft.com/v1.0/users/$userPrincipalName`?`$select=$selectClause"
                    $graphUsers.Add((Invoke-MgGraphRequest -Method GET -Uri $singleUserUri -ErrorAction Stop))
                }
            }
            else {
                $graphUsers = [System.Collections.Generic.List[PSObject]]::new()
                $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectClause&`$top=999"
                do {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                    foreach ($graphUser in @($response.value)) {
                        $graphUsers.Add($graphUser)
                    }
                    $uri = $response.'@odata.nextLink'
                } while ($uri)
            }
        }
        catch {
            Write-Warning "Unable to get the users from Microsoft Graph. $($_.Exception.Message)"
            return
        }

        foreach ($graphUser in $graphUsers) {
            $skuNames = [System.Collections.Generic.List[string]]::new()
            foreach ($assignedLicense in @($graphUser.assignedLicenses)) {
                $skuName = $skuPartNumberById["$($assignedLicense.skuId)"]
                if ($skuName) {
                    $skuNames.Add($skuName)
                }
            }

            # assignedPlans keeps the deleted plans too: only Enabled ones count
            $exchangePlanNames = [System.Collections.Generic.List[string]]::new()
            foreach ($assignedPlan in @($graphUser.assignedPlans)) {
                if ("$($assignedPlan.capabilityStatus)" -ne 'Enabled') {
                    continue
                }

                $planName = $servicePlanNameById["$($assignedPlan.servicePlanId)"]
                if ($planName -and ($planName -like 'BPOS_S_*' -or $planName -like 'EXCHANGE_*')) {
                    $exchangePlanNames.Add($planName)
                }
            }

            $licensesByUpn["$($graphUser.userPrincipalName)"] = [PSCustomObject]@{
                SkuNames          = $skuNames
                ExchangePlanNames = $exchangePlanNames
            }
        }

        if (-not $IncludeUsage.IsPresent) {
            Write-Host -ForegroundColor Yellow 'Current mailbox usage is not retrieved (one extra call per mailbox, slow on large tenants). Use -IncludeUsage to get the TotalItemSize and UsagePercent columns.'
        }

        Write-Host -ForegroundColor Cyan 'Getting the mailboxes and their quotas'

        # Get-EXOMailbox (REST) is much faster than Get-Mailbox on large tenants, but only returns
        # a minimal property set by default: the quotas must be requested explicitly
        $mailboxProperties = @('DisplayName', 'RecipientTypeDetails', 'ProhibitSendQuota', 'ProhibitSendReceiveQuota', 'IssueWarningQuota', 'UseDatabaseQuotaDefaults')

        try {
            if ($identityList.Count -gt 0) {
                $exoMailboxes = [System.Collections.Generic.List[PSObject]]::new()
                foreach ($userPrincipalName in $identityList) {
                    try {
                        $exoMailboxes.Add((Get-EXOMailbox -Identity $userPrincipalName -Properties $mailboxProperties -ErrorAction Stop))
                    }
                    catch {
                        Write-Warning "Mailbox not found: $userPrincipalName"
                    }
                }
            }
            else {
                $exoMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox -Properties $mailboxProperties -ErrorAction Stop)
            }
        }
        catch {
            Write-Warning "Unable to get the mailboxes. $($_.Exception.Message)"
            return
        }

        $entitlementArray = [System.Collections.Generic.List[PSCustomObject]]::new()
        $mailboxIndex = 0

        foreach ($mailbox in $exoMailboxes) {
            $mailboxIndex++

            # One statistics call per mailbox: without feedback a large tenant looks frozen.
            # Write-Host every 100 mailboxes rather than Write-Progress (too slow per-item)
            if ($IncludeUsage.IsPresent -and ($mailboxIndex % 100 -eq 0)) {
                Write-Host -ForegroundColor Cyan "$mailboxIndex/$($exoMailboxes.Count) mailbox(es) processed"
            }

            $userLicenses = $licensesByUpn["$($mailbox.UserPrincipalName)"]
            $skuNames = @(if ($userLicenses) { $userLicenses.SkuNames })
            $exchangePlanNames = @(if ($userLicenses) { $userLicenses.ExchangePlanNames })

            $hasBusinessSuite = @($skuNames | Where-Object { $_ -in $businessSuiteSkus }).Count -gt 0
            $hasStorageAddOn = $exchangePlanNames -contains 'EXCHANGE_STORAGE_50GB'

            # Quotas are not additive across products: the highest single entitlement wins.
            # Plans absent from the map are collected so the row is flagged instead of silently passing
            $baseQuotaGB = $null
            $unknownPlanNames = [System.Collections.Generic.List[string]]::new()

            foreach ($planName in $exchangePlanNames) {
                if ($exchangePlanQuotaGB.ContainsKey($planName)) {
                    if ($null -eq $baseQuotaGB -or $exchangePlanQuotaGB[$planName] -gt $baseQuotaGB) {
                        $baseQuotaGB = $exchangePlanQuotaGB[$planName]
                    }
                }
                elseif ($planName -notin $quotaNeutralPlans -and $planName -notlike 'EXCHANGE_S_ARCHIVE*' -and $planName -notlike 'EXCHANGE_ANALYTICS*') {
                    $unknownPlanNames.Add($planName)
                }
            }

            $expectedMaxQuotaGB = $baseQuotaGB
            if ($baseQuotaGB -eq 50 -and $hasStorageAddOn) {
                $expectedMaxQuotaGB = 100
            }

            # License-free mailboxes (shared, room, equipment) are entitled to 50 GB
            if ($null -eq $expectedMaxQuotaGB -and $unknownPlanNames.Count -eq 0 -and "$($mailbox.RecipientTypeDetails)" -in @('SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')) {
                $expectedMaxQuotaGB = 50
            }

            $prohibitSendReceiveQuotaGB = ConvertTo-QuotaGB -Quota "$($mailbox.ProhibitSendReceiveQuota)"

            $totalItemSize = ''
            $totalItemSizeGB = $null
            $usagePercent = $null

            if ($IncludeUsage.IsPresent) {
                try {
                    $mailboxStatistics = Get-EXOMailboxStatistics -Identity "$($mailbox.UserPrincipalName)" -ErrorAction Stop
                    $totalItemSize = "$($mailboxStatistics.TotalItemSize)"
                    $totalItemSizeGB = ConvertTo-QuotaGB -Quota $totalItemSize

                    # Zero is a valid result for very small mailboxes after rounding to GB.
                    # Test for missing values explicitly so UsagePercent becomes 0 instead of $null.
                    if ($null -ne $totalItemSizeGB -and $null -ne $prohibitSendReceiveQuotaGB -and $prohibitSendReceiveQuotaGB -gt 0) {
                        $usagePercent = [math]::Round(($totalItemSizeGB / $prohibitSendReceiveQuotaGB) * 100, 1)
                    }
                }
                catch {
                    Write-Warning "Unable to get the statistics for $($mailbox.UserPrincipalName). $($_.Exception.Message)"
                }
            }

            $issue = ''
            if ($null -eq $expectedMaxQuotaGB) {
                # An empty Issue must always mean 'evaluated and consistent', never 'not evaluated'
                if ($unknownPlanNames.Count -gt 0) {
                    $issue = 'EntitlementUnknown ({0})' -f (($unknownPlanNames | Sort-Object) -join ', ')
                }
                else {
                    $issue = 'NoExchangeLicense'
                }
            }
            elseif ($hasBusinessSuite -and -not $hasStorageAddOn) {
                $issue = 'StorageAddOnNotProvisioned'
            }
            elseif ($expectedMaxQuotaGB -and $prohibitSendReceiveQuotaGB -and $prohibitSendReceiveQuotaGB -lt $expectedMaxQuotaGB) {
                # Custom quotas set by an administrator are preserved by the rollout by design
                $issue = 'QuotaBelowEntitlement'
            }
            elseif ($expectedMaxQuotaGB -and $prohibitSendReceiveQuotaGB -and $prohibitSendReceiveQuotaGB -gt $expectedMaxQuotaGB) {
                # Seen in the field during the 2026 backfill: +50 GB applied on top of a mailbox already
                # at 100 GB (e.g. 150 GB with IssueWarningQuota still at 98), or a license since removed.
                # Not blocking (the mailbox has more than its entitlement), surfaced for awareness
                $issue = 'QuotaAboveEntitlement'
            }

            $object = [PSCustomObject][ordered]@{
                UserPrincipalName        = $mailbox.UserPrincipalName
                DisplayName              = $mailbox.DisplayName
                RecipientTypeDetails     = "$($mailbox.RecipientTypeDetails)"
                ProhibitSendQuota        = "$($mailbox.ProhibitSendQuota)"
                ProhibitSendReceiveQuota = "$($mailbox.ProhibitSendReceiveQuota)"
                IssueWarningQuota        = "$($mailbox.IssueWarningQuota)"
                UseDatabaseQuotaDefaults = $mailbox.UseDatabaseQuotaDefaults
                TotalItemSize            = $totalItemSize
                TotalItemSizeGB          = $totalItemSizeGB
                UsagePercent             = $usagePercent
                Licenses                 = ($skuNames | Sort-Object) -join '|'
                ExchangeServicePlans     = ($exchangePlanNames | Sort-Object) -join '|'
                HasBusinessSuite         = $hasBusinessSuite
                HasStorageAddOn          = $hasStorageAddOn
                ExpectedMaxQuotaGB       = $expectedMaxQuotaGB
                Issue                    = $issue
            }

            $entitlementArray.Add($object)
        }

        if ($OnlyIssues.IsPresent) {
            $entitlementArray = [System.Collections.Generic.List[PSCustomObject]]@(
                $entitlementArray | Where-Object { $_.Issue }
            )
        }

        $issueCount = @($entitlementArray | Where-Object { $_.Issue }).Count
        if ($issueCount -gt 0) {
            Write-Host -ForegroundColor Yellow "$issueCount mailbox(es) with a storage entitlement gap"
        }
        else {
            Write-Host -ForegroundColor Green 'No storage entitlement gap detected'
        }

        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $exportDirectory = if ($ExportPath) { $ExportPath } else { $env:userprofile }
            $excelFilePath = Join-Path -Path $exportDirectory -ChildPath "$now-ExMailboxQuotaInfo.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting the report to Excel file: $excelFilePath"
            $entitlementArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Ex-QuotaInfo' -TableStyle Light9
            Write-Host -ForegroundColor Green 'Export completed successfully!'
        }
        else {
            return $entitlementArray
        }
    }
}
