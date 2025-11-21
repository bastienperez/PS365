function New-MgAuditLogSignInHTMLReport {
    <#
    .SYNOPSIS
        Generate an HTML report from a collection of sign-in records using PSWriteHTML.

    .DESCRIPTION
        Accepts an array of sign-in objects (raw Graph objects or transformed PSCustomObjects)
        and writes an HTML report with PSWriteHTML containing a chart and detailed table.

    .PARAMETER SignIns
        Array of sign-in objects.

    .PARAMETER OutputFile
        Path to the HTML output file.
    #>
    param(
        [Parameter(Mandatory = $true)][object[]]$SignIns,
        [Parameter(Mandatory = $true)][string]$OutputFile
    )

    # Import PSWriteHTML (required)
    try {
        Import-Module PSWriteHTML -ErrorAction Stop
    }
    catch {
        throw 'PSWriteHTML module is required but not available. Install it with: Install-Module PSWriteHTML'
    }

    # Aggregate by day and category
    $categories = @('Success', 'ReportOnly', 'NotApplied', 'Failure', 'Other')
    $daily = @{}

    foreach ($s in $SignIns) {
        # try to get a Date key in yyyy-MM-dd
        try { $d = ([datetime]$s.CreatedDateTime).ToString('yyyy-MM-dd') } catch { $d = (Get-Date).ToString('yyyy-MM-dd') }

        if (-not $daily.ContainsKey($d)) { $daily[$d] = @{ Success = 0; ReportOnly = 0; NotApplied = 0; Failure = 0; Other = 0 } }

        # classify each sign-in
        $cat = 'Other'

        # Failure detection
        if ($s.Status -and $s.Status.ErrorCode -ne $null -and ($s.Status.ErrorCode -ne 0)) {
            $cat = 'Failure'
        }
        else {
            # Look for applied conditional access policy results if present
            if ($s.AppliedConditionalAccessPolicies) {
                $found = $false
                foreach ($p in $s.AppliedConditionalAccessPolicies) {
                    if ($p.Result -like 'reportOnly*' -and $p.Result -ne 'reportOnlyNotApplied') { $cat = 'ReportOnly'; $found = $true; break }
                    if ($p.Result -eq 'notApplied') { $cat = 'NotApplied'; $found = $true; break }
                    if ($p.Result -eq 'success') { $cat = 'Success'; $found = $true; break }
                }
                if (-not $found) { $cat = 'Other' }
            }
            else {
                # If no applied policies and no failure, consider it NotApplied (no CA enforcement)
                $cat = 'NotApplied'
            }
        }

        $daily[$d][$cat] += 1
    }

    # Build labels (sorted dates) and series arrays
    $labels = $daily.Keys | Sort-Object
    $seriesData = @{}
    foreach ($c in $categories) { $seriesData[$c] = @() }

    foreach ($dateKey in $labels) {
        foreach ($c in $categories) {
            $seriesData[$c] += $daily[$dateKey][$c]
        }
    }

    # Prepare table data for PSWriteHTML
    $tableData = $SignIns | Select-Object `
    @{Name = 'Time'; Expression = { [datetime]$_.CreatedDateTime } },
    UserDisplayName,
    UserPrincipalName,
    AppDisplayName,
    IpAddress,
    ClientAppUsed,
    @{Name = 'PolicyResult'; Expression = { $_.AppliedConditionalAccessPolicies -join '; ' } },
    @{Name = 'ErrorCode'; Expression = { $_.Status.ErrorCode } },
    @{Name = 'FailureReason'; Expression = { $_.Status.FailureReason } }

    # Generate PSWriteHTML report with chart and table
    New-HTML -Title 'Sign-in report' -FilePath $OutputFile -Show:$false -Content {
        New-HTMLSection -Title 'Policy impact (daily)' -Content {
            New-HTMLChart -ChartType 'line' -Labels ($labels) -Series @(
                @{ Name = 'Success'; Data = $seriesData['Success']; Color = '#20B2AA' },
                @{ Name = 'ReportOnly'; Data = $seriesData['ReportOnly']; Color = '#9B59B6' },
                @{ Name = 'NotApplied'; Data = $seriesData['NotApplied']; Color = '#D1D5DB' },
                @{ Name = 'Failure'; Data = $seriesData['Failure']; Color = '#E74C3C' }
            ) -Options @{ stacked = $true }
        }

        New-HTMLSection -Title 'Details' -Content {
            New-HTMLTable -DataTable $tableData -EnableFiltering
        }
    }

    Write-Host "HTML report written to: $OutputFile"
}