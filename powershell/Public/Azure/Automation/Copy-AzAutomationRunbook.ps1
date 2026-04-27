<#
    .SYNOPSIS
    Copies all runbooks from a source Azure Automation account to a destination account.

    .DESCRIPTION
    This function exports all runbooks from a source Azure Automation account and imports them
    into a destination Azure Automation account within the same resource group.

    .PARAMETER SourceAutomationAccount
    The name of the source Azure Automation account to export runbooks from.

    .PARAMETER DestinationAutomationAccount
    The name of the destination Azure Automation account to import runbooks into.

    .PARAMETER ResourceGroup
    The name of the resource group containing both Automation accounts.

    .PARAMETER TempFolder
    (Optional) The temporary folder path used to store exported runbooks. Defaults to a timestamped folder in the system temp directory.
    Note: Only auto-generated folders are deleted at the end. If you provide an existing folder, its contents are NOT cleaned up.

    .EXAMPLE
    Copy-AzAutomationRunbook -SourceAutomationAccount 'newauto' -DestinationAutomationAccount 'newauto1' -ResourceGroup 'my-rg'

    Copies all runbooks from 'newauto' to 'newauto1'.

    .EXAMPLE
    Copy-AzAutomationRunbook -SourceAutomationAccount 'newauto' -DestinationAutomationAccount 'newauto1' -ResourceGroup 'my-rg' -TempFolder 'C:\Temp\Runbooks'

    Copies all runbooks using a custom temporary folder.

    .LINK
    https://ps365.clidsys.com/docs/commands/Copy-AzAutomationRunbook
#>

function Copy-AzAutomationRunbook {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SourceAutomationAccount,

        [Parameter(Mandatory = $true)]
        [string]$DestinationAutomationAccount,

        [Parameter(Mandatory = $true)]
        [string]$ResourceGroup,

        [Parameter(Mandatory = $false)]
        [string]$TempFolder
    )

    $tempFolderCreated = $false
    if (-not $TempFolder) {
        $TempFolder = Join-Path $env:TEMP "AzAutomationRunbooks_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        $tempFolderCreated = $true
    }

    $isConnected = $null -ne (Get-AzContext -ErrorAction SilentlyContinue)

    if (-not $isConnected) {
        Write-Verbose 'Connecting to Azure...'
        Connect-AzAccount
    }

    # Create temp folder if it doesn't exist
    if (-not (Test-Path $TempFolder)) {
        $null = New-Item -ItemType Directory -Path $TempFolder -Force
        $tempFolderCreated = $true
        Write-Verbose "Created temp folder: $TempFolder"
    }

    # Get all runbooks from source account
    Write-Verbose "Retrieving runbooks from source account: $SourceAutomationAccount"
    $runbooks = Get-AzAutomationRunbook -AutomationAccountName $SourceAutomationAccount -ResourceGroupName $ResourceGroup

    if (-not $runbooks) {
        Write-Warning "No runbooks found in source account '$SourceAutomationAccount'."
        if ($tempFolderCreated -and (Test-Path $TempFolder)) {
            Remove-Item -Path $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        return
    }

    Write-Host "$($runbooks.Count) runbook(s) found in '$SourceAutomationAccount'" -ForegroundColor Cyan

    # Export runbooks from source
    $exportedRunbooks = [System.Collections.Generic.List[PSCustomObject]]@()

    $i = 0
    foreach ($runbook in $runbooks) {
        $i++
        Write-Host "($i/$($runbooks.Count)) Exporting: $($runbook.Name)" -ForegroundColor Cyan -NoNewline

        try {
            Export-AzAutomationRunbook -ResourceGroupName $ResourceGroup -AutomationAccountName $SourceAutomationAccount -Name $runbook.Name -OutputFolder $TempFolder -Force
            $exportedFile = Get-ChildItem -Path $TempFolder -Filter "$($runbook.Name).*" | Select-Object -First 1
            Write-Host ' OK' -ForegroundColor Green

            $exportedRunbooks.Add([PSCustomObject]@{
                Name     = $runbook.Name
                Type     = $runbook.RunbookType
                Path     = $exportedFile.FullName
                Exported = $true
                Imported = $null
            })
        }
        catch {
            Write-Host ' KO' -ForegroundColor Red
            Write-Warning "Failed to export '$($runbook.Name)': $($_.Exception.Message)"

            $exportedRunbooks.Add([PSCustomObject]@{
                Name     = $runbook.Name
                Type     = $runbook.RunbookType
                Path     = $null
                Exported = $false
                Imported = $null
            })
        }
    }

    # Import runbooks into destination
    Write-Host "`nImporting runbooks into '$DestinationAutomationAccount'..." -ForegroundColor Cyan

    $j = 0
    foreach ($runbook in $exportedRunbooks | Where-Object { $_.Exported }) {
        $j++
        Write-Host "($j/$($exportedRunbooks.Where({$_.Exported}).Count)) Importing: $($runbook.Name)" -ForegroundColor Cyan -NoNewline

        try {
            Import-AzAutomationRunbook -Path $runbook.Path -ResourceGroupName $ResourceGroup -AutomationAccountName $DestinationAutomationAccount -Name $runbook.Name -Type $runbook.Type -Force -ErrorAction Stop
            $runbook.Imported = $true
            Write-Host ' OK' -ForegroundColor Green
        }
        catch {
            $runbook.Imported = $false
            Write-Host ' KO' -ForegroundColor Red
            Write-Warning "Failed to import '$($runbook.Name)': $($_.Exception.Message)"
        }
    }

    # Summary
    $exportSuccess = ($exportedRunbooks | Where-Object { $_.Exported }).Count
    $exportFail    = ($exportedRunbooks | Where-Object { -not $_.Exported }).Count
    $importSuccess = ($exportedRunbooks | Where-Object { $_.Imported -eq $true }).Count
    $importFail    = ($exportedRunbooks | Where-Object { $_.Imported -eq $false }).Count

    Write-Host "`n--- Summary ---" -ForegroundColor Cyan
    Write-Host "Exported:  $exportSuccess/$($runbooks.Count)" -ForegroundColor $(if ($exportFail -eq 0) { 'Green' } else { 'Yellow' })
    Write-Host "Imported:  $importSuccess/$exportSuccess" -ForegroundColor $(if ($importFail -eq 0) { 'Green' } else { 'Yellow' })

    if ($exportFail -gt 0) {
        Write-Host "Export failures ($exportFail):" -ForegroundColor Red
        $exportedRunbooks | Where-Object { -not $_.Exported } | ForEach-Object {
            Write-Host "  - $($_.Name)" -ForegroundColor Red
        }
    }
    if ($importFail -gt 0) {
        Write-Host "Import failures ($importFail):" -ForegroundColor Red
        $exportedRunbooks | Where-Object { $_.Imported -eq $false } | ForEach-Object {
            Write-Host "  - $($_.Name)" -ForegroundColor Red
        }
    }

    # Cleanup temp folder (only if it was auto-generated by this function)
    if ($tempFolderCreated -and (Test-Path $TempFolder)) {
        Write-Verbose "Cleaning up temp folder: $TempFolder"
        Remove-Item -Path $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}
