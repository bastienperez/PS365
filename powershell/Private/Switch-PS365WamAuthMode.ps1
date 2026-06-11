function Switch-PS365WamAuthMode {
    <#
    .SYNOPSIS
    Shared implementation for the Switch-*Mode auth-mode togglers.

    .DESCRIPTION
    The Switch-AzureCliAuthMode / Switch-AzurePowerShellMode /
    Switch-MgGraphPowerShellMode functions were ~identical (only the get/set
    cmdlets and the product label differed). This helper centralizes the toggle/
    set/get-current logic; each public function passes its own GetState/SetState
    script blocks and label.

    .PARAMETER Label
    Product label used in the confirmation message (e.g. 'Azure CLI').

    .PARAMETER GetState
    Script block returning the current WAM-enabled state as a boolean.

    .PARAMETER SetState
    Script block that takes a single boolean (the desired WAM-enabled state).

    .PARAMETER Mode
    'Browser' or 'WAM'. When omitted, the current state is toggled.

    .PARAMETER GetCurrent
    Display the current mode without changing it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Label,

        [Parameter(Mandatory)]
        [scriptblock]$GetState,

        [Parameter(Mandatory)]
        [scriptblock]$SetState,

        [string]$Mode,

        [switch]$GetCurrent
    )

    $wamEnabled = [bool](& $GetState)

    if ($GetCurrent) {
        $currentMode = if ($wamEnabled) { 'WAM (Web Account Manager)' } else { 'Browser' }
        Write-Host "Current authentication mode: $currentMode" -ForegroundColor Cyan
        return
    }

    # Determine the target WAM state: explicit -Mode wins, otherwise toggle.
    $targetWam = if ($Mode) { $Mode -eq 'WAM' } else { -not $wamEnabled }

    if ($targetWam) {
        Write-Verbose 'Enabling Web Account Manager (WAM)...'
    }
    else {
        Write-Verbose 'Disabling Web Account Manager (WAM)...'
    }

    & $SetState $targetWam

    $targetName = if ($targetWam) { 'WAM' } else { 'browser' }
    Write-Host "$Label authentication switched to $targetName mode." -ForegroundColor Green
}
