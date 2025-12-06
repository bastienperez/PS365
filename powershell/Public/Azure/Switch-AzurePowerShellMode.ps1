<#
    .SYNOPSIS
    Switches Azure PowerShell authentication between browser and WAM modes.

    .DESCRIPTION
    Toggles the Windows broker (WAM) setting when using `Connect-AzAccount`.
    If no mode is specified, it toggles between modes.
    If a mode is specified, it switches to that mode.

    Using WAM has several benefits:
    - Enhanced security. See Conditional Access: Token protection (preview).
    - Support for Windows Hello, conditional access policies, and FIDO keys.
    - Streamlined single sign-on.
    - Bug fixes and enhancements shipped with Windows.

    Technically this opens a window for login.
    But, in some cases, especially when connecting to client environments, it might be preferable to use the browser-based authentication method.

    .PARAMETER Mode
    The authentication mode to switch to. Valid values are 'Browser' or 'WAM'.
    If not specified, the function will toggle between the current modes.

    .PARAMETER GetCurrent
    Displays the current authentication mode without making any changes.

    .EXAMPLE
    Switch-AzurePowerShellMode
    Toggles between browser and WAM authentication modes.

    .EXAMPLE
    Switch-AzurePowerShellMode -Mode Browser
    Switches to browser-based authentication.

    .EXAMPLE
    Switch-AzurePowerShellMode -Mode WAM
    Switches to Web Account Manager authentication.

    .EXAMPLE
    Switch-AzurePowerShellMode -GetCurrent
    Displays the current authentication mode.
#>

function Switch-AzurePowerShellMode {
    [CmdletBinding(DefaultParameterSetName = 'Toggle')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'Specify')]
        [ValidateSet('Browser', 'WAM')]
        [string]$Mode,

        [Parameter(Mandatory = $false, ParameterSetName = 'GetCurrent')]
        [switch]$GetCurrent
    )

    # Get current WAM setting
    #
    $currentSetting = (Get-AzConfig -EnableLoginByWam).Value

    if ($GetCurrent) {
        # Display current mode
        if ($currentSetting) {
            Write-Host 'Current authentication mode: WAM (Web Account Manager)' -ForegroundColor Cyan
        }
        else {
            Write-Host 'Current authentication mode: Browser' -ForegroundColor Cyan
        }
        return
    }
    
    if ($Mode) {
        # Switch to specified mode
        if ($Mode -eq 'Browser') {
            Write-Verbose 'Disabling Web Account Manager (WAM)...'
            $null = Update-AzConfig -EnableLoginByWam $false
            Write-Host 'Azure CLI authentication switched to browser mode.' -ForegroundColor Green
        }
        else {
            Write-Verbose 'Enabling Web Account Manager (WAM)...'
            $null = Update-AzConfig -EnableLoginByWam $true
            Write-Host 'Azure CLI authentication switched to WAM mode.' -ForegroundColor Green
        }
    }
    else {
        # Toggle mode
        if ($currentSetting) {
            Write-Verbose 'Disabling Web Account Manager (WAM)...'
            $null = Update-AzConfig -EnableLoginByWam $false
            Write-Host 'Azure CLI authentication switched to browser mode.' -ForegroundColor Green
        }
        else {
            Write-Verbose 'Enabling Web Account Manager (WAM)...'
            $null = Update-AzConfig -EnableLoginByWam $true
            Write-Host 'Azure CLI authentication switched to WAM mode.' -ForegroundColor Green
        }
    }
}