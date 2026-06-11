<#
    .SYNOPSIS
    Switches Microsoft Graph PowerShell authentication mode between Browser and WAM.

    .DESCRIPTION
    Toggles the Windows broker (WAM) setting when using `Connect-MgGraph`.
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
    Switch-MgGraphPowerShellMode

    Toggles between browser and WAM authentication modes.

    .EXAMPLE
    Switch-MgGraphPowerShellMode -Mode Browser

    Switches to browser-based authentication.

    .EXAMPLE
    Switch-MgGraphPowerShellMode -Mode WAM

    Switches to Web Account Manager authentication.

    .EXAMPLE
    Switch-MgGraphPowerShellMode -GetCurrent
    
    Displays the current authentication mode.

    .LINK
    https://ps365.clidsys.com/docs/commands/Switch-MgGraphPowerShellMode
#>

function Switch-MgGraphPowerShellMode {
    [CmdletBinding(DefaultParameterSetName = 'Toggle')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'Specify')]
        [ValidateSet('Browser', 'WAM')]
        [string]$Mode,

        [Parameter(Mandatory = $false, ParameterSetName = 'GetCurrent')]
        [switch]$GetCurrent
    )

    # EnableWAMForMSGraph = $true means WAM is enabled, $false means Browser.
    Switch-PS365WamAuthMode -Label 'Microsoft Graph PowerShell' `
        -GetState { [bool]((Get-MgGraphOption).EnableWAMForMSGraph) } `
        -SetState { param($enabled) Set-MgGraphOption -EnableLoginByWAM $enabled } `
        -Mode $Mode -GetCurrent:$GetCurrent
}