<#
    .SYNOPSIS
    Switches Azure CLI authentication between browser and WAM modes.

    .DESCRIPTION
    Toggles the Windows broker (WAM) setting when using `az login`.
    If no mode is specified, it toggles between modes.
    If a mode is specified, it switches to that mode.

    Beginning with Azure CLI version 2.61.0, Web Account Manager (WAM) is the default authentication method on Windows.
    WAM is a Windows 10+ component that acts as an authentication broker. An authentication broker is an application that runs on a user's machine. It manages the authentication handshakes and token maintenance for connected accounts.

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
    Switch-AzureCliAuthMode

    Toggles between browser and WAM authentication modes.

    .EXAMPLE
    Switch-AzureCliAuthMode -Mode Browser

    Switches to browser-based authentication.

    .EXAMPLE
    Switch-AzureCliAuthMode -Mode WAM

    Switches to Web Account Manager authentication.

    .EXAMPLE
    Switch-AzureCliAuthMode -GetCurrent
    
    Displays the current authentication mode.

    .LINK
    https://ps365.clidsys.com/docs/commands/Switch-AzureCliAuthMode
#>

function Switch-AzureCliAuthMode {
    [CmdletBinding(DefaultParameterSetName = 'Toggle')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'Specify')]
        [ValidateSet('Browser', 'WAM')]
        [string]$Mode,

        [Parameter(Mandatory = $false, ParameterSetName = 'GetCurrent')]
        [switch]$GetCurrent
    )

    Switch-PS365WamAuthMode -Label 'Azure CLI' `
        -GetState { (az config get core.enable_broker_on_windows --query value -o tsv 2>$null) -eq 'true' } `
        -SetState { param($enabled) az config set "core.enable_broker_on_windows=$($enabled.ToString().ToLower())" } `
        -Mode $Mode -GetCurrent:$GetCurrent
}