function Assert-PS365MicrosoftTeamsConnection {
    <#
    .SYNOPSIS
    Throws a clear error when there is no active Microsoft Teams session.

    .DESCRIPTION
    MicrosoftTeams cmdlets (Get-Team, Get-TeamUser, ...) emit a cryptic error when called without a
    session. Unlike Exchange Online, the MicrosoftTeams module does not expose a Get-ConnectionInformation
    cmdlet, so the connection is probed with Get-CsTenant (a lightweight tenant lookup that fails when no
    session is established). This helper tells the caller to run Connect-MicrosoftTeams first, mirroring
    the Exchange Online connection check used elsewhere in the module.
    #>
    [CmdletBinding()]
    param()

    $connected = $false
    if (Get-Command -Name Get-CsTenant -ErrorAction SilentlyContinue) {
        try {
            $null = Get-CsTenant -ErrorAction Stop
            $connected = $true
        }
        catch {
            $connected = $false
        }
    }

    if (-not $connected) {
        throw 'Not connected to Microsoft Teams. Run Connect-MicrosoftTeams first, then retry.'
    }
}
