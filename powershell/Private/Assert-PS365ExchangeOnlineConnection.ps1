function Assert-PS365ExchangeOnlineConnection {
    <#
    .SYNOPSIS
    Throws a clear error when there is no active Exchange Online session.

    .DESCRIPTION
    Exchange Online cmdlets (Get-AcceptedDomain, Get-Mailbox, ...) emit a cryptic
    "term not recognized" error when called without a session. This helper checks
    for an active connection up front and tells the caller to run
    Connect-ExchangeOnline first, mirroring the Graph connection check used
    elsewhere in the module.
    #>
    [CmdletBinding()]
    param()

    $connected = $false
    if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        $connected = [bool](Get-ConnectionInformation -ErrorAction SilentlyContinue |
                Where-Object { $_.State -eq 'Connected' -and $_.TokenStatus -eq 'Active' })
    }

    if (-not $connected) {
        throw 'Not connected to Exchange Online. Run Connect-ExchangeOnline first, then retry.'
    }
}
