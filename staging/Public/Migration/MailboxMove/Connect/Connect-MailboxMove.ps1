function Connect-MailboxMove {
    <#
    .SYNOPSIS
    Connect to Exchange Online and Microsoft Entra ID

    .DESCRIPTION
    Connect to Exchange Online and Microsoft Entra ID

    .PARAMETER Tenant
    if contoso.onmicrosoft.com use "Contoso"

    .EXAMPLE
    Connect-CloudMFA -Tenant

    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Position = 0, Mandatory)]
        [string] $Tenant
    )
    Connect-CloudMFA -Tenant $Tenant -ExchangeOnline -AzureAD
}
