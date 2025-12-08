<#
    .SYNOPSIS
    Retrieves mailbox protocol settings for Exchange Online mailboxes.

    .DESCRIPTION
    This function retrieves the protocol settings (MAPI, OWA, IMAP, POP, EWS, ActiveSync)
    for Exchange Online mailboxes. It can filter mailboxes by identity or by domain.
    Additionally, it provides information about SMTP Client Authentication settings
    at both the mailbox and tenant levels.

    .PARAMETER Identity
    The identity of the mailbox to retrieve protocol settings for.
    If not specified, the function retrieves settings for all mailboxes.

    .PARAMETER ByDomain
    The domain to filter mailboxes by. Only mailboxes with a primary SMTP address in this
    domain will be processed.

    .PARAMETER ExportToExcel
    If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-ExMailboxProtocol

    Retrieves mailbox protocol settings for all Exchange Online mailboxes.

    .EXAMPLE
    Get-ExMailboxProtocol -Identity "user@example.com"

    Retrieves mailbox protocol settings for the specified mailbox.

    .EXAMPLE
    Get-ExMailboxProtocol -ByDomain "example.com"

    Retrieves mailbox protocol settings for all mailboxes in the specified domain.

    .EXAMPLE
    Get-ExMailboxProtocol -ExportToExcel

    Exports mailbox protocol settings for all Exchange Online mailboxes to an Excel file in the user's profile directory.
#>

function Get-ExMailboxProtocol {
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [string]$ByDomain,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    [System.Collections.Generic.List[PSCustomObject]]$exoCasMailboxesArray = @()

    $tenantSmtpClientAuthenticationDisabled = (Get-TransportConfig).SmtpClientAuthenticationDisabled

    if ($tenantSmtpClientAuthenticationDisabled) {
        Write-Host 'SMTP Client Authentication is disabled' -ForegroundColor Green
        $tenantSmtpClientAuthenticationEnabled = $false
    }
    else {
        Write-Host 'SMTP Client Authentication is enabled' -ForegroundColor Yellow
        $tenantSmtpClientAuthenticationEnabled = $true
    }

    # PropertySets All because by default SMTPClientAuthenticationDisabled is not returned
    if ($ByDomain) {
        $casMailboxes = Get-EXOCasMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'" -PropertySets All | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
    }
    elseif ($Identity) {
        [System.Collections.Generic.List[PSCustomObject]]$casMailboxes = @()
        try {
            $mbx = Get-EXOCasMailbox -Identity $Identity -PropertySets All
            $casMailboxes.Add($mbx)
        }
        catch {
            Write-Warning "Mailbox not found: $Identity"
        }
    }
    else {
        $casMailboxes = Get-EXOCasMailbox -ResultSize Unlimited -PropertySets All
    }

    <#
    ECPEnabled        : True
    OWAEnabled        : True
    ImapEnabled       : True
    PopEnabled        : True
    MAPIEnabled       : True
    EwsEnabled        : True
    ActiveSyncEnabled : True
    #>

    foreach ($casMailbox in $casMailboxes) {
    
        $object = [PSCustomObject][ordered]@{ 
            PrimarySmtpAddress                    = $casMailbox.PrimarySmtpAddress
            DisplayName                           = $casMailbox.DisplayName
            ExchangeObjectId                      = $casMailbox.ExchangeObjectId
            MAPIEnabled                           = $casMailbox.MAPIEnabled
            OWAEnabled                            = $casMailbox.OWAEnabled
            UniversalOutlookEnabled               = $casMailbox.UniversalOutlookEnabled
            OutlookMobileEnabled                  = $casMailbox.OutlookMobileEnabled
            IMAPEnabled                           = $casMailbox.ImapEnabled
            POPEnabled                            = $casMailbox.PopEnabled
            EwsEnabled                            = $casMailbox.EwsEnabled
            ActiveSyncEnabled                     = $casMailbox.ActiveSyncEnabled
            # CMDlet returns SMTPClientAuthenticationDisabled but we want SMTPClientAuthenticationEnabled
            ECPEnabled                            = $casMailbox.ECPEnabled
            # we invert the value to provide SMTPClientAuthenticationEnabled because by default the cmdlet returns SMTPClientAuthentication*Disabled*
            SMTPClientAuthenticationEnabled       = if ($null -ne $casMailbox.SMTPClientAuthenticationDisabled) { -not $casMailbox.SMTPClientAuthenticationDisabled }else { '-' }
            TenantSmtpClientAuthenticationEnabled = $tenantSmtpClientAuthenticationEnabled
            MailboxWhenCreated                    = $casMailbox.WhenCreated
            MailboxWhenModified                   = $casMailbox.WhenChanged
        }

        $exoCasMailboxesArray.Add($object)
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-ExMailboxProtocol.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting mailbox protocols to Excel file: $excelFilePath"
        $exoCasMailboxesArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExchangeMailboxProtocols'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $exoCasMailboxesArray
    }
}