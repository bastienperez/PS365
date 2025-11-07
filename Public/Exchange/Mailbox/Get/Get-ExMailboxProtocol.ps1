function Get-ExMailboxProtocol {
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [string]$ByDomain
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
        }

        $exoCasMailboxesArray.Add($object)
    }

    return $exoCasMailboxesArray
}