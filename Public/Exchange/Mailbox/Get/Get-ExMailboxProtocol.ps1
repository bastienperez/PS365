function Get-ExMailboxProtocol {
    param (
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
            OWAEnabled                            = $casMailbox.OWAEnabled
            ImapEnabled                           = $casMailbox.ImapEnabled
            PopEnabled                            = $casMailbox.PopEnabled
            MAPIEnabled                           = $casMailbox.MAPIEnabled
            EwsEnabled                            = $casMailbox.EwsEnabled
            ActiveSyncEnabled                     = $casMailbox.ActiveSyncEnabled
            # CMDlet returns SMTPClientAuthenticationDisabled but we want SMTPClientAuthenticationEnabled
            UniversalOutlookEnabled               = $casMailbox.UniversalOutlookEnabled
            OutlookMobileEnabled                  = $casMailbox.OutlookMobileEnabled
            ECPEnabled                            = $casMailbox.ECPEnabled
            SMTPClientAuthenticationEnabled       = if ($null -ne $casMailbox.SMTPClientAuthenticationDisabled) { -not $casMailbox.SMTPClientAuthenticationDisabled }else { '-' }
            TenantSmtpClientAuthenticationEnabled = $tenantSmtpClientAuthenticationEnabled
        }

        $exoCasMailboxesArray.Add($object)
    }

    return $exoCasMailboxesArray
}