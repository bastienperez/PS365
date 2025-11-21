function Get-ExMailboxRegionalConfiguration {
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [string]$ByDomain
    )

    [System.Collections.Generic.List[PSCustomObject]]$exoMbxRegionalConfigArray = @()

    # PropertySets All because by default SMTPClientAuthenticationDisabled is not returned
    if ($ByDomain) {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'" | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
    }
    elseif ($Identity) {
        [System.Collections.Generic.List[PSCustomObject]]$mailboxes = @()
        try {
            $mbx = Get-EXOMailbox -Identity $Identity
            $mailboxes.Add($mbx)
        }
        catch {
            Write-Warning "Mailbox not found: $Identity"
        }
    }
    else {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited
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

    foreach ($mbx in $mailboxes) {
        $regionalConfig = Get-MailboxRegionalConfiguration -Identity $mbx.PrimarySmtpAddress

        $object = [PSCustomObject][ordered]@{ 
            DisplayName        = $mbx.DisplayName
            PrimarySmtpAddress = $mbx.PrimarySmtpAddress
            ExchangeObjectId   = $regionalConfig.Identity
            Language           = $regionalConfig.Language
            TimeZone           = $regionalConfig.TimeZone
            DateFormat         = $regionalConfig.DateFormat
            TimeFormat         = $regionalConfig.TimeFormat
        }

        $exoMbxRegionalConfigArray.Add($object)
    }

    return $exoMbxRegionalConfigArray
}