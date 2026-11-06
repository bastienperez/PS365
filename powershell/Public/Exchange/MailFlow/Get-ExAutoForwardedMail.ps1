<#
    .SYNOPSIS
    Get messages that were automatically forwarded to external recipients

    .DESCRIPTION
    Checks the Remote Domain and outbound spam policy auto-forwarding settings, then searches the message trace
    (last 10 days max) for messages handled by an auto-forward or a recipient forward rule.

    Uses Get-MessageTraceV2 and Get-MessageTraceDetailV2 (ExchangeOnlineManagement 3.7.0 or later). With ResultSize capped at 5000 and no paging,
    a busy tenant can have more messages than returned: narrow the search with -SenderAddress or -Days.

    Message trace data is only available for the last 10 days. For older data, use Start-HistoricalSearch and Get-HistoricalSearch.

    .PARAMETER SenderAddress
    Only return messages sent by this address.

    .PARAMETER FailedOnly
    Only return messages with a Failed status.

    .PARAMETER Days
    Number of days to search back (1 to 10). Default is 10.

    .EXAMPLE
    Get-ExAutoForwardedMail

    Returns auto-forwarded messages from the last 10 days.

    .EXAMPLE
    Get-ExAutoForwardedMail -SenderAddress user@example.com -Days 3 -FailedOnly

    Returns failed auto-forwarded messages sent by user@example.com in the last 3 days.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExAutoForwardedMail

    .NOTES
    Sources:
    https://office365itpros.com/2020/11/12/microsoft-clamps-down-mail-forwarding-exchange-online/
    http://blog.icewolf.ch/archive/2020/10/06/how-to-control-the-many-ways-of-email-forwarding-in.aspx
    https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/external-email-forwarding
    https://techcommunity.microsoft.com/t5/exchange-team-blog/all-you-need-to-know-about-automatic-email-forwarding-in/ba-p/2074888
    https://misstech.co.uk/2020/07/27/new-controls-available-to-block-automatic-email-forwarding/
    Report: https://admin.cloud.microsoft/exchange?ref=/homepage#/reports/autoforwardedmessages
#>

function Get-ExAutoForwardedMail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$SenderAddress,

        [Parameter(Mandatory = $false)]
        [switch]$FailedOnly,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 10)]
        [int]$Days = 10
    )

    Assert-PS365ExchangeOnlineConnection

    $remoteDomains = Get-RemoteDomain

    foreach ($remoteDomain in $remoteDomains) {
        Write-Host "Remote Domain '$($remoteDomain.Name)' AutoForwardEnabled: $($remoteDomain.AutoForwardEnabled)" -ForegroundColor Cyan
    }

    $outboundSpamPolicies = Get-HostedOutboundSpamFilterPolicy

    foreach ($outboundSpamPolicy in $outboundSpamPolicies) {
        Write-Host "OutboundSpamPolicy '$($outboundSpamPolicy.Name)' AutoForwardingMode: $($outboundSpamPolicy.AutoForwardingMode)" -ForegroundColor Cyan

        if ($outboundSpamPolicy.AutoForwardingMode -eq 'Automatic') {
            Write-Host "Careful, the value 'Automatic' is now the same as AutoForwardEnable=Off, meaning auto-forward is blocked even if the Remote Domain(s) are configured with AutoForwardEnable = `$true
        Sources:
        https://office365itpros.com/2020/11/12/microsoft-clamps-down-mail-forwarding-exchange-online/
        http://blog.icewolf.ch/archive/2020/10/06/how-to-control-the-many-ways-of-email-forwarding-in.aspx
        https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/external-email-forwarding
        https://techcommunity.microsoft.com/t5/exchange-team-blog/all-you-need-to-know-about-automatic-email-forwarding-in/ba-p/2074888

        RoadMap ID: MC221113" -ForegroundColor Yellow
        }
    }

    Write-Host "Get messages from the last $Days days" -ForegroundColor Cyan
    Write-Host "`nIf you prefer, you can use the report in https://admin.cloud.microsoft/exchange?ref=/homepage#/reports/autoforwardedmessages`n" -ForegroundColor Green

    $params = @{
        StartDate  = (Get-Date).ToUniversalTime().AddDays(-$Days)
        EndDate    = (Get-Date).ToUniversalTime()
        # Get-MessageTraceV2 has no paging: ResultSize accepts 1 to 5000 (default 1000)
        ResultSize = 5000
    }

    if ($SenderAddress) {
        $params.SenderAddress = $SenderAddress
    }

    if ($FailedOnly) {
        $params.Status = 'Failed'
    }

    $messages = Get-MessageTraceV2 @params

    Write-Host "Search in the $($messages.Count) messages to find auto-forward" -ForegroundColor Green

    $messagesAutoForwarded = $messages | ForEach-Object {
        Get-MessageTraceDetailV2 -RecipientAddress $_.RecipientAddress -MessageTraceId $_.MessageTraceId | Where-Object {
            $_.Detail -like '*LED=250 2.1.5 RESOLVER.MSGTYPE.AF; handled AutoForward addressed to external recipient*' -or
            $_.Detail -like '*LED=250 2.1.5 RESOLVER.FWD.Forwarded; recipient forward*'
        }
    }

    [System.Collections.Generic.List[PSCustomObject]]$emailsWithAutoForward = @()

    foreach ($messageAF in $messagesAutoForwarded) {
        # Get-MessageTraceDetailV2 does not return sender info, so we look the message up in $messages
        $matchingMessages = $messages | Where-Object { $_.MessageId -eq $messageAF.MessageId }

        foreach ($message in $matchingMessages) {
            $object = [PSCustomObject][ordered]@{
                SenderAddress    = $message.SenderAddress
                RecipientAddress = $message.RecipientAddress
                Subject          = $message.Subject
                Detail           = $messageAF.Detail
                Status           = $message.Status
                Received         = $message.Received
                FromIP           = $message.FromIP
                ToIP             = $message.ToIP
                MessageId        = $message.MessageId
            }

            $emailsWithAutoForward.Add($object)
        }
    }

    return $emailsWithAutoForward
}

