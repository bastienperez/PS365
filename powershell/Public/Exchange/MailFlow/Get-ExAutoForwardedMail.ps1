<#
    .SYNOPSIS
    Get messages that were automatically forwarded to external recipients

    .DESCRIPTION
    Checks the Remote Domain and outbound spam policy auto-forwarding settings, then searches the message trace
    (last 10 days max) for messages handled by an auto-forward or a recipient forward rule.

    Uses Get-MessageTraceV2 and Get-MessageTraceDetailV2 (ExchangeOnlineManagement 3.7.0 or later). Get-MessageTraceV2 returns 5000 messages at most,
    so the period is split into smaller time windows until each one fits. On a busy tenant this can be long: narrow the search with -SenderAddress or -Days.

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
        # Get-MessageTraceV2 has no paging: ResultSize accepts 1 to 5000 (default 1000)
        ResultSize = 5000
    }

    if ($SenderAddress) {
        $params.SenderAddress = $SenderAddress
    }

    if ($FailedOnly) {
        $params.Status = 'Failed'
    }

    # A window returning ResultSize items may be truncated: split it in two and query each half.
    # This needs no knowledge of senders or recipients and avoids -StartingRecipientAddress paging.
    function Get-TraceWindow {
        param([datetime]$Start, [datetime]$End)

        $result = @(Get-MessageTraceV2 @params -StartDate $Start -EndDate $End -WarningAction SilentlyContinue)

        if ($result.Count -lt $params.ResultSize) {
            return $result
        }

        if (($End - $Start).TotalSeconds -le 60) {
            Write-Warning "More than $($params.ResultSize) messages between $Start and $End (UTC): the results of this window are truncated."
            return $result
        }

        $middle = $Start.AddTicks([long](($End - $Start).Ticks / 2))
        Get-TraceWindow -Start $Start -End $middle
        Get-TraceWindow -Start $middle -End $End
    }

    $endDate = (Get-Date).ToUniversalTime()
    $startDate = $endDate.AddDays(-$Days)

    # Window boundaries can return the same message twice: deduplicate
    $messages = Get-TraceWindow -Start $startDate -End $endDate | Group-Object MessageTraceId, RecipientAddress | ForEach-Object { $_.Group[0] }

    Write-Host "Search in the $($messages.Count) messages to find auto-forward" -ForegroundColor Green

    # Get-MessageTraceDetailV2 is throttled: retry with an increasing delay when the limit is reached
    function Get-TraceDetailWithRetry {
        param([string]$RecipientAddress, [string]$MessageTraceId)

        $maxAttempts = 10
        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            try {
                return Get-MessageTraceDetailV2 -RecipientAddress $RecipientAddress -MessageTraceId $MessageTraceId -ErrorAction Stop
            }
            catch {
                $errorText = "$($_.Exception.Message) $($_.ErrorDetails.Message) $_"
                if ($errorText -notmatch 'surpassed the permitted limit|throttl|try again later') {
                    throw
                }
                if ($attempt -eq $maxAttempts) {
                    # Give up on this message only, the search goes on
                    $script:throttledMessages.Add($MessageTraceId)
                    Write-Warning "Still throttled after $maxAttempts attempts, message $MessageTraceId skipped."
                    return
                }
                $delay = [math]::Min(300, 15 * [math]::Pow(2, $attempt - 1))
                Write-Host "Throttled by Exchange Online, waiting $delay seconds (attempt $attempt/$maxAttempts)" -ForegroundColor Yellow
                Start-Sleep -Seconds $delay
            }
        }
    }

    $script:throttledMessages = [System.Collections.Generic.List[string]]::new()
    $messagesAutoForwarded = [System.Collections.Generic.List[PSObject]]::new()
    $index = 0

    foreach ($message in $messages) {
        $index++
        Write-Progress -Activity 'Searching auto-forwarded messages' -Status "$index / $($messages.Count)" -PercentComplete ($index / $messages.Count * 100)

        $details = Get-TraceDetailWithRetry -RecipientAddress $message.RecipientAddress -MessageTraceId $message.MessageTraceId | Where-Object {
            $_.Detail -like '*LED=250 2.1.5 RESOLVER.MSGTYPE.AF; handled AutoForward addressed to external recipient*' -or
            $_.Detail -like '*LED=250 2.1.5 RESOLVER.FWD.Forwarded; recipient forward*'
        }

        foreach ($detail in $details) {
            $messagesAutoForwarded.Add($detail)
        }
    }

    Write-Progress -Activity 'Searching auto-forwarded messages' -Completed

    if ($script:throttledMessages.Count -gt 0) {
        Write-Warning "$($script:throttledMessages.Count) message(s) could not be checked because of throttling: the result is incomplete. Retry later or reduce -Days."
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

