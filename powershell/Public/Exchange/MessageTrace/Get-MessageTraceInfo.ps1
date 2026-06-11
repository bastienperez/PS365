<#
    .SYNOPSIS
    Retrieves message trace information based on specified criteria.

    .DESCRIPTION
    This function retrieves message trace information from Exchange Online based on the specified criteria. It supports filtering by sender address, recipient address, start date, end date, connector name, transport rule name, and using connector flag.

    .PARAMETER SenderAddress
    Specifies the sender address to filter the message trace. This parameter accepts a string or an array of strings. Wildcards (*) can be used to match multiple sender addresses.

    .PARAMETER RecipientAddress
    Specifies the recipient address to filter the message trace. This parameter accepts a string or an array of strings. Wildcards (*) can be used to match multiple recipient addresses.

    .PARAMETER StartDate
    Specifies the start date to filter the message trace. This parameter accepts a DateTime object.

    .PARAMETER EndDate
    Specifies the end date to filter the message trace. This parameter accepts a DateTime object.

    .PARAMETER ByConnectorName
    Specifies the connector name to filter the message trace. This parameter accepts a string.

    .PARAMETER ByTransportRuleName
    Specifies the transport rule name to filter the message trace. This parameter accepts a string.

    .PARAMETER UsingConnector
    Specifies whether to filter the message trace using a connector. This parameter is a switch parameter.

    .EXAMPLE
    Get-MessageTraceInfo -SenderAddress "john.doe@example.com" -RecipientAddress "jane.doe@example.com" -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) -ByConnectorName "OutboundConnector" -UsingConnector
    
    Retrieves message trace information for messages sent by "john.doe@example.com" to "jane.doe@example.com" within the last 7 days, using the "OutboundConnector" connector.

    .EXAMPLE
    Get-MessageTraceInfo -RecipientAddress "jane.doe@example.com" -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) -ByTransportRuleName "Confidential" -UsingConnector
    
    Retrieves message trace information for messages received by "jane.doe@example.com" within the last 30 days, filtered by the "Confidential" transport rule, using a connector.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MessageTraceInfo

    .NOTES
    This function requires the Exchange Online PowerShell module to be installed and connected to an Exchange Online organization.

    For some event, it will fail to get messagetracedetail
#>

function Get-MessageTraceInfo {
    param(
        [Parameter(Mandatory = $false)]
        [String[]] $SenderAddress,
        [Parameter(Mandatory = $false)]
        [String[]] $RecipientAddress,
        [Parameter(Mandatory = $false)]
        [datetime] $StartDate,
        [Parameter(Mandatory = $false)]
        [datetime] $EndDate,
        [Parameter(Mandatory = $false)]
        [String] $ByConnectorName,
        [Parameter(Mandatory = $false)]
        [String]$ByTransportRuleName,
        [Parameter(Mandatory = $false)]
        [Switch] $UsingConnector
    )

    [System.Collections.Generic.List[PSObject]]$messagesInfo = @()
    [System.Collections.Generic.List[PSObject]]$messagesList = @()

    # PageSize maximum default = 1000; Max PageSize = 5000
    # There isn't really a reason to decrease this in this instance.
    <#
    not present with Get-MessageTraceV2
    $messageTraceParams = @{
        PageSize = 5000
    }
    #>

    if ($RecipientAddress -and $SenderAddress) {
        $messageTraceParams = @{
            RecipientAddress = $RecipientAddress
            SenderAddress    = $SenderAddress
            #Status           = 'Delivered'
        }
    }
    elseif ($RecipientAddress -and -not $SenderAddress) {
        $messageTraceParams = @{
            RecipientAddress = $RecipientAddress
        }
    }
    elseif ($SenderAddress -and -not $RecipientAddress) {
        $messageTraceParams = @{
            SenderAddress = $SenderAddress
        }
    }

    # convert to UTC because Exchange Online used UTC time
    if ($StartDate) {
        $messageTraceParams.Add('StartDate', $StartDate.ToUniversalTime())
        
        if ($EndDate) {
            $messageTraceParams.Add('EndDate', $EndDate.ToUniversalTime())
        }
        else {
            $messageTraceParams.Add('EndDate', $((Get-Date).ToUniversalTime()))
        }
    }

    <#
    Not present with Get-MessageTraceV2
    $messageTraceParams.Add('Page', 1)
    $maxPage = 1000
    
    
    $page = 1
    $maxPage = 1000
    $pageSize = 5000 # Max pagesize is 5000. There isn't really a reason to decrease this in this instance.
    #>

    $messages = Get-MessageTraceV2 @messageTraceParams | Select-Object MessageId, Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageTraceId, StartDate, EndDate
    
    foreach ($message in $messages) {
        $messagesList.Add($message)
    }

    Write-Host -ForegroundColor Cyan "Found $($messagesList.Count) messages (in total, without the filter)"
    
    $oldProgressPreference = $ProgressPreference

    $ProgressPreference = 'SilentlyContinue'

    foreach ($message in $messagesList) {
        # Ignore Journaling events if exist
        
        if ($UsingConnector.IsPresent) {
            $traceDetail = $message | Get-MessageTraceDetailV2 | Where-Object { $_.Action -eq 'RouteMessageUsingConnector' -and $_.Event -ne 'Journal' }
        }
        elseif ($ByTransportRuleName) {
            $traceDetail = $message | Get-MessageTraceDetailV2 | Where-Object { $_.Detail -like "*$ByTransportRuleName*" }
        }
        elseif ($ByConnectorName) {
            #$traceDetail = $message | Get-MessageTraceDetailV2 -Event RECEIVE | Where-Object { $_.Data -like "*S:InboundConnectorData=Name=$ByConnectorName*" }
            $traceDetail = $message | Get-MessageTraceDetailV2 | Where-Object { $_.Data -like "*S:Microsoft.Exchange.Hygiene.TenantOutboundConnectorCustomData=Name=$ByConnectorName*" }
        }
        else {
            #$traceDetail = $message | Get-MessageTraceDetailV2 -Event RECEIVE | Where-Object { $_.Event -ne 'Journal' }
            # Event: -Event RECEIVE, SEND, FAIL, DELIVER, EXPAND, TRANSFER, DEFER, DROP
            # event list: https://learn.microsoft.com/en-us/exchange/mail-flow/transport-logs/message-tracking?view=exchserver-2019#event-types-in-the-message-tracking-log
            
            if ($message.Count -eq 1) {
                $traceDetail = Get-MessageTraceDetailV2 -MessageId $message.MessageId -MessageTraceId $message.MessageTraceId -RecipientAddress $message.RecipientAddress
            }
            else {
                # we found the last event only, don't know if it's the best way to do it
                $traceDetail = ($message | Get-MessageTraceDetailV2 | Where-Object { $_.Event -ne 'Journal' })[-1]
            }
        }

        # redirect has no messagetracedetails in event receive
        if ($null -ne $traceDetail) {

            # Get-MessageTraceDetailV2 + Where-Object can return multiple rows. Casting an array to
            # [XML] fails with "This document already has a 'DocumentElement' node", so we iterate
            # one row at a time and emit one object per detail record.
            foreach ($detail in @($traceDetail)) {
                $hashTable = @{}

                $XMLDoc = $null
                try {
                    $XMLDoc = [XML]$detail.Data
                }
                catch {
                    Write-Verbose "Could not parse trace detail Data as XML: $($_.Exception.Message)"
                }

                $MEPNodes = $null
                if ($XMLDoc) {
                    try {
                        $MEPNodes = $XMLDoc.GetElementsByTagName('MEP')
                    }
                    catch {
                        Write-Verbose "Could not get MEP nodes: $($_.Exception.Message)"
                    }
                }

                if ($MEPNodes) {
                    for ($nodeval = 0; $nodeval -lt $MEPNodes.Count; $nodeval++) {
                        $key = $MEPNodes[$nodeval].Attributes[0].Value.ToString()
                        $value = $MEPNodes[$nodeval].Attributes[1].Value.ToString()
                        # Use [hashtable] indexer assignment instead of .Add() to tolerate duplicate keys.
                        $hashTable[$key] = $value
                    }
                }

                # Helper: extract a CustomData S:<Key>=<Value> fragment regardless of ordering.
                $customData = $hashTable['customdata']
                $extractFromCustomData = {
                    param([string]$Key)
                    if ([string]::IsNullOrWhiteSpace($customData)) { return $null }
                    $match = $customData -split ";'" -split ';' | Where-Object { $_ -match "^S:$Key=" } | Select-Object -First 1
                    if ($match) { return ($match -replace "^S:$Key=", '').Trim() }
                    return $null
                }

                # Mail client / user-agent best-effort extraction. The exact field that carries the
                # client name varies by event (SUBMIT, RECEIVE, DELIVER) and source (SMTP, OWA, mobile),
                # so we surface every known candidate so the user can pick whichever is populated.
                $userAgent         = & $extractFromCustomData 'UserAgent'
                $messageSourceName = & $extractFromCustomData 'MessageSourceName'
                $clientName        = & $extractFromCustomData 'ClientName'

                $object = [PSCustomObject] [ordered]@{
                    SenderAddress        = $message.SenderAddress
                    RecipientAddress     = $message.RecipientAddress
                    ReturnPath           = $hashTable['ReturnPath']
                    Subject              = $message.Subject
                    Detail               = $detail.Detail
                    Status               = $message.Status
                    'Received(UTC)'      = $message.Received
                    FromIP               = $message.FromIP
                    ToIP                 = $message.ToIP
                    ClientIP             = $hashTable['ClientIP']
                    UserAgent            = $userAgent
                    MessageSourceName    = $messageSourceName
                    ClientName           = $clientName
                    DeliveryPriority     = $hashTable['DeliveryPriority']
                    # CustomData is parsed below
                    #CustomData          = $hashTable['CustomData']
                    InboundConnectorData = (& $extractFromCustomData 'InboundConnectorData') -replace '^Name=', ''
                    ConnectorType        = ($customData -split ";'" -split ';' -match 'ConnectorType' -split 'ConnectorType=')[1]
                    TLSVersion           = & $extractFromCustomData 'tlsversion'
                    TLSCipher            = & $extractFromCustomData 'tlscipher'
                    OriginOrg            = & $extractFromCustomData 'Oorg'
                    MessageID            = $detail.MessageId
                    MessageTraceID       = $detail.MessageTraceId
                }

                $messagesInfo.Add($object)
            }
        }
    }

    $ProgressPreference = $oldProgressPreference
    
    return $messagesInfo
}