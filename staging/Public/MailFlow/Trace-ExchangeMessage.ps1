
function Trace-ExchangeMessage {
    <#
    .SYNOPSIS
    On-Premises Exchange Message Tracking Log made easy!
    You may see an error if you have Edge servers in your org. It is fine to ignore those.

    .DESCRIPTION
    Searches all Hub Transport and Mailbox Servers for messages. Once found, you can select one or more messages via Out-GridView to search by those MessageID's.

    .PARAMETER Sender
    Parameter description

    .PARAMETER Recipients
    Parameter description

    .PARAMETER StartSearchHoursAgo
    Parameter description

    .PARAMETER EndSearchHoursAgo
    Parameter description

    .PARAMETER Subject
    Parameter description

    .PARAMETER MessageID
    Parameter description

    .PARAMETER ResultSize
    Parameter description

    .PARAMETER Status
    Parameter description

    .PARAMETER ExportToExcel
    Export results to Excel file in PS365 folder on the desktop
    Can be used with ExportToCSV

    .PARAMETER ExportToCSV
    Export results to CSV file in PS365 folder on the desktop
    Can be used with ExportToExcel

    .EXAMPLE
    Trace-ExchangeMessage -StartSearchHoursAgo 48 -EndSearchHoursAgo 24 -Recipients "joe@contoso.com" -Subject "Forklift incident"

    .EXAMPLE
    Trace-ExchangeMessage -StartSearchHoursAgo .01

    .EXAMPLE
    Trace-ExchangeMessage -StartSearchHoursAgo 1 -ExportToExcel

    .NOTES
    General notes
    #>
    [CmdletBinding()]
    [Alias('New-MessageTrack')]
    param
    (
        [Parameter()]
        [string] $Sender,

        [Parameter()]
        [string] $Recipients,

        [Parameter()]
        [Double] $StartSearchHoursAgo = '.1',

        [Parameter()]
        [Double] $EndSearchHoursAgo = '0',

        [Parameter()]
        [string] $Subject,

        [Parameter()]
        [switch] $SkipHealthMessages,

        [Parameter()]
        [string] $MessageID,

        [Parameter()]
        [string] $ResultSize = 'Unlimited',

        [Parameter()]
        [string] $Status,

        [Parameter()]
        [switch] $ExportToExcel,

        [Parameter()]
        [switch] $ExportToCSV
    )
    $Servers = Get-TransportServer -WarningAction SilentlyContinue
    $currentErrorActionPrefs = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'

    if ($StartSearchHoursAgo) {
        [DateTime]$StartSearchHoursAgo = ((Get-Date).AddHours( - $StartSearchHoursAgo))
        $StartSearchHoursAgo = $StartSearchHoursAgo.ToUniversalTime()
    }

    if ($StartSearchHoursAgo) {
        [DateTime]$EndSearchHoursAgo = ((Get-Date).AddHours( - $EndSearchHoursAgo))
        $EndSearchHoursAgo = $EndSearchHoursAgo.ToUniversalTime()
    }

    $cmdletParams = (Get-Command $PSCmdlet.MyInvocation.InvocationName).Parameters.Keys

    $params = @{ }
    $NotArray = 'SkipHealthMessages', 'ExportToExcel', 'ExportToCSV', 'StartSearchHoursAgo', 'EndSearchHoursAgo', 'Subject', 'ResultSize', 'Debug', 'Verbose', 'ErrorAction', 'WarningAction', 'InformationAction', 'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable'
    foreach ($cmdletParam in $cmdletParams) {
        if ($cmdletParam -notin $NotArray) {
            if ([string]::IsNullOrWhiteSpace((Get-Variable -Name $cmdletParam).Value) -ne $true) {
                $params.Add($cmdletParam, (Get-Variable -Name $cmdletParam).Value)
            }
        }
    }
    $params.Add('Start', $StartSearchHoursAgo)
    $params.Add('End', $EndSearchHoursAgo)
    $params.Add('MessageSubject', $Subject)

    $allMessageTrackResults = New-Object 'System.Collections.Generic.List[PSObject]'

    try {
        $messageTrack = $Servers.name | Get-MessageTrackingLog @params -ResultSize $ResultSize
        if ($messageTrack) {
            $messageTrack | ForEach-Object {
                $messageTrackResults = [PSCustomObject][ordered]@{
                    Time             = $_.TimeStamp
                    Directionality   = $_.Directionality
                    EventID          = $_.EventID
                    Sender           = $_.Sender
                    Recipients       = ($_.Recipients) -ne '' -join '|'
                    Subject          = $_.MessageSubject
                    Connector        = $_.ConnectorID
                    SourceContext    = $_.SourceContext
                    EventData        = ($_.EventData) -ne '' -join '|'
                    ServerHostName   = $_.ServerHostName
                    ServerIP         = $_.ServerIP
                    ClientIP         = $_.ClientIP
                    OriginalClientIP = $_.OriginalClientIP
                    ClientHostName   = $_.ClientHostName
                    TotalBytes       = $_.TotalBytes
                    MessageId        = $_.MessageId
                }
                $allMessageTrackResults.Add($messageTrackResults)
            }
            else {
                Write-Verbose "`tNo results found"
            }
        }
    }
    catch {
        Write-Verbose "`tException gathering message trace data."
    }

    if ($allMessageTrackResults.count -gt 0) {
        if ($ExportToExcel -or $ExportToCSV) {
            $PoshPath = (Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath PS365)
            $null = New-Item $PoshPath -type Directory -Force -ErrorAction SilentlyContinue

            if ($ExportToExcel) {
                $FileStamp = 'OnPremises-MessageTrace_{0}.xlsx' -f [DateTime]::Now.ToString('yyyy-MM-dd-hhmm')
                $ExcelSplat = @{
                    Path                    = (Join-Path $PoshPath $FileStamp)
                    TableStyle              = 'Medium2'
                    FreezeTopRowFirstColumn = $true
                    AutoSize                = $true
                    BoldTopRow              = $true
                    ClearSheet              = $true
                }
                if ($SkipHealthMessages) {
                    $allMessageTrackResults.where{ $_.sender -notmatch 'contoso|MicrosoftExchange|HealthMailbox|maildeliveryprobe' -and $_.recipients -notmatch 'healthmailbox' } | Export-Excel @ExcelSplat
                }
                else {
                    $allMessageTrackResults | Export-Excel @ExcelSplat
                }

            }
            if ($ExportToCSV) {
                $FileStamp = 'OnPremises-MessageTrace_{0}.csv' -f [DateTime]::Now.ToString('yyyy-MM-dd-hhmm')
                $CsvSplat = @{
                    Path              = (Join-Path $PoshPath $FileStamp)
                    NoTypeInformation = $true
                    Encoding          = 'UTF8'
                }
                if ($SkipHealthMessages) {
                    $allMessageTrackResults.where{ $_.sender -notmatch 'contoso|MicrosoftExchange|HealthMailbox|maildeliveryprobe' -and $_.recipients -notmatch 'healthmailbox' } | Export-Csv @CsvSplat

                }
                else {
                    $allMessageTrackResults | Export-Csv @CsvSplat
                }
            }
            return
        }

        Write-Verbose "`n$($allMessageTrackResults.count) results returned."
        if ($SkipHealthMessages) {
            $WantsToTrackMoreSpecifically = $allMessageTrackResults.where{ $_.sender -notmatch 'contoso|MicrosoftExchange|HealthMailbox|maildeliveryprobe' -and $_.recipients -notmatch 'healthmailbox' } |
            Out-GridView -PassThru -Title 'Message Tracking Log. Select one or more then click OK to track by only those Message IDs.'
        }
        else {
            $WantsToTrackMoreSpecifically = $allMessageTrackResults |
            Out-GridView -PassThru -Title 'Message Tracking Log. Select one or more then click OK to track by only those Message IDs.'
        }

        if ($WantsToTrackMoreSpecifically) {
            foreach ($Wants in $WantsToTrackMoreSpecifically) {
                $allMessageTrackResults = New-Object 'System.Collections.Generic.List[PSObject]'
                try {
                    $messageTrack = $Servers.name | Get-MessageTrackingLog -MessageID $wants.MessageId -ResultSize $ResultSize
                    if ($messageTrack) {
                        $messageTrack | ForEach-Object {
                            $messageTrackResults = [PSCustomObject][ordered]@{
                                Time             = $_.TimeStamp
                                Directionality   = $_.Directionality
                                EventID          = $_.EventID
                                Sender           = $_.Sender
                                Recipients       = $_.Recipients
                                Subject          = $_.MessageSubject
                                Connector        = $_.ConnectorID
                                SourceContext    = $_.SourceContext
                                EventData        = $_.EventData
                                ServerHostName   = $_.ServerHostName
                                ServerIP         = $_.ServerIP
                                ClientIP         = $_.ClientIP
                                OriginalClientIP = $_.OriginalClientIP
                                ClientHostName   = $_.ClientHostName
                                TotalBytes       = $_.TotalBytes
                                MessageId        = $_.MessageId
                            }
                            $allMessageTrackResults.Add($messageTrackResults)
                        }
                        else {
                            Write-Verbose "`tNo results found"
                        }
                    }
                }
                catch {
                    Write-Verbose "`tException gathering message trace data."
                }
                $allMessageTrackResults | Out-GridView -Title "MessageID: $($Wants.MessageID)"
            }
        }
    }
    else {
        Write-Verbose "`nNo Results found."
    }
    $ErrorActionPreference = $currentErrorActionPrefs
}
