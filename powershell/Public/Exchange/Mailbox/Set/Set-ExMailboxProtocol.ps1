<#
    .SYNOPSIS
    Set Exchange mailbox protocol settings.

    .DESCRIPTION
    This function sets the protocol settings (MAPI, OWA, IMAP, POP, EWS, ActiveSync)
    for Exchange Online mailboxes. It can target mailboxes by identity, by domain, from a CSV file, or all mailboxes in the organization.
    It also allows enabling modern protocols only or modern protocols along with EWS.
    Additionally, it can configure SMTP Client Authentication settings at the mailbox level.

    .PARAMETER Identity
    The identity of the mailbox(es) to set protocol settings for. This can be an array of email addresses, usernames, or display names.

    .PARAMETER ByDomain
    Filter mailboxes by domain name. All mailboxes with a primary SMTP address in this domain will be processed.

    .PARAMETER FromCSV
    The path to a CSV file containing mailbox identities to process.

    .PARAMETER AllMailboxes
    If specified, all mailboxes in the organization will be processed.

    .PARAMETER GenerateCmdlets
    If specified, the function will generate the Set-EXOCasMailbox cmdlets and save them to a file instead of executing them.

    .PARAMETER OutputFile
    The path to the output file where generated cmdlets will be saved. Default is a timestamped file in the current directory.

    .PARAMETER ModernProtocolsOnly
    If specified, enables only modern protocols (MAPI, OWA, Outlook Mobile) and disables legacy protocols (POP, IMAP, ActiveSync, EWS) along with disabling SMTP Client Authentication
    at the mailbox level.

    .PARAMETER ModernProtocolsAndEws
    If specified, enables modern protocols (MAPI, OWA, Outlook Mobile) along with EWS, and disables legacy protocols (POP, IMAP, ActiveSync) along with disabling SMTP Client
    Authentication at the mailbox level.

    .PARAMETER MAPIEnabled
    Set MAPI protocol enabled or disabled.

    .PARAMETER OWAEnabled
    Set OWA protocol enabled or disabled.

    .PARAMETER PopEnabled
    Set POP protocol enabled or disabled.

    .PARAMETER ImapEnabled
    Set IMAP protocol enabled or disabled.

    .PARAMETER ActiveSyncEnabled
    Set ActiveSync protocol enabled or disabled.

    .PARAMETER EwsEnabled
    Set EWS protocol enabled or disabled.

    .PARAMETER SmtpClientAuthenticationDisabled
    Set SMTP Client Authentication disabled or enabled at the mailbox level.

    .PARAMETER UniversalOutlookEnabled
    Set Universal Outlook (the built-in Mail and Calendar app in Windows) enabled or disabled.

    .PARAMETER OutlookMobileEnabled
    Set Outlook Mobile access enabled or disabled.

    .EXAMPLE
    Set-ExMailboxProtocol -Identity "user@example.com" -ModernProtocolsOnly

    Sets modern protocols only for the specified mailbox.

    .EXAMPLE
    Set-ExMailboxProtocol -ByDomain "example.com" -GenerateCmdlets -OutputFile "C:\path\to\commands.ps1"

    Generates the Set-EXOCasMailbox cmdlets for all mailboxes in the specified domain to enable modern protocols only and saves them to the specified file without executing them.

    .EXAMPLE
    Set-ExMailboxProtocol -FromCSV "C:\path\to\file.csv"

    Sets protocol settings for mailboxes listed in the specified CSV file.

    .EXAMPLE
    Set-ExMailboxProtocol -AllMailboxes -ModernProtocolsAndEws

    Sets modern protocols along with EWS for all mailboxes in the organization.
#>
function Set-ExMailboxProtocol {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
        [string[]]$Identity,

        [Parameter(Mandatory = $true, ParameterSetName = 'ByDomain')]
        [string]$ByDomain,

        [Parameter(Mandatory = $true, ParameterSetName = 'FromCSV')]
        [string]$FromCSV,

        [Parameter(Mandatory = $true, ParameterSetName = 'AllMailboxes')]
        [switch]$AllMailboxes,

        [Parameter(Mandatory = $false)]
        [switch]$GenerateCmdlets,

        [Parameter(Mandatory = $false)]
        [string]$OutputFile = "$(Get-Date -Format 'yyyy-MM-dd_HHmmss')-SetExMailboxProtocol_Commands.txt",

        [Parameter(Mandatory = $false)]
        [switch]$ModernProtocolsOnly,

        # for migration projects where EWS is still needed
        [Parameter(Mandatory = $false)]
        [switch]$ModernProtocolsAndEws,

        <#Exchange on-premise only
        [Parameter(Mandatory = $false)]
        [bool]$ECPEnabled,
        #>

        [Parameter(Mandatory = $false)]
        [bool]$MAPIEnabled,

        [Parameter(Mandatory = $false)]
        [bool]$OWAEnabled,

        [Parameter(Mandatory = $false)]
        [bool]$PopEnabled,

        [Parameter(Mandatory = $false)]
        [bool]$ImapEnabled,

        [Parameter(Mandatory = $false)]
        [bool]$ActiveSyncEnabled,

        [Parameter(Mandatory = $false)]
        [bool]$EwsEnabled,

        [Parameter(Mandatory = $false)]
        [bool]$SmtpClientAuthenticationDisabled,

        [Parameter(Mandatory = $false)]
        [bool]$UniversalOutlookEnabled,

        [Parameter(Mandatory = $false)]
        [bool]$OutlookMobileEnabled
    )

    # Check for conflicting parameters
    $individualProtocolParams = @('ECPEnabled', 'MAPIEnabled', 'OWAEnabled', 'PopEnabled', 'ImapEnabled', 'ActiveSyncEnabled', 'EwsEnabled', 'SmtpClientAuthenticationDisabled', 'UniversalOutlookEnabled', 'OutlookMobileEnabled')
    $providedIndividualParams = $individualProtocolParams | Where-Object { $PSBoundParameters.ContainsKey($_) }
    
    if (($ModernProtocolsOnly -or $ModernProtocolsAndEws) -and $providedIndividualParams.Count -gt 0) {
        Write-Error 'Cannot use -ModernProtocolsOnly or -ModernProtocolsAndEws with individual protocol parameters. Choose either -ModernProtocolsOnly or -ModernProtocolsAndEws OR individual protocol parameters, not both.'
        return
    }

    if ($ModernProtocolsOnly -and $ModernProtocolsAndEws) {
        Write-Error 'Cannot use both -ModernProtocolsOnly and -ModernProtocolsAndEws switches together. Please choose one.'
        return
    }

    # Define protocol settings
    $protocolSettings = @{}
    
    if ($ModernProtocolsOnly) {
        $protocolSettings = @{
            MAPIEnabled                      = $true
            OWAEnabled                       = $true
            PopEnabled                       = $false
            ImapEnabled                      = $false
            ActiveSyncEnabled                = $false
            EwsEnabled                       = $false
            SmtpClientAuthenticationDisabled = $true
            # Admins can use the UniversalOutlookEnabled parameter value $false on the CASMailbox cmdlet to block organization accounts from using the built-in Mail and Calendar app in Windows (https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/manage/policy-management#disable-toggle-from-classic-outlook-for-windows)
            UniversalOutlookEnabled          = $false
            OutlookMobileEnabled             = $true
        }
    }
    elseif ($ModernProtocolsAndEws) {
        $protocolSettings = @{
            MAPIEnabled                      = $true
            OWAEnabled                       = $true
            PopEnabled                       = $false
            ImapEnabled                      = $false
            ActiveSyncEnabled                = $false
            EwsEnabled                       = $true
            SmtpClientAuthenticationDisabled = $true
            # Admins can use the UniversalOutlookEnabled parameter value $false on the CASMailbox cmdlet to block organization accounts from using the built-in Mail and Calendar app in Windows (https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/manage/policy-management#disable-toggle-from-classic-outlook-for-windows)
            UniversalOutlookEnabled          = $false
            OutlookMobileEnabled             = $true
        }
    }
    else {
        # Use individual parameters if specified
        if ($PSBoundParameters.ContainsKey('MAPIEnabled')) { $protocolSettings['MAPIEnabled'] = $MAPIEnabled }
        if ($PSBoundParameters.ContainsKey('OWAEnabled')) { $protocolSettings['OWAEnabled'] = $OWAEnabled }
        if ($PSBoundParameters.ContainsKey('PopEnabled')) { $protocolSettings['PopEnabled'] = $PopEnabled }
        if ($PSBoundParameters.ContainsKey('ImapEnabled')) { $protocolSettings['ImapEnabled'] = $ImapEnabled }
        if ($PSBoundParameters.ContainsKey('ActiveSyncEnabled')) { $protocolSettings['ActiveSyncEnabled'] = $ActiveSyncEnabled }
        if ($PSBoundParameters.ContainsKey('EwsEnabled')) { $protocolSettings['EwsEnabled'] = $EwsEnabled }
        if ($PSBoundParameters.ContainsKey('SmtpClientAuthenticationDisabled')) { $protocolSettings['SmtpClientAuthenticationDisabled'] = $SmtpClientAuthenticationDisabled }
        if ($PSBoundParameters.ContainsKey('UniversalOutlookEnabled')) { $protocolSettings['UniversalOutlookEnabled'] = $UniversalOutlookEnabled }
        if ($PSBoundParameters.ContainsKey('OutlookMobileEnabled')) { $protocolSettings['OutlookMobileEnabled'] = $OutlookMobileEnabled }
    }

    if ($PSCmdlet.ParameterSetName -eq 'ByDomain') {
        # we can't filter PrimarySmtpAddress with `-like '*domain'` so we first find mailboxes with emailaddresses matching the domain then filter primarySMTPAddress
        # https://michev.info/blog/post/2404/exchange-online-now-supports-the-like-operator-for-primarysmtpaddress-filtering-sort-of

        Write-Host -ForegroundColor Cyan "Searching mailboxes with PrimarySmtpAddress matching '*@$ByDomain'"
        $Mailboxes = Get-EXOCasMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'" | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'AllMailboxes') {
        $Mailboxes = Get-EXOCasMailbox -ResultSize Unlimited
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'FromCSV') {
        [System.Collections.Generic.List[PSCustomObject]]$Mailboxes = @()
        if (-not (Test-Path $FromCSV)) {
            Write-Error "CSV file not found: $FromCSV"
            return
        }

        $csvData = Import-Csv -Path $FromCSV

        foreach ($row in $csvData) {
            try {
                $Mailbox = Get-EXOCasMailbox -Identity $row.PrimarySmtpAddress -ErrorAction Stop
                $Mailbox | Add-Member -NotePropertyName 'CSVEnableProtocols' -NotePropertyValue $row.EnableProtocols -Force
                $Mailbox | Add-Member -NotePropertyName 'CSVDisableProtocols' -NotePropertyValue $row.DisableProtocols -Force

                $Mailboxes.Add($Mailbox)
            }
            catch {
                Write-Warning "Mailbox not found: $($row.PrimarySmtpAddress)"
            }
        }
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Identity') {
        [System.Collections.Generic.List[PSCustomObject]]$Mailboxes = @()

        foreach ($id in $Identity) {
            try {
                $mbx = Get-EXOCasMailbox -Identity $id -ErrorAction Stop
                $Mailboxes.Add($mbx)
            }
            catch {
                Write-Warning "Mailbox not found: $id"
            }
        }
    }
    else {
        Write-Warning 'ParameterSetName Identity is not supported in this function.'
        return 1
    }

    $Commands = @()


    foreach ($casMailbox in $Mailboxes) {
        $cmdParams = @{
            Identity = $casMailbox.PrimarySmtpAddress
        }

        # Apply protocol settings from ModernProtocolsOnly or individual parameters
        foreach ($setting in $protocolSettings.GetEnumerator()) {
            $cmdParams[$setting.Key] = $setting.Value
        }

        if ($casMailbox.CSVEnableProtocols) {
            $protocolsToEnable = $casMailbox.CSVEnableProtocols -split ',' | ForEach-Object { $_.Trim() }
            foreach ($protocol in $protocolsToEnable) {
                $cmdParams[$protocol] = $true
            }
        }
        if ($casMailbox.CSVDisableProtocols) {
            $protocolsToDisable = $casMailbox.CSVDisableProtocols -split ',' | ForEach-Object { $_.Trim() }
            foreach ($protocol in $protocolsToDisable) {
                $cmdParams[$protocol] = $false
            }
        }

        $command = 'Set-EXOCasMailbox ' + ($cmdParams.GetEnumerator() | ForEach-Object { 
                $value = if ($_.Value -is [bool]) { 
                    if ($_.Value) { '$true' } else { '$false' }
                }
                else { 
                    "`"$($_.Value)`"" 
                }
                "-$($_.Key) $value"
            }) -join ' '
        $Commands += $command

        if (-not $GenerateCmdlets -and $PSCmdlet.ShouldProcess($casMailbox.PrimarySmtpAddress, 'Set EXO Mailbox Protocols')) {
            Write-Host -ForegroundColor Cyan "$($casMailbox.PrimarySmtpAddress) - Setting protocols..."
            try {
                Set-CasMailbox @cmdParams -ErrorAction Stop
                Write-Host "$($casMailbox.PrimarySmtpAddress) - Configuration applied successfully." -ForegroundColor Green
            }
            catch {
                Write-Host "$($casMailbox.PrimarySmtpAddress) - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    if ($GenerateCmdlets) {
        $Commands | Out-File -FilePath $OutputFile -Encoding UTF8
        Write-Host "Generated cmdlets saved to $OutputFile" -ForegroundColor Yellow
    }
}