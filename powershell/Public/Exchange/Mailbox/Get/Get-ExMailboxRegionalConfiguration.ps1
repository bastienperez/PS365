<#
    .SYNOPSIS
    Get the regional configuration of Exchange Online mailboxes.

    .DESCRIPTION
    This function retrieves the regional configuration settings (language, time zone, date format, time format)
    for Exchange Online mailboxes. It can filter mailboxes by identity or by domain.

    .PARAMETER Identity
    The identity of the mailbox to retrieve the regional configuration for.
    If not specified, the function retrieves the configuration for all mailboxes.

    .PARAMETER ByDomain
    The domain to filter mailboxes by. Only mailboxes with a primary SMTP address in this
    domain will be processed.

    .EXAMPLE
    Get-ExMailboxRegionalConfiguration

    Retrieves the regional configuration for all Exchange Online mailboxes.

    .EXAMPLE
    Get-ExMailboxRegionalConfiguration -Identity "user@example.com"

    Retrieves the regional configuration for the specified mailbox.

    .EXAMPLE
    Get-ExMailboxRegionalConfiguration -ByDomain "example.com"

    Retrieves the regional configuration for all mailboxes in the specified domain.

    .PARAMETER ExportToExcel
    If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-ExMailboxRegionalConfiguration -ExportToExcel
    Exports results to an Excel file.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExMailboxRegionalConfiguration
#>

function Get-ExMailboxRegionalConfiguration {
    param (
        [Parameter(Mandatory = $false, position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [string]$ByDomain,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    [System.Collections.Generic.List[PSCustomObject]]$exoMbxRegionalConfigArray = @()

    # PropertySets All because by default SMTPClientAuthenticationDisabled is not returned
    if ($ByDomain) {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'" -Properties WhenCreated, WhenChanged | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
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
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties WhenCreated, WhenChanged
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
            DisplayName         = $mbx.DisplayName
            PrimarySmtpAddress  = $mbx.PrimarySmtpAddress
            ExchangeObjectId    = $regionalConfig.Identity
            Language            = $regionalConfig.Language
            TimeZone            = $regionalConfig.TimeZone
            DateFormat          = $regionalConfig.DateFormat
            TimeFormat          = $regionalConfig.TimeFormat
            MailboxWhenCreated  = $mbx.WhenCreated
            MailboxWhenModified = $mbx.WhenChanged
        }

        $exoMbxRegionalConfigArray.Add($object)
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-ExMailboxRegionalConfiguration.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $exoMbxRegionalConfigArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExMailboxRegionalConfiguration'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $exoMbxRegionalConfigArray
    }
}