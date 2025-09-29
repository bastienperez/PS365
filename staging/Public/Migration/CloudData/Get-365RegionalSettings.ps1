function Get-365RegionalSettings {
    <#
    .SYNOPSIS
    Export all Mailboxes Regional Configuration to a PSCustomObject

    .DESCRIPTION
    Export all Mailboxes Regional Configuration to a PSCustomObject.  Can then be exported to CSV etc.

    .EXAMPLE
    Get-365RegionalSettings | Export-Csv .\RegionalSettings.csv -notypeinformation

    .NOTES
    Results can be used with Set-365RegionalSettings during a migration.
    #>

    [CmdletBinding()]
    param (

    )

    $MailboxList = Get-Exoailbox -Properties ExchangeGuid -ResultSize Unlimited

    foreach ($Mailbox in $MailboxList) {
        try {
            $Config = Get-MailboxRegionalConfiguration -Identity $Mailbox.ExchangeGuid.ToString() -ErrorAction Stop
            [PSCustomObject][ordered]@{
                DisplayName        = $Mailbox.DisplayName
                PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                ExchangeGuid       = $Mailbox.ExchangeGuid
                Language           = $Config.Language
                TimeZone           = $Config.TimeZone
                DateFormat         = $Config.DateFormat
                TimeFormat         = $Config.TimeFormat
                Log                = 'SUCCESS'
            }
        }
        catch {
            [PSCustomObject][ordered]@{
                DisplayName        = $Mailbox.DisplayName
                PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                ExchangeGuid       = $Mailbox.ExchangeGuid
                Language           = ''
                TimeZone           = ''
                DateFormat         = ''
                TimeFormat         = ''
                Log                = $_.Exception.Message
            }
        }
    }
}
