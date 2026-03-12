<#
    .SYNOPSIS
    Get the working hours configuration of Exchange Online mailboxes.

    .DESCRIPTION
    This function retrieves the working hours configuration from the calendar settings
    for Exchange Online mailboxes. It returns the time zone, start/end times, and enabled
    days for each mailbox. The function can filter mailboxes by identity or by domain.

    .PARAMETER Identity
    The identity of the mailbox to retrieve the working hours configuration for.
    If not specified, the function retrieves the configuration for all mailboxes.

    .PARAMETER ByDomain
    The domain to filter mailboxes by. Only mailboxes with a primary SMTP address in this
    domain will be processed.

    .EXAMPLE
    Get-ExMailboxWorkingHours

    Retrieves the working hours configuration for all Exchange Online mailboxes.

    .EXAMPLE
    Get-ExMailboxWorkingHours -Identity "user@example.com"

    Retrieves the working hours configuration for the specified mailbox.

    .EXAMPLE
    Get-ExMailboxWorkingHours -ByDomain "example.com"

    Retrieves the working hours configuration for all mailboxes in the specified domain.

    .NOTES
    Requires connection to Exchange Online using Connect-ExchangeOnline.
    The function retrieves data from Get-MailboxCalendarConfiguration.

    .PARAMETER ExportToExcel
    If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-ExMailboxWorkingHours -ExportToExcel
    Exports results to an Excel file.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExMailboxWorkingHours
#>

function Get-ExMailboxWorkingHours {
    param (
        [Parameter(Mandatory = $false, position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [string]$ByDomain,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    [System.Collections.Generic.List[PSCustomObject]]$exoMbxWorkingHoursArray = @()

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

    foreach ($mbx in $mailboxes) {
        # This CMDlet returns warning "WARNING:[...] User Get-EventsFromEmailConfiguration [...]", but the CMDlet Get-MailboxCalendarConfiguration is the only way to get the working hours configuration, so we will ignore the warning for now.
        $calendarConfig = Get-MailboxCalendarConfiguration -Identity $mbx.PrimarySmtpAddress -WarningAction SilentlyContinue

        $object = [PSCustomObject][ordered]@{ 
            DisplayName           = $mbx.DisplayName
            PrimarySmtpAddress    = $mbx.PrimarySmtpAddress
            WorkingHoursTimeZone  = $calendarConfig.WorkingHoursTimeZone
            WorkingHoursStartTime = $calendarConfig.WorkingHoursStartTime
            WorkingHoursEndTime   = $calendarConfig.WorkingHoursEndTime
            WorkingHoursMonday    = $calendarConfig.WorkingHoursMonday
            WorkingHoursTuesday   = $calendarConfig.WorkingHoursTuesday
            WorkingHoursWednesday = $calendarConfig.WorkingHoursWednesday
            WorkingHoursThursday  = $calendarConfig.WorkingHoursThursday
            WorkingHoursFriday    = $calendarConfig.WorkingHoursFriday
            WorkingHoursSaturday  = $calendarConfig.WorkingHoursSaturday
            WorkingHoursSunday    = $calendarConfig.WorkingHoursSunday
            MailboxWhenCreated    = $mbx.WhenCreated
            MailboxWhenModified   = $mbx.WhenChanged
        }

        $exoMbxWorkingHoursArray.Add($object)
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-ExMailboxWorkingHours.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $exoMbxWorkingHoursArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExMailboxWorkingHours'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $exoMbxWorkingHoursArray
    }
}
