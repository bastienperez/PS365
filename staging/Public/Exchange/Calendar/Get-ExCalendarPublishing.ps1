<#
    .SYNOPSIS
    Get the calendar publishing status for one or more mailboxes.

    .DESCRIPTION
    Get the calendar publishing status for one or more mailboxes,
    including whether publishing is enabled, the detail level, and the published calendar URL.

    .PARAMETER Mailbox
    The mailbox or mailboxes to check the calendar publishing status for.
    If not specified, the status for all mailboxes will be returned.

    .EXAMPLE
    Get-CalendarPublishingStatus
    
    Gets the calendar publishing status for all mailboxes.

    .EXAMPLE
    Get-CalendarPublishingStatus -Mailbox john.doe@domain.com

    Gets the calendar publishing status for the mailbox john.doe@domain.com

    .OUTPUTS
    A list of custom objects with the following properties:
    - UserPrincipalName
    - PrimarySmtpAddress
    - DisplayName
    - CalendarName
    - PublishEnabled
    - DetailLevel
    - PublishedCalendarUrl
#>

#TODO: merge with `Get-ExCalendarInfo` if needed

function Get-CalendarPublishingStatus {
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$Mailbox
    )
    
    [System.Collections.Generic.List[PSCustomObject]]$calendarPublishingArray = @()
    [System.Collections.Generic.List[PSCustomObject]]$mailboxes = @()
    
    if ($Mailbox) {
        foreach ($mb in $Mailbox) {
            $mailboxObj = Get-EXOMailbox -Identity $mb -ErrorAction SilentlyContinue
            
            if ($mailboxObj) {
                $mailboxes.Add($mailboxObj)
            }
            else {
                Write-Host -ForegroundColor Red "Mailbox $mb not found"
            }
        }
    }
    else {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited
    }

    foreach ($mbx in $mailboxes) {
        # Get the name of the default calendar folder (depends on the mailbox's language)
        $calendarFolder = [string](Get-EXOMailboxFolderStatistics $mbx.ExchangeGuid -Folderscope Calendar | Where-Object { $_.FolderType -eq 'Calendar' }).Name 

        # Get users calendar folder settings for their default Calendar folder
        # calendar has the format identity:\<calendar folder name>
        $calendar = Get-MailboxCalendarFolder -Identity "$($mailbox.PrimarySmtpAddress):\$calendarFolder"
    
        $object = [PSCustomObject][ordered]@{
            UserPrincipalName    = $mailbox.UserPrincipalName
            PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
            DisplayName          = $mailbox.DisplayName
            CalendarName         = $calendar.Name
            PublishEnabled       = $calendar.PublishEnabled
            DetailLevel          = $calendar.DetailLevel
            PublishedCalendarUrl = $calendar.PublishedCalendarUrl        
        }

        $calendarPublishingArray.Add($object)
    }

    return $calendarPublishingArray
}