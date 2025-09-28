function Get-ExCalendarInfo {
    [System.Collections.Generic.List[Object]]$calendarInfoArray = @()

    if (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue) {
        # Get-EXOMailbox is available, so we are in EXO
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited
    }
    else {
        # Get-EXOMailbox is not available, so we are in on-premises Exchange
        $mailboxes = Get-Mailbox -ResultSize Unlimited
    }

    $i = 0
    foreach ($mbx in $mailboxes) {
        $i++
        Write-Host -ForegroundColor Cyan "Processing mailbox $($mbx.PrimarySmtpAddress) ($i / $($mailboxes.Count))"
	
        # Do not use (Get-MailboxFolderStatistics $mbx.PrimarySmtpAddress -FolderScope calendar).Name  because it returns other calendars like BirthdayCalendar or user created calendar
        $calFolderName = [string](Get-MailboxFolderStatistics $mbx.PrimarySmtpAddress -FolderScope Calendar | Where-Object { $_.FolderType -eq 'Calendar' }).Name 

        $calFolder = $mbx.PrimarySmtpAddress 
        $calFolder += ":\$calFolderName"
    
        #$calFolder += [string](Get-MailboxFolderStatistics $mbx.PrimarySmtpAddress | Where-Object {$_.FolderType -eq 'Calendar'}).Name 
    
        # select PrimarySmtpAddress but not exist
        Get-MailboxFolderPermission -Identity $calFolder | Select-Object PrimarySmtpAddress, Identity, FolderName, User, AccessRights | ForEach-Object {
            $_.PrimarySmtpAddress = $mbx.PrimarySmtpAddress
            $calendarInfoArray.Add($_)
        }

        $calFolderSharing = Get-MailboxCalendarFolder -Identity $calFolder

        $object = [PSCustomObject][ordered]@{
            PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
            FolderName           = "$calFolderName-Published"
            User                 = $calFolderSharing.PublishedCalendarUrl
            AccessRights         = if ($($calFolderSharing.PublishEnabled)) { 'Published' } else { 'None' }
            SearchableUrlEnabled = $calFolderSharing.SearchableUrlEnabled
        }

        $calendarInfoArray.Add($object)
        Start-Sleep -Milliseconds 10
    }

    return $calendarInfoArray
}