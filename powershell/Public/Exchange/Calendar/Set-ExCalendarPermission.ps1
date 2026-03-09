<#
    .SYNOPSIS
    Sets calendar folder permissions for a mailbox in Exchange Online.

    .DESCRIPTION
    Grants or updates calendar permissions for a delegate user on a target mailbox.
    The calendar folder is automatically resolved from the mailbox identity.

    .PARAMETER Identity
    The identity of the mailbox whose calendar permissions will be modified (email address, username, or display name).

    .PARAMETER AccessRights
    The permission level to grant to the delegate user. Accepted values:

    - None                              - The user has no permissions on the folder.
    - FreeBusyTimeOnly                  - The user can view only free/busy time within the calendar.
    - FreeBusyTimeAndSubjectAndLocation - The user can view free/busy time within the calendar and the subject and location of appointments.
    - Reviewer                          - The user can read all items in the folder.
    - Contributor                       - The user can create items in the folder. The contents of the folder do not appear.
    - NoneditingAuthor                  - The user can create and read all items in the folder, and delete only items that the user creates.
    - Author                            - The user can create and read all items in the folder, and edit and delete only items that the user creates.
    - PublishingAuthor                  - The user can create and read all items in the folder, edit and delete only items that the user creates, and create subfolders.
    - Editor                            - The user can create, read, edit and delete all items in the folder.
    - PublishingEditor                  - The user can create, read, edit, and delete all items in the folder, and create subfolders.
    - Owner                             - The user can create, read, edit, and delete all items in the folder, and create subfolders. The user is both folder owner and folder contact.
    - Custom                            - The user has custom access permissions on the folder.

    .PARAMETER DelegateUser
    The user to whom the calendar permissions will be granted.

    .EXAMPLE
    Set-ExCalendarPermission -Identity "john.doe@contoso.com" -AccessRights Reviewer -DelegateUser "jane.doe@contoso.com"

    Grants read-only access to Jane on John's calendar.

    .EXAMPLE
    Set-ExCalendarPermission -Identity "john.doe@contoso.com" -AccessRights AvailabilityOnly -DelegateUser "jane.doe@contoso.com"

    Grants free/busy visibility only to Jane on John's calendar.

    .NOTES
    Requires the ExchangeOnlineManagement module and an active connection to Exchange Online.
#>

function Set-ExCalendarPermission {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$Identity,
        [Parameter(Mandatory = $true)]
        [ValidateSet('None', 'FreeBusyTimeOnly', 'FreeBusyTimeAndSubjectAndLocation', 'Reviewer', 'Contributor', 'NoneditingAuthor', 'Author', 'PublishingAuthor', 'Editor', 'PublishingEditor', 'Owner', 'Custom')]
        [string]$AccessRights,
        [Parameter(Mandatory = $true)]
        [string]$DelegateUser
    )

    try {
        $mbx = Get-Mailbox -identity $Identity -ErrorAction Stop
        $calFolder = "$($mbx.PrimarySmtpAddress):\"
    }
    catch {
        Write-Warning "Mailbox not found: $Identity"
        return 1
    }
    

    try {
        $calFolder += [string](Get-MailboxFolderStatistics $mbx.PrimarySmtpAddress | Where-Object { $_.FolderType -eq 'Calendar' -and $_.Name -eq 'Calendar' }).Name
    }
    catch {
        Write-Warning "Calendar folder not found for mailbox: $Identity"
        return 1
    }
    
    Add-MailboxFolderPermission -Identity $calFolder -User $DelegateUser -AccessRights $AccessRights
}