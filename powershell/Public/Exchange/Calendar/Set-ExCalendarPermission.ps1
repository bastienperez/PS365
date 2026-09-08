<#
    .SYNOPSIS
    Sets calendar folder permissions for a mailbox in Exchange Online.

    .DESCRIPTION
    Grants or updates calendar permissions for a delegate user on a target mailbox.
    The calendar folder is automatically resolved from the mailbox identity.

    .PARAMETER Identity
    The identity of one or more mailboxes whose calendar permissions will be modified (email address, username, or display name).

    .PARAMETER AllMailboxes
    Apply the calendar permission to all mailboxes in the organization.

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
    Set-ExCalendarPermission -Identity "john.doe@contoso.com" -AccessRights FreeBusyTimeOnly -DelegateUser "jane.doe@contoso.com"

    Grants free/busy visibility only to Jane on John's calendar.

    .EXAMPLE
    Set-ExCalendarPermission -AllMailboxes -AccessRights Reviewer -DelegateUser "auditor@contoso.com" -WhatIf

    Previews granting Reviewer access to the delegate on every mailbox calendar.

    .NOTES
    Requires the ExchangeOnlineManagement module and an active connection to Exchange Online.
#>

function Set-ExCalendarPermission {
    [CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Identity')]
    param (
        [Parameter(Mandatory = $false, ParameterSetName = 'Identity', Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$Identity,

        [Parameter(Mandatory = $true, ParameterSetName = 'AllMailboxes')]
        [switch]$AllMailboxes,

        [Parameter(Mandatory = $true)]
        [ValidateSet('None', 'FreeBusyTimeOnly', 'FreeBusyTimeAndSubjectAndLocation', 'Reviewer', 'Contributor', 'NoneditingAuthor', 'Author', 'PublishingAuthor', 'Editor', 'PublishingEditor', 'Owner', 'Custom')]
        [string]$AccessRights,

        [Parameter(Mandatory = $true)]
        [string]$DelegateUser
    )

    begin {
        $pipelineIdentities = [System.Collections.Generic.List[string]]::new()
    }

    process {
        foreach ($pipelineIdentity in $Identity) {
            $pipelineIdentities.Add($pipelineIdentity)
        }
    }

    end {
        if ($PSCmdlet.ParameterSetName -eq 'AllMailboxes') {
            $mailboxes = @(Get-Mailbox -ResultSize Unlimited -ErrorAction Stop)
        }
        else {
            $requestedIdentities = $pipelineIdentities.ToArray()
            if ($requestedIdentities.Count -eq 0) {
                throw 'No mailbox target was specified. Use -Identity or -AllMailboxes.'
            }

            $mailboxes = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($requestedIdentity in $requestedIdentities) {
                try {
                    $mailboxes.Add((Get-Mailbox -Identity $requestedIdentity -ErrorAction Stop))
                }
                catch {
                    Write-Warning "Mailbox not found: $requestedIdentity. $($_.Exception.Message)"
                }
            }
        }

        $totalCount = $mailboxes.Count
        $currentCount = 0
        foreach ($mbx in $mailboxes) {
            $currentCount++
            $mailboxIdentity = [string]$mbx.PrimarySmtpAddress
            if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
                $mailboxIdentity = [string]$mbx.Identity
            }

            if ([string]::IsNullOrWhiteSpace($mailboxIdentity)) {
                Write-Warning "Skipping mailbox $currentCount/$totalCount because it has no usable identity."
                continue
            }

            try {
                $calendarFolder = Get-MailboxFolderStatistics -Identity $mailboxIdentity -FolderScope Calendar -ErrorAction Stop |
                    Where-Object { $_.FolderType -eq 'Calendar' } |
                    Select-Object -First 1

                if ($null -eq $calendarFolder -or [string]::IsNullOrWhiteSpace([string]$calendarFolder.Name)) {
                    throw 'No folder with FolderType Calendar was returned.'
                }

                $calendarFolderIdentity = "$mailboxIdentity`:\$($calendarFolder.Name)"
                $existingPermission = Get-MailboxFolderPermission -Identity $calendarFolderIdentity -User $DelegateUser -ErrorAction SilentlyContinue
                $permissionAction = if ($null -ne $existingPermission) { 'Update' } else { 'Add' }

                if (-not $PSCmdlet.ShouldProcess($calendarFolderIdentity, "$permissionAction calendar permission '$AccessRights' for '$DelegateUser'")) {
                    continue
                }

                if ($null -ne $existingPermission) {
                    Set-MailboxFolderPermission -Identity $calendarFolderIdentity -User $DelegateUser -AccessRights $AccessRights -ErrorAction Stop
                }
                else {
                    Add-MailboxFolderPermission -Identity $calendarFolderIdentity -User $DelegateUser -AccessRights $AccessRights -ErrorAction Stop
                }
            }
            catch {
                Write-Warning "Error processing calendar permissions for mailbox '$mailboxIdentity': $($_.Exception.Message)"
            }
        }
    }
}
