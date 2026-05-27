<#
    .SYNOPSIS
    Tests whether delegates with calendar folder permissions still exist.

    .DESCRIPTION
    Retrieves the calendar folder permissions for a given mailbox and checks
    whether each delegate (the person the calendar was shared with) is still
    a valid recipient. Default and anonymous entries are excluded.

    .PARAMETER Mailbox
    The mailbox to inspect (e.g. user@domain.com).

    .PARAMETER DelegateStatus
    Filters the results based on whether the delegate (the person with whom
    the calendar has been shared, i.e. who appears in the calendar folder
    permissions) still exists as a valid recipient in the tenant.
    - All       : returns every delegate, regardless of existence (default).
    - Existing  : returns only delegates that still exist as recipients.
    - Missing   : returns only delegates that no longer exist (orphaned
                  permissions left behind after the account was deleted).

    .EXAMPLE
    Test-ExCalendarPermission -Mailbox "user@domain.com"

    Tests all calendar permissions for the specified mailbox and returns a list of delegates with their existence status.

    .EXAMPLE
    Test-ExCalendarPermission -Mailbox "user@domain.com" -DelegateStatus Missing

    Tests calendar permissions for the specified mailbox and returns only delegates that no longer exist as recipients (orphaned permissions).
#>

function Test-ExCalendarPermission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Mailbox,

        [Parameter()]
        [ValidateSet('All', 'Existing', 'Missing')]
        [string]$DelegateStatus = 'All'
    )

    Write-Verbose "Inspecting mailbox '$Mailbox' with DelegateStatus filter '$DelegateStatus'"

    $excluded = @('Default', 'Anonymous', 'Par défaut', 'Anonyme')

    Write-Verbose "Resolving Calendar folder name for '$Mailbox'"
    $calendarFolder = (Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope Calendar |
        Where-Object { $_.FolderType -eq 'Calendar' }).Name
    $folder = "${Mailbox}:\$calendarFolder"
    Write-Verbose "Calendar folder resolved to '$folder'"

    [System.Collections.Generic.List[PSCustomObject]]$calendarPermissions = @()

    $permissions = Get-MailboxFolderPermission -Identity $folder |
    Where-Object { $_.User.DisplayName -notin $excluded }

    Write-Verbose "Found $($permissions.Count) permission entries to check"

    foreach ($permission in $permissions) {
        $delegate = $permission.User.DisplayName
        Write-Verbose "Checking delegate '$delegate'"
        $delegateExists = [bool](Get-Recipient -Identity $delegate -ErrorAction SilentlyContinue)
        Write-Verbose "  Exists = $delegateExists"

        $keep = switch ($DelegateStatus) {
            'Existing' { $delegateExists }
            'Missing' { -not $delegateExists }
            default { $true }
        }
        Write-Verbose "  DelegateStatus='$DelegateStatus' -> keep = $keep"

        if ($keep) {
            $object = [PSCustomObject][ordered]@{
                Identity                        = $folder
                Delegate                        = $delegate
                AccessRights                    = $permission.AccessRights -join ', '
                DelegateExistsInThisEnvironment = $delegateExists
            }
            $calendarPermissions.Add($object)
        }
    }

    Write-Verbose "Returning $($calendarPermissions.Count) result(s)"
    return $calendarPermissions
}