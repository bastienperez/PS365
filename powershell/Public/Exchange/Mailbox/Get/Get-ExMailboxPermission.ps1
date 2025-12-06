<#
    .SYNOPSIS
    Gets all permissions on an Exchange Online mailbox (Full Access, Send As, Send on Behalf)

    .DESCRIPTION
    This function retrieves and displays all permissions granted on a specific mailbox:
    - Full Access: Full access permissions to the mailbox
    - Send As: Permissions to send emails as the mailbox
    - Send on Behalf: Permissions to send emails on behalf of the mailbox

    .PARAMETER Identity
    The identity of the mailbox (email address, username, or display name)

    .PARAMETER ByDomain
    Filter mailboxes by domain name

    .PARAMETER UserPermission
    Find all permissions that a specific user has across all mailboxes (reverse lookup)

    .PARAMETER ExportToExcel
    Export the mailbox permissions to an Excel file

    .EXAMPLE
    Get-ExMailboxPermission -Identity "john.doe@contoso.com"
    Gets all permissions for the mailbox john.doe@contoso.com

    .EXAMPLE
    Get-ExMailboxPermission -Identity "john.doe@contoso.com" -ExportToExcel
    Gets all permissions for the mailbox and exports them to an Excel file

    .EXAMPLE
    Get-ExMailboxPermission -ByDomain "contoso.com"
    Gets all permissions for all mailboxes in the contoso.com domain

    .EXAMPLE
    Get-ExMailboxPermission -UserPermission "john.doe@contoso.com"
    Finds all mailbox permissions that john.doe@contoso.com has across all mailboxes

    .NOTES
    Requires the ExchangeOnlineManagement module and an active connection to Exchange Online
    For Excel export functionality, requires the ImportExcel module
#>

function Get-ExMailboxPermission {
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [string]$ByDomain,

        [Parameter(Mandatory = $false)]
        [string]$UserPermission,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    [System.Collections.Generic.List[PSCustomObject]] $allPermissions = @()

    # Determine which mailboxes to process
    if ($UserPermission) {
        Write-Host "Finding all permissions for user: $UserPermission" -ForegroundColor Green
        
        # Get all mailboxes to check permissions against
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited
        Write-Host "Checking permissions across $($mailboxes.Count) mailbox(es)" -ForegroundColor Yellow
        
        # Set flag for reverse lookup mode
        $isUserPermissionLookup = $true
    }
    elseif ($ByDomain) {
        Write-Host "Retrieving permissions for all mailboxes in domain: $ByDomain" -ForegroundColor Green
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'" | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
        Write-Host "Found $($mailboxes.Count) mailbox(es) in domain $ByDomain" -ForegroundColor Yellow
        $isUserPermissionLookup = $false
    }
    elseif ($Identity) {
        Write-Host "Retrieving permissions for mailbox: $Identity" -ForegroundColor Green
        try {
            $mailboxes = @(Get-EXOMailbox -Identity $Identity -ErrorAction Stop)
            Write-Host "Mailbox found: $($mailboxes[0].DisplayName) ($($mailboxes[0].PrimarySmtpAddress))" -ForegroundColor Yellow
        }
        catch {
            Write-Error "Error retrieving mailbox '$Identity': $($_.Exception.Message)"
            return $null
        }
        $isUserPermissionLookup = $false
    }
    else {
        Write-Host 'Retrieving permissions for all mailboxes' -ForegroundColor Green
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited
        Write-Host "Found $($mailboxes.Count) mailbox(es)" -ForegroundColor Yellow
        $isUserPermissionLookup = $false
    }

    $totalMailboxes = $mailboxes.Count
    $currentMailbox = 0

    foreach ($mailbox in $mailboxes) {
        $currentMailbox++
        Write-Host "Processing mailbox $currentMailbox/$totalMailboxes : $($mailbox.DisplayName)" -ForegroundColor Cyan

        try {
            # 1. Get Full Access permissions
            $fullAccessPerms = @(Get-EXOMailboxPermission -Identity $mailbox.PrimarySmtpAddress | Where-Object {
                    $_.AccessRights -contains 'FullAccess' -and 
                    $_.User -notlike 'NT AUTHORITY\*' -and 
                    $_.User -notlike 'S-1-*' -and
                    $_.Deny -eq $false
                })
            
            foreach ($perm in $fullAccessPerms) {
                # If UserPermission mode, only add permissions for the specified user
                if ($isUserPermissionLookup -and $perm.User -ne $UserPermission) {
                    continue
                }
                
                $allPermissions.Add([PSCustomObject]@{
                        MailboxIdentity    = $mailbox.PrimarySmtpAddress
                        MailboxDisplayName = $mailbox.DisplayName
                        MailboxEmail       = $mailbox.PrimarySmtpAddress
                        PermissionType     = 'Full Access'
                        User               = $perm.User
                        AccessRights       = ($perm.AccessRights -join ', ')
                        InheritanceType    = $perm.InheritanceType
                        IsInherited        = $perm.IsInherited
                    })
            }
            
            # 2. Get Send As permissions
            $sendAsPerms = @(Get-EXORecipientPermission -Identity $mailbox.PrimarySmtpAddress | Where-Object {
                    $_.AccessRights -contains 'SendAs' -and 
                    $_.Trustee -notlike 'NT AUTHORITY\*' -and 
                    $_.Trustee -notlike 'S-1-*'
                })
            
            foreach ($perm in $sendAsPerms) {
                # If UserPermission mode, only add permissions for the specified user
                if ($isUserPermissionLookup -and $perm.Trustee -ne $UserPermission) {
                    continue
                }
                
                $allPermissions.Add([PSCustomObject]@{
                        MailboxIdentity    = $mailbox.PrimarySmtpAddress
                        MailboxDisplayName = $mailbox.DisplayName
                        MailboxEmail       = $mailbox.PrimarySmtpAddress
                        PermissionType     = 'Send As'
                        User               = $perm.Trustee
                        AccessRights       = ($perm.AccessRights -join ', ')
                        InheritanceType    = $perm.InheritanceType
                        IsInherited        = $perm.IsInherited
                    })
            }
            
            # 3. Get Send on Behalf permissions
            $sendOnBehalfUsers = @($mailbox.GrantSendOnBehalfTo)
            if ($sendOnBehalfUsers -and $sendOnBehalfUsers.Count -gt 0) {
                foreach ($user in $sendOnBehalfUsers) {
                    # If UserPermission mode, only add permissions for the specified user
                    if ($isUserPermissionLookup -and $user -ne $UserPermission) {
                        continue
                    }
                    
                    $allPermissions.Add([PSCustomObject]@{
                            MailboxIdentity    = $mailbox.PrimarySmtpAddress
                            MailboxDisplayName = $mailbox.DisplayName
                            MailboxEmail       = $mailbox.PrimarySmtpAddress
                            PermissionType     = 'Send on Behalf'
                            User               = $user
                            AccessRights       = 'SendOnBehalf'
                            InheritanceType    = 'None'
                            IsInherited        = $false
                        })
                }
            }
        }
        catch {
            Write-Warning "Error processing mailbox $($mailbox.PrimarySmtpAddress): $($_.Exception.Message)"
        }
    }
    
    # Display results or export to Excel
    if ($allPermissions.Count -gt 0) {
        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $filenameSuffix = if ($UserPermission) { "User-$($UserPermission -replace '[<>:"/\\|?*]', '_')" } elseif ($ByDomain) { "Domain-$($ByDomain -replace '[<>:"/\\|?*]', '_')" } elseif ($Identity) { $Identity -replace '[<>:"/\\|?*]', '_' } else { 'AllMailboxes' }
            $excelFilePath = "$($env:userprofile)\$now-ExMailboxPermissions-$filenameSuffix.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting mailbox permissions to Excel file: $excelFilePath"
            $allPermissions | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExchangeMailboxPermissions'
            Write-Host -ForegroundColor Green 'Export completed successfully!'
        }
        else {
            Write-Host "`n=== PERMISSIONS SUMMARY ===" -ForegroundColor Yellow
            Write-Host "Total permissions found: $($allPermissions.Count)" -ForegroundColor Yellow
            
            # Group by permission type
            $groupedPerms = $allPermissions | Group-Object PermissionType
            foreach ($group in $groupedPerms) {
                Write-Host "`n--- $($group.Name) ---" -ForegroundColor Cyan
                $group.Group | Format-Table User, AccessRights, IsInherited -AutoSize
            }
            
            return $allPermissions
        }
    }
    else {
        Write-Host "`nNo permissions found." -ForegroundColor Yellow
        return $null
    }
}