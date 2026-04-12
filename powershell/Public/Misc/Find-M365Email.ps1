<#
    .SYNOPSIS
    Find email addresses across various services.

    .DESCRIPTION
    This function searches for email addresses across Exchange Online recipients,
    Microsoft 365 users, and deleted users. It can filter results by specific email
    addresses or by domain.

    .PARAMETER SearchEmail
    An array of email addresses to search for.

    .PARAMETER ByDomain
    An array of domains to search for.

    .PARAMETER NoPermissionCheck
    (Optional) Skip the Microsoft Graph scope verification performed against the current Get-MgContext token.

    .EXAMPLE
    Find-M365Email -SearchEmail "user@example.com"

    Searches for the specified email address and displays its details if found.

    .EXAMPLE
    Find-M365Email -ByDomain "example.com"

    Searches for all email addresses within the specified domain and displays their details.

    .LINK
    https://ps365.clidsys.com/docs/commands/Find-M365Email
#>

function Find-M365Email {
    [CmdletBinding()]
    param(
        [String[]]$SearchEmail,
        [String]$ByDomain,
        [switch]$NoPermissionCheck
    )

    if (-not $NoPermissionCheck.IsPresent) {
        $requiredScopes = @('User.Read.All')
        if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }
    }

    # $EmailIndex is a hashtable keyed on lower-cased email address pointing at the existing
    # PSCustomObject in $EmailObjects. It makes existence check and update O(1) instead of
    # a linear scan (-> O(N^2) on tens of thousands of addresses).
    function Add-EmailObjects {
        param
        (
            $EmailObjects,
            $EmailIndex,
            $Users
        )
        foreach ($user in $users) {
            foreach ($emailaddress in $user.emailaddresses) {
                $emailaddress = $emailaddress -replace 'X500:', ''
                $emailaddress = $emailaddress -replace 'smtp:', ''
                $emailaddress = $emailaddress -replace 'sip:', ''
                $emailaddress = $emailaddress -replace 'spo:', ''

                if ($ByDomain) {
                    if ($emailaddress -notlike "*$ByDomain") {
                        continue
                    }
                }

                $key = $emailaddress.ToLowerInvariant()
                $existingEmail = $EmailIndex[$key]

                if (-not $existingEmail) {
                    $emailObject = [PSCustomObject]@{
                        EmailAddress  = $emailaddress
                        ObjectID      = $user.objectID
                        DisplayName   = $user.DisplayName
                        RecipientType = $user.RecipientTypeDetails
                        WhenCreated   = $user.WhenCreated
                        Sources       = [System.Collections.Generic.List[string]]@($user.RecipientTypeDetails)
                    }
                    $EmailObjects.Add($emailObject)
                    $EmailIndex[$key] = $emailObject
                }
                else {
                    if (-not $existingEmail.Sources.Contains($user.RecipientTypeDetails)) {
                        $existingEmail.Sources.Add($user.RecipientTypeDetails)
                    }
                }
            }
        }
    }

    $modules = @(
        'ExchangeOnlineManagement'
        'Microsoft.Graph.Authentication'
        'Microsoft.Graph.Users'
    )

    foreach ($module in $modules) {
        try {
            Import-Module $module -ErrorAction Stop
        }
        catch {
            Write-Warning "First, install the $module module: Install-Module $module"
            return
        }
    }

    # Verify Exchange Online session — Get-Recipient will throw if not connected
    try {
        $null = Get-Recipient -ResultSize 1 -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch {
        Write-Warning "Exchange Online session required. Run: Connect-ExchangeOnline"
        return
    }

    [System.Collections.Generic.List[PSCustomObject]]$allM365EmailObjects = @()
    $allM365EmailIndex = @{}

    Write-Host 'Get All Exchange Online recipients...' -ForegroundColor Green
    $allExchangeRecipients = Get-Recipient * -ResultSize unlimited | Select-Object DisplayName, RecipientTypeDetails, EmailAddresses, @{Name = 'objectID'; Expression = { $_.ExternalDirectoryObjectId } }, @{Name = 'WhenCreated'; Expression = { $_.WhenCreatedUTC } }
    
    Write-Host 'Get All softDeletedMailbox...' -ForegroundColor Green
    $softDeleted = Get-Mailbox -SoftDeletedMailbox -ResultSize unlimited | Select-Object DisplayName, RecipientTypeDetails, EmailAddresses, @{Name = 'objectID'; Expression = { $_.ExternalDirectoryObjectId } }, @{Name = 'WhenCreated'; Expression = { $_.WhenCreatedUTC } }


    Write-Host 'Get All Microsoft 365 users...' -ForegroundColor Green
    $entraIDUsers = Get-MgUser -All -Property UserPrincipalName, ID, UserType, ProxyAddresses, CreatedDateTime
    
    $m365UPNUsers = $entraIDUsers | Select-Object DisplayName, @{Name = 'objectID'; Expression = { $_.ID } }, @{Name = 'EmailAddresses'; Expression = { $_.UserPrincipalName } }, @{Name = 'RecipientTypeDetails'; Expression = { if ($_.UserType -eq 'Member' -or $null -eq $_.UserType) { 'Microsoft365User' } else { 'GuestUser' } } }, @{Name = 'WhenCreated'; Expression = { $_.CreatedDateTime } }
    $m365Emails = $entraIDUsers | Select-Object DisplayName, @{Name = 'objectID'; Expression = { $_.ID } }, @{Name = 'EmailAddresses'; Expression = { $_.ProxyAddresses } }, @{Name = 'RecipientTypeDetails'; Expression = { if ($_.UserType -eq 'Member') { 'Microsoft365User' }else { 'GuestUser' } } }, @{Name = 'WhenCreated'; Expression = { $_.CreatedDateTime } }
    $m365AlternateEmails = $entraIDUsers | Select-Object DisplayName, @{Name = 'objectID'; Expression = { $_.ID } }, @{Name = 'EmailAddresses'; Expression = { $_.OtherMails } }, @{Name = 'RecipientTypeDetails'; Expression = { if ($_.UserType -eq 'Member' -or $null -eq $_.UserType) { 'O365UserAlternateEmailAddress' } else { 'GuestUserAlternateEmailAddress' } } }, @{Name = 'WhenCreated'; Expression = { $_.CreatedDateTime } }

    Write-Host 'Get All Microsoft 365 deleted users...' -ForegroundColor Green
    $entraIDDeletedUsers = Get-MgDirectoryDeletedItemAsUser -All -Property UserPrincipalName, ID, UserType, ProxyAddresses, CreatedDateTime

    $entraIDDeletedUsersUPN = $entraIDDeletedUsers | Select-Object DisplayName, @{Name = 'objectID'; Expression = { $_.ID } }, @{Name = 'EmailAddresses'; Expression = { $_.UserPrincipalName } }, @{Name = 'RecipientTypeDetails'; Expression = { if ($_.UserType -eq 'Member') { 'DeletedMicrosoft365User' }else { 'DeletedGuestUser' } } }, @{Name = 'WhenCreated'; Expression = { $_.CreatedDateTime } }
    $entraIDDeletedUsersEmails = $entraIDDeletedUsers | Select-Object DisplayName, @{Name = 'objectID'; Expression = { $_.ID } }, @{Name = 'EmailAddresses'; Expression = { $_.ProxyAddresses } }, @{Name = 'RecipientTypeDetails'; Expression = { if ($_.UserType -eq 'Member') { 'DeletedMicrosoft365User' }else { 'DeletedGuestUser' } } }, @{Name = 'WhenCreated'; Expression = { $_.CreatedDateTime } }

    # Creating email objects collection
    Write-Host 'Creating Email Objects Collection...' -ForegroundColor Green
    Add-EmailObjects -EmailObjects $allM365EmailObjects -EmailIndex $allM365EmailIndex -Users $allExchangeRecipients
    Add-EmailObjects -EmailObjects $allM365EmailObjects -EmailIndex $allM365EmailIndex -Users $softDeleted

    Add-EmailObjects -EmailObjects $allM365EmailObjects -EmailIndex $allM365EmailIndex -Users $m365UPNUsers
    Add-EmailObjects -EmailObjects $allM365EmailObjects -EmailIndex $allM365EmailIndex -Users $m365Emails
    Add-EmailObjects -EmailObjects $allM365EmailObjects -EmailIndex $allM365EmailIndex -Users $m365AlternateEmails

    Add-EmailObjects -EmailObjects $allM365EmailObjects -EmailIndex $allM365EmailIndex -Users $entraIDDeletedUsersUPN
    Add-EmailObjects -EmailObjects $allM365EmailObjects -EmailIndex $allM365EmailIndex -Users $entraIDDeletedUsersEmails

    if ($SearchEmail) {
        foreach ($email in $SearchEmail) {
            $foundEmail = $allM365EmailIndex[$email.ToLowerInvariant()]

            if ($foundEmail) {
                Write-Host "$email found:" -ForegroundColor Green
                $foundEmail | Format-Table -AutoSize
            }
            else {
                Write-Host "$email not found" -ForegroundColor Red
            }
        }
    }
    else {
        return $allM365EmailObjects
    }
}