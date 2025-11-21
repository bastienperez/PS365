function Find-M365Email {
    [CmdletBinding()]
    param(
        [string[]]$SearchEmails,
        [string[]]$SearchByDomain
    )

    function Add-EmailObjects {
        param
        (
            $EmailObjects,
            $Users
        )
        foreach ($user in $users) {
            foreach ($emailaddress in $user.emailaddresses) {
                #Write-Host 'Processing' $emailaddress -ForegroundColor green
                $emailaddress = $emailaddress -replace 'X500:', ''
                $emailaddress = $emailaddress -replace 'smtp:', ''
                $emailaddress = $emailaddress -replace 'sip:', ''
                $emailaddress = $emailaddress -replace 'spo:', ''

                if ($SearchByDomain) {
                    if ($emailaddress -notlike "*$SearchByDomain") {
                        continue
                    }
                }
				
                # Check if email already exists in our objects array
                $existingEmail = $EmailObjects | Where-Object { $_.EmailAddress -eq $emailaddress }
                
                if (-not $existingEmail) {
                    # Create new email object
                    $emailObject = [PSCustomObject]@{
                        EmailAddress  = $emailaddress
                        ObjectID      = $user.objectID
                        DisplayName   = $user.DisplayName
                        RecipientType = $user.RecipientTypeDetails
                        WhenCreated   = $user.WhenCreated
                        Sources       = @($user.RecipientTypeDetails)
                    }
                    $EmailObjects.Add($emailObject)
                }
                else {
                    # Add additional source/type if not already present
                    if ($existingEmail.Sources -notcontains $user.RecipientTypeDetails) {
                        $existingEmail.Sources += $user.RecipientTypeDetails
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
            Import-Module $modules -ErrorAction stop
        }
        catch {
            Write-Warning "First, install the Microsoft $modules module first : Install-Module $modules"
            return
        }
    }

    try {
        #Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome
    }
    catch {
        Write-Warning "Failed to connect Microsoft Graph. $($_.Exception.Message)"
        return
    }
    
    try {
        #WarningAction = silentlycontinue because of warning message when resultsize is bigger than 10
        $null = Get-Recipient -ResultSize 1 -ErrorAction Stop -WarningAction silentlycontinue
    }
    catch {
        Write-Host 'Connect Exchange Online' -ForegroundColor Green
        try {
            #Connect-ExchangeOnline -ErrorAction Stop -ShowBanner:$false
        }
        catch {
            Write-Warning "Failed to connect Exchange Online. $($_.Exception.Message)"
            return
        }
    }

    [System.Collections.Generic.List[PSCustomObject]]$allM365EmailObjects = @()

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
    Add-EmailObjects -EmailObjects $allM365EmailObjects -Users $allExchangeRecipients
    Add-EmailObjects -EmailObjects $allM365EmailObjects -Users $softDeleted

    Add-EmailObjects -EmailObjects $allM365EmailObjects -Users $m365UPNUsers 
    Add-EmailObjects -EmailObjects $allM365EmailObjects -Users $m365Emails
    Add-EmailObjects -EmailObjects $allM365EmailObjects -Users $m365AlternateEmails

    Add-EmailObjects -EmailObjects $allM365EmailObjects -Users $entraIDDeletedUsersUPN 
    Add-EmailObjects -EmailObjects $allM365EmailObjects -Users $entraIDDeletedUsersEmails

    if ($SearchEmails) {
        foreach ($SearchEmail in $SearchEmails) {
            $foundEmail = $allM365EmailObjects | Where-Object { $_.EmailAddress -eq $SearchEmail }
            
            if ($foundEmail) {
                Write-Host "$SearchEmail found:" -ForegroundColor Green
                $foundEmail | Format-Table -AutoSize
            }
            else {
                Write-Host "$SearchEmail not found" -ForegroundColor Red
            }
        }
    }
    else {
        return $allM365EmailObjects
    }
}