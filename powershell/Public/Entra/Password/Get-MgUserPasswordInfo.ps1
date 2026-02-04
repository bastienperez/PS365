<#
    .SYNOPSIS
    Retrieves and processes user password information from Microsoft Graph and get information about the user's password, such as the last password change date, on-premises sync status, and password policies.

    .DESCRIPTION
    The Get-MgUserPasswordInfo script collects details such as the user's principal name, last password change date, on-premises sync status, and password policies.

    .PARAMETER UserPrincipalName
    Specifies the user principal name(s) of the user(s) for which to retrieve password information.

    .PARAMETER OnlyDomainPasswordPolicies
    If specified, retrieves password policies for domains only, without retrieving individual user information.

    .PARAMETER OnlySyncedUsers
    If specified, retrieves password information for synchronized users only (OnPremisesSyncEnabled = $true).

    .PARAMETER FilterByDomain
    Specifies a domain name to filter users. Only users from the specified domain will be retrieved (excluding guest users).

    .PARAMETER IncludeExchangeDetails
    Include Exchange Online mailbox details in the output, useful to exclude shared mailboxes and others.

    .PARAMETER SimulatedMaxPasswordAgeDays
    An optional parameter to simulate password expiration based on a specified maximum password age in days.
    If provided, the function will calculate a simulated password expiration date and indicate whether the password would be expired based on this simulated age.

    .PARAMETER OnlyUsersWithForceChangePasswordNextSignIn
    If specified, retrieves password information for users who have ForceChangePasswordNextSignIn set to true only.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.
    
    .EXAMPLE
    Get-MgUserPasswordInfo

    Retrieves password information for all users and outputs it (default behavior).

    .EXAMPLE
    Get-MgUserPasswordInfo -UserPrincipalName xxx@domain.com

    Retrieves password information for the specified user and outputs it.

    .EXAMPLE
    Get-MgUserPasswordInfo -OnlyDomainPasswordPolicies

    Retrieves password policies for all domains only.

    .EXAMPLE
    Get-MgUserPasswordInfo -FilterByDomain "contoso.com"

    Retrieves password information for users in the contoso.com domain only.

    .EXAMPLE
    Get-MgUserPasswordInfo -SimulatedMaxPasswordAgeDays 180

    Retrieves password information for all users and simulates what would happen with a 180-day password expiration policy, showing both current and simulated expiration dates.
    
    .NOTES
    Ensure you have the necessary permissions and modules installed to run this script, such as the Microsoft Graph PowerShell module.
    The script assumes that the necessary authentication to Microsoft Graph has already been handled with the Connect-MgGraph function.
    Connect-MgGraph -Scopes 'User.Read.All', 'Domain.Read.All', 'OnPremDirectorySynchronization.Read.All'

    Password policies for *cloud-only* users:
    IF `PasswordPolicies` is 'DisablePasswordExpiration':
        THEN password never expires
    ELSEIF `PasswordPolicies` is 'None' or $null:
        IF domain's `PasswordValidityPeriodInDays` is 2147483647 or $null:
            THEN password never expires
        ELSE:
            password expires based on the domain's `PasswordValidityPeriodInDays`
    ELSE:
            IF domain's `PasswordValidityPeriodInDays` is 2147483647 or $null
                THEN password never expires
            ELSE
                password expires based on the domain's `PasswordValidityPeriodInDays`

    Password policies for *synchronized* users:
    IF `CloudPasswordPolicyForPasswordSyncedUsersEnabled` is enabled:
        IF `PasswordPolicies` is 'None' or $null:
            THEN password expires based on the domain's `PasswordValidityPeriodInDays` (same as cloud-only users above)
        ELSEIF `PasswordPolicies` is 'DisablePasswordExpiration':
            THEN password never expires
        ELSE:
            THEN password expires based on the domain's `PasswordValidityPeriodInDays` (same as cloud-only users above)
    ELSE (CloudPasswordPolicyForPasswordSyncedUsersEnabled is disabled):
        THEN password never expires

    Side note : When we manually want to set Password Policies to follow domain policies, we need to set PasswordPolicies 'None' via Microsoft Graph API because $null is not accepted.
    
    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgUserPasswordInfo
#>

function Get-MgUserPasswordInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$UserPrincipalName,

        [Parameter(Mandatory = $false)]
        [switch]$OnlyDomainPasswordPolicies,

        [Parameter(Mandatory = $false)]
        [switch]$OnlySyncedUsers,

        [Parameter(Mandatory = $false)]
        [string]$FilterByDomain,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeGuestUsers,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeExchangeDetails,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$SimulatedMaxPasswordAgeDays,

        [Parameter(Mandatory = $false)]
        [switch]$OnlyUsersWithForceChangePasswordNextSignIn
    )
    
    # Import required modules
    $modules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )
    
    foreach ($module in $modules) {
        try {
            $null = Import-Module $module -ErrorAction Stop
        }
        catch {
            Write-Warning "Please install $module first"
            return
        }							   
    }

    function Get-DomainPasswordPolicies {
        Write-Host -ForegroundColor Cyan 'Retrieving password policies for all domains'
        $domains = Get-MgDomain -All
        $domainPasswordPolicies = [System.Collections.Generic.List[PSCustomObject]]$domainPasswordPolicies = @()

        foreach ($domain in $domains) {
	
            if ($domain.PasswordValidityPeriodInDays -eq '2147483647') { 
                $validityPeriod = '2147483647 (Password never expire)'
            }
            elseif ($null -eq $domain.PasswordValidityPeriodInDays) {
                $validityPeriod = '2147483647 (Password never expire - currently null in the domain settings)'
            }
            else { 
                $validityPeriod = $domain.PasswordValidityPeriodInDays 
            }
            
            $object = [PSCustomObject][ordered]@{
                DomainName                       = $domain.ID
                AuthenticationType               = $domain.AuthenticationType
                PasswordValidityPeriodInDays     = $validityPeriod
                PasswordValidityInheritedFrom    = $null
                PasswordNotificationWindowInDays = $domain.PasswordNotificationWindowInDays
            }

            $domainPasswordPolicies.Add($object)
        }		   

        # Inherit password policies
        foreach ($domain in $domainPasswordPolicies) {
            $found = $false
            
            foreach ($policy in $domainPasswordPolicies) {
                if ($domain.DomainName.EndsWith($policy.DomainName) -and $domain.DomainName -ne $policy.DomainName -and -not $found) {
                    $domain.PasswordNotificationWindowInDays = $policy.PasswordNotificationWindowInDays
                    $domain.PasswordValidityPeriodInDays = $policy.PasswordValidityPeriodInDays
                    $domain.PasswordValidityInheritedFrom = "$($policy.DomainName) domain"

                    $found = $true
                }
            }
        }
        return $domainPasswordPolicies
    }

    if (-not (Get-MgContext)) {
        Write-Host -ForegroundColor Cyan 'Connecting to Microsoft Graph'
        Connect-MgGraph -Scopes 'User.Read.All', 'Domain.Read.All', 'OnPremDirectorySynchronization.Read.All' -NoWelcome
    }

    # Get tenant-level password policy for synced users
    Write-Host -ForegroundColor Cyan 'Retrieving tenant password policy settings for synchronized users'
    $onPremSyncResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/directory/onPremisesSynchronization' -OutputType PSObject
    $tenantEnforceCloudPasswordPolicyForPasswordSyncedUsers = $onPremSyncResponse.value[0].features.cloudPasswordPolicyForPasswordSyncedUsersEnabled

    if ($IncludeExchangeDetails) {
        Write-Host -ForegroundColor Cyan 'Connecting to Exchange Online'
        # Beginning in Exchange Online PowerShell module version 3.7.0, Microsoft is implementing Web Account Manager (WAM) as the default authentication broker for user authentication. 
        if ((Get-Command -Name Connect-ExchangeOnline).Version -ge [version]'3.7.0') {
            Connect-ExchangeOnline -ShowBanner:$false -DisableWAM
        }
        else {
            Connect-ExchangeOnline -ShowBanner:$false
        }

        Write-Host -ForegroundColor Cyan 'Getting all Exchange mailboxes'
        $mailboxesHashTable = @{}

        Get-EXOMailbox -ResultSize Unlimited | ForEach-Object {
            # ID = UserRecipientTypeDetails
            $mailboxesHashTable.Add($_.ExternalDirectoryObjectId, $_.RecipientTypeDetails)
        }        
    }
    
    # Retrieve domain password policies
    $domainPasswordPolicies = Get-DomainPasswordPolicies

    if ($OnlyDomainPasswordPolicies) {
        Write-Host -ForegroundColor Cyan "Note that if you have some federated domains, they don't have password policies because authentication is handled by another IDP (Identity Provider)"

        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $ExcelFilePath = "$($env:userprofile)\$now-MgDomainPasswordPolicies_Report.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting domain password policies to Excel file: $ExcelFilePath"
            $domainPasswordPolicies | Export-Excel -Path $ExcelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-DomainPasswordPolicies'

            return
        }
        else {
            return $domainPasswordPolicies
        }
    }

    $userParams = 'UserPrincipalName, LastPasswordChangeDateTime, OnPremisesLastSyncDateTime, OnPremisesSyncEnabled, PasswordProfile, PasswordPolicies, AccountEnabled, DisplayName, Id, SignInActivity, CreatedDateTime, OnPremisesDistinguishedName'

    if ($UserPrincipalName) {
        Write-Host -ForegroundColor Cyan "Retrieving password information for $($UserPrincipalName.Count) user(s)"
        [System.Collections.Generic.List[PSCustomObject]]$mgUsersList = @()
        foreach ($upn in $UserPrincipalName) {		
            # If we use Get-MgUser -UserId <upn> -Property <properties>, we get the error "Get-MgUser_Get: Get By Key only supports UserId and the key has to be a valid Guid".
            # It seems to be a problem with one of the propertys we are requesting.
            # So we use a filter instead.										 
            $mgUser = Get-MgUser -Filter "userPrincipalName eq '$upn'" -Property $userParams

            $mgUsersList.Add($mgUser)
        }
    }
    elseif ($FilterByDomain) {
        Write-Host -ForegroundColor Cyan "Retrieving password information for users in domain: $FilterByDomain (excluding guest users #EXT#@$FilterByDomain)"

        $mgUsersList = Get-MgUser -Filter "endswith(userPrincipalName,'$FilterByDomain') and not endswith(userPrincipalName,'#EXT#@$FilterByDomain')" -All -ConsistencyLevel eventual

    }
    else {
        Write-Host -ForegroundColor Cyan 'Retrieving password information for all users'
        $mgUsersList = Get-MgUser -All -Property $userParams
    }

    [System.Collections.Generic.List[PSCustomObject]]$passwordsInfoArray = @()

    if (-not $IncludeGuestUsers) {
        $mgUsersList = $mgUsersList | Where-Object { $_.UserPrincipalName -notmatch '#EXT#' }
    }

    if ($OnlySyncedUsers) {
        Write-Host -ForegroundColor Cyan 'Filtering synchronized users only (OnPremisesSyncEnabled = true)'
        $mgUsersList = $mgUsersList | Where-Object { $_.OnPremisesSyncEnabled}
    }

    foreach ($mgUser in $mgUsersList) {
        $userDomain = $mgUser.UserPrincipalName.Split('@')[1]
        $originalDomainPolicy = $domainPasswordPolicies | Where-Object { $_.DomainName -eq $userDomain }
        
        # Create a copy of the domain policy to avoid modifying the original object
        $userDomainPolicy = [PSCustomObject][ordered]@{
            DomainName                       = $originalDomainPolicy.DomainName
            AuthenticationType               = $originalDomainPolicy.AuthenticationType
            PasswordValidityPeriodInDays     = $originalDomainPolicy.PasswordValidityPeriodInDays
            PasswordValidityInheritedFrom    = $originalDomainPolicy.PasswordValidityInheritedFrom
            PasswordNotificationWindowInDays = $originalDomainPolicy.PasswordNotificationWindowInDays
        }

        $passwordExpired = $false 

        # Determine password policy inheritance
        if ($mgUser.OnPremisesSyncEnabled -and -not $tenantEnforceCloudPasswordPolicyForPasswordSyncedUsers) {
            # Synchronized user with cloud password policy NOT enforced
            $userDomainPolicy.PasswordValidityPeriodInDays = '2147483647 (Password never expire)'
            $userDomainPolicy.PasswordValidityInheritedFrom = 'Synchronized user - Cloud password policy NOT enforced (EnforceCloudPasswordPolicyForPasswordSyncedUsers is disabled)'
        }
        elseif ($mgUser.PasswordPolicies -eq 'DisablePasswordExpiration') {
            # User has DisablePasswordExpiration set at user level
            $userDomainPolicy.PasswordValidityPeriodInDays = '2147483647 (Password never expire)'
            $userDomainPolicy.PasswordValidityInheritedFrom = 'User password policy (DisablePasswordExpiration)'
        }
        elseif ($mgUser.PasswordPolicies -eq 'None' -or [string]::IsNullOrEmpty($mgUser.PasswordPolicies)) {
            # User follows domain policy (PasswordPolicies is 'None' or $null)
            if ([string]::IsNullOrEmpty($userDomainPolicy.PasswordValidityInheritedFrom)) {
                $userDomainPolicy.PasswordValidityInheritedFrom = 'Domain password policy'
            }
        }

        if ($userDomainPolicy.PasswordValidityPeriodInDays -ne '2147483647 (Password never expire)' -and $mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {

            if ($mgUser.LastPasswordChangeDateTime -lt (Get-Date).AddDays(-$userDomainPolicy.PasswordValidityPeriodInDays)) { 
                $passwordExpired = $true 
            }
        }

        if ($userDomainPolicy.PasswordValidityPeriodInDays -ne '2147483647 (Password never expire)' -and $mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
            $daysLeft = ($mgUser.LastPasswordChangeDateTime.AddDays($userDomainPolicy.PasswordValidityPeriodInDays) - (Get-Date)).Days
            if ($daysLeft -lt 0) {
                $daysLeft = 'Already expired'
            }
            
            $date = $mgUser.LastPasswordChangeDateTime.AddDays($userDomainPolicy.PasswordValidityPeriodInDays)
            $passwordExpirationDateUTC = $date.ToString('dd/MM/yyyy HH:mm:ss')
        }
        elseif ($userDomainPolicy.PasswordValidityPeriodInDays -ne '2147483647 (Password never expire)' -and (-not $mgUser.LastPasswordChangeDateTime -or $mgUser.LastPasswordChangeDateTime -eq [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc))) {
            $daysLeft = 'No password set date available'
            $passwordExpirationDateUTC = $null
        }
        else {
            $daysLeft = 'Password never expires'
            $passwordExpirationDateUTC = $null
        }


        if ($SimulatedMaxPasswordAgeDays -and $mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
            # Calculate simulated password expiration if SimulatedMaxPasswordAgeDays is provided
            $simulatedPasswordExpirationDateUTC = $null
            $simulatedPasswordExpired = $false
            $simulatedPasswordExpirationDateUTC = $mgUser.LastPasswordChangeDateTime.AddDays($SimulatedMaxPasswordAgeDays)
            if ($mgUser.LastPasswordChangeDateTime -lt (Get-Date).AddDays(-$SimulatedMaxPasswordAgeDays)) {
                $simulatedPasswordExpired = $true
            }
        }

        $object = [PSCustomObject][ordered]@{
            UserPrincipalName                                      = $mgUser.UserPrincipalName
            DisplayName                                            = $mgUser.DisplayName
            ID                                                     = $mgUser.Id
            AccountEnabled                                         = $mgUser.AccountEnabled
            PasswordPolicies                                       = $mgUser.PasswordPolicies
            PasswordLastSetUTCTime                                 = if ($mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) { $mgUser.LastPasswordChangeDateTime } else { $null }
            PasswordPolicyMaxPasswordAgeInDays                     = $userDomainPolicy.PasswordValidityPeriodInDays
            PasswordExpirationDateUTC                              = $passwordExpirationDateUTC
            DaysUntilPasswordExpiration                            = $daysLeft
            PasswordExpired                                        = $passwordExpired
            #The last interactive sign-in date and time for a specific user. This property records the last time a user attempted an interactive sign-in to the directory—whether the attempt was successful or not. Note: Since unsuccessful attempts are also logged, this value might not accurately reflect actual system usage.
            LastInteractiveSignInDateTime                          = $mgUser.signInActivity.LastSignInDateTime
            # The date and time of the user's most recent successful interactive or non-interactive sign-in
            LastSuccessfulSignInDateTime                           = $mgUser.signInActivity.LastSuccessfulSignInDateTime
            ForceChangePasswordNextSignIn                          = if ($mgUser.PasswordProfile) { $mgUser.PasswordProfile.ForceChangePasswordNextSignIn } else { $null }
            ForceChangePasswordNextSignInWithMfa                   = if ($mgUser.PasswordProfile) { $mgUser.PasswordProfile.ForceChangePasswordNextSignInWithMfa } else { $null }
            OnPremisesSyncEnabled                                  = if($null -eq $mgUser.OnPremisesSyncEnabled) { $false } else { $mgUser.OnPremisesSyncEnabled }
            OnPremisesLastSyncDateTimeUTC                          = if ($mgUser.OnPremisesLastSyncDateTime -and $mgUser.OnPremisesLastSyncDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) { $mgUser.OnPremisesLastSyncDateTime } else { $null }
            OnPremisesDistinguishedName                            = if ($mgUser.OnPremisesDistinguishedName) { $mgUser.OnPremisesDistinguishedName } else { $null }        
            Domain                                                 = $userDomain
            DomainAuthenticationType                               = $userDomainPolicy.AuthenticationType
            PasswordValidityInheritedFrom                          = $userDomainPolicy.PasswordValidityInheritedFrom
            PasswordNotificationWindowInDays                       = $userDomainPolicy.PasswordNotificationWindowInDays
            TenantEnforceCloudPasswordPolicyForPasswordSyncedUsers = $tenantEnforceCloudPasswordPolicyForPasswordSyncedUsers
            CreatedDateTime                                        = $mgUser.CreatedDateTime
        }

        # Add simulation columns only when parameter is provided, positioned after PasswordExpired
        if ($SimulatedMaxPasswordAgeDays) {
            # Create ordered hashtable with simulation columns inserted at correct position
            $orderedProperties = [ordered]@{}
            foreach ($prop in $object.PSObject.Properties) {
                $orderedProperties[$prop.Name] = $prop.Value
                if ($prop.Name -eq 'PasswordExpired') {
                    $orderedProperties['SimulatedPasswordExpirationDateUTC'] = $simulatedPasswordExpirationDateUTC
                    $orderedProperties['SimulatedPasswordExpired'] = $simulatedPasswordExpired
                }
            }
            $object = [PSCustomObject]$orderedProperties
        }
    
        $passwordsInfoArray.Add($object)
    }

    if ($OnlyUsersWithForceChangePasswordNextSignIn) {
        Write-Host -ForegroundColor Cyan 'Filtering users with ForceChangePasswordNextSignIn = true'
        $passwordsInfoArray = $passwordsInfoArray | Where-Object { $_.ForceChangePasswordNextSignIn }
    }

    if ($IncludeExchangeDetails) {
        Write-Host -ForegroundColor Cyan 'Adding Exchange details'
        foreach ($user in $passwordsInfoArray) {
            
            $recipientTypeDetails = $mailboxesHashTable[$user.Id]

            if ($null -eq $recipientTypeDetails) {
                $recipientTypeDetails = 'No Mailbox'
            }

            $user | Add-Member -MemberType NoteProperty -Name 'RecipientTypeDetails' -Value $recipientTypeDetails
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $ExcelFilePath = "$($env:userprofile)\$now-MgUserPasswordInfo_Report.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting password information to Excel file: $ExcelFilePath"
        $passwordsInfoArray | Export-Excel -Path $ExcelFilePath -AutoSize -AutoFilter -WorksheetName 'Entra-PasswordInfo'
    }
    else {
        return $passwordsInfoArray
    }
}