<#
    .SYNOPSIS
    Retrieves and processes user password information from Microsoft Graph and get information about the user's password, such as the last password change date, on-premises sync status, and password policies.

    .DESCRIPTION
    The Get-MgUserPasswordInfo script collects details such as the user's principal name, last password change date, on-premises sync status, and password policies.

    .PARAMETER UserPrincipalName
    Specifies the user principal name(s) of the user(s) for which to retrieve password information.

    .PARAMETER PasswordPoliciesByDomainOnly
    If specified, retrieves password policies for domains only, without retrieving individual user information.

    .PARAMETER IncludeExchangeDetails
    Include Exchange Online mailbox details in the output, useful to exclude shared mailboxes and others.

    .EXAMPLE
    Get-MgUserPasswordInfo
    Retrieves password information for all users and outputs it (default behavior).

    .EXAMPLE
    Get-MgUserPasswordInfo -UserPrincipalName xxx@domain.com
    Retrieves password information for the specified user and outputs it.

    .EXAMPLE
    Get-MgUserPasswordInfo -PasswordPoliciesByDomainOnly
    Retrieves password policies for all domains only.

    .EXAMPLE
    Get-MgUserPasswordInfo -SimulatedMaxPasswordAgeDays 180

    Retrieves password information for all users and simulates what would happen with a 180-day password expiration policy, showing both current and simulated expiration dates.
    
    .NOTES
    Ensure you have the necessary permissions and modules installed to run this script, such as the Microsoft Graph PowerShell module.
    The script assumes that the necessary authentication to Microsoft Graph has already been handled with the Connect-MgGraph function.
    Connect-MgGraph -Scopes 'User.Read.All', 'Domain.Read.All'

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgUserPasswordInfo

    .CHANGELOG
    [2.1.0] - 2025-12-03
    # Added
    - Add parameter `SimulatedMaxPasswordAgeDays` to simulate password expiration with different password age policies.
    - Add property `SimulatedPasswordExpirationDateUTC` to the output object showing simulated password expiration dates.

    [2.0.0] - 2025-11-27
    # Added
    - Add properties `ID`, `LastSignInDateTime` and `LastSuccessfulSignInDateTime ` to the output object.
    - Add parameter `ExportToExcel` to export the report to an Excel file.
    - Add parameter `IncludeGuestUsers` to include guest users in the report.
    - Add parameter `IncludeExchangeDetails` to include Exchange Online mailbox details in the output.

    # Changed
    - Modified property `Enabled` to `AccountEnabled` in the output object.
#>

function Get-MgUserPasswordInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$UserPrincipalName,

        [Parameter(Mandatory = $false)]
        [switch]$PasswordPoliciesByDomainOnly,

        [Parameter(Mandatory = $false)]
        [string]$ByDomain,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeGuestUsers,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeExchangeDetails,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$SimulatedMaxPasswordAgeDays
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
	
            $validityPeriod = if ($domain.PasswordValidityPeriodInDays -eq '2147483647') { 
                '2147483647 (Password never expire)' 
            }
            else { 
                $domain.PasswordValidityPeriodInDays 
            }
            
            $object = [PSCustomObject][ordered]@{
                DomainName                       = $domain.ID
                AuthenticationType               = $domain.AuthenticationType
                PasswordValidityPeriod           = $validityPeriod
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
                    $domain.PasswordValidityPeriod = $policy.PasswordValidityPeriod
                    $domain.PasswordValidityInheritedFrom = "$($policy.DomainName) domain"

                    $found = $true
                }
            }
        }
        return $domainPasswordPolicies
    }

    if (-not (Get-MgContext)) {
        Write-Host -ForegroundColor Cyan 'Connecting to Microsoft Graph'
        Connect-MgGraph -Scopes 'User.Read.All' -NoWelcome
    }

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

    if ($PasswordPoliciesByDomainOnly) {
        Write-Host -ForegroundColor Cyan "Note that if you have some federated domains, they don't have password policies because authentication is handled by another IDP (Identity Provider)"

        return $domainPasswordPolicies
    }

    $userParams = 'UserPrincipalName, LastPasswordChangeDateTime, OnPremisesLastSyncDateTime, OnPremisesSyncEnabled, PasswordProfile, PasswordPolicies, AccountEnabled, DisplayName, Id, SignInActivity, CreatedDateTime'

    if ($UserPrincipalName) {
        Write-Host -ForegroundColor Cyan "Retrieving password information for $($UserPrincipalName.Count) user(s)"
        [System.Collections.Generic.List[PSCustomObject]]$mgUsersList = @()
        foreach ($upn in $UserPrincipalName) {												 
            $mgUser = Get-MgUser -UserId $upn -Property $userParams

            $mgUsersList.Add($mgUser)
        }
    }
    elseif ($ByDomain) {
        Write-Host -ForegroundColor Cyan "Retrieving password information for users in domain: $ByDomain (exluding guest users #EXT#@$ByDomain)"

        $mgUsersList = Get-MgUser -Filter "endswith(userPrincipalName,'$ByDomain') and not endswith(userPrincipalName,'#EXT#@$ByDomain')" -All -ConsistencyLevel eventual

    }
    else {
        Write-Host -ForegroundColor Cyan 'Retrieving password information for all users'
        $mgUsersList = Get-MgUser -All -Property $userParams
    }

    [System.Collections.Generic.List[PSCustomObject]]$passwordsInfoArray = @()

    if (-not $IncludeGuestUsers) {
        $mgUsersList = $mgUsersList | Where-Object { $_.UserPrincipalName -notmatch '#EXT#' }
    }

    foreach ($mgUser in $mgUsersList) {
        $userDomain = $mgUser.UserPrincipalName.Split('@')[1]
        $userDomainPolicy = $domainPasswordPolicies | Where-Object { $_.DomainName -eq $userDomain }

        $passwordExpired = $false 

        if ($mgUser.PasswordPolicies -eq 'DisablePasswordExpiration') {
            $userDomainPolicy.PasswordValidityPeriod = '2147483647 (Password never expire)'
            $userDomainPolicy.PasswordValidityInheritedFrom = 'User password policy'
        }

        if ($userDomainPolicy.PasswordValidityPeriod -ne '2147483647 (Password never expire)' -and $mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {

            if ($mgUser.LastPasswordChangeDateTime -lt (Get-Date).AddDays(-$userDomainPolicy.PasswordValidityPeriod)) { 
                $passwordExpired = $true 
            }
        }

        if ($userDomainPolicy.PasswordValidityPeriod -ne '2147483647 (Password never expire)' -and $mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
            $daysLeft = ($mgUser.LastPasswordChangeDateTime.AddDays($userDomainPolicy.PasswordValidityPeriod) - (Get-Date)).Days
            if ($daysLeft -lt 0) {
                $daysLeft = 'Already expired'
            }
            
            $date = $mgUser.LastPasswordChangeDateTime.AddDays($userDomainPolicy.PasswordValidityPeriod)
            $passwordExpirationDateUTC = $date.ToString('dd/MM/yyyy HH:mm:ss')
        }
        elseif ($userDomainPolicy.PasswordValidityPeriod -ne '2147483647 (Password never expire)' -and (-not $mgUser.LastPasswordChangeDateTime -or $mgUser.LastPasswordChangeDateTime -eq [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc))) {
            $daysLeft = 'No password set date available'
            $passwordExpirationDateUTC = $null
        }
        else {
            $daysLeft = 'Password never expires'
            $passwordExpirationDateUTC = $null
        }

        # Calculate simulated password expiration if SimulatedMaxPasswordAgeDays is provided
        $simulatedPasswordExpirationDateUTC = $null
        $simulatedPasswordExpired = $false
        if ($SimulatedMaxPasswordAgeDays -and $mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) {
            $simulatedPasswordExpirationDateUTC = $mgUser.LastPasswordChangeDateTime.AddDays($SimulatedMaxPasswordAgeDays)
            if ($mgUser.LastPasswordChangeDateTime -lt (Get-Date).AddDays(-$SimulatedMaxPasswordAgeDays)) {
                $simulatedPasswordExpired = $true
            }
        }

        $object = [PSCustomObject][ordered]@{
            UserPrincipalName                    = $mgUser.UserPrincipalName
            DisplayName                          = $mgUser.DisplayName
            ID                                   = $mgUser.Id
            AccountEnabled                       = $mgUser.AccountEnabled
            PasswordPolicies                     = $mgUser.PasswordPolicies
            PasswordLastSetUTCTime               = if ($mgUser.LastPasswordChangeDateTime -and $mgUser.LastPasswordChangeDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) { $mgUser.LastPasswordChangeDateTime } else { $null }
            PasswordPolicyMaxPasswordAgeInDays   = $userDomainPolicy.PasswordValidityPeriod
            PasswordExpirationDateUTC            = $passwordExpirationDateUTC
            DaysLeftBeforePasswordChangeUTC      = $daysLeft
            PasswordExpired                      = $passwordExpired
            LastSignInDateTime                   = $mgUser.signInActivity.LastSignInDateTime
            LastSuccessfulSignInDateTime         = $mgUser.signInActivity.LastSuccessfulSignInDateTime
            OnPremisesLastSyncDateTimeUTC        = if ($mgUser.OnPremisesLastSyncDateTime -and $mgUser.OnPremisesLastSyncDateTime -ne [datetime]::new(1601, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)) { $mgUser.OnPremisesLastSyncDateTime } else { $null }
            ForceChangePasswordNextSignIn        = if ($mgUser.PasswordProfile) { $mgUser.PasswordProfile.ForceChangePasswordNextSignIn } else { $null }
            ForceChangePasswordNextSignInWithMfa = if ($mgUser.PasswordProfile) { $mgUser.PasswordProfile.ForceChangePasswordNextSignInWithMfa } else { $null }
            OnPremisesSyncEnabled                = $mgUser.OnPremisesSyncEnabled
            Domain                               = $userDomain
            PasswordValidityInheritedFrom        = $userDomainPolicy.PasswordValidityInheritedFrom
            PasswordNotificationWindowInDays     = $userDomainPolicy.PasswordNotificationWindowInDays
            CreatedDateTime                      = $mgUser.CreatedDateTime
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