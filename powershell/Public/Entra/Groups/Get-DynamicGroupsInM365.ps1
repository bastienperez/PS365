<#
    .SYNOPSIS
    Retrieves dynamic groups from Microsoft 365, Exchange Online, and Entra ID with attribute analysis.

    .DESCRIPTION
    The Get-DynamicGroupsInM365 function retrieves all dynamic groups from Microsoft 365 environments,
    including Exchange Online Dynamic Distribution Groups and Entra ID Dynamic Security/M365 Groups.
    It analyzes membership rules to extract attributes and provides security warnings for attributes
    in the "Personal-Information" property set that users can modify themselves.

    .PARAMETER ExchangeOnlineOnly
    When specified, retrieves only Exchange Online Dynamic Distribution Groups.
    Requires an active Exchange Online session (use Connect-ExchangeOnline).

    .PARAMETER EntraIDOnly
    When specified, retrieves only Entra ID Dynamic Groups (Security and M365 groups).
    Requires an active Microsoft Graph connection (use Connect-MgGraph).

    .EXAMPLE
    PS C:\> Get-DynamicGroupsInM365
    
    Retrieves all dynamic groups from both Exchange Online and Entra ID.

    .EXAMPLE
    PS C:\> Get-DynamicGroupsInM365 -ExchangeOnlineOnly
    
    Retrieves only Exchange Online Dynamic Distribution Groups.

    .EXAMPLE
    PS C:\> Get-DynamicGroupsInM365 -EntraIDOnly
    
    Retrieves only Entra ID Dynamic Groups.

    .OUTPUTS
    System.Collections.Generic.List[Object]

    .NOTES
    OUTPUT PROPERTIES
    Returns a collection of custom objects with the following properties:
    - Name: Display name of the group
    - Type: Type of dynamic group (Exchange Dynamic Distribution Group, M365 Dynamic Group, Entra ID Dynamic Security Group)
    - Email: Primary email address of the group
    - Filter: The membership rule or LDAP filter used for dynamic membership
    - UserAttributes: Pipe-separated list of user attributes referenced in the membership rule
    - GroupAttributes: Pipe-separated list of group attributes referenced in the membership rule (Entra ID only)
    - DeviceAttributes: Pipe-separated list of device attributes referenced in the membership rule (Entra ID only)
    - Warning: Security warning if any attribute is in the "Personal-Information" property set

    Security Considerations:
    This function identifies attributes in the "Personal-Information" property set that users can
    modify themselves, potentially allowing unauthorized group membership. Review warnings carefully.

    More information on: https://itpro-tips.com/property-set-personal-information-and-active-directory-security-and-governance/

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-DynamicGroupsInM365
#>

<#
Active Directory to Entra ID Attribute Mapping Reference:
=========================================================
Attribut AD                           Attribut Entra ID
assistant                             assistant
c                                     country
facsimileTelephoneNumber             faxNumber
homePhone                            homePhone
homePostalAddress                    homePostalAddress
info                                 notes
ipPhone                              ipPhone
l                                    city
mobile                               mobile
otherFacsimileTelephoneNumber        otherFaxNumbers
otherHomePhone                       otherHomePhones
otherIpPhone                         otherIpPhones
otherMobile                          otherMobiles
otherPager                           otherPagers
otherTelephone                       otherPhones
pager                                pager
personalTitle                        personalTitle
physicalDeliveryOfficeName           officeLocation
postalAddress                        streetAddress
postalCode                           postalCode
postOfficeBox                        postOfficeBox
st                                   state
street                               streetAddress
streetAddress                        streetAddress
telephoneNumber                      telephoneNumber
thumbnailPhoto                       thumbnailPhoto
userCertificate                      userCertificate
msDS-cloudExtensionAttribute1        extensionAttribute1
msDS-cloudExtensionAttribute2        extensionAttribute2
msDS-cloudExtensionAttribute3        extensionAttribute3
msDS-cloudExtensionAttribute4        extensionAttribute4
msDS-cloudExtensionAttribute5        extensionAttribute5
msDS-cloudExtensionAttribute6        extensionAttribute6
msDS-cloudExtensionAttribute7        extensionAttribute7
msDS-cloudExtensionAttribute8        extensionAttribute8
msDS-cloudExtensionAttribute9        extensionAttribute9
msDS-cloudExtensionAttribute10       extensionAttribute10
msDS-cloudExtensionAttribute11       extensionAttribute11
msDS-cloudExtensionAttribute12       extensionAttribute12
msDS-cloudExtensionAttribute13       extensionAttribute13
msDS-cloudExtensionAttribute14       extensionAttribute14
msDS-cloudExtensionAttribute15       extensionAttribute15
msDS-ExternalDirectoryObjectId       externalDirectoryObjectId
#>

function Get-DynamicGroupsInM365 {
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ExchangeOnlineOnly,

        [Parameter(Mandatory = $false)]
        [switch]$EntraIDOnly,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    $propertySetsAttribute = @(
        'assistant'
        'country'
        'faxNumber'
        'homePhone'
        'homePostalAddress'
        'notes'
        'ipPhone'
        'city'
        'mobile'
        'MobilePhone'
        'otherFaxNumbers'
        'otherHomePhones'
        'otherIpPhones'
        'otherMobiles'
        'otherPagers'
        'otherPhones'
        'pager'
        'personalTitle'
        'officeLocation'
        'streetAddress'
        'postalCode'
        'postOfficeBox'
        'state'
        'streetAddress'
        'streetAddress'
        'telephoneNumber'
        'thumbnailPhoto'
        'userCertificate'
    )

    # Initialize an array list for better performance
    [System.Collections.Generic.List[Object]]$dynGroupArray = @()

    # Check Exchange Online for Dynamic Distribution Groups
    if (-not $EntraIDOnly) {
        # Retrieve dynamic distribution groups
        try {
            $exchangeGroups = Get-DynamicDistributionGroup -ErrorAction Stop
        }
        catch {
            Write-Error "Error retrieving Exchange Online groups: $($_.Exception.Message)"
            $object = [PSCustomObject][ordered]@{
                GroupId                       = 'Error'
                Name                          = 'Error'
                Type                          = 'Exchange Dynamic Distribution Group'
                MembershipRule                = 'Error retrieving groups. Ensure you are connected to Exchange Online.'
                MembershipRuleProcessingState = 'Error'
                UserAttributes                = 'N/A'
                GroupAttributes               = 'N/A'
                DeviceAttributes              = 'N/A'
                MemberOf                      = 'N/A'
                DisplayName                   = 'Error'
                Description                   = 'Error'
                Mail                          = 'Error'
                MailEnabled                   = 'Error'
                MailNickname                  = 'Error'
                SecurityEnabled               = 'Error'
                GroupTypes                    = 'Error'
                CreatedDateTime               = 'Error'
                RenewedDateTime               = 'Error'
                OnPremisesSyncEnabled         = 'Error'
                SecurityIdentifier            = 'Error'
                Classification                = 'Error'
                Visibility                    = 'Error'
            }

            $dynGroupArray.Add($object)
        }

        foreach ($group in $exchangeGroups) {
            try {
                $memberOf = Get-DistributionGroup -Identity $group.Identity | Select-Object -ExpandProperty MemberOfGroup -ErrorAction SilentlyContinue
            }
            catch {
                $memberOf = $null
            }

            $object = [PSCustomObject][ordered]@{
                GroupId                       = $group.Guid
                Name                          = $group.Name
                Type                          = 'Exchange Dynamic Distribution Group'
                MembershipRule                = (($group.LdapRecipientFilter.Replace("`r`n", '').Replace("`n", '').Replace("`r", '') -replace '\s+', ' ') -replace '(\))(\w)', '$1 $2' -replace '(\w)(\()', '$1 $2' -replace '\(\s+', '(').Trim()
                MembershipRuleProcessingState = $null
                UserAttributes                = (($group.LdapRecipientFilter | 
                        Select-String -Pattern '\(([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | Select-Object -ExpandProperty Value -Unique).Replace('(', '') -join ' | '
                GroupAttributes               = $null
                DeviceAttributes              = $null
                MemberOf                      = $memberOf -join '|'
                DisplayName                   = $group.DisplayName
                Description                   = $group.Notes
                Mail                          = $group.PrimarySmtpAddress
                MailEnabled                   = $true
                MailNickname                  = $group.Alias
                SecurityEnabled               = $false
                GroupTypes                    = 'DynamicDistribution'
                CreatedDateTime               = $group.WhenCreated
                RenewedDateTime               = $null
                OnPremisesSyncEnabled         = $group.IsDirSynced
                SecurityIdentifier            = $null
                Classification                = $null
                Visibility                    = if ($group.HiddenFromAddressListsEnabled) { 'Private' } else { 'Public' }
            }
            
            $dynGroupArray.Add($object)
        }
    }

    # Check Entra ID for Dynamic Groups
    if (-not $ExchangeOnlineOnly) {
        try {
            $entraGroups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'DynamicMembership')" -ErrorAction Stop
        }
        catch {
            Write-Error "Error retrieving Entra ID groups: $($_.Exception.Message)"

            $object = [PSCustomObject][ordered]@{
                GroupId                       = 'Error'
                Name                          = 'Error'
                Type                          = 'Entra ID Dynamic Group'
                MembershipRule                = 'Error retrieving groups. Ensure you are connected to Microsoft Graph.'
                MembershipRuleProcessingState = 'Error'
                UserAttributes                = 'N/A'
                GroupAttributes               = 'N/A'
                DeviceAttributes              = 'N/A'
                MemberOf                      = 'N/A'
                DisplayName                   = 'Error'
                Description                   = 'Error'
                Mail                          = 'Error'
                MailEnabled                   = 'Error'
                MailNickname                  = 'Error'
                SecurityEnabled               = 'Error'
                GroupTypes                    = 'Error'
                CreatedDateTime               = 'Error'
                RenewedDateTime               = 'Error'
                OnPremisesSyncEnabled         = 'Error'
                SecurityIdentifier            = 'Error'
                Classification                = 'Error'
                Visibility                    = 'Error'
            }

            $dynGroupArray.Add($object)
        }

        foreach ($group in $entraGroups) {
            try { 
                $memberOf = Get-MgGroupMemberOf -GroupId $group.Id -ErrorAction SilentlyContinue

                $memberOf = ($memberOf | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } | 
                ForEach-Object { $_.AdditionalProperties.displayName }) -join '|'
            }
            catch { 
                $memberOf = $null 
            }

            $object = [PSCustomObject][ordered]@{
                GroupId                       = $group.Id
                Name                          = $group.DisplayName
                Type                          = if ($group.groupTypes -contains 'Unified') { 'M365 Dynamic Group' } else { 'Entra ID Dynamic Security Group' }
                # MembershipRule can be multiline, so we need to clean it up
                # Remove line breaks and extra spaces
                MembershipRule                = (($group.MembershipRule.Replace("`r`n", '').Replace("`n", '').Replace("`r", '') -replace '\s+', ' ') -replace '(\))(\w)', '$1 $2' -replace '(\w)(\()', '$1 $2' -replace '\(\s+', '(').Trim()
                MembershipRuleProcessingState = $group.MembershipRuleProcessingState
                UserAttributes                = (($group.MembershipRule | 
                        Select-String -Pattern 'user\.([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | 
                    Select-Object -ExpandProperty Value | 
                    ForEach-Object { $_.Replace('user.', '') } | 
                    Sort-Object -Unique) -join '| '
                GroupAttributes               = (($group.MembershipRule | 
                        Select-String -Pattern 'group\.([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | 
                    Select-Object -ExpandProperty Value | 
                    ForEach-Object { $_.Replace('group.', '') } | 
                    Sort-Object -Unique) -join '| '
                DeviceAttributes              = (($group.MembershipRule | 
                        Select-String -Pattern 'device\.([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | 
                    Select-Object -ExpandProperty Value | 
                    ForEach-Object { $_.Replace('device.', '') } | 
                    Sort-Object -Unique) -join '| '
                MemberOf                      = $memberOf -join '|'
                DisplayName = $group.DisplayName
                Description                   = $group.Description
                Mail                          = $group.Mail
                MailEnabled                   = $group.MailEnabled
                MailNickname                  = $group.MailNickname
                SecurityEnabled               = $group.SecurityEnabled
                GroupTypes                    = ($group.GroupTypes -join ', ')
                CreatedDateTime               = $group.CreatedDateTime
                RenewedDateTime               = $group.RenewedDateTime
                OnPremisesSyncEnabled         = $group.OnPremisesSyncEnabled
                SecurityIdentifier            = $group.SecurityIdentifier
                Classification                = $group.Classification
                Visibility                    = $group.Visibility
            }

            $dynGroupArray.Add($object)
        }
    }

    # foreach attribute, check if it's in the `Personal-Information` property set attribute
    # if it's, add a warning to the object 
    foreach ($group in $dynGroupArray) {
        $group | Add-Member -MemberType NoteProperty -Name Warning -Value $null
        foreach ($attribute in $group.UserAttributes.Split('|')) {
            if ($propertySetsAttribute -contains $attribute) {
                $group.Warning = "'$attribute' is in the `Personal-Information` property set, the user can modify it and add himself to the group. See https://itpro-tips.com/property-set-personal-information-and-active-directory-security-and-governance/"
            }
        }
    }
    
    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-DynamicGroupsInM365.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting dynamic groups to Excel file: $excelFilePath"
        $dynGroupArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'DynamicGroupsInM365'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $dynGroupArray
    }
}