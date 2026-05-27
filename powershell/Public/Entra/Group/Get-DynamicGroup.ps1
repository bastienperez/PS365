<#
    .SYNOPSIS
    Retrieves dynamic groups from Microsoft 365, Exchange Online, and Entra ID with attribute analysis.

    .DESCRIPTION
    The Get-DynamicGroup function retrieves all dynamic groups from Microsoft 365 environments,
    including Exchange Online Dynamic Distribution Groups and Entra ID Dynamic Security/M365 Groups.
    It analyzes membership rules to extract attributes and provides security warnings for attributes
    in the "Personal-Information" property set that users can modify themselves.

    By default, only the member count is returned (MembersCount). Use -IncludeMembers to also get
    the concatenated lists of member names and IDs, or -MemberReport to get one row per (group, member).

    .PARAMETER GroupId
    When specified, retrieves dynamic groups by their unique GroupId.
    If not specified, retrieves all dynamic groups (both Exchange Online and Entra ID).

    .PARAMETER EntraIDOnly
    When specified, retrieves only Entra ID Dynamic Groups (Security and M365 groups).
    Requires an active Microsoft Graph connection (use Connect-MgGraph).

    .PARAMETER ExchangeOnlineOnly
    When specified, retrieves only Exchange Online Dynamic Distribution Groups.
    Requires an active Exchange Online session (use Connect-ExchangeOnline).

    .PARAMETER ExportToExcel
    When specified, exports the result to an Excel file in the user's profile directory instead of
    returning it. When combined with -MemberReport, the member rows are written to a second worksheet
    named "Members" in the same workbook.

    .PARAMETER IncludeMembers
    When specified, populates the MembersName (pipe-separated display names) and MembersId
    (pipe-separated object IDs) properties in addition to MembersCount.
    By default, only MembersCount is populated to avoid slow enumeration on large groups.
    Ignored when -MemberReport is used.

    .PARAMETER MemberReport
    When specified, the function returns one row per (group, member) pair instead of the per-group
    aggregated objects. Useful to export a flat list of members with their parent group.
    Output columns: GroupId, GroupName, GroupType, MembershipRule, MemberId, MemberDisplayName,
    MemberPrincipal, MemberType.

    .EXAMPLE
    Get-DynamicGroup

    Retrieves all dynamic groups from both Exchange Online and Entra ID (members count only).

    .EXAMPLE
    Get-DynamicGroup -ExchangeOnlineOnly

    Retrieves only Exchange Online Dynamic Distribution Groups.

    .EXAMPLE
    Get-DynamicGroup -EntraIDOnly

    Retrieves only Entra ID Dynamic Groups.

    .EXAMPLE
    Get-DynamicGroup -IncludeMembers

    Retrieves all dynamic groups with MembersName and MembersId populated (pipe-separated).

    .EXAMPLE
    Get-DynamicGroup -MemberReport

    Returns one row per (group, member). Each row contains the group context and a single member.

    .EXAMPLE
    Get-DynamicGroup -MemberReport -ExportToExcel

    Exports both the per-group sheet ("DynamicGroups") and the per-member sheet ("Members") in the
    same Excel workbook.

    .OUTPUTS
    System.Collections.Generic.List[Object]

    .NOTES
    OUTPUT PROPERTIES (default / -IncludeMembers mode)
    Returns a collection of custom objects with the following properties:
    - GroupId: Unique identifier of the group
    - Name: Display name of the group
    - Type: Type of dynamic group (Exchange Dynamic Distribution Group, M365 Dynamic Group, Entra ID Dynamic Security Group)
    - MembershipRule: The membership rule or LDAP filter used for dynamic membership
    - MembershipRuleProcessingState: Processing state of the membership rule (Entra ID only)
    - UserAttributes: Pipe-separated list of user attributes referenced in the membership rule
    - GroupAttributes: Pipe-separated list of group attributes referenced in the membership rule (Entra ID only)
    - DeviceAttributes: Pipe-separated list of device attributes referenced in the membership rule (Entra ID only)
    - MemberOf: Pipe-separated list of parent groups
    - MembersCount: Number of current members of the group
    - MembersName: Pipe-separated list of member display names (only with -IncludeMembers)
    - MembersId: Pipe-separated list of member IDs (only with -IncludeMembers)
    - DisplayName, Description, Mail, MailEnabled, MailNickname, SecurityEnabled, GroupTypes,
      CreatedDateTime, RenewedDateTime, OnPremisesSyncEnabled, SecurityIdentifier, Classification, Visibility
    - Warning: Security warning if any attribute is in the "Personal-Information" property set

    OUTPUT PROPERTIES (-MemberReport mode)
    - GroupId, GroupName, GroupType, MembershipRule
    - MemberId, MemberDisplayName, MemberPrincipal, MemberType

    Security Considerations:
    This function identifies attributes in the "Personal-Information" property set that users can
    modify themselves, potentially allowing unauthorized group membership. Review warnings carefully.

    More information on: https://itpro-tips.com/property-set-personal-information-and-active-directory-security-and-governance/


    .LINK
    https://ps365.clidsys.com/docs/commands/Get-DynamicGroup
#>

function Get-DynamicGroup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)]
        [String]$GroupId,

        [Parameter(Mandatory = $false)]
        [switch]$ExchangeOnlineOnly,

        [Parameter(Mandatory = $false)]
        [switch]$EntraIDOnly,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeMembers,

        [Parameter(Mandatory = $false)]
        [switch]$MemberReport
    )

    if ($MemberReport.IsPresent) {
        Write-Verbose 'MemberReport mode: returning one row per (group, member). MembersCount/MembersName/MembersId columns are not produced.'
        [System.Collections.Generic.List[Object]]$memberReportArray = @()
    }
    elseif (-not $IncludeMembers.IsPresent) {
        Write-Warning 'Members enumeration is disabled by default (only MembersCount is populated). Use -IncludeMembers to also retrieve MembersName and MembersId, or -MemberReport to get one row per member. This can be slow on large groups.'
    }

    
    # Active Directory to Entra ID Attribute Mapping Reference:
    # =========================================================
    # Attribut AD                           Attribut Entra ID
    # assistant                             assistant
    # c                                     country
    # facsimileTelephoneNumber             faxNumber
    # homePhone                            homePhone
    # homePostalAddress                    homePostalAddress
    # info                                 notes
    # ipPhone                              ipPhone
    # l                                    city
    # mobile                               mobile
    # otherFacsimileTelephoneNumber        otherFaxNumbers
    # otherHomePhone                       otherHomePhones
    # otherIpPhone                         otherIpPhones
    # otherMobile                          otherMobiles
    # otherPager                           otherPagers
    # otherTelephone                       otherPhones
    # pager                                pager
    # personalTitle                        personalTitle
    # physicalDeliveryOfficeName           officeLocation
    # postalAddress                        streetAddress
    # postalCode                           postalCode
    # postOfficeBox                        postOfficeBox
    # st                                   state
    # street                               streetAddress
    # streetAddress                        streetAddress
    # telephoneNumber                      telephoneNumber
    # thumbnailPhoto                       thumbnailPhoto
    # userCertificate                      userCertificate
    # msDS-cloudExtensionAttribute1        xx
    # msDS-cloudExtensionAttribute2        xx
    # msDS-cloudExtensionAttribute3        xx
    # msDS-cloudExtensionAttribute4        xx
    # msDS-cloudExtensionAttribute5        xx
    # msDS-cloudExtensionAttribute6        xx
    # msDS-cloudExtensionAttribute7        xx
    # msDS-cloudExtensionAttribute8        xx
    # msDS-cloudExtensionAttribute9        xx
    # msDS-cloudExtensionAttribute10       xx
    # msDS-cloudExtensionAttribute11       xx
    # msDS-cloudExtensionAttribute12       xx
    # msDS-cloudExtensionAttribute13       xx
    # msDS-cloudExtensionAttribute14       xx
    # msDS-cloudExtensionAttribute15       xx
    # msDS-ExternalDirectoryObjectId       externalDirectoryObjectId
    
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
        Write-Verbose 'Retrieving Exchange Online Dynamic Distribution Groups...'
        # Retrieve dynamic distribution groups
        if ([string]::IsNullOrWhitespace($GroupId) -eq $false) {
            try {
                $exchangeGroups = @(Get-DynamicDistributionGroup -Identity $GroupId -ErrorAction Stop)
            }
            catch {
                Write-Error "Error retrieving Exchange Online group with GroupId $GroupId. $($_.Exception.Message)"
                return
            }
        }
        else {
            try {
                $exchangeGroups = Get-DynamicDistributionGroup -ErrorAction Stop
                Write-Verbose "Found $($exchangeGroups.Count) Exchange Online Dynamic Distribution Groups"
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
                    MembersCount                  = 'N/A'
                    MembersName                   = 'N/A'
                    MembersId                     = 'N/A'
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
        }

        $exchangeCounter = 0
        $exchangeTotal = @($exchangeGroups).Count
        foreach ($group in $exchangeGroups) {
            $exchangeCounter++
            Write-Verbose "[$exchangeCounter/$exchangeTotal] Processing Exchange group: $($group.Name)"
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
                MembersCount                  = 0
                MembersName                   = $null
                MembersId                     = $null
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
            
            try {
                if ($MemberReport.IsPresent) {
                    $ddgMembers = @(Get-DynamicDistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
                    $object.MembersCount = $ddgMembers.Count
                    foreach ($m in $ddgMembers) {
                        $memberReportArray.Add([PSCustomObject][ordered]@{
                                GroupId           = $group.Guid
                                GroupName         = $group.Name
                                GroupType         = 'Exchange Dynamic Distribution Group'
                                MembershipRule    = $object.MembershipRule
                                MemberId          = $m.Guid
                                MemberDisplayName = $m.DisplayName
                                MemberPrincipal   = $m.PrimarySmtpAddress
                                MemberType        = $m.RecipientType
                            })
                    }
                }
                elseif ($IncludeMembers.IsPresent) {
                    $ddgMembers = @(Get-DynamicDistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop)
                    $object.MembersCount = $ddgMembers.Count
                    $object.MembersName = ($ddgMembers | ForEach-Object { $_.DisplayName }) -join '|'
                    $object.MembersId = ($ddgMembers | ForEach-Object { $_.Guid }) -join '|'
                }
                else {
                    $object.MembersCount = @(Get-DynamicDistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop).Count
                }
            }
            catch {
                Write-Verbose "Error retrieving members for $($group.Name): $($_.Exception.Message)"
            }

            $dynGroupArray.Add($object)
        }
    }

    # Check Entra ID for Dynamic Groups
    if (-not $ExchangeOnlineOnly) {
        Write-Verbose 'Retrieving Entra ID Dynamic Groups...'
        if ([string]::IsNullOrWhitespace($GroupId) -eq $false) {
            try {
                $entraGroups = @(Get-MgGroup -GroupId $GroupId -ErrorAction Stop)
                if ($entraGroups.groupTypes -notcontains 'DynamicMembership') {
                    Write-Error "The specified GroupId $GroupId is not a dynamic group."
                    return
                }
            }
            catch {
                Write-Error "Error retrieving Entra ID group with GroupId $GroupId. $($_.Exception.Message)"
                return
            }
        }
        else {
            try {
                $entraGroups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'DynamicMembership')" -ErrorAction Stop
                Write-Verbose "Found $($entraGroups.Count) Entra ID Dynamic Groups"
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
                    MembersCount                  = 'N/A'
                    MembersName                   = 'N/A'
                    MembersId                     = 'N/A'
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
        }

        $entraCounter = 0
        $entraTotal = @($entraGroups).Count
        foreach ($group in $entraGroups) {
            $entraCounter++
            Write-Verbose "[$entraCounter/$entraTotal] Processing Entra ID group: $($group.DisplayName)"
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
                MembersCount                  = 0
                MembersName                   = $null
                MembersId                     = $null
                DisplayName                   = $group.DisplayName
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

            try {
                $mgMembers = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)
                $object.MembersCount = $mgMembers.Count
                if ($MemberReport.IsPresent) {
                    foreach ($m in $mgMembers) {
                        $memberReportArray.Add([PSCustomObject][ordered]@{
                                GroupId           = $group.Id
                                GroupName         = $group.DisplayName
                                GroupType         = $object.Type
                                MembershipRule    = $object.MembershipRule
                                MemberId          = $m.Id
                                MemberDisplayName = $m.AdditionalProperties['displayName']
                                MemberPrincipal   = $m.AdditionalProperties['userPrincipalName']
                                MemberType        = ($m.AdditionalProperties['@odata.type'] -replace '#microsoft\.graph\.', '')
                            })
                    }
                }
                elseif ($IncludeMembers.IsPresent) {
                    $object.MembersName = ($mgMembers | ForEach-Object { $_.AdditionalProperties['displayName'] }) -join '|'
                    $object.MembersId = ($mgMembers | ForEach-Object { $_.Id }) -join '|'
                }
            }
            catch {
                Write-Verbose "Error retrieving members for $($group.DisplayName): $($_.Exception.Message)"
            }

            $dynGroupArray.Add($object)
        }
    }

    # foreach attribute, check if it's in the `Personal-Information` property set attribute
    # if it's, add a warning to the object 
    Write-Verbose "Analyzing security attributes for $($dynGroupArray.Count) groups..."
    foreach ($group in $dynGroupArray) {
        $group | Add-Member -MemberType NoteProperty -Name Warning -Value $null
        foreach ($attribute in $group.UserAttributes.Split('|')) {
            if ($propertySetsAttribute -contains $attribute) {
                $group.Warning = "'$attribute' is in the `Personal-Information` property set, the user can modify it and add himself to the group. See https://itpro-tips.com/property-set-personal-information-and-active-directory-security-and-governance/"
            }
        }
    }
    
    if ($ExportToExcel.IsPresent) {
        Write-Verbose 'Preparing Excel export...'
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-DynamicGroups.xlsx"
        Write-Verbose "Excel file path: $excelFilePath"
        Write-Host -ForegroundColor Cyan "Exporting dynamic groups to Excel file: $excelFilePath"

        # Excel hard limit: 32,767 characters per cell. Also strip illegal XML control chars.
        $maxCellLength = 32700
        $illegalXmlChars = '[\x00-\x08\x0B\x0C\x0E-\x1F]'
        $sanitizeForExcel = {
            param($collection)
            foreach ($item in $collection) {
                foreach ($prop in $item.PSObject.Properties) {
                    if ($prop.Value -is [string]) {
                        $value = $prop.Value -replace $illegalXmlChars, ''
                        if ($value.Length -gt $maxCellLength) {
                            $value = $value.Substring(0, $maxCellLength) + '...[TRUNCATED]'
                        }
                        $prop.Value = $value
                    }
                }
            }
        }

        & $sanitizeForExcel $dynGroupArray
        $dynGroupArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'DynamicGroups'
        if ($MemberReport.IsPresent) {
            & $sanitizeForExcel $memberReportArray
            $memberReportArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'Members'
        }
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif ($MemberReport.IsPresent) {
        Write-Verbose "Returning $($memberReportArray.Count) member rows"
        return $memberReportArray
    }
    else {
        Write-Verbose "Returning $($dynGroupArray.Count) dynamic groups"
        return $dynGroupArray
    }
}