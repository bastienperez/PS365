<#
Attribut AD	Attribut Entra ID
assistant	assistant
c	country
facsimileTelephoneNumber	faxNumber
homePhone	homePhone
homePostalAddress	homePostalAddress
info	notes
ipPhone	ipPhone
l	city
mobile	mobile
otherFacsimileTelephoneNumber	otherFaxNumbers
otherHomePhone	otherHomePhones
otherIpPhone	otherIpPhones
otherMobile	otherMobiles
otherPager	otherPagers
otherTelephone	otherPhones
pager	pager
personalTitle	personalTitle
physicalDeliveryOfficeName	officeLocation
postalAddress	streetAddress
postalCode	postalCode
postOfficeBox	postOfficeBox
st	state
street	streetAddress
streetAddress	streetAddress
telephoneNumber	telephoneNumber
thumbnailPhoto	thumbnailPhoto
userCertificate	userCertificate
msDS-cloudExtensionAttribute1	extensionAttribute1
msDS-cloudExtensionAttribute2	extensionAttribute2
msDS-cloudExtensionAttribute3	extensionAttribute3
msDS-cloudExtensionAttribute4	extensionAttribute4
msDS-cloudExtensionAttribute5	extensionAttribute5
msDS-cloudExtensionAttribute6	extensionAttribute6
msDS-cloudExtensionAttribute7	extensionAttribute7
msDS-cloudExtensionAttribute8	extensionAttribute8
msDS-cloudExtensionAttribute9	extensionAttribute9
msDS-cloudExtensionAttribute10	extensionAttribute10
msDS-cloudExtensionAttribute11	extensionAttribute11
msDS-cloudExtensionAttribute12	extensionAttribute12
msDS-cloudExtensionAttribute13	extensionAttribute13
msDS-cloudExtensionAttribute14	extensionAttribute14
msDS-cloudExtensionAttribute15	extensionAttribute15
msDS-ExternalDirectoryObjectId	externalDirectoryObjectId
https://itpro-tips.com/property-set-personal-information-and-active-directory-security-and-governance/
#>
function Get-DynamicGroupsInM365 {
    param(
        [Parameter(Mandatory = $false)]
        [switch]$ExchangeOnlineOnly,
        [Parameter(Mandatory = $false)]
        [switch]$EntraIDOnly
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
        try {
            # Ensure Exchange Online session is active
            if (-not (Get-ConnectionInformation)) {
                Write-Warning 'Connexion à Exchange Online requise. Utilisez Connect-ExchangeOnline'
                return
            }

            # Retrieve dynamic distribution groups
            $exchangeGroups = Get-DynamicDistributionGroup -ErrorAction Stop

            foreach ($group in $exchangeGroups) {
                $object = [PSCustomObject][ordered]@{
                    Name             = $group.Name
                    Type             = 'Exchange Dynamic Distribution Group'
                    Email            = $group.PrimarySmtpAddress
                    Filter           = $group.LdapRecipientFilter
                    # (&(company=Diptyque)(st=France)(!(!(objectClass=user)))(objectCategory=person)(mailNickname=*)(msExchHomeServerName=*)(!(name=SystemMailbox{*))(!(name=CAS_{*))(!(msExchRecipientTypeDetails=16777216))(!(msExchRecipientTypeDetails=536870912))(!(msExchRecipientTypeDetails=68719476736))(!(msExchRecipientTypeDetails=8388608))(!(msExchRecipientTypeDetails=4398046511104))(!(msExchRecipientTypeDetails=70368744177664))(!(msExchRecipientTypeDetails=140737488355328))(!(msExchRecipientTypeDetails=35184372088832)))
                    # I want to only get the name of the attribute after any '('
                    UserAttributes   = (($group.LdapRecipientFilter | 
                            Select-String -Pattern '\(([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | Select-Object -ExpandProperty Value -Unique).Replace('(', '') -join ' | '
                    GroupAttributes  = $null
                    DeviceAttributes = $null
                }
                $dynGroupArray.Add($object)
            }
        }
        catch {
            Write-Error "Error retrieving Exchange Online groups: $_"
        }
    }

    # Check Entra ID for Dynamic Groups
    if (-not $ExchangeOnlineOnly) {
        try {
            # Ensure Graph module is connected
            if (-not (Get-MgContext)) {
                Write-Warning 'Connection to Entra ID required. Use Connect-MgGraph'
                return
            }

            # Retrieve dynamic groups from Microsoft Graph
            $entraGroups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'DynamicMembership')" -ErrorAction Stop

            foreach ($group in $entraGroups) {
                $object = [PSCustomObject][ordered]@{
                    Name             = $group.DisplayName
                    Type             = if ($group.groupTypes -contains 'Unified') { 'M365 Dynamic Group' } else { 'Entra ID Dynamic Security Group' }
                    Email            = $group.Mail
                    Filter           = $group.MembershipRule
                    #(user.accountEnabled -eq true) and (user.dirSyncEnabled -eq True) and (user.mail -contains "xx")
                    # want to get the attribute user.xxx
                    UserAttributes   = (($group.MembershipRule | 
                            Select-String -Pattern 'user\.([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | 
                        Select-Object -ExpandProperty Value | 
                        ForEach-Object { $_.Replace('user.', '') } | 
                        Sort-Object -Unique) -join '| '
                    GroupAttributes  = (($group.MembershipRule | 
                            Select-String -Pattern 'group\.([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | 
                        Select-Object -ExpandProperty Value | 
                        ForEach-Object { $_.Replace('group.', '') } | 
                        Sort-Object -Unique) -join '| '
                    DeviceAttributes = (($group.MembershipRule | 
                            Select-String -Pattern 'device\.([a-zA-Z][a-zA-Z0-9]*)' -AllMatches).Matches | 
                        Select-Object -ExpandProperty Value | 
                        ForEach-Object { $_.Replace('device.', '') } | 
                        Sort-Object -Unique) -join '| '
                }
                $dynGroupArray.Add($object)
            }
        }
        catch {
            Write-Error "Error retrieving groups Entra ID: $_"
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
    
    return $dynGroupArray
}