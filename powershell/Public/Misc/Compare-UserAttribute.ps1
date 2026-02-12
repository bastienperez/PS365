<#
    .SYNOPSIS
    Compare two user attributes from different sources.

    .DESCRIPTION
    This function compares two specified user attributes from Active Directory, Entra ID (Microsoft Graph),
    or Exchange. It can filter users based on email domain or specific user identities and returns users
    whose attributes either match or do not match.

    .PARAMETER Attribute1
    The first user attribute to compare.

    .PARAMETER Attribute2
    The second user attribute to compare.

    .PARAMETER Source
    The source from which to retrieve user information.
    Valid options are 'AD' for Active Directory, 'EntraID' for Microsoft Entra ID, and 'Exchange' for Exchange.

    .PARAMETER User
    An array of user identities to filter the comparison.

    .PARAMETER ByDomain
    An array of email domains to filter users.

    .PARAMETER Return
    Specifies whether to return users with 'Matching' or 'NotMatching' attributes.

    .EXAMPLE
    Compare-UserAttribute -Attribute1 "mail" -Attribute2 "proxyAddresses" -

    Compares the 'mail' and 'proxyAddresses' attributes for users in Active Directory
    and returns those with matching values.

    .EXAMPLE
    Compare-UserAttribute -Attribute1 "userPrincipalName" -Attribute2 "mail" -Source "EntraID" -Return "NotMatching"

    Compares the 'userPrincipalName' and 'mail' attributes for users in Entra ID
    and returns those with non-matching values.

    .LINK
    https://ps365.clidsys.com/docs/commands/Compare-UserAttribute
#>

function Compare-UserAttribute {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$Attribute1,

        [Parameter(Mandatory = $true)]
        [String]$Attribute2,

        [Parameter(Mandatory = $true)]
        [ValidateSet('AD', 'EntraID', 'Exchange')]
        [String]$Source,

        [Parameter(Mandatory = $false)]
        [String[]]$User,

        [Parameter(Mandatory = $false)]
        [String[]]$ByDomain,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Matching', 'NotMatching')]
        [String]$Return
    )

    switch ($Source) {
        'AD' {
            if ($ByDomain) {
                $users = Get-ADUser -LDAPFilter "(mail=*$ByDomain)" -Properties $Attribute1, $Attribute2
            }
            elseif ($User) {
                [System.Collections.Generic.List[PSCustomObject]]$users = @()

                foreach ($u in $User) {
                    $adUser = Get-ADUser -Identity $u -Properties $Attribute1, $Attribute2
                    $users.Add($adUser)
                }
            }
            else {
                $users = Get-ADUser -Filter * -Properties $Attribute1, $Attribute2
            }

            break
        }

        'EntraID' {
            Write-Verbose "Using Microsoft Graph to compare attributes '$Attribute1' and '$Attribute2' in Entra ID."
            if ($ByDomain) {
                $users = (Invoke-PS365GraphRequest -Uri '/v1.0/users' -All) | Where-Object { $_.mail -like "*$ByDomain" }
            }
            elseif ($User) {
                [System.Collections.Generic.List[PSCustomObject]]$users = @()
                
                foreach ($u in $User) {
                    $mgUser = Invoke-PS365GraphRequest -Uri "/v1.0/users/$u" -Select (@($Attribute1, $Attribute2) -join ',')
                    $users.Add($mgUser)
                }
            }
            else {
                $users = Invoke-PS365GraphRequest -Uri '/v1.0/users' -All -Filter 'mail ne null' -ConsistencyLevel 'eventual' -Select (@($Attribute1, $Attribute2) -join ',')
            }

            break
        }

        'Exchange' {
            if ($Attribute1 -eq 'UserPrincipalName') {
                Write-Warning 'Attribute1 "UserPrincipalName" is not available in Exchange. Using "WindowsLiveID" instead.'
                $Attribute1 = 'WindowsLiveID'
            }

            if ($Attribute2 -eq 'UserPrincipalName') {
                Write-Warning 'Attribute2 "UserPrincipalName" is not available in Exchange. Using "WindowsLiveID" instead.'
                $Attribute2 = 'WindowsLiveID'
            }
            
            if ($ByDomain) {
                $users = Get-Recipient -Filter "EmailAddresses -like '*$ByDomain'" -Properties $Attribute1, $Attribute2 | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
            }
            elseif ($User) {
                [System.Collections.Generic.List[PSCustomObject]]$users = @()
                
                foreach ($u in $User) {
                    $exchUser = Get-Recipient -Identity $u -Properties $Attribute1, $Attribute2
                    $users.Add($exchUser)
                }
            }
            else {
                $users = Get-Recipient -ResultSize unlimited -Properties $Attribute1, $Attribute2
            }

            break
        }

        default {
            Write-Error "Unsupported source: $Source. Supported sources are 'AD', 'EntraID', and 'Exchange'."
            return
        }
    }

    # Retourner les objets selon le choix de l'utilisateur
    switch ($Return) {
        'Matching' {
            $result = $users | Where-Object { $_.$Attribute1 -eq $_.$Attribute2 }
        }
        'NotMatching' {
            $result = $users | Where-Object { $_.$Attribute1 -ne $_.$Attribute2 }
        }
    }

    return $result
}