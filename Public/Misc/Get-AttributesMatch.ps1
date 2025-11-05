function Get-AttributeMatching {
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
        [String[]]$FromEmailDomain,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Matching', 'NotMatching')]
        [String]$Return
    )

    switch ($Source) {
        'AD' {
            if ($FromEmailDomain) {
                $users = Get-ADUser -LDAPFilter "(mail=*$FromEmailDomain)" -Properties $Attribute1, $Attribute2
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
            if ($FromEmailDomain) {
                $users = Get-MgUser -All | Where-Object { $_.mail -like "*$FromEmailDomain" }
            }
            elseif ($User) {
                [System.Collections.Generic.List[PSCustomObject]]$users = @()
                
                foreach ($u in $User) {
                    $mgUser = Get-MgUser -UserId $u -Property $Attribute1, $Attribute2
                    $users.Add($mgUser)
                }
            }
            else {
                $users = Get-MgUser -Filter 'mail ne null' -Property $Attribute1, $Attribute2
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
            
            if ($FromEmailDomain) {
                $users = Get-Recipient -Filter "EmailAddresses -like '*$FromEmailDomain'" -Properties $Attribute1, $Attribute2 | Where-Object { $_.PrimarySmtpAddress -like "*@$FromEmailDomain" }
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