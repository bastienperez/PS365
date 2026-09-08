<#
    .SYNOPSIS
    Compare two user attributes from different sources.

    .DESCRIPTION
    This function compares two specified user attributes from Active Directory, Entra ID (Microsoft Graph),
    or Exchange. It can filter users based on email domain or specific user identities and returns users
    whose attributes either match or do not match.

    PowerShell 7 and Active Directory:
    PS365 requires PowerShell 7 or later. Source AD can still be used on Windows with the ActiveDirectory
    module supplied by compatible RSAT tools. Calling Get-ADUser does not inherently require PowerShell 5.1.
    For example, Microsoft lists the ActiveDirectory module on Windows Server 2019 with RSAT-AD-PowerShell
    as natively compatible with PowerShell 7. Import-Module ActiveDirectory -ErrorAction Stop loads it.

    If your installed AD module needs Windows PowerShell compatibility, run
    Import-Module ActiveDirectory -UseWindowsPowerShell -ErrorAction Stop from PowerShell 7 on Windows.
    This runs the module in a background Windows PowerShell 5.1 session and returns deserialized objects.
    Simple property comparisons generally work, but live AD object methods are not preserved.
    This does not provide ActiveDirectory support on Linux or macOS.

    Only Source AD needs this module, domain connectivity and appropriate AD read permissions.
    Source EntraID and Source Exchange do not depend on ActiveDirectory. The function does not install RSAT
    or explicitly select a compatibility mode; prepare the required module/session before calling it.

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
    Import-Module ActiveDirectory -ErrorAction Stop
    Compare-UserAttribute -Source AD -Attribute1 mail -Attribute2 UserPrincipalName -Return NotMatching

    From PowerShell 7 on Windows with compatible RSAT tools, compares the 'mail' and 'UserPrincipalName'
    attributes in Active Directory and returns users whose values differ.

    .EXAMPLE
    Compare-UserAttribute -Attribute1 "userPrincipalName" -Attribute2 "mail" -Source "EntraID" -Return "NotMatching"

    Compares the 'userPrincipalName' and 'mail' attributes for users in Entra ID
    and returns those with non-matching values.

    .LINK
    https://ps365.clidsys.com/docs/commands/Compare-UserAttribute

    .LINK
    https://learn.microsoft.com/en-us/powershell/windows/module-compatibility?view=windowsserver2019-ps

    .LINK
    https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_windows_powershell_compatibility
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

    # Microsoft Graph is only used when the source is EntraID
    if ($Source -eq 'EntraID') {
        $requiredScopes = @('User.Read.All')
        if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }
    }

    switch ($Source) {
        'AD' {
            if ($ByDomain) {
                $users = Get-ADUser -LDAPFilter "(mail=*$ByDomain)" -Properties $Attribute1, $Attribute2
            }
            elseif ($User) {
                $users = [System.Collections.Generic.List[PSCustomObject]]::new()

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
                $users = Get-MgUser -All | Where-Object { $_.mail -like "*$ByDomain" }
            }
            elseif ($User) {
                $users = [System.Collections.Generic.List[PSCustomObject]]::new()
                
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
            
            if ($ByDomain) {
                $users = Get-Recipient -Filter "EmailAddresses -like '*$ByDomain'" -Properties $Attribute1, $Attribute2 | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
            }
            elseif ($User) {
                $users = [System.Collections.Generic.List[PSCustomObject]]::new()
                
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
