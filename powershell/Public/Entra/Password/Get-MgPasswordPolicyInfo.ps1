<#
    .SYNOPSIS
    Retrieves password policy details for all verified domains in the Microsoft Entra tenant.

    .DESCRIPTION
    This function connects to Microsoft Entra (Azure AD) using the Microsoft Graph PowerShell module
    and retrieves password policy details for all verified domains in the tenant. It provides information
    such as the password validity period and notification window for each domain.

    .EXAMPLE
    Get-MgPasswordPolicyInfo

    Retrieves password policy details for all verified domains in the tenant.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgPasswordPolicyInfo

    .NOTES
#>

function Get-MgPasswordPolicyInfo { 

    [System.Collections.Generic.List[PSCustomObject]]$pwdPolicies = @()

    $domains = Get-MgDomain
    
    foreach ($domain in $domains) {
        
        if ($domain.PasswordValidityPeriodInDays -eq '2147483647' -or $null -eq $domain.PasswordValidityPeriodInDays) {
            $pwddValidityPeriodInDays = 'Password never expire'
        }
        else {
            $pwddValidityPeriodInDays = $domain.PasswordValidityPeriodInDays
        }
        
        $object = [PSCustomObject][ordered]@{
            Domain                           = $domain.Id
            PasswordValidityPeriodInDays     = $pwddValidityPeriodInDays
            PasswordNotificationWindowInDays = $domain.PasswordNotificationWindowInDays
        }

        $pwdPolicies.Add($object)
    }

    return $pwdPolicies
} 