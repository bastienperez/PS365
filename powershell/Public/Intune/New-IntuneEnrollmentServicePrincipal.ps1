<#
    .SYNOPSIS
    Creates the Microsoft Intune Enrollment Service Principal if it does not already exist.

    .DESCRIPTION
    Creates the Microsoft Intune Enrollment Service Principal if it does not already exist in your Microsoft 365 tenant.

    The Microsoft Intune Enrollment Service Principal (Application ID: d4ebce55-015a-49b5-a083-c84d1797ae8c) is essential for
    properly configuring Conditional Access policies that target device enrollment scenarios

    .EXAMPLE
    New-IntuneEnrollmentServicePrincipal

    Creates the Microsoft Intune Enrollment Service Principal if it doesn't already exist in the tenant.

    .LINK
    https://ps365.clidsys.com/docs/commands/New-IntuneEnrollmentServicePrincipal

    .NOTES
    Scope(s) required:
    - ServicePrincipal.ReadWrite.All
    Microsoft documentation:
    https://learn.microsoft.com/en-us/intune/intune-service/enrollment/multi-factor-authentication
    > The Microsoft Intune Enrollment cloud app isn't created automatically for new tenants. To add the app for new tenants, a Microsoft Entra administrator must create a service principal object, with app ID d4ebce55-015a-49b5-a083-c84d1797ae8c, in PowerShell or Microsoft Graph.
#>

function New-IntuneEnrollmentServicePrincipal {

    $intuneEnrollmentAppExists = [bool](Invoke-MgGraphRequest -Method GET -Uri $intuneEnrollmentAppUri -ContentType 'PSObject' -OutputType PSObject).value.Count -gt 0

    if (-not $intuneEnrollmentAppExists) {
        Write-Host -ForegroundColor Magenta 'Creating Microsoft Intune Enrollment'
        $body = @{ appId = 'd4ebce55-015a-49b5-a083-c84d1797ae8c' } | ConvertTo-Json
        $null = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body $body -ContentType 'application/json'
        Write-Host -ForegroundColor Green 'Microsoft Intune Enrollment created'
    }
    else {
        Write-Host -ForegroundColor Green 'Microsoft Intune Enrollment already exists'
    }
}