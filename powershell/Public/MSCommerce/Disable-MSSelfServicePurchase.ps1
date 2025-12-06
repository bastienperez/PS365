<#
    .SYNOPSIS
    Disables self-service purchase for all products in Microsoft 365.

    .DESCRIPTION
    This function connects to the Microsoft Commerce service and disables the self-service purchase
    option for all products that currently have it enabled. This is useful for organizations that
    want to prevent users from purchasing additional services or products on their own.

    .EXAMPLE
    Disable-MSSelfServicePurchase

    Disables self-service purchase for all products in the Microsoft 365 tenant.
#>

function Disable-MSSelfServicePurchase {
    # Install-Module -Name MSCommerce -Scope CurrentUser
    # Install-PSRessource -Name MSCommerce -Scope CurrentUser
    Import-Module -Name MSCommerce 
    # Global Administrator or Billing Administrator permissions are required to run this script
    Connect-MSCommerce
    Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | Where-Object { $_.PolicyValue -eq 'Enabled' } | ForEach-Object {
        Write-Host -ForegroundColor Cyan "Disabling self-service purchase for product: $($_.ProductID)"
        Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $_.ProductID -Enabled $false  
    }
}   