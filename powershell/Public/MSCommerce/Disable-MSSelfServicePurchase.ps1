<#
    .SYNOPSIS
    Disables self-service purchase for ALL products in Microsoft 365.

    .DESCRIPTION
    This function connects to the Microsoft Commerce service and disables the self-service purchase
    option for all products that currently have it enabled. This is useful for organizations that
    want to prevent users from purchasing additional services or products on their own.

    .PARAMETER Force
    Skips confirmation prompts and disables self-service purchase for all products without asking.

    .EXAMPLE
    Disable-MSSelfServicePurchase

    Disables self-service purchase for all products in the Microsoft 365 tenant, prompting for confirmation for each product.

    .EXAMPLE
    Disable-MSSelfServicePurchase -Force

    Disables self-service purchase for all products without prompting for confirmation.

    .LINK
    https://ps365.clidsys.com/docs/commands/Disable-MSSelfServicePurchase
#>

function Disable-MSSelfServicePurchase {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [switch]$Force
    )

    # Install-Module -Name MSCommerce -Scope CurrentUser
    # Install-PSRessource -Name MSCommerce -Scope CurrentUser
    Import-Module -Name MSCommerce
    # Global Administrator or Billing Administrator permissions are required to run this script
    Connect-MSCommerce
    $enabledProducts = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | Where-Object { $_.PolicyValue -eq 'Enabled' }
    foreach ($product in $enabledProducts) {
        Write-Host -ForegroundColor Cyan "Disabling self-service purchase for product: $($product.ProductID) - $($product.ProductName)"

        if ($Force -or $PSCmdlet.ShouldProcess($product.ProductID, 'Disable self-service purchase')) {
            try {
                # $null because Update-MSCommerceProductPolicy returns by default the updated product policy, but we don't need it here
                $null = Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $product.ProductID -Enabled $false -InformationAction SilentlyContinue -ErrorAction Stop
                Write-Host -ForegroundColor Green "Successfully disabled self-service purchase for product: $($product.ProductID) - $($product.ProductName)"
            }
            catch {
                Write-Error "Failed to disable self-service purchase for product: $($product.ProductID) - $($product.ProductName). Error: $_"
            }
        }
    }
}   