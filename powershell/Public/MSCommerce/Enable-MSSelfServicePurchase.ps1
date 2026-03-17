<#
    .SYNOPSIS
    Enables self-service purchase for all products in Microsoft 365.

    .DESCRIPTION
    This function connects to the Microsoft Commerce service and enables the self-service purchase
    option for all products that currently have it disabled. This is useful for organizations that
    want to allow users to purchase additional services or products on their own.

    .PARAMETER Force
    Skips confirmation prompts and enables self-service purchase for all products without asking.

    .EXAMPLE
    Enable-MSSelfServicePurchase

    Enables self-service purchase for all products in the Microsoft 365 tenant, prompting for confirmation for each product.

    .EXAMPLE
    Enable-MSSelfServicePurchase -Force

    Enables self-service purchase for all products without prompting for confirmation.

    .LINK
    https://ps365.clidsys.com/docs/commands/Enable-MSSelfServicePurchase
#>

function Enable-MSSelfServicePurchase {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [switch]$Force
    )

    # Install-Module -Name MSCommerce -Scope CurrentUser
    # Install-PSRessource -Name MSCommerce -Scope CurrentUser
    Import-Module -Name MSCommerce
    # Global Administrator or Billing Administrator permissions are required to run this script
    Connect-MSCommerce
    $disabledProducts = Get-MSCommerceProductPolicies -PolicyId AllowSelfServicePurchase | Where-Object { $_.PolicyValue -eq 'Disabled' }
    foreach ($product in $disabledProducts) {
        Write-Host -ForegroundColor Cyan "Enabling self-service purchase for product: $($product.ProductID) - $($product.ProductName)"

        if ($Force -or $PSCmdlet.ShouldProcess($product.ProductID, 'Enable self-service purchase')) {
            try {
                # $null because Update-MSCommerceProductPolicy returns by default the updated product policy, but we don't need it here
                $null = Update-MSCommerceProductPolicy -PolicyId AllowSelfServicePurchase -ProductId $product.ProductID -Enabled $true -InformationAction SilentlyContinue -ErrorAction Stop
                Write-Host -ForegroundColor Green "Successfully enabled self-service purchase for product: $($product.ProductID) - $($product.ProductName)"
            } catch {
                Write-Error "Failed to enable self-service purchase for product: $($product.ProductID) - $($product.ProductName). Error: $_"
            }
        }
    }
}
