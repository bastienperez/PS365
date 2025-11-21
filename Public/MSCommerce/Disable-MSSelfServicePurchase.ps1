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