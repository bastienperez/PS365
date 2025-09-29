function Get-LicenseDecision {
    [CmdletBinding()]
    param (

    )
    end {

        $LicenseSkuSplat = @{
            Title      = 'Choose one or more options and click OK'
            OutputMode = 'Multiple'
        }
        $LicenseSkuDecision = @(
            [PSCustomObject][ordered]@{
                'Options' = 'AddSkus'
            },
            [PSCustomObject][ordered]@{
                'Options' = 'AddOptions'
            },
            [PSCustomObject][ordered]@{
                'Options' = 'RemoveSkus'
            },
            [PSCustomObject][ordered]@{
                'Options' = 'RemoveOptions'
            }
        )
        $LicenseSkuDecision | Out-GridView @LicenseSkuSplat
    }
}