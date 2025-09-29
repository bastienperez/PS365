function Get-PermissionDecision {
    [CmdletBinding()]
    param (

    )
    end {

        $PermissionSplat = @{
            Title      = 'Choose one or more options and click OK'
            OutputMode = 'Multiple'
        }
        $PermissionDecision = @(
            [PSCustomObject][ordered]@{
                'Options' = 'FullAccess'
            },
            [PSCustomObject][ordered]@{
                'Options' = 'SendAs'
            },
            [PSCustomObject][ordered]@{
                'Options' = 'SendOnBehalf'
            },
            [PSCustomObject][ordered]@{
                'Options' = 'Folder'
            }
        )
        $PermissionDecision | Out-GridView @PermissionSplat
    }
}