function Get-PermissionDecisionBatch {
    [CmdletBinding()]
    param (

    )
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
        },
        [PSCustomObject][ordered]@{
            'Options' = 'AddToBatch'
        }
    )
    $PermissionDecision | Out-GridView @PermissionSplat
}