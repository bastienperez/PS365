function Invoke-AddX500FromContact {
    [CmdletBinding()]
    param (

        [Parameter(Mandatory)]
        $MatchingPrimary
    )
    $AllFound = $MatchingPrimary.where{ $_.Found -eq 'TRUE' } | Sort-Object TargetDisplayName
    $Count = @($AllFound).Count
    $i = 0
    foreach ($Item in $AllFound) {
        $i++
        [PSCustomObject][ordered]@{
            Num                  = '[{0} of {1}]' -f $i, $Count
            TargetDisplayName    = $Item.TargetDisplayName
            SourceDisplayName    = $Item.SourceDisplayName
            TargetType           = $Item.TargetType
            PrimarySmtpAddress   = $Item.PrimarySmtpAddress
            LegacyExchangeDN     = $Item.LegacyExchangeDN
            X500                 = $Item.X500
            TargetGUID           = $Item.TargetGUID
            TargetIdentity       = $Item.TargetIdentity
            SourceName           = $Item.SourceName
            SourceEmailAddresses = $Item.SourceEmailAddresses
            TargetEmailAddresses = $Item.TargetEmailAddresses
        }
    }
}