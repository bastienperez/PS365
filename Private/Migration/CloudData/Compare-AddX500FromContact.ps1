function Compare-Addx500FromContact {
    [CmdletBinding()]
    param (

        [Parameter(Mandatory)]
        $Source,

        [Parameter(Mandatory)]
        $Target
    )
    foreach ($Key in $Source.Keys) {
        if ($Target.ContainsKey($Key)) {
            [PSCustomObject][ordered]@{
                FOUND                = 'TRUE'
                TargetDisplayName    = $Target[$Key]['DisplayName']
                SourceDisplayName    = $Source[$Key]['DisplayName']
                TargetType           = $Target[$Key]['RecipientTypeDetails']
                PrimarySmtpAddress   = $Key
                LegacyExchangeDN     = $Source[$Key]['LegacyExchangeDN']
                X500                 = $Source[$Key]['X500']
                TargetGUID           = $Target[$Key]['GUID']
                TargetIdentity       = $Target[$Key]['Identity']
                SourceName           = $Source[$Key]['Name']
                TargetEmailAddresses = $Target[$Key]['EmailAddresses']
                SourceEmailAddresses = $Source[$Key]['EmailAddresses']
            }
        }
        else {
            [PSCustomObject][ordered]@{
                FOUND                = 'FALSE'
                TargetDisplayName    = 'NOTFOUND'
                SourceDisplayName    = $Source[$Key]['DisplayName']
                TargetType           = 'NOTFOUND'
                PrimarySmtpAddress   = $Key
                LegacyExchangeDN     = $Source[$Key]['LegacyExchangeDN']
                X500                 = $Source[$Key]['X500']
                TargetGUID           = 'NOTFOUND'
                TargetIdentity       = 'NOTFOUND'
                SourceName           = $Source[$Key]['Name']
                TargetEmailAddresses = 'NOTFOUND'
                SourceEmailAddresses = $Source[$Key]['EmailAddresses']
            }
        }
    }
}