<#
    .SYNOPSIS
    Converts between Microsoft Entra ID SID and Object ID formats.

    .DESCRIPTION
    Automatically detects whether the input is a SID (S-1-12-1-...) or an Object ID (GUID)
    and converts it to the other format.

    - SID (S-1-12-1-...) → returns the corresponding Object ID (GUID)
    - Object ID (GUID)   → returns the corresponding SID (S-1-12-1-...)

    .PARAMETER Value
    The SID or Object ID to convert. Accepts pipeline input.
    - SID format:       S-1-12-1-{4 decimal numbers separated by dashes}
    - Object ID format: GUID (e.g., xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)

    .EXAMPLE
    Convert-EntraObjectIDAndSID -Value "S-1-12-1-1234567890-987654321-1122334455-5544332211"

    Converts the given SID to its corresponding Object ID (GUID).

    .EXAMPLE
    Convert-EntraObjectIDAndSID -Value "a1b2c3d4-e5f6-7890-abcd-ef1234567890"

    Converts the given Object ID to its corresponding SID.

    .EXAMPLE
    "S-1-12-1-1234567890-987654321-1122334455-5544332211", "a1b2c3d4-e5f6-7890-abcd-ef1234567890" | Convert-EntraObjectIDAndSID

    Converts multiple values via pipeline input.

    .LINK
    https://ps365.clidsys.com/docs/commands/Convert-EntraObjectIDAndSID

    .NOTES

    Entra ID uses a specific SID format: S-1-12-1-{4 UInt32 components derived from the GUID bytes}.
    This conversion is bijective: SID <-> ObjectID.
#>

function Convert-EntraObjectIDAndSID {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, HelpMessage = 'SID (S-1-12-1-...) or Object ID (GUID) to convert')]
        [ValidateNotNullOrEmpty()]
        [string]$Value
    )

    process {
        if ($Value -match '^S-1-12-1-') {
            # SID → ObjectID
            $text = $Value -replace '^S-1-12-1-', ''
            $array = [UInt32[]]$text.Split('-')

            $bytes = [byte[]]::new(16)
            [System.Buffer]::BlockCopy($array, 0, $bytes, 0, 16)
            [Guid]$guid = $bytes

            return $guid.Guid
        }
        else {
            # ObjectID → SID
            try {
                $sidComponents = [UInt32[]]::new(4)
                [System.Buffer]::BlockCopy([Guid]::Parse($Value).ToByteArray(), 0, $sidComponents, 0, 16)
                return "S-1-12-1-$sidComponents".Replace(' ', '-')
            }
            catch {
                Write-Error "Invalid value: '$Value' is neither a valid Entra ID SID (S-1-12-1-...) nor a valid GUID."
            }
        }
    }
}
