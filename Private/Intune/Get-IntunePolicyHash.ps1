function Get-IntunePolicyHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Policy
    )

    $PropertyHash = @{ }

    foreach ($Item in $Policy.PSObject.properties) {
        $PropertyHash[$Item.Name] = $Item.value

    }

    $PropertyHash
}