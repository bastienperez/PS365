function Connect-CloudDeleteCredential {
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory)]
        [string]
        $CredFile
    )
    end {
        try {
            Remove-Item $CredFile -Force -ErrorAction Stop
        }
        catch {
            $_.Exception.Message
        }
    }
}
