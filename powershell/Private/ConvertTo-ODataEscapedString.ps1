function ConvertTo-ODataEscapedString {
    <#
    .SYNOPSIS
    Escapes a value for safe use inside an OData filter string literal.

    .DESCRIPTION
    Doubles single quotes per the OData specification so that a value containing
    an apostrophe cannot break out of a '...' literal in a Graph -Filter
    expression (which would break the query or, worse, alter the filter).

    .PARAMETER Value
    The raw string value to escape. Empty strings are allowed.

    .EXAMPLE
    $escaped = ConvertTo-ODataEscapedString -Value $DeviceName
    Get-MgDevice -Filter "displayName eq '$escaped'"
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Value
    )

    return ($Value -replace "'", "''")
}
