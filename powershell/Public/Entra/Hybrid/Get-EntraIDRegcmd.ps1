<#
    .SYNOPSIS
    Retrieves the current device registration status by invoking the 'dsregcmd /status' command and
    parsing its output.

    .DESCRIPTION
    Retrieves the current device registration status by invoking the 'dsregcmd /status' command and parsing its output.

    .EXAMPLE
    Get-EntraIDRegcmd
#>

function Get-EntraIDRegcmd {
    $dsregcmd = dsregcmd /status

    $object = New-Object -TypeName PSObject

    $dsregcmd | Select-String -Pattern ' *[A-z]+ : *' | ForEach-Object {
        $object | Add-Member -MemberType NoteProperty -Name (([String]$_).Trim() -split ' : ')[0] -Value (([String]$_).Trim() -split ' : ')[1]
    }

    return $object
}