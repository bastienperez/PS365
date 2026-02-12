<#
.SYNOPSIS
    Retrieves Entra ID Hybrid Join registry keys from the local computer.

.DESCRIPTION
    This function reads registry keys from HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ\AAD
    to retrieve TenantId, TenantName, and other hybrid join configuration values.

.EXAMPLE
    Get-EntraIDHybridJoinComputerRegistryKey

    Returns an object containing TenantId, TenantName, and other registry keys if present.

.NOTES
    Returns a PSCustomObject with ComputerName, TenantId, TenantName, and OtherKeys properties.
    Properties will be null if the registry path or values do not exist.
#>

function Get-EntraIDHybridJoinComputerRegistryKey {

    $path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ\AAD'

    $object = [PSCustomObject][ordered]@{
        ComputerName = $env:COMPUTERNAME
        TenantId     = $null
        TenantName   = $null
        OtherKeys    = ''
    }

    if (-not (Test-Path -Path $path)) {
        return $object
    }

    try {
        $reg = Get-ItemProperty -Path $path -ErrorAction Stop
    }
    catch {
        return $object
    }

    $valueProps = $reg.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' } | Select-Object -ExpandProperty Name

    if ($valueProps -contains 'TenantId') { $object.TenantId = $reg.TenantId }
    if ($valueProps -contains 'TenantName') { $object.TenantName = $reg.TenantName }

    $otherKeys = $valueProps | Where-Object { $_ -notin @('TenantId', 'TenantName') }
    if ($otherKeys) {
        $object.OtherKeys = ($otherKeys | ForEach-Object {
                $value = $reg.$_
                if ($value -is [array]) {
                    $value = $value -join ','
                }
                "$_=$value"
            }) -join '|'
    }

    return $object
}