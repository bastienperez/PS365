function ConvertTo-FlatObject {
    <#
    .SYNOPSIS
    Flattens a nested PSObject (audit record) into a single-level hashtable.

    .DESCRIPTION
    Extracted from Search-UnifiedAuditLogCustom so the flattening logic is
    reusable and unit-testable on its own. Recursively expands nested objects,
    dictionaries and lists into prefixed keys, and special-cases the audit
    Parameters array (ParameterString / Param_* / FullCommand).
    #>
    param (
        [Parameter(Mandatory = $true)]
        [PSObject]$InputObject,

        [Parameter(Mandatory = $false)]
        [string]$Prefix = '',

        [Parameter(Mandatory = $false)]
        [switch]$PreserveTypes
    )

    $flatProperties = @{}

    foreach ($property in $InputObject.PSObject.Properties) {
        $key = if ($Prefix) { "${Prefix}_$($property.Name)" } else { $property.Name }

        if ($property.Name -eq 'Parameters' -and $property.Value -is [Array]) {
            $parameterStrings = foreach ($parameter in $property.Value) {
                "$($parameter.Name)=$($parameter.Value)"
            }
            $flatProperties['ParameterString'] = $parameterStrings -join ' | '

            foreach ($parameter in $property.Value) {
                $parameterKey = "Param_$($parameter.Name)"
                $flatProperties[$parameterKey] = $parameter.Value
            }

            if ($InputObject.Operation) {
                $parameterStrings = foreach ($parameter in $property.Value) {
                    $parameterValue = switch -Regex ($parameter.Value) {
                        '\s' { "'$($parameter.Value)'" }
                        '^True$|^False$' { "`$$($parameter.Value.ToLower())" }
                        ';' { "'$($parameter.Value)'" }
                        default { $parameter.Value }
                    }
                    "-$($parameter.Name) $parameterValue"
                }
                $flatProperties['FullCommand'] = "$($InputObject.Operation) $($parameterStrings -join ' ')"
            }

            continue
        }

        switch ($property.Value) {
            { $_ -is [System.Collections.IDictionary] } {
                $nestedProperties = ConvertTo-FlatObject -InputObject $_ -Prefix $key -PreserveTypes:$PreserveTypes
                foreach ($nestedKey in $nestedProperties.Keys) {
                    $uniqueKey = if ($flatProperties.ContainsKey($nestedKey)) {
                        $counter = 1
                        while ($flatProperties.ContainsKey("${nestedKey}_$counter")) {
                            $counter++
                        }
                        "${nestedKey}_$counter"
                    }
                    else {
                        $nestedKey
                    }
                    $flatProperties[$uniqueKey] = $nestedProperties[$nestedKey]
                }
            }
            { $_ -is [System.Collections.IList] -and $property.Name -ne 'Parameters' } {
                if ($_.Count -gt 0) {
                    if ($_[0] -is [PSObject]) {
                        # If every item has a recipient-like shape (Address/EmailAddress/SmtpAddress
                        # +/- Name), collapse the array into a single readable pipe-joined string
                        # instead of exploding it into Recipients_0_Address, Recipients_0_Name, ...
                        $addressProp = $null
                        foreach ($candidate in @('Address', 'EmailAddress', 'SmtpAddress')) {
                            if ($_[0].PSObject.Properties.Name -contains $candidate) {
                                $addressProp = $candidate
                                break
                            }
                        }
                        $everyItemHasAddress = $false
                        if ($addressProp) {
                            $everyItemHasAddress = -not ($_ | Where-Object { -not ($_.PSObject.Properties.Name -contains $addressProp) } | Select-Object -First 1)
                        }

                        if ($everyItemHasAddress) {
                            $joined = ($_ | ForEach-Object {
                                    $addr = $_.$addressProp
                                    $name = if ($_.PSObject.Properties.Name -contains 'Name') { $_.Name } else { $null }
                                    if ($name -and ($name -ne $addr)) { "$name <$addr>" } else { $addr }
                                }) -join '|'
                            $flatProperties[$key] = $joined
                        }
                        else {
                            for ($i = 0; $i -lt $_.Count; $i++) {
                                $nestedProperties = ConvertTo-FlatObject -InputObject $_[$i] -Prefix "${key}_${i}" -PreserveTypes:$PreserveTypes
                                foreach ($nestedKey in $nestedProperties.Keys) {
                                    $uniqueKey = if ($flatProperties.ContainsKey($nestedKey)) {
                                        $counter = 1
                                        while ($flatProperties.ContainsKey("${nestedKey}_$counter")) {
                                            $counter++
                                        }
                                        "${nestedKey}_$counter"
                                    }
                                    else {
                                        $nestedKey
                                    }
                                    $flatProperties[$uniqueKey] = $nestedProperties[$nestedKey]
                                }
                            }
                        }
                    }
                    else {
                        $flatProperties[$key] = $_ -join '|'
                    }
                }
                else {
                    $flatProperties[$key] = [string]::Empty
                }
            }
            { $_ -is [PSObject] } {
                $nestedProperties = ConvertTo-FlatObject -InputObject $_ -Prefix $key -PreserveTypes:$PreserveTypes
                foreach ($nestedKey in $nestedProperties.Keys) {
                    $uniqueKey = if ($flatProperties.ContainsKey($nestedKey)) {
                        $counter = 1
                        while ($flatProperties.ContainsKey("${nestedKey}_$counter")) {
                            $counter++
                        }
                        "${nestedKey}_$counter"
                    }
                    else {
                        $nestedKey
                    }
                    $flatProperties[$uniqueKey] = $nestedProperties[$nestedKey]
                }
            }
            default {
                if ($PreserveTypes) {
                    $flatProperties[$key] = $_
                }
                else {
                    $flatProperties[$key] = switch ($_) {
                        { $_ -is [datetime] } { $_ }
                        { $_ -is [bool] } { $_ }
                        { $_ -is [int] } { $_ }
                        { $_ -is [long] } { $_ }
                        { $_ -is [decimal] } { $_ }
                        { $_ -is [double] } { $_ }
                        default { [string]$_ }
                    }
                }
            }
        }
    }

    return $flatProperties
}
