function Format-TransportRuleValue {
	param(
		[Parameter(Mandatory = $false)]
		$Value
	)

	if ($null -eq $Value) {
		return '$null'
	}
	if ($Value -is [bool]) {
		return ('${0}' -f $Value.ToString().ToLower())
	}
	if ($Value -is [int] -or $Value -is [long] -or $Value -is [double]) {
		return [string]$Value
	}
	if ($Value -is [System.Collections.IDictionary]) {
		$nestedKeys = @($Value.Keys) | Sort-Object
		$parts = foreach ($k in $nestedKeys) {
			'{0} = {1}' -f $k, (Format-TransportRuleValue -Value $Value[$k])
		}
		return '@{ ' + ($parts -join '; ') + ' }'
	}
	if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
		$parts = foreach ($item in $Value) { Format-TransportRuleValue -Value $item }
		return '@(' + ($parts -join ', ') + ')'
	}

	$escaped = ([string]$Value).Replace("'", "''")
	return "'$escaped'"
}
