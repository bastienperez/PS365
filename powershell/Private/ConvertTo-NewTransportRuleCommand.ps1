function ConvertTo-NewTransportRuleCommand {
	param(
		[Parameter(Mandatory = $true)]
		[System.Collections.IDictionary]$RuleParams
	)

	$sb = [System.Text.StringBuilder]::new()
	$null = $sb.Append('New-TransportRule')

	# Iterate keys in a deterministic order so -GenerateCmdlets output is stable across runs.
	$orderedKeys = @($RuleParams.Keys) | Sort-Object
	foreach ($key in $orderedKeys) {
		$value = $RuleParams[$key]
		$null = $sb.Append(" -$key ")
		$null = $sb.Append((Format-TransportRuleValue -Value $value))
	}

	return $sb.ToString()
}
