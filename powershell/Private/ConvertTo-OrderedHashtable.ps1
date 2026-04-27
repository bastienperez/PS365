function ConvertTo-OrderedHashtable {
	<#
		Recursively converts the output of ConvertFrom-Json (PSCustomObject tree) into an [ordered] hashtable so the
		result can be splatted to New-TransportRule. Used as a Windows PowerShell 5.1 substitute for
		ConvertFrom-Json -AsHashtable (PowerShell 6.0+).
	#>
	param(
		[Parameter(Mandatory = $false)]
		$InputObject
	)

	if ($null -eq $InputObject) {
		return $null
	}
	if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
		$result = [ordered]@{}
		foreach ($prop in $InputObject.PSObject.Properties) {
			$result[$prop.Name] = ConvertTo-OrderedHashtable -InputObject $prop.Value
		}
		return $result
	}
	if ($InputObject -is [System.Collections.IList] -and -not ($InputObject -is [string])) {
		$list = New-Object System.Collections.ArrayList
		foreach ($item in $InputObject) {
			$null = $list.Add((ConvertTo-OrderedHashtable -InputObject $item))
		}
		return , $list.ToArray()
	}
	return $InputObject
}
