function Set-ExMailboxRegionalConfiguration {
	[CmdletBinding(SupportsShouldProcess)]
	param(
		[Parameter(Mandatory = $true, ParameterSetName = 'Identity')]
		[string[]]$Identity,

		[Parameter(Mandatory = $true, ParameterSetName = 'ByDomain')]
		[string]$ByDomain,

		[Parameter(Mandatory = $true, ParameterSetName = 'FromCSV')]
		[string]$FromCSV,

		[Parameter(Mandatory = $true, ParameterSetName = 'AllMailboxes')]
		[switch]$AllMailboxes,

		[Parameter(Mandatory = $false, ParameterSetName = 'Identity')]
		[Parameter(Mandatory = $false, ParameterSetName = 'ByDomain')]
		[Parameter(Mandatory = $false, ParameterSetName = 'AllMailboxes')]
		[ValidateSet('af-ZA', 'am-ET', 'ar-AE', 'ar-BH', 'ar-DZ', 'ar-EG', 'ar-IQ', 'ar-JO', 'ar-KW', 'ar-LB', 'ar-LY', 'ar-MA', 'ar-OM', 'ar-QA', 'ar-SA', 'ar-SY', 'ar-TN', 'ar-YE', 'as-IN', 'az-Cyrl-AZ', 'az-Latn-AZ', 'ba-RU', 'be-BY', 'bg-BG', 'bn-BD', 'bn-IN', 'bo-CN', 'br-FR', 'bs-Cyrl-BA', 'bs-Latn-BA', 'ca-ES', 'co-FR', 'cs-CZ', 'cy-GB', 'da-DK', 'de-AT', 'de-CH', 'de-DE', 'de-LI', 'de-LU', 'el-GR', 'en-029', 'en-AU', 'en-BZ', 'en-CA', 'en-GB', 'en-IE', 'en-IN', 'en-JM', 'en-MY', 'en-NZ', 'en-PH', 'en-SG', 'en-TT', 'en-US', 'en-ZA', 'en-ZW', 'es-AR', 'es-BO', 'es-CL', 'es-CO', 'es-CR', 'es-DO', 'es-EC', 'es-ES', 'es-GT', 'es-HN', 'es-MX', 'es-NI', 'es-PA', 'es-PE', 'es-PR', 'es-PY', 'es-SV', 'es-US', 'es-UY', 'es-VE', 'et-EE', 'eu-ES', 'fa-IR', 'fi-FI', 'fil-PH', 'fo-FO', 'fr-BE', 'fr-CA', 'fr-CH', 'fr-FR', 'fr-LU', 'fr-MC', 'fy-NL', 'ga-IE', 'gd-GB', 'gl-ES', 'gu-IN', 'ha-Latn-NG', 'he-IL', 'hi-IN', 'hr-BA', 'hr-HR', 'hu-HU', 'hy-AM', 'id-ID', 'ig-NG', 'is-IS', 'it-CH', 'it-IT', 'ja-JP', 'ka-GE', 'kk-KZ', 'km-KH', 'kn-IN', 'ko-KR', 'ky-KG', 'lb-LU', 'lo-LA', 'lt-LT', 'lv-LV', 'mi-NZ', 'mk-MK', 'ml-IN', 'mn-MN', 'mr-IN', 'ms-BN', 'ms-MY', 'mt-MT', 'nb-NO', 'ne-NP', 'nl-BE', 'nl-NL', 'nn-NO', 'or-IN', 'pa-IN', 'pl-PL', 'ps-AF', 'pt-BR', 'pt-PT', 'rm-CH', 'ro-RO', 'ru-RU', 'rw-RW', 'sa-IN', 'se-FI', 'se-NO', 'se-SE', 'si-LK', 'sk-SK', 'sl-SI', 'sq-AL', 'sr-Cyrl-BA', 'sr-Cyrl-RS', 'sr-Latn-BA', 'sr-Latn-RS', 'sv-FI', 'sv-SE', 'sw-KE', 'ta-IN', 'te-IN', 'th-TH', 'tk-TM', 'tr-TR', 'tt-RU', 'uk-UA', 'ur-PK', 'uz-Cyrl-UZ', 'uz-Latn-UZ', 'vi-VN', 'wo-SN', 'xh-ZA', 'yo-NG', 'zh-CN', 'zh-HK', 'zh-MO', 'zh-SG', 'zh-TW', 'zu-ZA')]
		[string]$Language,

		[Parameter(Mandatory = $false, ParameterSetName = 'Identity')]
		[Parameter(Mandatory = $false, ParameterSetName = 'ByDomain')]
		[Parameter(Mandatory = $false, ParameterSetName = 'AllMailboxes')]
		[ValidateSet('+1', '+2', '+3', '+4', '+5', '+6', '+7', '+8', '+9', '+10', '+11', '+12', '0', '-1', '-2', '-3', '-4', '-5', '-6', '-7', '-8', '-9', '-10', '-11', '-12')]
		[string]$TimeZone,

		[Parameter(Mandatory = $false, ParameterSetName = 'Identity')]
		[Parameter(Mandatory = $false, ParameterSetName = 'ByDomain')]
		[Parameter(Mandatory = $false, ParameterSetName = 'AllMailboxes')]
		[Parameter(Mandatory = $false)]
		[string]$DateFormat,

		[Parameter(Mandatory = $false, ParameterSetName = 'Identity')]
		[Parameter(Mandatory = $false, ParameterSetName = 'ByDomain')]
		[Parameter(Mandatory = $false, ParameterSetName = 'AllMailboxes')]
		[Parameter(Mandatory = $false)]
		[string]$TimeFormat,

		[Parameter(Mandatory = $false)]
		[switch]$GenerateCmdlets,

		[Parameter(Mandatory = $false)]
		[string]$OutputFile = "$(Get-Date -Format 'yyyy-MM-dd_HHmmss')-SetMailboxRegionalConfig_Commands.txt"
	)

	# Time zone mappings for simplified input (e.g., '+1' -> 'Romance Standard Time')
	# Based on Exchange data from https://github.com/bastienperez/AD-Exchange-M365-Localization/blob/main/Exchange_TimeZones.csv
	$TimeZoneMappings = @{
		'+1'  = 'Romance Standard Time'
		'+2'  = 'GTB Standard Time'
		'+3'  = 'Arab Standard Time'
		'+4'  = 'Arabian Standard Time'
		'+5'  = 'West Asia Standard Time'
		'+6'  = 'Central Asia Standard Time'
		'+7'  = 'SE Asia Standard Time'
		'+8'  = 'China Standard Time'
		'+9'  = 'Tokyo Standard Time'
		'+10' = 'AUS Eastern Standard Time'
		'+11' = 'Central Pacific Standard Time'
		'+12' = 'New Zealand Standard Time'
		'0'   = 'GMT Standard Time'
		'-1'  = 'Azores Standard Time'
		'-2'  = 'Mid-Atlantic Standard Time'
		'-3'  = 'E. South America Standard Time'
		'-4'  = 'Atlantic Standard Time'
		'-5'  = 'Eastern Standard Time'
		'-6'  = 'Central Standard Time'
		'-7'  = 'Mountain Standard Time'
		'-8'  = 'Pacific Standard Time'
		'-9'  = 'Alaskan Standard Time'
		'-10' = 'Hawaiian Standard Time'
		'-11' = 'UTC-11'
		'-12' = 'Dateline Standard Time'
	}

	if ($PSCmdlet.ParameterSetName -eq 'ByDomain') {
		# we can't filter PrimarySmtpAddress with `-like '*domain'` so we first find mailboxes with emailaddresses matching the domain then filter primarySMTPAddress
		# https://michev.info/blog/post/2404/exchange-online-now-supports-the-like-operator-for-primarysmtpaddress-filtering-sort-of

		Write-Host -ForegroundColor Cyan "Searching mailboxes with PrimarySmtpAddress matching '*@$ByDomain'"
		$Mailboxes = Get-Mailbox -Filter "EmailAddresses -like '*@$ByDomain'" -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
	}
	elseif ($PSCmdlet.ParameterSetName -eq 'AllMailboxes') {
		$Mailboxes = Get-Mailbox -ResultSize Unlimited
	}
	elseif ($PSCmdlet.ParameterSetName -eq 'FromCSV') {
		[System.Collections.Generic.List[PSCustomObject]]$Mailboxes = @()
		if (-not (Test-Path $FromCSV)) {
			Write-Error "CSV file not found: $FromCSV"
			return
		}

		$csvData = Import-Csv -Path $FromCSV

		foreach ($row in $csvData) {
			try {
				$Mailbox = Get-Mailbox -Identity $row.PrimarySmtpAddress -ErrorAction Stop
				$Mailbox | Add-Member -NotePropertyName 'CSVLanguage' -NotePropertyValue $row.Language -Force
				$Mailbox | Add-Member -NotePropertyName 'CSVTimeZone' -NotePropertyValue $row.TimeZone -Force

				if ($row.DateFormat) {
					$Mailbox | Add-Member -NotePropertyName 'CSVDateFormat' -NotePropertyValue $row.DateFormat -Force
				}
				if ($row.TimeFormat) {
					$Mailbox | Add-Member -NotePropertyName 'CSVTimeFormat' -NotePropertyValue $row.TimeFormat -Force
				}
				
				$Mailboxes.Add($Mailbox)
			}
			catch {
				Write-Warning "Mailbox not found: $($row.PrimarySmtpAddress)"
			}
		}
	}
	elseif ($PSCmdlet.ParameterSetName -eq 'Identity') {
		[System.Collections.Generic.List[PSCustomObject]]$Mailboxes = @()

		foreach ($id in $Identity) {
			try {
				$mbx = Get-Mailbox -Identity $id -ErrorAction Stop
				$Mailboxes.Add($mbx)
			}
			catch {
				Write-Warning "Mailbox not found: $id"
			}
		}
	}
	else {
		Write-Warning 'ParameterSetName Identity is not supported in this function.'
		return 1
	}

	$Commands = @()

	foreach ($mbx in $Mailboxes) {
		$cmdParams = @{
			Identity = $mbx.PrimarySmtpAddress
		}

		# Use CSV-specific settings if available, otherwise use parameter values
		if ($PSCmdlet.ParameterSetName -eq 'FromCSV') {
			if ($mbx.CSVLanguage) { 
				$cmdParams['Language'] = $mbx.CSVLanguage 
				# Auto-detect date and time formats based on language culture
			}

			if ($mbx.CSVDateFormat) {
				$cmdParams['DateFormat'] = $mbx.CSVDateFormat
			}

			if ($mbx.CSVTimeFormat) {
				$cmdParams['TimeFormat'] = $mbx.CSVTimeFormat
			}

			if ($mbx.CSVTimeZone) {
				$TZ = $mbx.CSVTimeZone
				if ($TimeZoneMappings.ContainsKey($TZ)) {
					$TZ = $TimeZoneMappings[$TZ]
				}
				$cmdParams['TimeZone'] = $TZ
			}
		}
		else {
			if ($Language) {
				$cmdParams['Language'] = $Language
			}
			# Auto-detect date and time formats based on language culture
			if ($DateFormat) {
				$cmdParams['DateFormat'] = $DateFormat
			}

			if ($TimeFormat) {
				$cmdParams['TimeFormat'] = $TimeFormat
			}
		
			if ($TimeZone) {
				$TZ = $TimeZone
				if ($TimeZoneMappings.ContainsKey($TZ)) {
					$TZ = $TimeZoneMappings[$TZ]
				}
				$cmdParams['TimeZone'] = $TZ
			}
		}

		$cmdletString = 'Set-MailboxRegionalConfiguration'
		
		$cmdParamstring = ($cmdParams.GetEnumerator() | ForEach-Object { 
				if ($_.Value -match '\s') {
					"-$($_.Key) '$($_.Value)'"
				}
				else {
					"-$($_.Key) $($_.Value)"
				}
			}) -join ' '

		# The LocalizeDefaultFolderName switch localizes the default folder names of the mailbox in the current or specified language.
		# You don't need to specify a value with this switch.
		$cmdParams['LocalizeDefaultFolderName'] = $true
		
		$cmdParamstring = ($cmdParams.GetEnumerator() | ForEach-Object { 
				if ($_.Key -eq 'LocalizeDefaultFolderName') {
					"-$($_.Key)"
				}
				elseif ($_.Value -match '\s') {
					"-$($_.Key) '$($_.Value)'"
				}
				else {
					"-$($_.Key) $($_.Value)"
				}
			}) -join ' '
		
		$FullCommand = "$cmdletString $cmdParamstring"

		if ($GenerateCmdlets) {
			$Commands += $FullCommand
		}

		if (-not $GenerateCmdlets -and $PSCmdlet.ShouldProcess($mbx.PrimarySmtpAddress, 'Set regional configuration')) {
			Write-Host -ForegroundColor Cyan "$($mbx.PrimarySmtpAddress) - Setting configuration..."
			try {
				Set-MailboxRegionalConfiguration @cmdParams -ErrorAction Stop
				Write-Host "$($mbx.PrimarySmtpAddress) - Configuration applied successfully." -ForegroundColor Green
			}
			catch {
				Write-Host "$($mbx.PrimarySmtpAddress) - $($_.Exception.Message)" -ForegroundColor Red
			}
		}
	}

	if ($GenerateCmdlets -and $Commands.Count -gt 0) {
		$Commands | Out-File -FilePath $OutputFile -Encoding UTF8
		Write-Host "Commands generated in file: $OutputFile" -ForegroundColor Cyan
	}
}