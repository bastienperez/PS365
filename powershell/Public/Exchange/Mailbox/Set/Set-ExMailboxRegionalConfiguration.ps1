<#
	.SYNOPSIS
	Sets the regional configuration for Exchange mailboxes, including language, time zone, date format, and time format.

	.DESCRIPTION
	This function sets the regional configuration settings (language, time zone, date format, time format)
	for Exchange Online mailboxes. It can target mailboxes by identity, by domain, from a CSV file, or all mailboxes in the organization.
	It also supports generating the corresponding Set-MailboxRegionalConfiguration cmdlets without executing them.

	.PARAMETER Identity
	The identity of the mailbox(es) to set the regional configuration for. This can be an array of email addresses, usernames, or display names.

	.PARAMETER ByDomain
	Filter mailboxes by domain name. All mailboxes with a primary SMTP address in this domain will be processed.

	.PARAMETER FromCSV
	The path to a CSV file containing mailbox identities and optional regional configuration settings to process.

	.PARAMETER AllMailboxes
	If specified, all mailboxes in the organization will be processed.

	.PARAMETER Language
	The language code to set for the mailbox regional configuration (e.g., 'en-US', 'fr-FR').
	See the ValidateSet in the parameter definition for supported language codes.

	.PARAMETER TimeZone
	The time zone to set for the mailbox regional configuration. This can be a time zone ID (e.g., 'Pacific Standard Time')
	or a simplified offset format (e.g., '+1', '-5').
	See the ValidateSet in the parameter definition for supported offset values.

	.PARAMETER DateFormat
	The date format to set for the mailbox regional configuration (e.g., 'MM/dd/yyyy', 'dd/MM/yyyy').
	Auto-detected based on the specified language if not provided.

	.PARAMETER TimeFormat
	The time format to set for the mailbox regional configuration (e.g., 'HH:mm', 'hh:mm tt').
	Auto-detected based on the specified language if not provided.

	.PARAMETER OnlyIfLanguageEmpty
	If specified, the function will only modify mailboxes that currently have no language configured (empty Language property).
	Useful to avoid overwriting existing configurations set by users or administrators.

	.PARAMETER GenerateCmdlets
	If specified, the function will generate the cmdlets and save them to a file instead of executing them.

	.PARAMETER OutputFile
	The path to the output file where generated cmdlets will be saved. Default is a timestamped file in the current directory.

	.EXAMPLE
	Set-ExMailboxRegionalConfiguration -Identity "user@example.com" -Language "en-US" -TimeZone "+1" -DateFormat "MM/dd/yyyy" -TimeFormat "HH:mm" -GenerateCmdlets -OutputFile "C:\temp\cmdlets.txt"

	Generates the Set-MailboxRegionalConfiguration cmdlet for the specified mailbox with the given regional settings and saves it to the specified file without executing it.

	.EXAMPLE
	Set-ExMailboxRegionalConfiguration -ByDomain "example.com" -Language "fr-FR" -TimeZone "Romance Standard Time"

	Sets the regional configuration for all mailboxes in the specified domain to French language and Romance Standard Time zone.

	.EXAMPLE
	Set-ExMailboxRegionalConfiguration -FromCSV "C:\path\to\file.csv" -GenerateCmdlets -OutputFile "C:\temp\cmdlets.txt"
	
	Generates the Set-MailboxRegionalConfiguration cmdlets for mailboxes listed in the specified CSV file and saves them to the specified file without executing them.

	.EXAMPLE
	Set-ExMailboxRegionalConfiguration -ByDomain "example.com" -Language "en-US" -TimeZone "+1" -OnlyIfLanguageEmpty

	Sets the regional configuration only for mailboxes in the specified domain that currently have no language configured.

	.LINK
	https://ps365.clidsys.com/docs/commands/Set-ExMailboxRegionalConfiguration
#>

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
		[switch]$OnlyIfLanguageEmpty,

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

	$commands = @()
	$totalCount = @($Mailboxes).Count
	$currentCount = 0

	foreach ($mbx in $Mailboxes) {
		$currentCount++

		# Check if OnlyIfLanguageEmpty is specified and skip if language is already configured
		if ($OnlyIfLanguageEmpty) {
			try {
				$currentConfig = Get-MailboxRegionalConfiguration -Identity $mbx.PrimarySmtpAddress -ErrorAction Stop
				if ($currentConfig.Language -and $currentConfig.Language -ne '') {
					Write-Host "[$currentCount/$totalCount] $($mbx.PrimarySmtpAddress) - Skipped (Language already configured: $($currentConfig.Language))" -ForegroundColor Yellow
					continue
				}
			}
			catch {
				Write-Warning "[$currentCount/$totalCount] $($mbx.PrimarySmtpAddress) - Could not retrieve current configuration: $($_.Exception.Message)"
				continue
			}
		}

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
			Write-Host -ForegroundColor Cyan "[$currentCount/$totalCount] $($mbx.PrimarySmtpAddress) - Setting configuration..."
			try {
				Set-MailboxRegionalConfiguration @cmdParams -ErrorAction Stop
				Write-Host "[$currentCount/$totalCount] $($mbx.PrimarySmtpAddress) - Configuration applied successfully." -ForegroundColor Green
			}
			catch {
				Write-Host "[$currentCount/$totalCount] $($mbx.PrimarySmtpAddress) - $($_.Exception.Message)" -ForegroundColor Red
			}
		}
	}

	if ($GenerateCmdlets -and $Commands.Count -gt 0) {
		$Commands | Out-File -FilePath $OutputFile -Encoding UTF8
		Write-Host "Commands generated in file: $OutputFile" -ForegroundColor Cyan
	}
}