<#
	.SYNOPSIS
	Creates Exchange Online transport (mail flow) rules from one or more JSON definition files.

	.DESCRIPTION
	This function reads JSON files describing Exchange Online transport rules and provisions them by splatting their content
	to New-TransportRule. The recommended layout is one JSON file per rule, which makes the rule set version-controllable and
	reproducible across tenants. Each JSON file must be a single object whose keys match New-TransportRule parameter names.

	If a rule with the same Name already exists in the tenant, the rule is skipped (use -Force to recreate it after deletion).
	Supports -WhatIf and -Confirm via SupportsShouldProcess. Use -GenerateCmdlets to emit the equivalent New-TransportRule
	cmdlets to a file instead of executing them.

	The function expects a connected ExchangeOnlineManagement session (Connect-ExchangeOnline).

	.PARAMETER Path
	Path to a JSON file or to a directory containing one or more JSON files. When a directory is provided, every *.json file
	in the directory is processed in alphabetical order.

	.PARAMETER Force
	If a rule with the same Name already exists, remove it before creating the new one. Without this switch, existing rules
	are skipped with a warning.

	.PARAMETER GenerateCmdlets
	If specified, the function generates the New-TransportRule cmdlets and saves them to a file instead of executing them.

	.PARAMETER OutputFile
	Path to the output file used by -GenerateCmdlets. If omitted while -GenerateCmdlets is set, defaults to a timestamped
	file in the user profile (cross-platform). Ignored unless -GenerateCmdlets is specified.

	.EXAMPLE
	Connect-ExchangeOnline

	New-ExTransportRuleFromJson -Path "C:\eop-rules\Block-Outbound-OnMicrosoft.json"

	Creates a single transport rule from the specified JSON file.

	.EXAMPLE
	New-ExTransportRuleFromJson -Path "C:\eop-rules" -WhatIf

	Lists every transport rule that would be created from the JSON files in the directory without applying any change.

	.EXAMPLE
	New-ExTransportRuleFromJson -Path "C:\eop-rules" -Force

	Creates every rule defined in the directory; pre-existing rules with the same Name are removed first.

	.EXAMPLE
	New-ExTransportRuleFromJson -Path "C:\eop-rules" -GenerateCmdlets -OutputFile "C:\temp\eop-rules.ps1"

	Emits the equivalent New-TransportRule cmdlets to the specified file without executing them. Useful for review or for
	running the rule provisioning from another host.

	.NOTES
	Prerequisites:
	- PowerShell 5.1 or later. JSON content is converted to an ordered hashtable through a local helper to stay
	  compatible with Windows PowerShell 5.1, which does not support ConvertFrom-Json -AsHashtable.
	- ExchangeOnlineManagement module installed and an active session opened with Connect-ExchangeOnline before
	  running the function (unless -GenerateCmdlets is specified, which only emits the cmdlet text).
	- The signed-in account must hold a role with permission to manage mail flow rules, typically Exchange
	  Administrator or a member of the Organization Management role group.

	.LINK
	https://ps365.clidsys.com/docs/commands/New-ExTransportRuleFromJson
#>

function New-ExTransportRuleFromJson {
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
	param(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()]
		[string[]]$Path,

		[Parameter(Mandatory = $false)]
		[switch]$Force,

		[Parameter(Mandatory = $false)]
		[switch]$GenerateCmdlets,

		[Parameter(Mandatory = $false)]
		[string]$OutputFile
	)

	begin {
		[System.Collections.Generic.List[string]]$commands = @()
		[System.Collections.Generic.List[PSCustomObject]]$resultsArray = @()

		if ($GenerateCmdlets.IsPresent) {
			if ([string]::IsNullOrWhiteSpace($OutputFile)) {
				$userProfile = [Environment]::GetFolderPath('UserProfile')
				if ([string]::IsNullOrWhiteSpace($userProfile)) { $userProfile = $HOME }
				$OutputFile = Join-Path $userProfile "$(Get-Date -Format 'yyyy-MM-dd_HHmmss')-NewExTransportRuleFromJson-Commands.ps1"
			}
		}
		else {
			if (-not (Get-Command -Name New-TransportRule -ErrorAction SilentlyContinue)) {
				Write-Error 'New-TransportRule is not available. Connect to Exchange Online first with Connect-ExchangeOnline.'
				return
			}
		}
	}

	process {
		foreach ($p in $Path) {
			if (-not (Test-Path -LiteralPath $p)) {
				Write-Warning "Path not found: $p"
				continue
			}

			$item = Get-Item -LiteralPath $p
			if ($item.PSIsContainer) {
				$jsonFiles = Get-ChildItem -LiteralPath $p -Filter '*.json' -File | Sort-Object Name
			}
			else {
				$jsonFiles = @($item)
			}

			if (-not $jsonFiles -or $jsonFiles.Count -eq 0) {
				Write-Warning "No JSON file found in: $p"
				continue
			}

			foreach ($file in $jsonFiles) {
				Write-Host -ForegroundColor Cyan "Processing: $($file.FullName)"

				try {
					$jsonObject = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
						$ruleParams = ConvertTo-OrderedHashtable -InputObject $jsonObject
				}
				catch {
					Write-Host -ForegroundColor Red "[$($file.Name)] Invalid JSON: $($_.Exception.Message)"
					$resultsArray.Add([PSCustomObject]@{
						File   = $file.FullName
						Name   = $null
						Status = 'InvalidJson'
						Error  = $_.Exception.Message
					})
					continue
				}

				if (-not $ruleParams.Contains('Name') -or [string]::IsNullOrWhiteSpace([string]$ruleParams['Name'])) {
					Write-Host -ForegroundColor Red "[$($file.Name)] Missing mandatory 'Name' property."
					$resultsArray.Add([PSCustomObject]@{
						File   = $file.FullName
						Name   = $null
						Status = 'MissingName'
						Error  = "Missing 'Name' property"
					})
					continue
				}

				$ruleName = [string]$ruleParams['Name']

				if ($GenerateCmdlets.IsPresent) {
					$commands.Add((ConvertTo-NewTransportRuleCommand -RuleParams $ruleParams))
					$resultsArray.Add([PSCustomObject]@{
						File   = $file.FullName
						Name   = $ruleName
						Status = 'CmdletGenerated'
						Error  = $null
					})
					continue
				}

				$existing = Get-TransportRule -Identity $ruleName -ErrorAction SilentlyContinue
				if ($existing) {
					if ($Force.IsPresent) {
						if ($PSCmdlet.ShouldProcess($ruleName, 'Remove existing transport rule')) {
							try {
								Remove-TransportRule -Identity $ruleName -Confirm:$false -ErrorAction Stop
								Write-Host -ForegroundColor Yellow "[$ruleName] Existing rule removed (-Force)."
							}
							catch {
								Write-Host -ForegroundColor Red "[$ruleName] Failed to remove existing rule: $($_.Exception.Message)"
								$resultsArray.Add([PSCustomObject]@{
									File   = $file.FullName
									Name   = $ruleName
									Status = 'RemoveFailed'
									Error  = $_.Exception.Message
								})
								continue
							}
						}

						if (Get-TransportRule -Identity $ruleName -ErrorAction SilentlyContinue) {
							Write-Host -ForegroundColor Yellow "[$ruleName] Removal not performed (declined or skipped); creation aborted."
							$resultsArray.Add([PSCustomObject]@{
								File   = $file.FullName
								Name   = $ruleName
								Status = 'RemoveDeclined'
								Error  = $null
							})
							continue
						}
					}
					else {
						Write-Host -ForegroundColor Yellow "[$ruleName] Already exists. Use -Force to recreate."
						$resultsArray.Add([PSCustomObject]@{
							File   = $file.FullName
							Name   = $ruleName
							Status = 'AlreadyExists'
							Error  = $null
						})
						continue
					}
				}

				if ($PSCmdlet.ShouldProcess($ruleName, 'New-TransportRule')) {
					try {
						$null = New-TransportRule @ruleParams -ErrorAction Stop
						Write-Host -ForegroundColor Green "[$ruleName] Created."
						$resultsArray.Add([PSCustomObject]@{
							File   = $file.FullName
							Name   = $ruleName
							Status = 'Created'
							Error  = $null
						})
					}
					catch {
						Write-Host -ForegroundColor Red "[$ruleName] $($_.Exception.Message)"
						$resultsArray.Add([PSCustomObject]@{
							File   = $file.FullName
							Name   = $ruleName
							Status = 'Failed'
							Error  = $_.Exception.Message
						})
					}
				}
			}
		}
	}

	end {
		if ($GenerateCmdlets.IsPresent -and $commands.Count -gt 0) {
			$commands | Out-File -FilePath $OutputFile -Encoding UTF8
			Write-Host -ForegroundColor Cyan "Commands generated in file: $OutputFile"
		}

		return $resultsArray
	}
}
