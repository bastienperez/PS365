function Set-ExMailboxMaxSize {
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
		[Parameter(Mandatory = $false)]
		[switch]$GenerateCmdlets,

		[Parameter(Mandatory = $false)]
		[string]$OutputFile = "$(Get-Date -Format 'yyyy-MM-dd_HHmmss')-SetExMailboxMaxSize-Commands.ps1"
	)

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
		Write-Warning 'NotImplementedYet'
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

		$cmdletString = 'Set-Mailbox -Identity'
		# param are -MaxReceiveSize 150MB -MaxSendSize 150MB -ErrorAction Stop
		$cmdParamstring = '-MaxReceiveSize 150MB -MaxSendSize 150MB -ErrorAction Stop'
		$fullCommand = "$cmdletString $($cmdParams.Identity) $cmdParamstring"
		if ($GenerateCmdlets) {
			$Commands += $fullCommand
		}

		if (-not $GenerateCmdlets -and $PSCmdlet.ShouldProcess($mbx.PrimarySmtpAddress, 'Set regional configuration')) {
			Write-Host -ForegroundColor Cyan "$($mbx.PrimarySmtpAddress) - Setting configuration..."
			try {
				Set-Mailbox -Identity $mbx.PrimarySmtpAddress -MaxReceiveSize 150MB -MaxSendSize 150MB -ErrorAction Stop
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