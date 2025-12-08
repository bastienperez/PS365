<#
	.SYNOPSIS
	Set maximum mailbox size for specified mailboxes.

	.DESCRIPTION
	This function sets the maximum mailbox size (MaxReceiveSize and MaxSendSize) for Exchange Online mailboxes.
	It can target mailboxes by identity, by domain, from a CSV file, or all mailboxes in the organization.
	
	.PARAMETER Identity
	The identity of the mailbox(es) to set the maximum size for. This can be an array of email addresses, usernames, or display names.
	
	.PARAMETER ByDomain
	Filter mailboxes by domain name. All mailboxes with a primary SMTP address in this domain will be processed.
	
	.PARAMETER FromCSV
	The path to a CSV file containing mailbox identities to process.
	
	.PARAMETER AllMailboxes
	If specified, all mailboxes in the organization will be processed.
	
	.PARAMETER GenerateCmdlets
	If specified, the function will generate the Set-Mailbox cmdlets and save them to a file instead of executing them.
	
	.PARAMETER OutputFile
	The path to the output file where generated cmdlets will be saved. Default is a timestamped file in the current directory.
	
	.EXAMPLE
	Set-ExMailboxMaxSize -Identity "user@example.com" -MaxReceiveSize 150MB -MaxSendSize 150MB
	Sets the maximum mailbox size for the specified mailbox.
	
	.EXAMPLE
	Set-ExMailboxMaxSize -ByDomain "example.com"
	Sets the maximum mailbox size for all mailboxes in the specified domain.
	
	.EXAMPLE
	Set-ExMailboxMaxSize -AllMailboxes
	Sets the maximum mailbox size for all mailboxes in the organization.
	
	.EXAMPLE
	Set-ExMailboxMaxSize -FromCSV "C:\path\to\file.csv"
	Sets the maximum mailbox size for mailboxes listed in the specified CSV file.
	
	.EXAMPLE
	Set-ExMailboxMaxSize -ByDomain "example.com" -GenerateCmdlets -OutputFile "C:\path\to\commands.ps1"
	Generates the Set-Mailbox cmdlets for all mailboxes in the specified domain and saves them to the specified file without executing them.
#>
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

	$commands = @()
	$totalCount = @($Mailboxes).Count
	$currentCount = 0

	foreach ($mbx in $Mailboxes) {
		$currentCount++
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
			Write-Host -ForegroundColor Cyan "[$CurrentCount/$TotalCount] $($mbx.PrimarySmtpAddress) - Setting configuration..."
			try {
				Set-Mailbox -Identity $mbx.PrimarySmtpAddress -MaxReceiveSize 150MB -MaxSendSize 150MB -ErrorAction Stop
				Write-Host "[$CurrentCount/$TotalCount] $($mbx.PrimarySmtpAddress) - Configuration applied successfully." -ForegroundColor Green
			}
			catch {
				Write-Host "[$CurrentCount/$TotalCount] $($mbx.PrimarySmtpAddress) - $($_.Exception.Message)" -ForegroundColor Red
			}
		}
	}
	
	if ($GenerateCmdlets -and $Commands.Count -gt 0) {
		$Commands | Out-File -FilePath $OutputFile -Encoding UTF8
		Write-Host "Commands generated in file: $OutputFile" -ForegroundColor Cyan
	}
}