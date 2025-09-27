function Search-DistributionGroupMembers {
	Param
	(
		[Parameter(Mandatory = $false)]
		[string]$FilterByDomain,
		[Parameter(Mandatory = $false)]
		[ValidateNotNull()]
		[string[]]$FilterByEmailAddresses,
		[Parameter(Mandatory = $false)]
		[boolean]$FilterByExternalDomains,
		[Parameter(Mandatory = $false)]
		[string]$FilterRecipientTypeDetails
	)

	$distributionGroups = Get-DistributionGroup -ResultSize unlimited -RecipientTypeDetails MailUniversalDistributionGroup
	
	[System.Collections.Generic.List[PSCustomObject]]$foundMembers = @()

	$i = 0

	if (-not ($FilterByDomain -or $FilterByEmailAddresses -or $FilterByExternalDomains)) {
		Write-Warning 'Please fill at least one parameter FilterByDomain or FilterByEmailAddresses'
		return
	}

	foreach ($dg in $distributionGroups) {
		$i++
		$members = @()
		Write-Host "Processing $($dg.Name) ($($dg.PrimarySmtpAddress)) [$i/$($distributionGroups.count)]" -ForegroundColor Cyan
	
		if ($FilterByEmailAddresses) {
		
			$members = @()

			foreach ($emailAddress in $FilterByEmailAddresses) {
				$members += Get-DistributionGroupMember $dg.PrimarySMTPAddress | Where-Object { $_.PrimarySmtpAddress -eq $emailAddress } | Select-Object @{Name = 'DGName'; expression = { $DG.Identity } }, @{Name = 'DGPrimaryStmpAddress'; expression = { $DG.PrimarySmtpAddress } }, Name, PrimarySMTPAddress, RecipientTypeDetails	
			}
		}
		elseif ($FilterByDomain) {
			$members = Get-DistributionGroupMember $dg.PrimarySMTPAddress | Where-Object { $_.EmailAddresses -like "*@$FilterByDomain*" } | Select-Object @{Name = 'DGName'; expression = { $DG.Identity } }, @{Name = 'DGPrimaryStmpAddress'; expression = { $DG.PrimarySmtpAddress } }, Name, PrimarySMTPAddress, RecipientTypeDetails
		}
		elseif ($FilterByExternalDomains) {
			# Get domain managed in this Exchange Online
			$acceptedDomains = (Get-AcceptedDomain).DomainName

			# Find messaging objects with a domain outside the domain managed in this Exchange Online
			$members = Get-DistributionGroupMember $dg.PrimarySMTPAddress | Where-Object { $acceptedDomains -notcontains $_.PrimarySmtpAddress.split('@')[1] } | Select-Object @{Name = 'DGName'; expression = { $DG.Identity } }, @{Name = 'DGPrimaryStmpAddress'; expression = { $DG.PrimarySmtpAddress } }, Name, PrimarySMTPAddress, RecipientTypeDetails
			#$members = Get-Recipient | Where-Object { $acceptedDomains -notcontains $_.PrimarySmtpAddress.split('@')[1] }
		}

		
		if ($FilterRecipientTypeDetails) {
			$members = $members | Where-Object { $_.RecipientTypeDetails -eq $FilterRecipientTypeDetails }
		}
    
		foreach ($member in $members) {
			Write-Host "DG: $($dg.Name) $($dg.PrimarySmtpAddress) - member: $($member.PrimarySmtpAddress) $($member.RecipientTypeDetails)" -ForegroundColor Yellow
			
			$object = [PSCustomObject][ordered]@{
				DGName               = $dg.Name
				DGPrimarySmtpAddress = $dg.PrimarySmtpAddress
				Name                 = $member.Name
				PrimarySmtpAddress   = $member.PrimarySmtpAddress
				RecipientTypeDetails = $member.RecipientTypeDetails
			}

			$foundMembers.Add($object)
		}
	}

	return $foundMembers
}