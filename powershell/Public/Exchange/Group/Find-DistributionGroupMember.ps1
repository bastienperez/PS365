<#
    .SYNOPSIS
    Find members of distribution groups based on email address or domain.

    .DESCRIPTION
    This function retrieves members of all distribution groups in Exchange Online that match specified criteria such as email addresses or domains.

    .PARAMETER FilterByDomain
    Specify a domain to filter distribution group members by their email addresses.

    .PARAMETER FilterByEmailAddresses
    Provide specific email addresses to filter distribution group members.

    .PARAMETER FilterByExternalDomains
    If set to true, filters members whose email addresses belong to external domains not managed in the Exchange Online environment (based on accepted domains).

    .PARAMETER FilterRecipientTypeDetails
    Specify a recipient type detail to further filter the members (e.g., MailContact, MailUser).

    .EXAMPLE
    Find-DistributionGroupMember -FilterByDomain "contoso.com"

    Searches all distribution groups for members with email addresses containing "contoso.com" and exports the results to a CSV file.
    
    .EXAMPLE
    Find-DistributionGroupMember -FilterByEmailAddresses "user@contoso.com"

    Searches all distribution groups for members with the email address "user@contoso.com" and exports the results to a CSV file.   

    .EXAMPLE
    Find-DistributionGroupMember -FilterByExternalDomains -FilterRecipientTypeDetails "MailContact"

    Searches all distribution groups for members whose email addresses belong to external domains and are of the type "MailContact", then exports the results to a CSV file.

    .LINK
    https://ps365.clidsys.com/docs/commands/Find-DistributionGroupMember
#>
function Find-DistributionGroupMember {
    param
    (
        [Parameter(Mandatory = $false)]
        [string]$FilterByDomain,
        [Parameter(Mandatory = $false)]
        [ValidateNotNull()]
        [string[]]$FilterByEmailAddresses,
        [Parameter(Mandatory = $false)]
        [switch]$FilterByExternalDomains,
        [Parameter(Mandatory = $false)]
        [string]$FilterRecipientTypeDetails
    )

    # Usage example : 
    #	Connect-ExchangeOnline
    #	Find-DistributionGroupMember Export-Csv "AllContactsMember.csv" -NoTypeInformation -Encoding utf8 -Delimiter ';'
    # Ignore GroupMailbox type (Office 365 groups-unified groups)
    $distributionGroups = Get-DistributionGroup -ResultSize unlimited -RecipientTypeDetails MailUniversalDistributionGroup
	
    [System.Collections.Generic.List[PSObject]]$foundMembers = @()

    $i = 0

    if (-not ($FilterByDomain -or $FilterByEmailAddresses -or $FilterByExternalDomains.IsPresent)) {
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
            $members = Get-DistributionGroupMember $dg.PrimarySMTPAddress | Where-Object { $_.EmailAddresses -like "*$FilterByDomain*" } | Select-Object @{Name = 'DGName'; expression = { $DG.Identity } }, @{Name = 'DGPrimaryStmpAddress'; expression = { $DG.PrimarySmtpAddress } }, Name, PrimarySMTPAddress, RecipientTypeDetails
        }
        elseif ($FilterByExternalDomains.IsPresent) {
            # Get domain managed in this Exchange Online
            $acceptedDomains = (Get-AcceptedDomain).DomainName

            # Find messaging objects with a domain outside the domain managed in this Exchange Online
            $members = Get-DistributionGroupMember $dg.PrimarySMTPAddress | Where-Object { $acceptedDomains -notcontains $_.PrimarySmtpAddress.split('@')[1] } | Select-Object @{Name = 'DGName'; expression = { $DG.Identity } }, @{Name = 'DGPrimaryStmpAddress'; expression = { $DG.PrimarySmtpAddress } }, Name, PrimarySMTPAddress, RecipientTypeDetails
            #$members = Get-Recipient | Where-Object { $acceptedDomains -notcontains $_.PrimarySmtpAddress.split('@')[1] }
        }

		
        if ($FilterRecipientTypeDetails) {
            $members = $members | Where-Object { $_.RecipientTypeDetails -eq $FilterRecipientTypeDetails }
        }
		
        #$foundMembers.Add($members)
    
        foreach ($member in $members) {
            Write-Host "DG: $($dg.Name) $($dg.PrimarySmtpAddress) - member: $($member.PrimarySmtpAddress) $($member.RecipientTypeDetails)" -ForegroundColor Yellow
			
            $object = [PSCustomObject][ordered]@{
                DGName               = $dg.Name
                DGPrimaryStmpAddress = $dg.PrimarySmtpAddress
                Name                 = $member.Name
                PrimarySmtpAddress   = $member.PrimarySmtpAddress
                RecipientTypeDetails = $member.RecipientTypeDetails
            }

            $foundMembers.Add($object)
        }
    }

    return $foundMembers
}