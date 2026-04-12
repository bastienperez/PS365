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

    if (-not ($FilterByDomain -or $FilterByEmailAddresses -or $FilterByExternalDomains.IsPresent)) {
        Write-Warning 'Please fill at least one parameter FilterByDomain or FilterByEmailAddresses'
        return
    }

    # Usage example :
    #	Connect-ExchangeOnline
    #	Find-DistributionGroupMember Export-Csv "AllContactsMember.csv" -NoTypeInformation -Encoding utf8 -Delimiter ';'
    # Ignore GroupMailbox type (Office 365 groups-unified groups)
    $distributionGroups = Get-DistributionGroup -ResultSize unlimited -RecipientTypeDetails MailUniversalDistributionGroup

    [System.Collections.Generic.List[PSObject]]$foundMembers = @()

    # Fetch accepted domains once outside the loop (was a per-DG remote Exchange round-trip, 30-100 ms each).
    # Use a HashSet for O(1) -notcontains checks.
    $acceptedDomainSet = $null
    if ($FilterByExternalDomains.IsPresent) {
        $acceptedDomainSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($d in (Get-AcceptedDomain).DomainName) {
            [void]$acceptedDomainSet.Add($d)
        }
    }

    # HashSet of requested email addresses (case-insensitive) for O(1) membership test
    $emailAddressSet = $null
    if ($FilterByEmailAddresses) {
        $emailAddressSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($e in $FilterByEmailAddresses) {
            [void]$emailAddressSet.Add($e)
        }
    }

    $i = 0
    foreach ($dg in $distributionGroups) {
        $i++
        $members = @()
        Write-Host "Processing $($dg.Name) ($($dg.PrimarySmtpAddress)) [$i/$($distributionGroups.count)]" -ForegroundColor Cyan

        # Single Get-DistributionGroupMember call per DG (previously: one call per requested email).
        $dgMembers = Get-DistributionGroupMember $dg.PrimarySMTPAddress

        if ($FilterByEmailAddresses) {
            $members = $dgMembers | Where-Object { $emailAddressSet.Contains([string]$_.PrimarySmtpAddress) }
        }
        elseif ($FilterByDomain) {
            $members = $dgMembers | Where-Object { $_.EmailAddresses -like "*$FilterByDomain*" }
        }
        elseif ($FilterByExternalDomains.IsPresent) {
            # Find messaging objects with a domain outside the domain managed in this Exchange Online
            $members = $dgMembers | Where-Object { -not $acceptedDomainSet.Contains([string]$_.PrimarySmtpAddress.Split('@')[1]) }
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