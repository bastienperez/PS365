<#
	.SYNOPSIS
	Retrieves all Exchange mailboxes for a specified domain.

	.DESCRIPTION
	This function retrieves all Exchange mailboxes that belong to a specified domain.
	It filters the mailboxes based on their *primarySMTP* address to ensure they match the provided domain.

	.PARAMETER Domain
	The domain for which to retrieve the mailboxes. This should be a string representing the domain (e.g., "example.com").

	.EXAMPLE
	Get-ExMailboxByDomain -Domain "example.com"

	Returns all Exchange mailboxes associated with the domain "example.com".
#>
function Get-ExMailboxByDomain {
	param (
		[Parameter(Mandatory = $true, Position = 0)]    
		[string]$Domain,
		[Parameter(Mandatory = $false)]
		[ValidateSet('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'LinkedMailbox', 'SchedulingMailbox')]
		[string]$RecipientTypeDetails
	)

	$mailboxesArray = Get-ExoMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$Domain'" -Properties WhenCreated, WhenChanged | Where-Object { $_.PrimarySmtpAddress -like "*@$Domain" }
    
	if ($RecipientTypeDetails) {
		$mailboxesArray = $mailboxesArray | Where-Object { $RecipientTypeDetails -contains $_.RecipientTypeDetails }
	}

	return $mailboxesArray
}