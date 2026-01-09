<#
	.SYNOPSIS
	Retrieves all Exchange mailboxes for a specified domain.

	.DESCRIPTION
	This function retrieves all Exchange mailboxes that belong to a specified domain.
	It filters the mailboxes based on their *primarySMTP* address to ensure they match the provided domain.

	.PARAMETER Domain
	The domain for which to retrieve the mailboxes. This should be a string representing the domain (e.g., "example.com").

	.PARAMETER RecipientTypeDetails
	Filters the mailboxes by their recipient type details. Acceptable values include:
	- UserMailbox
	- SharedMailbox
	- RoomMailbox
	- EquipmentMailbox
	- LinkedMailbox
	- SchedulingMailbox

	.PARAMETER ExportToExcel
	If specified, exports the retrieved mailbox information to an Excel file in the user's profile directory.
	
	.EXAMPLE
	Get-ExMailboxByDomain -Domain "example.com"

	Returns all Exchange mailboxes associated with the domain "example.com".

	.LINK
	https://ps365.clidsys.com/docs/commands/Get-ExMailboxByDomain
#>
function Get-ExMailboxByDomain {
	param (
		[Parameter(Mandatory = $true, Position = 0)]    
		[string]$Domain,

		[Parameter(Mandatory = $false)]
		[ValidateSet('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox', 'LinkedMailbox', 'SchedulingMailbox')]
		[string]$RecipientTypeDetails,

		[Parameter(Mandatory = $false)]
		[switch]$ExportToExcel
	)

	$mailboxesArray = Get-EXOMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$Domain'" -Properties WhenCreated, WhenChanged | Where-Object { $_.PrimarySmtpAddress -like "*@$Domain" }
    
	if ($RecipientTypeDetails) {
		$mailboxesArray = $mailboxesArray | Where-Object { $RecipientTypeDetails -contains $_.RecipientTypeDetails }
	}

	if ($ExportToExcel.IsPresent) {
		$now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
		$ExcelFilePath = "$($env:userprofile)\$now-MailboxesByDomain_Report.xlsx"
		Write-Host -ForegroundColor Cyan "Exporting mailboxes information to Excel file: $ExcelFilePath"
		$mailboxesArray | Select-Object PrimarySmtpAddress, DisplayName, RecipientTypeDetails, WhenCreated, WhenChanged | Export-Excel -Path $ExcelFilePath -AutoSize -AutoFilter -WorksheetName 'MailboxesByDomain'
	}
	else {
		return $mailboxesArray
	}
}