<#
	.SYNOPSIS
	Gets detailed mailbox statistics for Exchange Online including item count by type.

	.DESCRIPTION
	This function retrieves comprehensive statistics for an Exchange Online mailbox, including:
	- General statistics (size, total item count, etc.)
	- Item count by folder (Inbox, Sent Items, etc.)
	- Number of contacts
	- Number of calendar items
	- Number of tasks
	- Number of notes

	.PARAMETER Identity
	Mailbox identity (email address, username, etc.). If omitted, statistics are retrieved for all Exchange Online mailboxes.

	.EXAMPLE
	Get-ExMailboxStatisticsInfo

	Gets detailed statistics for all Exchange Online mailboxes.

	.PARAMETER IncludeFolderDetails
	Include folder details in the results

	.EXAMPLE
	Get-ExMailboxStatisticsInfo -Identity "user@domain.com"

	Gets detailed statistics for the specified mailbox.

	.EXAMPLE
	Get-ExMailboxStatisticsInfo -Identity "user@domain.com" -IncludeFolderDetails

	Gets statistics with folder details included.

	.PARAMETER ExportToExcel
	If specified, exports the results to an Excel file in the user's profile directory.

	.EXAMPLE
	Get-ExMailboxStatisticsInfo -ExportToExcel

	Exports results to an Excel file.

	.LINK
	https://ps365.clidsys.com/docs/commands/Get-ExMailboxStatisticsInfo

	.NOTES
	Author: Bastien Perez
	Version: 1.0.0
#>

function Get-ExMailboxStatisticsInfo {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $false, position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[ValidateNotNullOrEmpty()] 
		[string]$Identity,
        
		[Parameter(Mandatory = $false)]
		[switch]$IncludeFolderDetails,

		[Parameter(Mandatory = $false)]
		[switch]$ExportToExcel,

		[Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
		[string]$ExportPath
	)

	begin {
		# Check if the commands used for statistics collection are available.
		foreach ($requiredCommand in @('Get-MailboxStatistics', 'Get-MailboxFolderStatistics')) {
			if (-not (Get-Command $requiredCommand -ErrorAction SilentlyContinue)) {
				throw "Required Exchange command '$requiredCommand' is not available. Please connect using Connect-ExchangeOnline."
			}
		}

		$resultsArray = [System.Collections.Generic.List[PSCustomObject]]::new()

		Write-Verbose "Starting statistics retrieval for mailbox $Identity"
	}

	process {
		if ([string]::IsNullOrWhiteSpace($Identity)) {
			if (-not (Get-Command 'Get-EXOMailbox' -ErrorAction SilentlyContinue)) {
				throw "Required Exchange command 'Get-EXOMailbox' is not available to enumerate mailboxes."
			}

			Write-Host -ForegroundColor Cyan 'Retrieving all Exchange Online mailboxes'
			$requestedIdentities = @(Get-EXOMailbox -ResultSize Unlimited | ForEach-Object {
					$mailbox = $_
					$resolvedIdentity = @(
						$mailbox.UserPrincipalName
						$mailbox.PrimarySmtpAddress
						$mailbox.ExternalDirectoryObjectId
						$mailbox.Identity
					) | Where-Object { -not [string]::IsNullOrWhiteSpace("$_") } | Select-Object -First 1

					if ($null -ne $resolvedIdentity) {
						[string]$resolvedIdentity
					}
					else {
						Write-Warning "Skipping a mailbox because none of UserPrincipalName, PrimarySmtpAddress, ExternalDirectoryObjectId or Identity is populated. DisplayName: '$($mailbox.DisplayName)'."
					}
				})
		}
		else {
			$requestedIdentities = @($Identity)
		}

		foreach ($mailboxIdentity in $requestedIdentities) {
		try {
			# Get general mailbox statistics
			Write-Verbose "Retrieving general statistics for mailbox $mailboxIdentity..."
			$mailboxStats = Get-MailboxStatistics -Identity $mailboxIdentity -ErrorAction Stop
            
			# Get folder statistics
			Write-Verbose "Retrieving folder statistic for mailbox $mailboxIdentity..."
			$folderStats = Get-MailboxFolderStatistics -Identity $mailboxIdentity -ErrorAction Stop
            
			# Calculate statistics by item type
			$inboxItems = ($folderStats | Where-Object { $_.FolderType -eq 'Inbox' } | Measure-Object ItemsInFolder -Sum).Sum
			$sentItems = ($folderStats | Where-Object { $_.FolderType -eq 'SentItems' } | Measure-Object ItemsInFolder -Sum).Sum
			$deletedItems = ($folderStats | Where-Object { $_.FolderType -eq 'DeletedItems' } | Measure-Object ItemsInFolder -Sum).Sum
			$drafts = ($folderStats | Where-Object { $_.FolderType -eq 'Drafts' } | Measure-Object ItemsInFolder -Sum).Sum
			$junkEmail = ($folderStats | Where-Object { $_.FolderType -eq 'JunkEmail' } | Measure-Object ItemsInFolder -Sum).Sum
			$outbox = ($folderStats | Where-Object { $_.FolderType -eq 'Outbox' } | Measure-Object ItemsInFolder -Sum).Sum
            
			# Specific items
			$contacts = ($folderStats | Where-Object { $_.FolderType -eq 'Contacts' } | Measure-Object ItemsInFolder -Sum).Sum
			$calendar = ($folderStats | Where-Object { $_.FolderType -eq 'Calendar' } | Measure-Object ItemsInFolder -Sum).Sum
			$tasks = ($folderStats | Where-Object { $_.FolderType -eq 'Tasks' } | Measure-Object ItemsInFolder -Sum).Sum
			$notes = ($folderStats | Where-Object { $_.FolderType -eq 'Notes' } | Measure-Object ItemsInFolder -Sum).Sum
            
			# Custom/other folders
			$otherFolders = $folderStats | Where-Object { 
				$_.FolderType -notin @('Inbox', 'SentItems', 'DeletedItems', 'Drafts', 'JunkEmail', 'Outbox', 'Contacts', 'Calendar', 'Tasks', 'Notes', 'Root') 
			}
			$otherItems = ($otherFolders | Measure-Object ItemsInFolder -Sum).Sum
            
			# Create result object
			$result = [PSCustomObject]@{
				Identity             = $mailboxIdentity
				DisplayName          = $mailboxStats.DisplayName
				TotalItemSize        = $mailboxStats.TotalItemSize
				TotalDeletedItemSize = $mailboxStats.TotalDeletedItemSize
				ItemCount            = $mailboxStats.ItemCount
				DeletedItemCount     = $mailboxStats.DeletedItemCount
				LastLogonTime        = $mailboxStats.LastLogonTime
				LastUserActionTime   = $mailboxStats.LastUserActionTime
                
				# Detail by folder type
				InboxItems           = if ($null -eq $inboxItems) { 0 } else { $inboxItems }
				SentItems            = if ($null -eq $sentItems) { 0 } else { $sentItems }
				DeletedItems         = if ($null -eq $deletedItems) { 0 } else { $deletedItems }
				DraftsItems          = if ($null -eq $drafts) { 0 } else { $drafts }
				JunkEmailItems       = if ($null -eq $junkEmail) { 0 } else { $junkEmail }
				OutboxItems          = if ($null -eq $outbox) { 0 } else { $outbox }
                
				# Specific requested items
				ContactsCount        = if ($null -eq $contacts) { 0 } else { $contacts }
				CalendarItemsCount   = if ($null -eq $calendar) { 0 } else { $calendar }
				TasksCount           = if ($null -eq $tasks) { 0 } else { $tasks }
				NotesCount           = if ($null -eq $notes) { 0 } else { $notes }
                
				# Other folders
				OtherFoldersItems    = if ($null -eq $otherItems) { 0 } else { $otherItems }
                
				# Additional information
				DatabaseName         = $mailboxStats.Database
				ServerName           = $mailboxStats.ServerName
				MailboxGuid          = $mailboxStats.MailboxGuid
				MailboxWhenCreated   = $mailboxStats.WhenCreated
				MailboxWhenModified  = $mailboxStats.WhenChanged
			}
            
			# Add folder details if requested
			if ($IncludeFolderDetails) {
				$folderDetails = $folderStats | Select-Object FolderName, FolderType, ItemsInFolder, FolderSize | Sort-Object FolderType, FolderName
				$result | Add-Member -MemberType NoteProperty -Name 'FolderDetails' -Value $folderDetails
			}
            
			$resultsArray.Add($result)
            
		}
		catch {
			Write-Error "Error retrieving statistics for mailbox $mailboxIdentity : $($_.Exception.Message)"
		}
		}
	}

	end {
		Write-Verbose 'Statistics retrieval completed'

		if ($ExportToExcel.IsPresent) {
			$now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
			$excelFilePath = "$(if ($ExportPath) { $ExportPath } else { $env:userprofile })\$now-ExMailboxStatisticsInfo.xlsx"
			Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
			$resultsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExMailboxStatisticsInfo' -TableStyle Light9
			Write-Host -ForegroundColor Green 'Export completed successfully!'
		}
		else {
			return $resultsArray
		}
	}
}
