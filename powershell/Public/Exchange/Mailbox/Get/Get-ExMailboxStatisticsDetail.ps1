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
	Mailbox identity (email address, username, etc.)

	.PARAMETER IncludeFolderDetails
	Include folder details in the results

	.EXAMPLE
	Get-ExMailboxStatisticsDetail -Identity "user@domain.com"

	Gets detailed statistics for the specified mailbox.

	.EXAMPLE
	Get-ExMailboxStatisticsDetail -Identity "user@domain.com" -IncludeFolderDetails

	Gets statistics with folder details included.

	.NOTES
	Author: Bastien Perez
	Version: 1.0.0
#>

function Get-ExMailboxStatisticsDetail {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[string]$Identity,
        
		[Parameter(Mandatory = $false)]
		[switch]$IncludeFolderDetails
	)

	begin {
		# Check if Exchange Online module is available
		if (-not (Get-Command 'Get-MailboxStatistics' -ErrorAction SilentlyContinue)) {
			throw 'Exchange Online PowerShell module is not available. Please connect using Connect-ExchangeOnline.'
		}
        
		Write-Verbose "Starting statistics retrieval for mailbox $Identity"
	}

	process {
		try {
			# Get general mailbox statistics
			Write-Verbose 'Retrieving general statistics...'
			$mailboxStats = Get-MailboxStatistics -Identity $Identity -ErrorAction Stop
            
			# Get folder statistics
			Write-Verbose "Retrieving folder statistic for mailbox $Identity..."
			$folderStats = Get-MailboxFolderStatistics -Identity $Identity -ErrorAction Stop
            
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
				Identity             = $Identity
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
			}
            
			# Add folder details if requested
			if ($IncludeFolderDetails) {
				$folderDetails = $folderStats | Select-Object FolderName, FolderType, ItemsInFolder, FolderSize | Sort-Object FolderType, FolderName
				$result | Add-Member -MemberType NoteProperty -Name 'FolderDetails' -Value $folderDetails
			}
            
			Write-Output $result
            
		}
		catch {
			Write-Error "Error retrieving statistics for mailbox $Identity : $($_.Exception.Message)"
		}
	}

	end {
		Write-Verbose 'Statistics retrieval completed'
	}
}