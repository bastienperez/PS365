<#
.SYNOPSIS
    Analyzes Exchange Server SMTP receive logs to identify unique sender IP addresses and their connection frequency.

.DESCRIPTION
    This function parses Exchange Server SMTP receive protocol logs to extract and analyze remote endpoint IP addresses.
    It processes logs from both Hub and Frontend transport roles, providing insights into which external systems
    are connecting to send email through your Exchange infrastructure.

    The function is particularly useful for:
    - Exchange server decommissioning scenarios to identify remaining SMTP dependencies
    - Security auditing to identify unusual or suspicious connection patterns  
    - Capacity planning and connection monitoring
    - Identifying applications that may need configuration updates during migrations

    Log locations by Exchange version:
    - Exchange 2010: C:\Program Files\Microsoft\Exchange Server\V14\TransportRoles\Logs\ProtocolLog\SmtpReceive
    - Exchange 2013/2016/2019: C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive

.EXAMPLE
    Get-SMTPLog
    
    Analyzes SMTP receive logs from all transport services and returns a summary of unique sender IPs with DNS resolution and connection counts.

.NOTES
    Inspired by SMTP-Review 1.0 by r-milne
    Source: https://github.com/r-milne/Check-Exchange-SMTP-Logs-To-Get-Unique-Sender-IP-Addresses/blob/master/SMTP-Review-1.0.ps1

    Requirements:
    - SMTP receive logging must be enabled on Exchange servers
    - Appropriate permissions to access log files on Exchange servers
    - Network connectivity to Exchange servers for UNC path access

    Limitations:
    - Only analyzes logs retained by Exchange (default 30 days)
    - May not capture infrequent senders (quarterly/yearly processes)
    - Requires manual verification of application dependencies

.CHANGELOG

[1.2.0] - 2025-08-14
# Modified
- Updated synopsis, documentation and array handling

[1.1.0] - 2020-06-01
# Modified by B.Perez: Performance improvements

[1.0.0] - 2019-01-25
# Initial release
#>

function Get-SMTPLog {
	[CmdletBinding()]
	param()

	# Declare an array to hold the output
	[System.Collections.Generic.List[PSCustomObject]]$smtpLogsArray = @()
    
	# Look for UNC path of SmtpReceive. You also can directly put
	$TransportServices = Get-TransportService

	foreach ($ts in $transportServices) {
		$serverName = $ts.Name
		#$logFilePath = "C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive"
		# we can't get frontend path (or I not found it), so we get hub path and we assume the frontend path is in the same folder so we replace. Examples:
		#	hub : C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive
		#	frontend : C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive
		$logFilePath = "$($ts.ReceiveProtocolLogPath)"
	
		$logFilePath = $logFilePath.replace('\Hub\', '\FrontEnd\')
		# convert to UNC
		$logFilePath = $logFilePath -replace '^(.):', "\\$serverName\`$1$"
	
		# Change to suit the particular input location etc.  
		$logFiles = Get-Item "$logFilePath\*.log"
	
		# Use an array around the $logFiles to work out the count.  Needed for later to display the progress bar
		$Count = @($logfiles).count
	
		$i = 0
	
		foreach ($log in $logFiles) {
			$i++
			Write-Host "Processing Log File on $serverName - $log [$i of $Count]" -ForegroundColor Magenta
		
			# Skip the first 5 lines of the file as they are headers we do not want to review
			#$FileContent = Get-Content $Log | Select-Object -Skip 5
			try {
				$fileContent = [System.IO.File]::ReadAllLines($log) | Select-Object -Skip 5
			}
			catch {
				# if file is currently in used, ReadAllLines can't read it, so Get-Content is used (much slower than ReadAllLines but only the current file of the day is locked)
				$fileContent = Get-Content -Path $log | Select-Object -Skip 5
			}
	
			# Retrieve Element Number 5 from the log - this will be the field named "remote-endpoint" 
			# Note that the array is zero based 

			foreach ($line in $fileContent) {
				$socket = $line.Split(',')[5]
	
				# This will return data in the form of the socket used - IP and ephemeral TCP port
				# 185.234.217.220:61061
				$IP = $socket.Split(':')[0]
	
				$smtpLogsArray.Add($IP)
			} 
		} 
	}

	$smtpLogsArray | Group-Object | Select-Object Name, @{Name = 'DNSName'; Expression = { ([System.Net.Dns]::GetHostEntry($_.Name).HostName) } }, Count | Sort-Object Count -Descending
}