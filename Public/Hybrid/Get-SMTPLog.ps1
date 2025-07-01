<# 
Source : SMTP-Review1.0
.SYNOPSIS
    Script is intended to help determine servers that are using an Exchange server to connect and send email.
    
    This is especially pertinent in a decomission scenario, where the logs are to be checked to ensure that all SMTP traffic has been moved to the correct endpoint.



.DESCRIPTION

    Logs on an Exchange 2010 servers are here by default.
    C:\Program Files\Microsoft\Exchange Server\V14\TransportRoles\Logs\ProtocolLog\SmtpReceive

    Note that the script can be easily modified for other versions, or to look at the SMTPSend logs instead.  

	
    An empty array is declared that will be used to hold the data gathered during each iteration. 
    This allows for the additional information to be easily added on, and then either echo it to the screen or export to a CSV file



	# Sample Exchange 2010 SMTP Receive log

	#	#Software: Microsoft Exchange Server
	#	#Version: 14.0.0.0
	#	#Log-type: SMTP Receive Protocol Log
	#	#Date: 2019-01-25T00:03:58.478Z
	#	#Fields: date-time,connector-id,session-id,sequence-number,local-endpoint,remote-endpoint,event,data,context
	#	2019-01-25T00:03:58.478Z,TAIL-EXCH-1\Internet Mail,08D675E58CA1DA38,0,10.0.0.6:25,185.234.217.220:61061,+,,
	#	2019-01-25T00:03:58.494Z,TAIL-EXCH-1\Internet Mail,08D675E58CA1DA38,1,10.0.0.6:25,185.234.217.220:61061,*,SMTPSubmit SMTPAcceptAnySender SMTPAcceptAuthoritativeDomainSender AcceptRoutingHeaders,Set Session Permissions


.ASSUMPTIONS
    Logging was enabled to generate the required log files.
    Logging was enabled previously, and time was allowed to colled the data in the logs

    Not all activity will be present on a given server.  Will have to check multiple in most deployments.
    Not all activity will be present in the logs.  For example, Exchange maintains 30 days of logs by default.  This will not catch connections for processes which
    send email once a quarter or once a fiscal year.

    Assuption is that something will likely be negatively impacted.  Application ownsers should have been told to update their config, so we can say "unlucky" to them...

	You can live with the Write-Host cmdlets :) 

	You can add your error handling if you need it.  

.VERSION
  
	1.0  25-1-2019 -- Initial script released to the scripting gallery
    1.1	 06-2020 - Modified by B.Perez : speed up the script
#>

# Declare an empty array to hold the output
$output = New-Object 'System.Collections.Generic.List[String]'
# Look for UNC path of SmtpReceive. You also can directly put
$transportServices = Get-TransportService

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
	
		ForEach ($line IN $fileContent) {
			$Socket = $line.split(",")[5]
	
			# This will return data in the form of the socket used - IP and ephemeral TCP port
			# 185.234.217.220:61061
			# Split this so that only the IP is retained, burn the rest!
			$IP = $Socket.Split(":")[0]
	
			# Append  current results to final output
			$null = $Output.Add($IP)
		} 
	} 
}

$Output | Group-Object | Select-Object Name, @{Name = 'DNSName'; Expression = { ([System.Net.Dns]::GetHostEntry($_.Name).HostName) } }, Count | Sort-Object Count -Descending