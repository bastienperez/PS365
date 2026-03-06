<#
.SYNOPSIS
    Retrieves and parses event logs from the Microsoft Intune ODJ (Offline Domain Join) Connector Service.

.DESCRIPTION
    This function uses the Get-WinEvent cmdlet to filter and retrieve specific event logs related
    to the Microsoft Intune ODJ (Offline Domain Join) Connector Service. It parses the event messages
    to extract structured information such as ActivityId, DeviceId, DeviceName, and other diagnostic data.
    Intune ODJ Connector Service is responsible for creating Windows computer objects in Active Directory 
    during the device enrollment process via Windows Autopilot.

.PARAMETER DeviceId
    Filters the results to show only events for the specified Device ID (GUID).

.PARAMETER DeviceName
    Filters the results to show only events for the specified Device Name. Supports wildcards.

.EXAMPLE
    Get-IntuneODJConnectorServiceEventLog

    Retrieves and parses the event logs for the Microsoft Intune ODJ Connector Service, returning structured objects with parsed properties.

.EXAMPLE
    Get-IntuneODJConnectorServiceEventLog | Where-Object { $_.DiagnosticCode -ne 0 }

    Retrieves parsed event logs and filters for events with non-zero diagnostic codes (potential errors).

.EXAMPLE
    Get-IntuneODJConnectorServiceEventLog -DeviceId 'e078603c-fe1a-488a-a760-415622ff9f45'

    Retrieves event logs for a specific device ID.

.EXAMPLE
    Get-IntuneODJConnectorServiceEventLog -DeviceName 'DESKTOP-*'

    Retrieves event logs for devices with names matching the wildcard pattern.

.LINK
    https://ps365.clidsys.com/docs/commands/Get-IntuneODJConnectorServiceEventLog

.NOTES
    Must be executed on the server where the Microsoft Intune ODJ Connector Service is installed to access the relevant event logs.
    Requires local administrator privileges to read the event logs.
    Automatically resolves UserID to Active Directory account information (User, gMSA, Computer) when the ActiveDirectory module is available.
#>

function Get-IntuneODJConnectorServiceEventLog {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$DeviceId,

        [Parameter()]
        [string]$DeviceName
    )

    $filterHashTable = @{
        LogName = 'Microsoft-Intune-ODJConnectorService/Operational'
        Id      = 30130, 30140
    }

    try {
        Write-Host -ForegroundColor Cyan 'Querying Windows Event Log...'
        $events = Get-WinEvent -FilterHashtable $filterHashTable -ErrorAction Stop
        Write-Host -ForegroundColor Green "Found $($events.Count) raw events"
    }
    catch {
        Write-Warning "Failed to retrieve event logs: $($_.Exception.Message)"
        return
    }

    # Parse the event messages and return structured objects
    $parsedEvents = [System.Collections.Generic.List[PSCustomObject]]@()
    Write-Host -ForegroundColor Cyan "Parsing $($events.Count) events..."

    foreach ($event in $events) {
        $parsedEvent = [PSCustomObject][ordered]@{
            TimeCreated         = $event.TimeCreated
            DeviceName          = $null
            DeviceId            = $null
            EventId             = $event.Id
            UserId              = $event.UserId
            CreatedBy           = $null
            ElapsedMilliseconds = $null
            DiagnosticCode      = $null
            DiagnosticText      = $null
            EventType           = $null
        }

        # Parse the message content
        $message = $event.Message
        if ($message) {
            # Extract event type from the first line
            $firstLine = ($message -split "`n")[0]
            if ($firstLine -match '^([^:]+):') {
                $parsedEvent.EventType = $matches[1].Trim()
            }

            # Extract structured properties using regex patterns
            if ($message -match 'ElapsedMilliseconds:\s*(\d+)') {
                $parsedEvent.ElapsedMilliseconds = [int]$matches[1]
            }
            if ($message -match 'DeviceId:\s*([a-fA-F0-9-]+)') {
                $parsedEvent.DeviceId = $matches[1]
            }
            if ($message -match 'DeviceName:\s*(.+)') {
                $deviceName = $matches[1].Trim()
                if ($deviceName -ne '') {
                    $parsedEvent.DeviceName = $deviceName
                }
            }
            if ($message -match 'DiagnosticCode:\s*(\d+)') {
                $parsedEvent.DiagnosticCode = [int]$matches[1]
            }
            if ($message -match 'DiagnosticText:\s*(.+)') {
                $parsedEvent.DiagnosticText = $matches[1].Trim()
            }
        }

        # Resolve UserID to Active Directory account if available
        if ($event.UserId -and $event.UserId.Value -ne $null) {
            try {
                $identity = $event.UserId.Value
                
                # Try gMSA first with Get-ADServiceAccount
                try {
                    $adObject = Get-ADServiceAccount -Identity $identity -Properties SamAccountName -ErrorAction Stop
                    $parsedEvent.CreatedBy = $adObject.SamAccountName
                }
                catch {
                    # Fall back to Get-ADObject for other account types
                    try {
                        $adObject = Get-ADObject -Identity $identity -Properties SamAccountName -ErrorAction Stop
                        $parsedEvent.CreatedBy = $adObject.SamAccountName
                    }
                    catch {
                        # Account not found
                    }
                }
            }
            catch {
                # Silently ignore if AD module not available
            }
        }

        $parsedEvents.Add($parsedEvent)
    }

    Write-Host -ForegroundColor Green "Parsed $($parsedEvents.Count) events successfully"

    # Apply filters only if explicitly provided
    if ($PSBoundParameters.ContainsKey('DeviceId') -and $DeviceId) {
        Write-Host -ForegroundColor Cyan "Filtering events for Device ID: $DeviceId"
        $parsedEvents = $parsedEvents | Where-Object { $_.DeviceId -eq $DeviceId }
        Write-Host -ForegroundColor Green "Found $($parsedEvents.Count) matching events"
    }

    if ($PSBoundParameters.ContainsKey('DeviceName') -and $DeviceName) {
        Write-Host -ForegroundColor Cyan "Filtering events for Device Name: $DeviceName"
        $parsedEvents = $parsedEvents | Where-Object { $_.DeviceName -like $DeviceName }
        Write-Host -ForegroundColor Green "Found $($parsedEvents.Count) matching events"
    }

    return $parsedEvents
}