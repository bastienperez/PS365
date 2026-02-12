<#
.SYNOPSIS
    Retrieves detailed information about Entra ID Hybrid Join configuration on the local computer.

.DESCRIPTION
    This function gathers comprehensive diagnostics for Entra ID Hybrid Join status including:
    - Service Connection Point (SCP) details
    - Network connectivity to Microsoft Entra ID endpoints
    - Registry keys related to hybrid join
    - Scheduled tasks for Automatic Device Join
    - Computer certificates in AD and local store
    - Event logs for device registration
    - dsregcmd status output

.PARAMETER ExportToExcel
    Optional path to export the results to an Excel file.

.EXAMPLE
    Get-EntraIDHybridJoinComputerInfo

    Retrieves all hybrid join diagnostic information for the local computer.

.EXAMPLE
    Get-EntraIDHybridJoinComputerInfo -ExportToExcel "C:\Reports\HybridJoinInfo.xlsx"

    Retrieves hybrid join information and exports it to an Excel file.

.NOTES
    Must be run on a domain-joined computer.
#>

function Get-EntraIDHybridJoinComputerInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )
    
    # Initialize results collection
    $results = @{}
    
    Write-Host -ForegroundColor Cyan 'Looking for Service Connection Point (SCP)'
    $scp = Get-EntraIDHybridJoinSCP
    $results.SCP = $scp

    if (-not ($ExportToExcel.IsPresent)) {
        $scp
    }

    # Urls
    $urls = @(
        'login.microsoftonline.com'
        'enterpriseregistration.windows.net'
        'device.login.microsoftonline.com'
    )

    Write-Host -ForegroundColor Cyan 'Testing connectivity to Microsoft Entra ID'
    
    [System.Collections.Generic.List[Object]]$connectivityResults = @()
    
    foreach ($url in $urls) {
        $result = Test-NetConnection $url -Port 443
        $connectivityResult = [PSCustomObject][ordered]@{
            Url           = $url
            Connected     = $result.TcpTestSucceeded
            RemoteAddress = $result.RemoteAddress
        }

        $connectivityResults.Add($connectivityResult)
        if ($result.TcpTestSucceeded -eq $false) {
            Write-Warning -ForegroundColor Red "Failed to connect to $url"
        }
        else {
            Write-Host -ForegroundColor Green "Successfully connected to $url - $($result.RemoteAddress)"
        }
    }

    $results.Connectivity = $connectivityResults

    Write-Host -ForegroundColor Cyan 'Testing Microsoft Entra ID registry keys'
    $registryInfo = Get-EntraIDHybridJoinComputerRegistryKey
    $results.RegistryInfo = $registryInfo

    if ($registryInfo.TenantId) {
        Write-Host -ForegroundColor Green "TenantID: $($registryInfo.TenantId)"
    }
    if ($registryInfo.TenantName) {
        Write-Host -ForegroundColor Green "TenantName: $($registryInfo.TenantName)"
    }
    if ($registryInfo.OtherKeys) {
        Write-Host -ForegroundColor Yellow "Other registry keys: $($registryInfo.OtherKeys)"
    }

    Write-Host -ForegroundColor Cyan 'Get scheduled tasks for Workplace Join'
    $deviceJoinTask = Get-ScheduledTask -TaskPath '\Microsoft\Windows\Workplace Join\' -TaskName 'Automatic-Device-Join' 

    $task = Get-ScheduledTaskInfo -TaskPath ($deviceJoinTask.TaskPath) -TaskName $deviceJoinTask.TaskName
    $taskInfo = [PSCustomObject][ordered]@{
        TaskPath       = $task.TaskPath
        TaskName       = $task.TaskName
        LastTaskResult = $task.LastTaskResult
        LastRunTime    = $task.LastRunTime
        NextRunTime    = $task.NextRunTime
    }
    $results.ScheduledTask = $taskInfo

    Write-Host -ForegroundColor Cyan "Task: $($task.TaskPath)$($task.TaskName) - Task last result: $($task.LastTaskResult) - LastRunTime: $($task.LastRunTime) - NextRunTime: $($task.NextRunTime)"

    Write-Host -ForegroundColor Cyan 'Get usercertificate in AD computer object'

    $computerName = $env:COMPUTERNAME

    $search = [adsisearcher]"(&(ObjectClass=Computer)(cn=$env:computername))"

    [System.Collections.Generic.List[Object]]$adCertificates = @()
    try {
        $computer = $search.FindAll()
    
        foreach ($certificate in $computer.Properties.usercertificate) {  
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]$certificate 

            $adCertificate = [PSCustomObject][ordered]@{ 
                DnsNameList          = $cert.DnsNameList -join '|' 
                FriendlyName         = $cert.FriendlyName
                Subject              = $cert.Subject  
                Issuer               = $cert.Issuer
                NotAfter             = $cert.NotAfter
                NotBefore            = $cert.NotBefore
                Thumbprint           = $cert.Thumbprint
                ExpiredData          = $cert.NotAfter  
                SerialNumber         = $cert.SerialNumber  
                EnhancedKeyUsageList = $cert.EnhancedKeyUsageList -join '|'
            }
            $adCertificates.Add($adCertificate) 
        }
    }
    catch {
        Write-Warning "Failed to get $computerName object from AD - $($_.Exception.Message)"
    }

    if (-not ($ExportToExcel.IsPresent)) {
        $adCertificates
    }
    else {
        $results.ADCertificates = $adCertificates
    }

    Write-Host -ForegroundColor Cyan 'Get computer certificates in Personal store'
    
    [System.Collections.Generic.List[Object]]$localEntraCertificates = @()
    
    $entraCertificates = Get-ChildItem -Path cert:\LocalMachine\My | Where-Object { $_.Issuer -like '*MS-Organization-Access*' -or $_.Issuer -like '*MS-Organization-P2P-Access*' }
    
    foreach ($certificate in $entraCertificates) {  
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]$certificate 

        $localCertificate = [PSCustomObject][ordered]@{ 
            SerialNumber         = $cert.SerialNumber  
            FriendlyName         = $cert.FriendlyName
            Subject              = $cert.Subject  
            Issuer               = $cert.Issuer
            NotAfter             = $cert.NotAfter
            NotBefore            = $cert.NotBefore
            Thumbprint           = $cert.Thumbprint
            ExpiredData          = $cert.NotAfter  
            EnhancedKeyUsageList = $cert.EnhancedKeyUsageList -join '|'
            DnsNameList          = $cert.DnsNameList -join '|' 

        }

        $localEntraCertificates.Add($localCertificate)
    }

    if (-not ($ExportToExcel.IsPresent)) {
        $localEntraCertificates
    }
    else {
        $results.LocalCertificates = $localEntraCertificates
    }

    Write-Host -ForegroundColor Cyan 'Get Hybrid Join details'
    [System.Collections.Generic.List[Object]]$DeviceRegistrationEvent = @()

    $eventObject = [PSCustomObject][ordered]@{
        TimeCreated = $null
        Message     = 'No event found'
    }

    
    try {
        $eventLogs = Get-WinEvent -FilterHashtable @{LogName = 'Microsoft-Windows-User Device Registration/Admin'; Id = 306 } -ErrorAction SilentlyContinue

        if ($eventLogs) {
            foreach ($event in $eventLogs) {
                $eventObject = [PSCustomObject][ordered]@{
                    TimeCreated = $event.TimeCreated
                    Message     = $event.Message
                }
            }
        }

        else {
            Write-Warning 'No event in Microsoft-Windows-User Device Registration/Admin was found'
            # Sometimes, we just need to wait after the logon
            # I have some issue with 360 events and Windows Hello Enterprise provisioning or not domain controller line of sight
            Write-Warning "You can logoff/login or wait or runStart-ScheduledTask -TaskPath '\Microsoft\Windows\Workplace Join\' -TaskName 'Automatic-Device-Join'"
        }

        $DeviceRegistrationEvent.Add($eventObject)
    }
    catch {
        Write-Warning "Failed to get events: $($_.Exception.Message)"
        $DeviceRegistrationEvent.Add($eventObject)
    }
    
    $results.DeviceRegistrationEvent = $DeviceRegistrationEvent

    Write-Host -ForegroundColor Cyan 'Get dsregcmd status'

    # Computer state
    $dsregcmd = (dsregcmd /status | Where-Object { $_ -match ' : ' } | ForEach-Object { $_.Trim() } | ConvertFrom-String -PropertyNames 'Name', 'Value' -Delimiter ' : ')

    $dsregcmdObject = [PSCustomObject][ordered]@{}

    foreach ($property in $dsregcmd) {
        $dsregcmdObject | Add-Member -NotePropertyName $property.Name -NotePropertyValue $property.Value
    }

    $joinType = 0
    if ( ($dsregcmd | Where-Object { $_.Name -eq 'EnterpriseJoined' }).Value -eq 'YES' ) {
        $joinType = 4
    }
    else {
        if ( ($dsregcmd | Where-Object { $_.Name -eq 'AzureAdJoined' }).Value -eq 'YES' ) {
            if ( ($dsregcmd | Where-Object { $_.Name -eq 'DomainJoined' }).Value -eq 'YES' ) {
                $joinType = 3
            }
            else {
                $joinType = 1
            }
        }
        else {
            if ( ($dsregcmd | Where-Object { $_.Name -eq 'DomainJoined' }).Value -eq 'YES' ) {
                if ( ($dsregcmd | Where-Object { $_.Name -eq 'WorkplaceJoined' }).Value -eq 'YES' ) {
                    $joinType = 6
                }
                else {
                    $joinType = 2
                }
            }
            else {
                if ( ($dsregcmd | Where-Object { $_.Name -eq 'WorkplaceJoined' }).Value -eq 'YES' ) {
                    $joinType = 5
                }
            }
        }
    }

    switch ($joinType) {
        0 {
            $joinTypeText = 'Local Only'
            break 
        }
        1 {
            $joinTypeText = 'AAD Joined'
            break 
        }
        2 {
            $joinTypeText = 'Domain Only'
            break
        }
        3 {
            $joinTypeText = 'Hybrid AAD'
            break
        }
        4 {
            $joinTypeText = 'DRS'
            break 
        }
        5 {
            $joinTypeText = 'Local;AAD Reg'
            break 
        }
        6 {
            $joinTypeText = 'Domain;AAD Reg'
            break 
        }
        default {
            $joinTypeText = "Unknown join Type - $joinType"
            break
        }
    }

    $dsregcmdObject | Add-Member -NotePropertyName 'JoinType' -NotePropertyValue $joinTypeText
    $results.DsRegCmd = $dsregcmdObject

    $dsregcmdObject

    # Export to Excel if requested
    if ($ExportToExcel.IsPresent) {

        Write-Verbose 'Preparing Excel export...'
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-EntraIDHybridJoinInfo.xlsx"
        Write-Verbose "Excel file path: $excelFilePath"

        $workbook = @{}
        if ($results.SCP) {
            $workbook.SCP = @($results.SCP)
        }
        if ($results.Connectivity) {
            $workbook.Connectivity = $results.Connectivity
        }
        if ($results.RegistryInfo) {
            $workbook.RegistryInfo = @($results.RegistryInfo)
        }
        if ($results.ScheduledTask) {
            $workbook.ScheduledTask = @($results.ScheduledTask)
        }
        if ($results.ADCertificates -and $results.ADCertificates.Count -gt 0) {
            $workbook.ADCertificates = $results.ADCertificates
        }
        if ($results.LocalCertificates -and $results.LocalCertificates.Count -gt 0) {
            $workbook.LocalCertificates = $results.LocalCertificates
        }
        if ($results.DeviceRegistrationEvent -and $results.DeviceRegistrationEvent.Count -gt 0) {
            $workbook.DeviceRegistrationEvent = $results.DeviceRegistrationEvent
        }
        if ($results.DsRegCmd) {
            # Convert dsregcmd object to array for Excel export
            [System.Collections.Generic.List[Object]]$dsregArray = @()
            foreach ($property in $results.DsRegCmd.PSObject.Properties) {
                $dsregProperty = [PSCustomObject]@{
                    Property = $property.Name
                    Value    = $property.Value
                }
                $dsregArray.Add($dsregProperty)
            }
            $workbook.DsRegCmd = $dsregArray
        }

        # Export each section to a separate worksheet
        Write-Host -ForegroundColor Cyan "Exporting dynamic groups to Excel file: $excelFilePath"

        foreach ($sheetName in $workbook.Keys) {
            Write-Verbose "Exporting $sheetName to Excel..."
            $workbook[$sheetName] | Export-Excel -Path $excelFilePath -WorksheetName $sheetName -AutoSize -TableStyle Medium9
        }
    }
}