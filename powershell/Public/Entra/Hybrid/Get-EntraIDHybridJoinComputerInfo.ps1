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

.EXAMPLE
    Get-EntraIDHybridJoinComputerInfo

    Retrieves all hybrid join diagnostic information for the local computer.

.NOTES
    Must be run on a domain-joined computer.
#>

function Get-EntraIDHybridJoinComputerInfo {
    [CmdletBinding()]
    param ()
    Write-Host -ForegroundColor Cyan 'Looking for Service Connection Point (SCP)'
    $scp = Get-EntraIDHybridJoinSCP

    Write-Host -ForegroundColor Cyan "SCP Details: $scp"

    # Urls
    $urls = @(
        'login.microsoftonline.com'
        'enterpriseregistration.windows.net'
        'device.login.microsoftonline.com'
    )

    Write-Host -ForegroundColor Cyan 'Testing connectivity to Microsoft Entra ID'
    foreach ($url in $urls) {
        $result = Test-NetConnection $url -Port 443
        if ($result.TcpTestSucceeded -eq $false) {
            Write-Warning -ForegroundColor Red "Failed to connect to $url"
        }
        else {
            Write-Host -ForegroundColor Green "Successfully connected to $url - $($result.RemoteAddress)"
        }
    }

    Write-Host -ForegroundColor Cyan "`nTesting Microsoft Entra ID registry keys"
    $registryInfo = Get-EntraIDHybridJoinComputerRegistryKeys

    if ($registryInfo.TenantId) {
        Write-Host -ForegroundColor Green "TenantID: $($registryInfo.TenantId)"
    }
    if ($registryInfo.TenantName) {
        Write-Host -ForegroundColor Green "TenantName: $($registryInfo.TenantName)"
    }
    if ($registryInfo.OtherKeys) {
        Write-Host -ForegroundColor Yellow "Other registry keys: $($registryInfo.OtherKeys)"
    }

    Write-Host -ForegroundColor Cyan "`nGet scheduled tasks for Workplace Join"
    $deviceJoinTask = Get-ScheduledTask -TaskPath '\Microsoft\Windows\Workplace Join\' -TaskName 'Automatic-Device-Join' 

    $task = Get-ScheduledTaskInfo -TaskPath ($deviceJoinTask.TaskPath) -TaskName $deviceJoinTask.TaskName

    Write-Host -ForegroundColor Cyan "Task: $($task.TaskPath)$($task.TaskName) - Task last result: $($task.LastTaskResult) - LastRunTime: $($task.LastRunTime) - NextRunTime: $($task.NextRunTime)"

    Write-Host -ForegroundColor Cyan "`nGet usercertificate in AD computer object"

    $computerName = $env:COMPUTERNAME

    $search = [adsisearcher]"(&(ObjectClass=Computer)(cn=$env:computername))"

    try {
        $computer = $search.FindAll()
    
        $computer.Properties.usercertificate | ForEach-Object {  
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]$_ 

            [PSCustomObject][ordered]@{ 
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
        
        }
    }
    catch {
        Write-Warning "Failed to get $computername  object from AD - $($_.Exception.Message)"
    }

    Write-Host -ForegroundColor Cyan "`nGet computer certificates in Personal store"
    Get-ChildItem -Path cert:\LocalMachine\My | ForEach-Object {  
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]$_ 

        $object = [PSCustomObject][ordered]@{ 
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

        return $object
    }

    Write-Host -ForegroundColor Cyan "`nGet Hybrid Join details"
    $event = Get-WinEvent -FilterHashtable @{LogName = 'Microsoft-Windows-User Device Registration/Admin'; Id = 306 }

    if ($event) {
        $event | ForEach-Object {
            $eventObject = [PSCustomObject][ordered]@{
                TimeCreated = $_.TimeCreated
                Message     = $_.Message
            }
            $eventObject
        }
    }
    else {
        Write-Warning 'No event in Microsoft-Windows-User Device Registration/Admin was found'
        # Sometimes, we just need to wait after the logon
        # I have some issue with 360 events and Windows Hello Enterprise provisioning or not domain controller line of sight
        Write-Warning "You can logoff/login or wait or run`nStart-ScheduledTask -TaskPath '\Microsoft\Windows\Workplace Join\' -TaskName 'Automatic-Device-Join'"
    }

    Write-Host -ForegroundColor Cyan "`nGet dsregcmd status"

    # Computer state
    $dsregcmd = (dsregcmd /status | Where-Object { $_ -match ' : ' } | ForEach-Object { $_.Trim() } | ConvertFrom-String -PropertyNames 'Name', 'Value' -Delimiter ' : ')

    $dsregcmdObject = [PSCustomObject][ordered]@{}

    foreach ($property in $dsregcmd) {
        #    $dsregcmdObject.Add($property.Name, $property.Value)
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

    $dsregcmdObject

    # Start task Automatic-Device-Join in \Microsoft\Windows\Workplace Join
    # Start-ScheduledTask -TaskPath '\Microsoft\Windows\Workplace Join\' -TaskName 'Automatic-Device-Join'

    # For Intune, only need to wait the next gpupdate cycle (connected to the AD domain) or run gpupdate
}