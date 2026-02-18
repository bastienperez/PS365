<#
    .SYNOPSIS
    Invokes the backup of BitLocker recovery keys to Intune for all Intune managed devices.

    .DESCRIPTION
    This script connects to Microsoft Intune and retrieves BitLocker recovery keys from all
    devices managed by Intune. It requires the Microsoft Graph PowerShell SDK to be installed and
    appropriate permissions to access device management and BitLocker key data.

    .PARAMETER IncludeDeviceInfo
    Switch to include device information such as device name, OS, compliance status, etc.
    
    .PARAMETER IncludeDeviceOwner
    Switch to include device owner information (UPN). Requires IncludeDeviceInfo parameter.
    
    .PARAMETER ExportToExcel
    Switch to export the results to an Excel file in the user profile directory.
    If not specified, the function returns the data objects.
    
    .PARAMETER RevealKeys
    Switch to display BitLocker recovery keys in plain text format in the CSV export.
    WARNING: This will expose sensitive BitLocker recovery keys in the output file!
    Use only when necessary and ensure secure storage of the exported file.
    
    .PARAMETER BackupToKeyVault
    Specify the name of Azure Key Vault to backup BitLocker recovery keys.
    Requires Azure PowerShell module and appropriate permissions to access Key Vault.
    Keys will be stored with device name and BitLocker key ID as the secret name.
    Example: -BackupToKeyVault "MyBitLockerVault"
    
    .PARAMETER DeviceName
    Filter results to a specific device by its display name.
    Cannot be used together with DeviceID parameter.
    Example: -DeviceName "LAPTOP-ABC123"
    
    .PARAMETER DeviceID
    Filter results to a specific device by its device ID (GUID).
    Cannot be used together with DeviceName parameter.
    Example: -DeviceID "12345678-1234-1234-1234-123456789abc"

    .EXAMPLE
    Get-MgBitlockerKeyInfo -IncludeDeviceInfo -IncludeDeviceOwner

    This command retrieves BitLocker recovery keys for all Intune managed devices with device and owner information.

    .EXAMPLE
    Get-MgBitlockerKeyInfo -IncludeDeviceInfo -ExportToExcel

    This command retrieves BitLocker keys with device information and exports to Excel.
    
    .EXAMPLE
    Get-MgBitlockerKeyInfo -IncludeDeviceInfo -IncludeDeviceOwner -RevealKeys -ExportToExcel

    This command generates a comprehensive report with BitLocker keys visible in plain text and exports to Excel.
    WARNING: Use with extreme caution as this exposes sensitive recovery keys!
    
    .EXAMPLE
    Get-MgBitlockerKeyInfo -IncludeDeviceInfo -BackupToKeyVault "MyBitLockerVault" -ExportToExcel

    This command retrieves BitLocker keys, backs them up to the specified Azure Key Vault, and exports to Excel.

    .EXAMPLE
    Get-MgBitlockerKeyInfo -DeviceName "LAPTOP-ABC123" -IncludeDeviceInfo -RevealKeys

    This command retrieves BitLocker keys for a specific device by name, includes device information, and displays keys in plain text.

    .EXAMPLE
    Get-MgBitlockerKeyInfo -DeviceID "12345678-1234-1234-1234-123456789abc" -ExportToExcel

    This command retrieves BitLocker keys for a specific device by ID and exports the results to Excel.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgBitlockerKeyInfo
    
    .NOTES
    Author: Bastien Perez (adapted from Vasil Michev)
    Source: https://github.com/michevnew/PowerShell/blob/master/GraphSDK_Bitlocker_report.ps1
    
    The script requires the following Microsoft Graph permissions:
    - BitLockerKey.Read.All (required) - Allows the app to read BitLocker keys on behalf of the signed-in user, 
    for their owned devices. Allows read of the recovery key.
    - Device.Read.All (optional) - Needed to retrieve device details like name, OS, compliance status
    - User.ReadBasic.All (optional) - Needed to retrieve device owner UPN information
    
    PERMISSION SCOPE CONSIDERATIONS:
    - User context: Can only read BitLocker keys for devices owned by the signed-in user (if you have admin permissions, you can read all devices and all bitlocker keys)
    - Application context: Can read BitLocker keys for all devices in the organization (requires admin consent)
    - Managed Identity: Same as application context when properly configured with admin consent
    
    SECURITY WARNING: The exported CSV file contains sensitive BitLocker recovery keys. 
    Store it in a secure location and limit access appropriately!
#>

function Get-MgBitlockerKeyInfo {
    [CmdletBinding()]
    param(
        [Parameter(HelpMessage = 'Include device information in the output')]
        [switch]$IncludeDeviceInfo,
        
        [Parameter(HelpMessage = 'Include device owner information (requires IncludeDeviceInfo)')]
        [switch]$IncludeDeviceOwner,
        
        [Parameter(HelpMessage = 'Export results to Excel file in user profile directory')]
        [switch]$ExportToExcel,
        
        [Parameter(HelpMessage = 'Show BitLocker recovery keys in plain text (/!\ SECURITY RISK: Use with caution!)')]
        [switch]$RevealKeys,
        
        [Parameter(HelpMessage = 'Specify Azure Key Vault name to backup BitLocker keys')]
        [ValidateNotNullOrEmpty()]
        [string]$BackupToKeyVault,
        
        [Parameter(HelpMessage = 'Filter by device name (cannot be used with DeviceID)')]
        [ValidateNotNullOrEmpty()]
        [string]$DeviceName,
        
        [Parameter(HelpMessage = 'Filter by device ID (cannot be used with DeviceName)')]
        [ValidateNotNullOrEmpty()]
        [string]$DeviceID
    )

    # Validate that only one device filter is specified
    if ($PSBoundParameters.ContainsKey('DeviceName') -and $PSBoundParameters.ContainsKey('DeviceID')) {
        Write-Error 'Cannot specify both DeviceName and DeviceID parameters. Please use only one.' -ErrorAction Stop
    }

    function Get-DriveTypeName {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [int]$DriveType
        )
        
        switch ($DriveType) {
            1 { return 'operatingSystemVolume' }
            2 { return 'fixedDataVolume' }
            3 { return 'removableDataVolume' }
            4 { return 'unknownFutureValue' }
            default { return 'Unknown' }
        }
    }

    # Handle parameter dependencies for comprehensive reporting
    # When exporting to Excel, we typically want full device information
    if ($PSBoundParameters.ContainsKey('ExportToExcel') -and $PSBoundParameters['ExportToExcel']) {
        if (-not $PSBoundParameters.ContainsKey('IncludeDeviceInfo')) {
            Write-Verbose 'ExportToExcel specified - automatically including device information for comprehensive report'
            $PSBoundParameters['IncludeDeviceInfo'] = $true
        }
    }
    if ($PSBoundParameters.ContainsKey('IncludeDeviceOwner') -and $PSBoundParameters['IncludeDeviceOwner']) {
        $PSBoundParameters['IncludeDeviceInfo'] = $true
    }

    # Determine the required scopes, based on the parameters passed to the script
    $requiredScopes = switch ($PSBoundParameters.Keys) {
        'IncludeDeviceInfo' { if ($PSBoundParameters['IncludeDeviceInfo']) { 'Device.Read.All' } }
        'IncludeDeviceOwner' { if ($PSBoundParameters['IncludeDeviceOwner']) { 'User.ReadBasic.All' } } # Otherwise we only get the UserId
        default { 'BitLockerKey.Read.All' }
    }

    Write-Verbose 'Importing required PowerShell modules...'
    try {
        Import-Module Microsoft.Graph.Identity.SignIns -Force -ErrorAction Stop
        Write-Verbose 'Microsoft.Graph.Identity.SignIns module imported successfully'
    }
    catch {
        Write-Error "Failed to import Microsoft.Graph.Identity.SignIns module: $($_.Exception.Message)" -ErrorAction Stop
    }
    
    # Verify we have all required permissions
    Write-Verbose 'Verifying Microsoft Graph permissions...'
    try {
        $currentScopes = (Get-MgContext).Scopes
        $missingScopes = $requiredScopes | Where-Object { $_ -notin $currentScopes }
        
        if ($missingScopes) {
            $missingScopesString = $missingScopes -join ', '
            Write-Error "Missing required Microsoft Graph permissions: $missingScopesString. Please rerun the script and consent to the missing scopes." -ErrorAction Stop
        }
        
        Write-Verbose 'All required permissions are available'
    }
    catch {
        Write-Error "Failed to verify permissions: $($_.Exception.Message)" -ErrorAction Stop
    }
    
    # Setup Azure Key Vault connection if backup is requested
    if ($PSBoundParameters['BackupToKeyVault']) {
        Write-Verbose 'Setting up Azure Key Vault connection for BitLocker keys backup...'
        
        # Key Vault configuration from parameter
        $keyVaultName = $BackupToKeyVault
        Write-Verbose "Using Key Vault: $keyVaultName"
        
        try {
            # Check if Azure PowerShell module is available
            if (-not (Get-Module -ListAvailable -Name Az.KeyVault)) {
                Write-Error 'Az.KeyVault module is required for Key Vault backup. Install it with: Install-Module Az.KeyVault' -ErrorAction Stop
            }
            
            # Connect to Azure (assumes Managed Identity or existing connection)
            Write-Verbose 'Connecting to Azure for Key Vault access...'
            try {
                # Try to get current context first
                $azContext = Get-AzContext -ErrorAction SilentlyContinue
                if (-not $azContext) {
                    # Attempt Managed Identity connection
                    Connect-AzAccount -Identity -ErrorAction Stop
                    Write-Verbose 'Connected to Azure using Managed Identity'
                }
                else {
                    Write-Verbose 'Using existing Azure connection'
                }
            }
            catch {
                Write-Warning 'Failed to connect to Azure automatically. Please ensure you are logged in with Connect-AzAccount or using Managed Identity.'
                Write-Error "Azure connection required for Key Vault backup: $($_.Exception.Message)" -ErrorAction Stop
            }
            
            # Verify Key Vault access
            Write-Verbose "Verifying access to Key Vault: $keyVaultName"
            try {
                $keyVault = Get-AzKeyVault -VaultName $keyVaultName -ErrorAction Stop
                Write-Verbose "Successfully verified access to Key Vault: $($keyVault.VaultName)"
            }
            catch {
                Write-Error "Cannot access Key Vault '$keyVaultName'. Please ensure it exists and you have appropriate permissions: $($_.Exception.Message)" -ErrorAction Stop
            }
        }
        catch {
            Write-Error "Failed to setup Azure Key Vault connection: $($_.Exception.Message)" -ErrorAction Stop
        }
    }

    # Retrieve device details if requested
    if ($PSBoundParameters['IncludeDeviceInfo']) {
        Write-Verbose 'Retrieving device details from Microsoft Graph...'
        
        try {
            [System.Collections.Generic.List[Object]]$devices = @()
            
            # Determine the appropriate filter based on provided parameters
            if ($PSBoundParameters.ContainsKey('DeviceName')) {
                Write-Verbose "Filtering devices by name: $DeviceName"
                $filter = "displayName eq '$DeviceName'"
            }
            elseif ($PSBoundParameters.ContainsKey('DeviceID')) {
                Write-Verbose "Filtering devices by ID: $DeviceID"
                $filter = "deviceId eq '$DeviceID'"
            }
            else {
                $filter = $null
            }
            
            # Retrieve devices with or without filter
            if ($PSBoundParameters['IncludeDeviceOwner']) {
                Write-Verbose 'Including device owner information...'
                if ($filter) {
                    $deviceResults = Get-MgDevice -Filter $filter -ExpandProperty registeredOwners -ErrorAction Stop
                }
                else {
                    $deviceResults = Get-MgDevice -All -ExpandProperty registeredOwners -ErrorAction Stop
                }
            }
            else {
                if ($filter) {
                    $deviceResults = Get-MgDevice -Filter $filter -ErrorAction Stop
                }
                else {
                    $deviceResults = Get-MgDevice -All -ErrorAction Stop
                }
            }
            
            # Convert to list and add only valid results
            foreach ($deviceResult in $deviceResults) {
                # Only add valid devices (filter out invalid/dummy devices)
                if ($deviceResult.Id -ne '00000000-0000-0000-0000-000000000000' -and 
                    $deviceResult.DeviceId -ne '00000000-0000-0000-0000-000000000000') {
                    $devices.Add($deviceResult)
                }
            }
            
            $originalDeviceCount = $deviceResults.Count
            if ($devices.Count -gt 0) { 
                Write-Verbose "Successfully retrieved $($devices.Count) valid devices (filtered out $($originalDeviceCount - $devices.Count) invalid devices)"
            }
            else { 
                if ($PSBoundParameters.ContainsKey('DeviceName') -or $PSBoundParameters.ContainsKey('DeviceID')) {
                    Write-Warning 'No valid devices found matching the specified criteria'
                }
                else {
                    Write-Warning 'No valid devices found in Microsoft Graph'
                }
                return
            }
        }
        catch {
            Write-Error "Failed to retrieve device information: $($_.Exception.Message)" -ErrorAction Stop
        }

        # Prepare the device object to be used later on
        if ($PSBoundParameters['ExportToExcel']) {
            foreach ($device in $devices) {
                $device | Add-Member -MemberType NoteProperty -Name 'BitLockerKeyId' -Value $null
                $device | Add-Member -MemberType NoteProperty -Name 'BitLockerRecoveryKey' -Value $null
                $device | Add-Member -MemberType NoteProperty -Name 'BitLockerDriveType' -Value $null
                $device | Add-Member -MemberType NoteProperty -Name 'BitLockerBackedUp' -Value $null
            }
        }
        foreach ($device in $devices) {
            $device | Add-Member -MemberType NoteProperty -Name 'DeviceOwner' -Value (& { if ($device.registeredOwners) { $device.registeredOwners[0].AdditionalProperties.userPrincipalName } else { '$null' } })
        }    
    }

    # Get the list of BitLocker keys
    [System.Collections.Generic.List[Object]]$keys = @()
    
    # If device filtering is specified, only get keys for those specific devices
    if ($PSBoundParameters.ContainsKey('DeviceName') -or $PSBoundParameters.ContainsKey('DeviceID')) {
        Write-Verbose 'Filtering BitLocker keys for specified devices only...'
        
        # If devices were not retrieved yet (because IncludeDeviceInfo was not specified), get them now for filtering
        if (-not $PSBoundParameters['IncludeDeviceInfo']) {
            try {
                # Determine the appropriate filter based on provided parameters
                if ($PSBoundParameters.ContainsKey('DeviceName')) {
                    Write-Verbose "Retrieving device by name: $DeviceName"
                    $filter = "displayName eq '$DeviceName'"
                }
                elseif ($PSBoundParameters.ContainsKey('DeviceID')) {
                    Write-Verbose "Retrieving device by ID: $DeviceID"
                    $filter = "deviceId eq '$DeviceID'"
                }
                
                $deviceResults = Get-MgDevice -Filter $filter -ErrorAction Stop
                [System.Collections.Generic.List[Object]]$devices = @()
                
                foreach ($deviceResult in $deviceResults) {
                    # Only add valid devices (filter out invalid/dummy devices)
                    if ($deviceResult.Id -ne '00000000-0000-0000-0000-000000000000' -and 
                        $deviceResult.DeviceId -ne '00000000-0000-0000-0000-000000000000') {
                        $devices.Add($deviceResult)
                    }
                }
                
                if ($devices.Count -eq 0) {
                    Write-Warning 'No valid devices found matching the specified criteria'
                    return
                }
            }
            catch {
                Write-Error "Failed to retrieve device for filtering: $($_.Exception.Message)" -ErrorAction Stop
            }
        }
        
        # Get BitLocker keys for the specific devices
        foreach ($device in $devices) {
            try {
                Write-Verbose "Retrieving BitLocker keys for device: $($device.DisplayName) (ID: $($device.DeviceId))..."
                
                # Create base object for this device (no BitLocker key found by default)  
                $deviceKeyInfo = [PSCustomObject]@{
                    Id                   = $null
                    DeviceId             = $device.DeviceId
                    VolumeType           = $null
                    Key                  = 'No BitLocker key found'
                    CreatedDateTime      = $null
                    AdditionalProperties = @{}
                }
                
                # Try to find BitLocker keys for this device
                $deviceKeys = Get-MgInformationProtectionBitlockerRecoveryKey -Filter "deviceId eq '$($device.DeviceId)'" -ErrorAction Stop -Verbose:$false
                
                if ($deviceKeys -and $deviceKeys.Count -gt 0) {
                    # Device has BitLocker keys - add each key as separate object
                    foreach ($key in $deviceKeys) {
                        $keys.Add($key)
                    }
                }
                else {
                    # No BitLocker keys found - add the base object
                    Write-Verbose "No BitLocker keys found for device: $($device.DisplayName)"
                    $keys.Add($deviceKeyInfo)
                }
            }
            catch {
                Write-Warning "Failed to retrieve BitLocker keys for device $($device.DisplayName): $($_.Exception.Message)"
                # Create base object for device with retrieval error
                $deviceKeyInfo = [PSCustomObject]@{
                    Id                   = $null
                    DeviceId             = $device.DeviceId
                    VolumeType           = $null
                    Key                  = 'Error retrieving BitLocker keys'
                    CreatedDateTime      = $null
                    AdditionalProperties = @{}
                }
                $keys.Add($deviceKeyInfo)
            }
        }
    }
    else {
        Write-Verbose 'Retrieving all BitLocker keys...'
        # Get all BitLocker keys (original behavior)
        $bitlockerKeys = Get-MgInformationProtectionBitlockerRecoveryKey -All -ErrorAction Stop -Verbose:$false
        foreach ($key in $bitlockerKeys) {
            $keys.Add($key)
        }
    }
    
    Write-Verbose "Found $($keys.Count) BitLocker keys to process"

    # Cycle through the keys and retrieve the key
    Write-Verbose 'Processing BitLocker Recovery keys...'
    $keyCounter = 0
    $totalKeys = $keys.Count
    
    foreach ($key in $keys) {
        $keyCounter++
        Write-Verbose "[$keyCounter/$totalKeys] Processing BitLocker key for device $($key.DeviceId)..."
        
        # Skip stale/dummy devices
        if ($key.DeviceId -eq '00000000-0000-0000-0000-000000000000') {
            Write-Warning "[$keyCounter/$totalKeys] BitLocker key with ID $($key.Id) has a device ID of 00000000-0000-0000-0000-000000000000, skipping..."
            continue
        }

        # Handle devices without BitLocker keys (dummy objects)
        if ($null -eq $key.Id) {
            Write-Verbose "[$keyCounter/$totalKeys] Processing device without BitLocker key: $($key.DeviceId)"
            $keyValue = 'No BitLocker key found'
            $actualKeyValue = $null
        }
        else {
            # Get the BitLocker key details (plain text required for Key Vault backup or if explicitly requested)
            if ($PSBoundParameters['RevealKeys'] -or $PSBoundParameters['BackupToKeyVault']) {
                Write-Verbose "[$keyCounter/$totalKeys] Retrieving plain text BitLocker key for device $($key.DeviceId)..."
                $recoveryKey = Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $key.Id -Property key -ErrorAction Stop -Verbose:$false
                $actualKeyValue = if ($recoveryKey.Key) { $recoveryKey.Key } else { '$null' }
                
                # For display purposes, hide the key unless explicitly requested
                if ($PSBoundParameters['RevealKeys']) {
                    $keyValue = $actualKeyValue
                }
                else {
                    $keyValue = '[HIDDEN - Backed up to Key Vault]'
                }
            }
            else {
                $keyValue = '[HIDDEN - Use -RevealKeys to display]'
                $actualKeyValue = $null
            }
        }
        
        # Backup to Key Vault if requested
        if ($PSBoundParameters['BackupToKeyVault'] -and $actualKeyValue -and $actualKeyValue -ne '$null') {
            try {
                Write-Verbose "[$keyCounter/$totalKeys] Backing up BitLocker key to Key Vault..."
                
                # Get device information for Key Vault secret name
                $deviceInfo = Get-MgDevice -Filter "DeviceId eq '$($key.DeviceId)'" -ErrorAction SilentlyContinue
                $deviceName = if ($deviceInfo -and $deviceInfo.DisplayName) { $deviceInfo.DisplayName } else { "UnknownDevice-$($key.DeviceId)" }
                
                # Create Key Vault secret name: DeviceName-BitLockerKeyID-KeyId
                $secretName = "$deviceName-BitLockerKeyID-$($key.Id)"
                
                # Convert to secure string and save to Key Vault
                $secretValue = ConvertTo-SecureString $actualKeyValue -AsPlainText -Force
                $null = Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName -SecretValue $secretValue -ErrorAction Continue
                
                Write-Verbose "[$keyCounter/$totalKeys] Successfully backed up key for device $deviceName to Key Vault"
            }
            catch {
                Write-Warning "[$keyCounter/$totalKeys] Failed to backup BitLocker key to Key Vault for device $($key.DeviceId): $($_.Exception.Message)"
            }
        }
        
        $key.Key = $keyValue
        $key | Add-Member -MemberType NoteProperty -Name 'BitLockerKeyId' -Value (& { if ($null -eq $key.Id) { '$null' } else { $key.Id } })
        $key | Add-Member -MemberType NoteProperty -Name 'BitLockerRecoveryKey' -Value $keyValue
        $key | Add-Member -MemberType NoteProperty -Name 'BitLockerDriveType' -Value (& { if ($null -eq $key.VolumeType) { '$null' } else { Get-DriveTypeName -DriveType $key.VolumeType } })
        $key | Add-Member -MemberType NoteProperty -Name 'BitLockerBackedUp' -Value (& { if ($key.CreatedDateTime) { Get-Date($key.CreatedDateTime) -Format g } else { '$null' } })

        # If requested, include the device details
        if ($PSBoundParameters['IncludeDeviceInfo']) {
            Write-Verbose "[$keyCounter/$totalKeys] Looking up device information for $($key.DeviceId)..."
            $device = $devices | Where-Object { $key.DeviceId -eq $_.DeviceId }
            if (-not $device) {
                Write-Warning "[$keyCounter/$totalKeys] Device with ID $($key.DeviceId) not found!"
                $key | Add-Member -MemberType NoteProperty -Name 'DeviceName' -Value 'Device not found'
                continue
            }
            if ($device.Id -eq '00000000-0000-0000-0000-000000000000' -or $device.DeviceId -eq '00000000-0000-0000-0000-000000000000') {
                Write-Warning "[$keyCounter/$totalKeys] Stale/dummy device found for key $($key.DeviceId), skipping..."
                $key | Add-Member -MemberType NoteProperty -Name 'DeviceName' -Value 'Stale/Dummy Device'
                continue
            }

            # If exporting to Excel, add the BitLocker key details to the device object for comprehensive reporting
            if ($PSBoundParameters['ExportToExcel']) {
                $device.BitLockerKeyId = $key.Id
                $device.BitLockerRecoveryKey = $keyValue
                $device.BitLockerDriveType = (Get-DriveTypeName -DriveType $key.VolumeType)
                $device.BitLockerBackedUp = (& { if ($key.CreatedDateTime) { Get-Date($key.CreatedDateTime) -Format g } else { '$null' } })
            }

            $key | Add-Member -MemberType NoteProperty -Name 'DeviceName' -Value $device.DisplayName
            $key | Add-Member -MemberType NoteProperty -Name 'DeviceGUID' -Value $device.Id # key actually used by the stupid module...
            $key | Add-Member -MemberType NoteProperty -Name 'DeviceOS' -Value $device.OperatingSystem
            $key | Add-Member -MemberType NoteProperty -Name 'DeviceTrustType' -Value $device.TrustType
            $key | Add-Member -MemberType NoteProperty -Name 'DeviceMDM' -Value $device.AdditionalProperties.managementType # can be null! ALWAYS null when using a filter...
            $key | Add-Member -MemberType NoteProperty -Name 'DeviceCompliant' -Value $device.isCompliant # can be null!
            $key | Add-Member -MemberType NoteProperty -Name 'DeviceRegistered' -Value (& { if ($device.registrationDateTime) { Get-Date($device.registrationDateTime) -Format g } else { '$null' } })
            $key | Add-Member -MemberType NoteProperty -Name 'DeviceLastActivity' -Value (& { if ($device.approximateLastSignInDateTime) { Get-Date($device.approximateLastSignInDateTime) -Format g } else { '$null' } })

            # If requested, include the device owner
            if ($PSBoundParameters['IncludeDeviceOwner']) {
                $key | Add-Member -MemberType NoteProperty -Name 'DeviceOwner' -Value (& { if ($device.registeredOwners) { $device.registeredOwners[0].AdditionalProperties.userPrincipalName } else { '$null' } })
            }
        }
    }

    # Export results or return data
    if ($PSBoundParameters['ExportToExcel']) {
        try {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $excelFilePath = "$($env:userprofile)\${now}_BitLockerKeys_Report.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting BitLocker keys report to Excel file: $excelFilePath"
            
            if ($PSBoundParameters['IncludeDeviceInfo']) {
                Write-Verbose 'Exporting comprehensive device report to Excel...'
                # Exclude internal properties from the export
                $excludeProps = @(
                    'AdditionalProperties', 'AlternativeSecurityIds', 'complianceExpirationDateTime',
                    'deviceMetadata', 'deviceVersion', 'memberOf', 'PhysicalIds', 'SystemLabels',
                    'transitiveMemberOf', 'RegisteredOwners', 'RegisteredUsers'
                )
                $devices | Select-Object * -ExcludeProperty $excludeProps | 
                Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'BitLocker-DeviceReport'
            }
            else {
                Write-Verbose 'Exporting BitLocker keys report to Excel...'
                $keys | Select-Object * -ExcludeProperty Id, VolumeType, AdditionalProperties, CreatedDateTime, Key | 
                Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'BitLocker-Keys'
            }
            
            Write-Host "Report successfully exported to: $excelFilePath" -ForegroundColor Green
            
            if ($PSBoundParameters['RevealKeys']) {
                Write-Warning 'SECURITY ALERT: The Excel file contains BitLocker recovery keys in PLAIN TEXT!'
                Write-Warning 'Ensure this file is stored securely and access is properly restricted!'
            }
            else {
                Write-Host 'BitLocker keys are hidden in the export. Use -RevealKeys to display them.' -ForegroundColor Cyan
            }
        }
        catch {
            Write-Error "Failed to export results to Excel: $($_.Exception.Message)" -ErrorAction Stop
        }
    }
    else {
        # Return the data objects
        if ($PSBoundParameters['IncludeDeviceInfo']) {
            Write-Verbose 'Returning device objects with BitLocker key information'
            return $devices
        }
        else {
            Write-Verbose 'Returning BitLocker key objects'
            return $keys
        }
    }
}