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
    ⚠️  WARNING: This will expose sensitive BitLocker recovery keys in the output file!
    Use only when necessary and ensure secure storage of the exported file.
    
    .PARAMETER BackupToKeyVault
    Specify the name of Azure Key Vault to backup BitLocker recovery keys.
    Requires Azure PowerShell module and appropriate permissions to access Key Vault.
    Keys will be stored with device name and BitLocker key ID as the secret name.
    Example: -BackupToKeyVault "MyBitLockerVault"

    .EXAMPLE
    Get-BitlockerKeyFromIntune -IncludeDeviceInfo -IncludeDeviceOwner

    This command retrieves BitLocker recovery keys for all Intune managed devices with device and owner information.

    .EXAMPLE
    Get-BitlockerKeyFromIntune -IncludeDeviceInfo -ExportToExcel

    This command retrieves BitLocker keys with device information and exports to Excel.
    
    .EXAMPLE
    Get-BitlockerKeyFromIntune -IncludeDeviceInfo -IncludeDeviceOwner -RevealKeys -ExportToExcel

    This command generates a comprehensive report with BitLocker keys visible in plain text and exports to Excel.
    ⚠️  WARNING: Use with extreme caution as this exposes sensitive recovery keys!
    
    .EXAMPLE
    Get-BitlockerKeyFromIntune -IncludeDeviceInfo -BackupToKeyVault "MyBitLockerVault" -ExportToExcel

    This command retrieves BitLocker keys, backs them up to the specified Azure Key Vault, and exports to Excel.

    .NOTES
    Author: Bastien Perez (adapted from Vasil Michev)
    Source: https://github.com/michevnew/PowerShell/blob/master/GraphSDK_Bitlocker_report.ps1
    Date: December 2025
    
    The script requires the following Microsoft Graph permissions:
    - BitLockerKey.Read.All (required) - Allows the app to read BitLocker keys on behalf of the signed-in user, 
    for their owned devices. Allows read of the recovery key.
    - Device.Read.All (optional) - Needed to retrieve device details like name, OS, compliance status
    - User.ReadBasic.All (optional) - Needed to retrieve device owner UPN information
    
    PERMISSION SCOPE CONSIDERATIONS:
    - User context: Can only read BitLocker keys for devices owned by the signed-in user (if you have admin permissions, you can read all devices and all bitlocker keys)
    - Application context: Can read BitLocker keys for all devices in the organization (requires admin consent)
    - Managed Identity: Same as application context when properly configured with admin consent
    
    ⚠️  SECURITY WARNING: The exported CSV file contains sensitive BitLocker recovery keys. 
    Store it in a secure location and limit access appropriately!
#>

function Get-BitlockerKeyInfo {
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
        [string]$BackupToKeyVault
    )

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

    #Determine the required scopes, based on the parameters passed to the script
    $RequiredScopes = switch ($PSBoundParameters.Keys) {
        'IncludeDeviceInfo' { if ($PSBoundParameters['IncludeDeviceInfo']) { 'Device.Read.All' } }
        'IncludeDeviceOwner' { if ($PSBoundParameters['IncludeDeviceOwner']) { 'User.ReadBasic.All' } } #Otherwise we only get the UserId
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
        $CurrentScopes = (Get-MgContext).Scopes
        $MissingScopes = $RequiredScopes | Where-Object { $_ -notin $CurrentScopes }
        
        if ($MissingScopes) {
            $MissingScopesString = $MissingScopes -join ', '
            Write-Error "Missing required Microsoft Graph permissions: $MissingScopesString. Please rerun the script and consent to the missing scopes." -ErrorAction Stop
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
            $Devices = @()
            if ($PSBoundParameters['IncludeDeviceOwner']) {
                Write-Verbose 'Including device owner information...'
                $Devices = Get-MgDevice -All -ExpandProperty registeredOwners -ErrorAction Stop
            }
            else {
                $Devices = Get-MgDevice -All -ErrorAction Stop
            }
            
            # Filter out invalid/dummy devices
            $OriginalDeviceCount = $Devices.Count
            $Devices = $Devices | Where-Object { 
                $_.Id -ne '00000000-0000-0000-0000-000000000000' -and 
                $_.DeviceId -ne '00000000-0000-0000-0000-000000000000' 
            }
            
            if ($Devices) { 
                Write-Verbose "Successfully retrieved $($Devices.Count) valid devices (filtered out $($OriginalDeviceCount - $Devices.Count) invalid devices)"
            }
            else { 
                Write-Warning 'No valid devices found in Microsoft Graph'
                return
            }
        }
        catch {
            Write-Error "Failed to retrieve device information: $($_.Exception.Message)" -ErrorAction Stop
        }

        # Prepare the device object to be used later on
        if ($PSBoundParameters['ExportToExcel']) {
            $Devices | Add-Member -MemberType NoteProperty -Name 'BitLockerKeyId' -Value $null
            $Devices | Add-Member -MemberType NoteProperty -Name 'BitLockerRecoveryKey' -Value $null
            $Devices | Add-Member -MemberType NoteProperty -Name 'BitLockerDriveType' -Value $null
            $Devices | Add-Member -MemberType NoteProperty -Name 'BitLockerBackedUp' -Value $null
        }
        $Devices | ForEach-Object { Add-Member -InputObject $_ -MemberType NoteProperty -Name 'DeviceOwner' -Value (& { if ($_.registeredOwners) { $_.registeredOwners[0].AdditionalProperties.userPrincipalName } else { 'N/A' } }) }    
    }

    #Get the list of application objects within the tenant.
    $Keys = @()

    #Get the list of BitLocker keys
    Write-Verbose 'Retrieving BitLocker keys...'
    $Keys = Get-MgInformationProtectionBitlockerRecoveryKey -All -ErrorAction Stop -Verbose:$false
    Write-Verbose "Found $($Keys.Count) BitLocker keys to process"

    #Cycle through the keys and retrieve the key
    Write-Verbose 'Processing BitLocker Recovery keys...'
    $KeyCounter = 0
    $TotalKeys = $Keys.Count
    
    foreach ($Key in $Keys) {
        $KeyCounter++
        Write-Verbose "[$KeyCounter/$TotalKeys] Processing BitLocker key for device $($Key.DeviceId)..."
        #Skip stale/dummy devices
        if ($Key.DeviceId -eq '00000000-0000-0000-0000-000000000000') {
            Write-Warning "[$KeyCounter/$TotalKeys] BitLocker key with ID $($Key.Id) has a device ID of 00000000-0000-0000-0000-000000000000, skipping..."
            continue
        }

        # Get the BitLocker key details (plain text required for Key Vault backup or if explicitly requested)
        if ($PSBoundParameters['RevealKeys'] -or $PSBoundParameters['BackupToKeyVault']) {
            Write-Verbose "[$KeyCounter/$TotalKeys] Retrieving plain text BitLocker key for device $($Key.DeviceId)..."
            $RecoveryKey = Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $Key.Id -Property key -ErrorAction Stop -Verbose:$false
            $ActualKeyValue = if ($RecoveryKey.Key) { $RecoveryKey.Key } else { 'N/A' }
            
            # For display purposes, hide the key unless explicitly requested
            if ($PSBoundParameters['RevealKeys']) {
                $KeyValue = $ActualKeyValue
            }
            else {
                $KeyValue = '[HIDDEN - Backed up to Key Vault]'
            }
        }
        else {
            $KeyValue = '[HIDDEN - Use -RevealKeys to display]'
            $ActualKeyValue = $null
        }
        
        # Backup to Key Vault if requested
        if ($PSBoundParameters['BackupToKeyVault'] -and $ActualKeyValue -and $ActualKeyValue -ne 'N/A') {
            try {
                Write-Verbose "[$KeyCounter/$TotalKeys] Backing up BitLocker key to Key Vault..."
                
                # Get device information for Key Vault secret name
                $deviceInfo = Get-MgDevice -Filter "DeviceId eq '$($Key.DeviceId)'" -ErrorAction SilentlyContinue
                $deviceName = if ($deviceInfo -and $deviceInfo.DisplayName) { $deviceInfo.DisplayName } else { "UnknownDevice-$($Key.DeviceId)" }
                
                # Create Key Vault secret name: DeviceName-BitLockerKeyID-KeyId
                $secretName = "$deviceName-BitLockerKeyID-$($Key.Id)"
                
                # Convert to secure string and save to Key Vault
                $secretValue = ConvertTo-SecureString $ActualKeyValue -AsPlainText -Force
                Set-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName -SecretValue $secretValue -ErrorAction Continue
                
                Write-Verbose "[$KeyCounter/$TotalKeys] Successfully backed up key for device $deviceName to Key Vault"
            }
            catch {
                Write-Warning "[$KeyCounter/$TotalKeys] Failed to backup BitLocker key to Key Vault for device $($Key.DeviceId): $($_.Exception.Message)"
            }
        }
        
        $Key.Key = $KeyValue
        $Key | Add-Member -MemberType NoteProperty -Name 'BitLockerKeyId' -Value $Key.Id
        $Key | Add-Member -MemberType NoteProperty -Name 'BitLockerRecoveryKey' -Value $KeyValue
        $Key | Add-Member -MemberType NoteProperty -Name 'BitLockerDriveType' -Value (Get-DriveTypeName -DriveType $Key.VolumeType)
        $Key | Add-Member -MemberType NoteProperty -Name 'BitLockerBackedUp' -Value (& { if ($Key.CreatedDateTime) { Get-Date($Key.CreatedDateTime) -Format g } else { 'N/A' } })

        #If requested, include the device details
        if ($PSBoundParameters['IncludeDeviceInfo']) {
            Write-Verbose "[$KeyCounter/$TotalKeys] Looking up device information for $($Key.DeviceId)..."
            $Device = $Devices | Where-Object { $Key.DeviceId -eq $_.DeviceId }
            if (!$Device) {
                Write-Warning "[$KeyCounter/$TotalKeys] Device with ID $($Key.DeviceId) not found!"
                $Key | Add-Member -MemberType NoteProperty -Name 'DeviceName' -Value 'Device not found'
                continue
            }
            if ($Device.Id -eq '00000000-0000-0000-0000-000000000000' -or $Device.DeviceId -eq '00000000-0000-0000-0000-000000000000') {
                Write-Warning "[$KeyCounter/$TotalKeys] Stale/dummy device found for key $($Key.DeviceId), skipping..."
                $Key | Add-Member -MemberType NoteProperty -Name 'DeviceName' -Value 'Stale/Dummy Device'
                continue
            }

            # If exporting to Excel, add the BitLocker key details to the device object for comprehensive reporting
            if ($PSBoundParameters['ExportToExcel']) {
                $Device.BitLockerKeyId = $Key.Id
                $Device.BitLockerRecoveryKey = $KeyValue
                $Device.BitLockerDriveType = (Get-DriveTypeName -DriveType $Key.VolumeType)
                $Device.BitLockerBackedUp = (& { if ($Key.CreatedDateTime) { Get-Date($Key.CreatedDateTime) -Format g } else { 'N/A' } })
            }

            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceName' -Value $Device.DisplayName
            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceGUID' -Value $Device.Id #key actually used by the stupid module...
            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceOS' -Value $Device.OperatingSystem
            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceTrustType' -Value $Device.TrustType
            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceMDM' -Value $Device.AdditionalProperties.managementType #can be null! ALWAYS null when using a filter...
            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceCompliant' -Value $Device.isCompliant #can be null!
            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceRegistered' -Value (& { if ($Device.registrationDateTime) { Get-Date($Device.registrationDateTime) -Format g } else { 'N/A' } })
            $Key | Add-Member -MemberType NoteProperty -Name 'DeviceLastActivity' -Value (& { if ($Device.approximateLastSignInDateTime) { Get-Date($Device.approximateLastSignInDateTime) -Format g } else { 'N/A' } })

            #If requested, include the device owner
            if ($PSBoundParameters['IncludeDeviceOwner']) {
                $Key | Add-Member -MemberType NoteProperty -Name 'DeviceOwner' -Value (& { if ($Device.registeredOwners) { $Device.registeredOwners[0].AdditionalProperties.userPrincipalName } else { 'N/A' } })
            }
        }
    }

    # Export results or return data
    if ($PSBoundParameters['ExportToExcel']) {
        try {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $ExcelFilePath = "$($env:userprofile)\${now}_BitLockerKeys_Report.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting BitLocker keys report to Excel file: $ExcelFilePath"
            
            if ($PSBoundParameters['IncludeDeviceInfo']) {
                Write-Verbose 'Exporting comprehensive device report to Excel...'
                # Exclude internal properties from the export
                $ExcludeProps = @(
                    'AdditionalProperties', 'AlternativeSecurityIds', 'complianceExpirationDateTime',
                    'deviceMetadata', 'deviceVersion', 'memberOf', 'PhysicalIds', 'SystemLabels',
                    'transitiveMemberOf', 'RegisteredOwners', 'RegisteredUsers'
                )
                $Devices | Select-Object * -ExcludeProperty $ExcludeProps | 
                Export-Excel -Path $ExcelFilePath -AutoSize -AutoFilter -WorksheetName 'BitLocker-DeviceReport'
            }
            else {
                Write-Verbose 'Exporting BitLocker keys report to Excel...'
                $Keys | Select-Object * -ExcludeProperty Id, VolumeType, AdditionalProperties, CreatedDateTime, Key | 
                Export-Excel -Path $ExcelFilePath -AutoSize -AutoFilter -WorksheetName 'BitLocker-Keys'
            }
            
            Write-Host "✅ Report successfully exported to: $ExcelFilePath" -ForegroundColor Green
            
            if ($PSBoundParameters['RevealKeys']) {
                Write-Warning '⚠️  SECURITY ALERT: The Excel file contains BitLocker recovery keys in PLAIN TEXT!'
                Write-Warning '⚠️  Ensure this file is stored securely and access is properly restricted!'
            }
            else {
                Write-Host '🔒 BitLocker keys are hidden in the export. Use -RevealKeys to display them.' -ForegroundColor Cyan
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
            return $Devices
        }
        else {
            Write-Verbose 'Returning BitLocker key objects'
            return $Keys
        }
    }
}