<#
    .SYNOPSIS
    Retrieves the LAPS password for a Microsoft Entra ID device.

    .DESCRIPTION
    Gets the Windows Local Administrator Password Solution (LAPS) password for one or all devices in Microsoft Entra ID (formerly Azure AD).
    By default, only metadata is returned (no password). Use -ShowPassword to retrieve the password in plain text.
    Passwords can optionally be backed up to an Azure Key Vault.

    .PARAMETER DeviceID
    The Microsoft Entra ID (Azure AD) Device ID for which you want to retrieve the LAPS password.
    If not specified, retrieves LAPS passwords for all devices.

    .PARAMETER ShowPassword
    Retrieve and display the LAPS password in plain text.
    By default, only metadata (expiration time, etc.) is returned.
    Use with caution, as this will expose the password in the console output.

    .PARAMETER IncludeHistory
    Include previous LAPS passwords in the output, in addition to the current one.
    Only applicable when -ShowPassword or -BackupToKeyVault is specified. Has no effect otherwise.
    The output includes an IsCurrent property to identify the active password.

    .PARAMETER RunFromAzureAutomation
    Use managed identity authentication instead of interactive authentication.
    Suitable for Azure Automation runbooks, Azure Functions, or VMs with managed identity enabled.

    .PARAMETER BackupToKeyVault
    Enable backup of LAPS passwords to Azure Key Vault.
    Must be used together with -KeyVaultName.
    The secret name is the device name; the Content Type field contains the account name and backup date.

    .PARAMETER KeyVaultName
    Name of the Azure Key Vault to back up LAPS passwords to.
    Mandatory when -BackupToKeyVault is specified.
    Requires the Az.KeyVault module and appropriate permissions.

    .EXAMPLE
    Get-MgLAPSPassword

    Retrieves metadata (no password) for all devices with LAPS configured.

    .EXAMPLE
    Get-MgLAPSPassword -DeviceID "12345678-1234-1234-1234-123456789012"

    Retrieves metadata (no password) for the specified device.

    .EXAMPLE
    Get-MgLAPSPassword -ShowPassword

    Retrieves the current LAPS password in plain text for all devices.

    .EXAMPLE
    Get-MgLAPSPassword -DeviceID "12345678-1234-1234-1234-123456789012" -ShowPassword

    Retrieves the current LAPS password in plain text for the specified device.

    .EXAMPLE
    Get-MgLAPSPassword -DeviceID "12345678-1234-1234-1234-123456789012" -ShowPassword -IncludeHistory

    Retrieves the current and historical LAPS passwords for the specified device.
    The IsCurrent property indicates which entry is the active password.

    .EXAMPLE
    Get-MgLAPSPassword -BackupToKeyVault -KeyVaultName "MyLAPSVault"

    Backs up LAPS passwords for all devices to Azure Key Vault.

    .EXAMPLE
    Get-MgLAPSPassword -DeviceID "12345678-1234-1234-1234-123456789012" -BackupToKeyVault -KeyVaultName "MyLAPSVault"

    Backs up the LAPS password for the specified device to Azure Key Vault.

    .EXAMPLE
    Get-MgLAPSPassword -RunFromAzureAutomation -BackupToKeyVault -KeyVaultName "MyLAPSVault"

    Backs up LAPS passwords for all devices using managed identity authentication. Suitable for Azure Automation runbooks.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgLAPSPassword

    .NOTES
    Requires the DeviceLocalCredential.Read.All and Device.Read.All permissions in Microsoft Entra ID.
#>

function Get-MgLAPSPassword {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Mandatory = $false, HelpMessage = 'The Microsoft Entra ID (Azure AD) Device ID for which you want to retrieve the LAPS password. If not specified, retrieves LAPS passwords for all devices.')]
        [string]$DeviceID,
        [switch]$ShowPassword,
        [switch]$IncludeHistory,

        [Parameter(HelpMessage = 'Use managed identity authentication (for Azure Automation)')]
        [switch]$RunFromAzureAutomation,

        [Parameter(Mandatory, ParameterSetName = 'KeyVault', HelpMessage = 'Enable backup of LAPS passwords to Azure Key Vault')]
        [switch]$BackupToKeyVault,

        [Parameter(Mandatory, ParameterSetName = 'KeyVault', HelpMessage = 'Azure Key Vault name to backup LAPS passwords')]
        [ValidateNotNullOrEmpty()]
        [string]$KeyVaultName
    )

    $fetchPasswords = $ShowPassword.IsPresent -or $BackupToKeyVault.IsPresent

    if ($IncludeHistory.IsPresent -and -not $fetchPasswords) {
        Write-Warning '-IncludeHistory has no effect without -ShowPassword or -BackupToKeyVault.'
    }

    $requiredScopes = @('DeviceLocalCredential.Read.All', 'Device.Read.All')

    # Version check for Azure Automation before connecting
    if ($RunFromAzureAutomation.IsPresent) {
        if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
            $mgAuth = Get-Module 'Microsoft.Graph.Authentication' -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            if ($mgAuth -and [version]$mgAuth.Version -gt [version]'2.25.0') {
                Write-Error "Microsoft.Graph.Authentication v$($mgAuth.Version) is not compatible with Azure Automation on PowerShell $($PSVersionTable.PSVersion). Maximum supported version is 2.25.0. Script execution stopped." -ErrorAction Stop
                return
            }
        }
    }

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)

    if (-not $isConnected) {
        if ($RunFromAzureAutomation.IsPresent) {
            Write-Verbose 'Connecting to Microsoft Graph using Managed Identity'
            Connect-MgGraph -Identity -NoWelcome
        }
        else {
            Write-Verbose "Connecting to Microsoft Graph. Scopes: $($requiredScopes -join ',')"
            $null = Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        }
    }

    # Setup Azure Key Vault connection if backup is requested
    if ($BackupToKeyVault.IsPresent) {
        Write-Verbose 'Setting up Azure Key Vault connection for LAPS password backup...'
        $keyVaultName = $KeyVaultName
        Write-Verbose "Using Key Vault: $keyVaultName"

        try {
            if (-not (Get-Module -ListAvailable -Name Az.KeyVault)) {
                Write-Error 'Az.KeyVault module is required for Key Vault backup. Install it with: Install-Module Az.KeyVault' -ErrorAction Stop
            }

            Write-Verbose 'Connecting to Azure for Key Vault access...'
            try {
                $azContext = Get-AzContext -ErrorAction SilentlyContinue
                if (-not $azContext) {
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

    #Define the URI path
    # Build request headers (used for both list and individual calls)
    $headers = @{
        'ocp-client-name'    = 'Get-LapsAADPassword Windows LAPS Cmdlet'
        'ocp-client-version' = '1.0'
    }

    # Build the list of device credential IDs to process
    $deviceCredentialIds = [System.Collections.Generic.List[string]]@()

    if ($DeviceID) {
        $deviceCredentialIds.Add($DeviceID)
    }
    else {
        Write-Verbose 'No DeviceID specified - retrieving all device LAPS credentials...'
        $listUri = 'v1.0/directory/deviceLocalCredentials'
        $listResponse = Invoke-MgGraphRequest -Method GET -Uri $listUri -Headers $headers -OutputType PSObject
        foreach ($item in $listResponse.value) {
            $deviceCredentialIds.Add($item.id)
        }
        Write-Verbose "Found $($deviceCredentialIds.Count) devices with LAPS credentials"
    }

    foreach ($deviceCredentialId in $deviceCredentialIds) {

        # New correlation ID per request
        $headers['client-request-id'] = [System.Guid]::NewGuid().ToString()

        $uri = 'v1.0/directory/deviceLocalCredentials/' + $deviceCredentialId
        # ?$select=credentials will cause the server to return all credentials, ie latest plus history

        if ($fetchPasswords) {
            $uri = $uri + '?$select=credentials'
        }

        #Initation the request to Microsoft Graph for the LAPS password
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $URI -Headers $headers -OutputType Json
        }
        catch {
            Write-Warning "Device ID: $deviceCredentialId $($_.Exception.Message -replace "`n", ' ' -replace "`r", ' ')"
            $object = [PSCustomObject][ordered]@{
                DeviceName             = '$null'
                DeviceId               = $deviceCredentialId
                PasswordExpirationTime = $null
            }

            $object
            continue
        }

        if ([string]::IsNullOrWhitespace($response)) {
            $object = [PSCustomObject][ordered]@{
                DeviceName             = '$null'
                DeviceId               = $deviceCredentialId
                PasswordExpirationTime = $null
            }

            $object
            continue
        }

        # Build custom PS output object
        $resultsJson = ConvertFrom-Json $response
    
        $lapsDeviceId = $resultsJson.deviceName

        $lapsDeviceId = New-Object([System.Guid])
        $lapsDeviceId = [System.Guid]::Parse($resultsJson.id)

        # Grab password expiration time (only applies to the latest password)
        $lapsPasswordExpirationTime = Get-Date $resultsJson.refreshDateTime

        if ($fetchPasswords) {
            # Copy the credentials array
            $credentials = $resultsJson.credentials

            # Sort the credentials array by backupDateTime.
            $credentials = $credentials | Sort-Object -Property backupDateTime -Descending

            # Note: current password (ie, the one most recently set) is now in the zero position of the array

            # If history was not requested, truncate the credential array down to just the latest one
            if (-not $IncludeHistory) {
                $credentials = @($credentials[0])
            }

            $currentCredential = $credentials[0]

            # When backing up history to Key Vault, process oldest first so the most recent
            # password is written last and becomes the active version of the secret
            $credentialsToProcess = if ($BackupToKeyVault.IsPresent -and $IncludeHistory) {
                @($credentials | Sort-Object -Property backupDateTime)
            }
            else {
                $credentials
            }

            foreach ($credential in $credentialsToProcess) {

                # Cloud returns passwords in base64, decode to plain text
                $plainText = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($credential.passwordBase64))

                # Backup to Key Vault if requested
                if ($BackupToKeyVault.IsPresent) {
                    try {
                        $backupDate = if ($credential.backupDateTime) { (Get-Date $credential.backupDateTime -Format 'yyyy-MM-dd-HHmmss') } else { 'unknown' }
                        $secretName = "LAPS-$($resultsJson.deviceName)" -replace '[^0-9a-zA-Z-]', '-'
                        $contentType = "$backupDate-$($credential.accountName)"

                        Write-Verbose "Backing up LAPS password for $($resultsJson.deviceName) ($($credential.accountName)) to Key Vault '$keyVaultName' with secret name '$secretName' and content type '$contentType'"
                        $existingVersions = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName -IncludeVersions -ErrorAction SilentlyContinue
                        $alreadyBackedUp = $existingVersions | Where-Object { $_.ContentType -eq $contentType }
                        if ($alreadyBackedUp) {
                            Write-Host "Secret '$secretName' with ContentType '$contentType' already exists in Key Vault, skipping..." -ForegroundColor Yellow
                        }
                        else {
                            $secretValue = ConvertTo-SecureString $plainText -AsPlainText -Force
                            $notBefore = if ($credential.backupDateTime) { (Get-Date $credential.backupDateTime).ToUniversalTime() } else { $null }
                            $setParams = @{
                                VaultName   = $keyVaultName
                                Name        = $secretName
                                SecretValue = $secretValue
                                ContentType = $contentType
                                ErrorAction = 'Continue'
                            }
                            if ($notBefore) { $setParams['NotBefore'] = $notBefore }
                            $null = Set-AzKeyVaultSecret @setParams
                            Write-Verbose "Successfully backed up LAPS password for $($resultsJson.deviceName) ($($credential.accountName)) to Key Vault"
                        }
                    }
                    catch {
                        Write-Warning "Failed to backup LAPS password to Key Vault: $($_.Exception.Message)"
                    }
                }
                else {
                    $object = [PSCustomObject][ordered]@{
                        DeviceName             = $resultsJson.deviceName
                        DeviceId               = $lapsDeviceId
                        Account                = $credential.accountName
                        IsCurrent              = ($credential -eq $currentCredential)
                        Password               = $plainText
                        PasswordExpirationTime = $lapsPasswordExpirationTime
                        PasswordUpdateTime     = if ($credential.backupDateTime) { Get-Date $credential.backupDateTime } else { $null }
                    }

                    $object
                }
            }
        }
        else {
            # Output a single object that just displays latest password expiration time
            # Note, $IncludeHistory is ignored even if specified in this case
            $object = [PSCustomObject][ordered]@{
                DeviceName             = $resultsJson.deviceName
                DeviceId               = $lapsDeviceId
                Password               = '[HIDDEN - Use -ShowPassword to display]'
                PasswordExpirationTime = $lapsPasswordExpirationTime
            }

            $object
        }
    } # end foreach deviceCredentialId
}