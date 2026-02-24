<#
    .SYNOPSIS
    Retrieves all Microsoft Entra ID applications and their credentials (key and password).

    .DESCRIPTION
    This function returns a list of all Microsoft Entra ID applications with their credentials information,
    including key credentials and password credentials, along with their validity status.
    The function also retrieves the owners of each application.
    
    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER ForceNewToken
    (Optional) Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .PARAMETER RunFromAzureAutomation
    (Optional) If specified, uses managed identity authentication instead of interactive authentication.
    This is useful when running the script in Azure environments like Azure Functions, Logic Apps, or VMs with managed identity enabled.
    When this parameter is used, ExpirationThresholdDays, NotificationRecipient and NotificationSender are required.

    PowerShell modules used in Azure Automation must be a MAXIMUM of version 2.25.0 when using PowerShell < 7.4.0, because starting from version 2.26.0, PowerShell 7.4.0 is required, and Azure Automation does not support it yet as of February 2026. For PowerShell 7.4.0+, there are no version restrictions.
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3147
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3151
    https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3166

    .PARAMETER ExpirationThresholdDays
    (Required when RunFromAzureAutomation is enabled) Number of days threshold for expiration notification. Default is 30 days.

    .PARAMETER NotificationRecipient
    (Required when RunFromAzureAutomation is enabled) Email address to receive expiration notifications.

    .PARAMETER NotificationSender
    (Required when RunFromAzureAutomation is enabled) Email address of the sender for expiration notifications.

    .EXAMPLE
    Get-MgApplicationCredential
    Retrieves all Microsoft Entra ID applications and their credentials.

    .EXAMPLE
    Get-MgApplicationCredential -ForceNewToken

    Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .EXAMPLE
    Get-MgApplicationCredential -ExportToExcel

    Gets all application credentials and exports them to an Excel file.

    .EXAMPLE
    Gets all application credentials using managed identity, exports them to an Excel file, and sends notification for credentials expiring within 15 days.

    .EXAMPLE
    Get-MgApplicationCredential -RunFromAzureAutomation -ExpirationThresholdDays 30 -NotificationRecipient 'admin@company.com' -NotificationSender 'automation@company.com'

    Gets all application credentials using managed identity authentication and sends notification for credentials expiring within 30 days.

    .EXAMPLE
    Get-MgApplicationCredential -RunFromAzureAutomation -ExpirationThresholdDays 7 -NotificationRecipient 'admin@company.com' -NotificationSender 'automation@company.com'

    Gets all application credentials using managed identity and sends email notification for credentials expiring within 7 days.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationCredential

#>

function Get-MgApplicationCredential {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$RunFromAzureAutomation,

        [Parameter(Mandatory = $false)]
        [int]$ExpirationThresholdDays = 30,

        [Parameter(Mandatory = $false)]
        [string]$NotificationRecipient,

        [Parameter(Mandatory = $false)]
        [string]$NotificationSender
    )

    # Validate notification parameters
    if ($RunFromAzureAutomation.IsPresent) {
        if ([string]::IsNullOrWhiteSpace($NotificationRecipient)) {
            Write-Error 'NotificationRecipient parameter is required when RunFromAzureAutomation is enabled.'
            return
        }
        if ([string]::IsNullOrWhiteSpace($NotificationSender)) {
            Write-Error 'NotificationSender parameter is required when RunFromAzureAutomation is enabled.'
            return
        }
        if ($ExpirationThresholdDays -le 0) {
            Write-Error 'ExpirationThresholdDays must be greater than 0 when RunFromAzureAutomation is enabled.'
            return
        }

        try {
            Import-Module 'Microsoft.Graph.Users.Actions' -ErrorAction Stop -ErrorVariable mgGraphMailMissing
        }
        catch {
            if ($mgGraphMailMissing) {
                Write-Warning "Failed to import Microsoft.Graph.Users.Actions module: $($mgGraphMailMissing.Exception.Message)"
            }

            return
        }
    }

    try {
        # for Get-MgApplication/Get-MgApplicationOwner
        Import-Module 'Microsoft.Graph.Applications' -ErrorAction Stop -ErrorVariable mgGraphAppsMissing
        # for Send-MgMail (when RunFromAzureAutomation is enabled)
    }
    catch {
        if ($mgGraphAppsMissing) {
            Write-Warning "Failed to import Microsoft.Graph.Applications module: $($mgGraphAppsMissing.Exception.Message)"
        }

        return
    }

    $isConnected = $false

    $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
    
    if ($ForceNewToken.IsPresent) {
        Write-Verbose 'Disconnecting from Microsoft Graph'
        $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
        $isConnected = $false
    }
    
    $scopes = (Get-MgContext).Scopes

    $permissionsNeeded = @('Application.Read.All')
    if ($RunFromAzureAutomation.IsPresent) {
        $permissionsNeeded += 'Mail.Send'
    }
    
    $permissionMissing = $permissionsNeeded | Where-Object { $_ -notin $scopes }

    if ($permissionMissing) {
        Write-Verbose "You need to have the $permissionsNeeded permission in the current token, disconnect to force getting a new token with the right permissions"
    }

    # Version check for Azure Automation before connecting
    if ($RunFromAzureAutomation.IsPresent) {
        # Only check module version if PowerShell < 7.4 (Azure Automation limitation)
        if ($PSVersionTable.PSVersion -lt [version]'7.4.0') {
            $mgAuth = Get-Module 'Microsoft.Graph.Authentication' -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            if ($mgAuth -and [version]$mgAuth.Version -gt [version]'2.25.0') {
                Write-Error "Microsoft.Graph.Authentication v$($mgAuth.Version) is not compatible with Azure Automation on PowerShell $($PSVersionTable.PSVersion). Maximum supported version is 2.25.0. Script execution stopped."
                return
            }
        }
    }

    if (-not $isConnected) {
        if ($RunFromAzureAutomation.IsPresent) {
            Write-Verbose 'Connecting to Microsoft Graph using Managed Identity'
            Connect-MgGraph -Identity -NoWelcome
        }
        else {
            Write-Verbose "Connecting to Microsoft Graph. Scopes: $($permissionsNeeded -join ',')"
            $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
        }
    }

    [System.Collections.Generic.List[PSCustomObject]]$credentialsArray = @()

    $mgApps = Get-MgApplication -All

    foreach ($mgApp in $mgApps) {
        $owner = Get-MgApplicationOwner -ApplicationId $mgApp.Id

        # if severral owners, join them with '|'

        foreach ($keyCredential in $mgApp.KeyCredentials) {
            $object = [PSCustomObject][ordered]@{
                DisplayName             = $mgApp.DisplayName
                CredentialType          = 'KeyCredentials'
                AppId                   = $mgApp.AppId
                CredentialDescription   = $keyCredential.DisplayName
                CredentialStartDate     = $keyCredential.StartDateTime
                CredentialExpiryDate    = $keyCredential.EndDateTime
                # CredentialExpiryDate is date time, compared to now and see it is valid
                CredentialValid         = $keyCredential.EndDateTime -gt (Get-Date)
                CredentialExpiresInDays = if ($keyCredential.EndDateTime) { [int](New-TimeSpan -Start (Get-Date) -End $keyCredential.EndDateTime).TotalDays } else { $null }
                Type                    = $keyCredential.Type
                Usage                   = $keyCredential.Usage
                Owners                  = $owner.AdditionalProperties.userPrincipalName -join '|'
            }

            $credentialsArray.Add($object)
        }

        foreach ($passwordCredential in $mgApp.PasswordCredentials) {
            $object = [PSCustomObject][ordered]@{
                DisplayName             = $mgApp.DisplayName
                CredentialType          = 'PasswordCredentials'
                AppId                   = $mgApp.AppId
                CredentialDescription   = $passwordCredential.DisplayName
                CredentialStartDate     = $passwordCredential.StartDateTime
                CredentialExpiryDate    = $passwordCredential.EndDateTime
                # CredentialExpiryDate is date time, compared to now and see it is valid
                CredentialValid         = $passwordCredential.EndDateTime -gt (Get-Date)
                CredentialExpiresInDays = if ($passwordCredential.EndDateTime) { [int](New-TimeSpan -Start (Get-Date) -End $passwordCredential.EndDateTime).TotalDays } else { $null }
                Type                    = 'NA'
                Usage                   = 'NA'
                Owners                  = $owner.AdditionalProperties.userPrincipalName -join '|'
            }

            $credentialsArray.Add($object)
        }
    }
    
    # Check for expiring credentials and send notification if enabled
    if ($RunFromAzureAutomation.IsPresent) {
        $expiringCredentials = $credentialsArray | Where-Object { 
            $null -ne $_.CredentialExpiresInDays -and 
            $_.CredentialExpiresInDays -le $ExpirationThresholdDays
        }
        
        # Calculate statistics for different expiration categories
        $expiredCredentials = $credentialsArray | Where-Object { 
            $null -ne $_.CredentialExpiresInDays -and $_.CredentialExpiresInDays -le 0 
        }
        $credentialsExpiring15Days = $credentialsArray | Where-Object { 
            $null -ne $_.CredentialExpiresInDays -and $_.CredentialExpiresInDays -le 15 -and $_.CredentialExpiresInDays -gt 0 
        }
        $credentialsExpiring30Days = $credentialsArray | Where-Object { 
            $null -ne $_.CredentialExpiresInDays -and $_.CredentialExpiresInDays -le 30 -and $_.CredentialExpiresInDays -gt 0 
        }
        
        if ($expiringCredentials.Count -gt 0) {
            Write-Verbose "Found $($expiringCredentials.Count) credentials expiring within $ExpirationThresholdDays days. Sending notification email."
            
            $expiringCredentials = $expiringCredentials | Sort-Object CredentialExpiresInDays
            
            $emailBody = @"
<!DOCTYPE html>
<html>
<head>
<title>Microsoft Entra ID Application Credentials Expiration Alert</title>
<style>
    body { 
        font-family: Segoe UI, SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif; 
        margin: 0; 
        padding: 20px; 
        color: #11100f; 
        font-size: 14px;
        line-height: 20px;
        background-color: #ffffff;
    }
    
    .summary-container {
        width: 100%;
        margin-bottom: 30px;
    }
    
    .summary-table {
        width: 80%;
        border-collapse: separate;
        border-spacing: 8px;
        margin: 0 auto;
    }
    
    .summary-box {
        background-color: #fff9f5;
        border: 1px solid #8a6700;
        padding: 10px;
        border-radius: 6px;
        text-align: center;
        width: 25%;
    }
    
    .summary-box.expired {
        background-color: #fde7e9;
        border-color: #a4262c;
    }
    
    .summary-box.warning {
        background-color: #fff4ce;
        border-color: #8a6700;
    }
    
    .summary-count {
        font-size: 20px;
        font-weight: bold;
        color: #d73502;
        margin: 6px 0;
    }
    
    h2 { 
        padding-top: 0; 
        margin: 0 0 16px 0; 
        font-family: "Segoe UI Semibold", SegoeUISemibold, "Segoe UI", SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif;
        font-weight: 600; 
        font-size: 20px; 
        line-height: 28px;
        color: #323130;
    }
    
    table { 
        border-spacing: 0; 
        border-collapse: collapse; 
        width: 100%; 
        margin-bottom: 20px;
        background-color: #ffffff;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }
    
    th { 
        vertical-align: bottom;
        color: #ffffff;
        background-color: #323130;
        padding: 12px 8px;
        text-align: left;
        font-family: "Segoe UI Semibold", SegoeUISemibold, "Segoe UI", SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif;
        font-weight: 600;
        font-size: 13px;
        word-wrap: break-word;
    }
    
    td { 
        vertical-align: top;
        color: #11100f;
        padding: 12px 8px;
        border-bottom: solid 1px #edebe9;
        word-wrap: break-word;
        font-size: 13px;
    }
    
    .critical { background-color: #fde7e9; color: #a4262c; }
    .warning { background-color: #fff4ce; color: #8a6700; }
    .caution { background-color: #fff9f5; color: #8a6700; }
    
    .footer {
        margin-top: 30px;
        padding: 20px;
        background-color: #faf9f8;
        border-radius: 8px;
        border-top: 3px solid #0078d4;
    }
    
    .footer p {
        margin: 8px 0;
        font-size: 13px;
        color: #605e5c;
    }
    
    .action-required {
        font-weight: 600;
        color: #d73502;
    }
</style>
</head>
<body>
    <div class="summary-container">
        <table class="summary-table">
            <tr>
                <td class="summary-box expired">
                    <div class="summary-count">$($expiredCredentials.Count)</div>
                    <div>secrets already expired</div>
                </td>
                <td class="summary-box warning">
                    <div class="summary-count">$($credentialsExpiring15Days.Count)</div>
                    <div>secrets expire within 15 days</div>
                </td>
                <td class="summary-box">
                    <div class="summary-count">$($credentialsExpiring30Days.Count)</div>
                    <div>secrets expire within 30 days</div>
                </td>
            </tr>
        </table>
    </div>
    
    <h2>Credentials Requiring Attention</h2>
    <table>
        <tr>
            <th>Application</th>
            <th>Credential Type</th>
            <th>Description</th>
            <th>Expires In (Days)</th>
            <th>Expiry Date</th>
            <th>Owners</th>
        </tr>
"@
            
            foreach ($cred in $expiringCredentials) {
                $rowClass = if ($cred.CredentialExpiresInDays -le 0) { 'critical' } elseif ($cred.CredentialExpiresInDays -le 14) { 'warning' } else { 'caution' }
                $expiresInDaysDisplay = if ($cred.CredentialExpiresInDays -le 0) { "$($cred.CredentialExpiresInDays) (already expired)" } else { $cred.CredentialExpiresInDays }
                $emailBody += "<tr class=`"$rowClass`"><td>$($cred.DisplayName)</td><td>$($cred.CredentialType)</td><td>$($cred.CredentialDescription)</td><td><strong>$expiresInDaysDisplay</strong></td><td>$($cred.CredentialExpiryDate)</td><td>$($cred.Owners)</td></tr>"
            }
            
            $emailBody += @"
    </table>
    
    <div class="footer">
        <p class="action-required">Action Required:</p>
        <p>Please review and renew these credentials before they expire to avoid service disruptions.</p>
        <hr style="border: none; border-top: 1px solid #d2d0ce; margin: 15px 0;">
        <p><em>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Get-MgApplicationCredential</em></p>
    </div>
</body>
</html>
"@
            
            try {
                $params = @{
                    Message         = @{
                        Subject      = "Microsoft Entra ID Application Credentials Expiring ($($expiringCredentials.Count) credentials)"
                        Body         = @{
                            ContentType = 'HTML'
                            Content     = $emailBody
                        }
                        ToRecipients = @(
                            @{
                                EmailAddress = @{
                                    Address = $NotificationRecipient
                                }
                            }
                        )
                    }
                    SaveToSentItems = 'false'
                }
                
                Send-MgUserMail -UserId $NotificationSender -BodyParameter $params
                Write-Host -ForegroundColor Green "✓ Expiration notification email sent successfully to $NotificationRecipient"
            }
            catch {
                Write-Warning "Failed to send notification email: $($_.Exception.Message)"
            }
        }
        else {
            Write-Verbose "No credentials found expiring within $ExpirationThresholdDays days."
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationCredential.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting application credentials to Excel file: $excelFilePath"
        $credentialsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraApplicationCredentials'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif (-not $RunFromAzureAutomation.IsPresent) {
        return $credentialsArray
    }
}