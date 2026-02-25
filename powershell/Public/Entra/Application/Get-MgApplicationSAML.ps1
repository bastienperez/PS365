<#
    .SYNOPSIS
    Retrieves all Entra ID applications configured for SAML SSO.

    .DESCRIPTION
    This function returns a list of all Entra ID applications configured for SAML Single Sign-On
    along with their SAML-related properties, including the PreferredTokenSigningKeyEndDateTime
    and its validity status.

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
    Get-MgApplicationSAML

    Retrieves all Entra ID applications configured for SAML SSO.

    .EXAMPLE
    Get-MgApplicationSAML -ForceNewToken

    Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .EXAMPLE
    Get-MgApplicationSAML -ExportToExcel

    Gets all SAML applications and exports them to an Excel file.

    .EXAMPLE
    Get-MgApplicationSAML -RunFromAzureAutomation -ExpirationThresholdDays 30 -NotificationRecipient 'admin@company.com' -NotificationSender 'automation@company.com'

    Gets all SAML applications using managed identity authentication and sends notification for certificates expiring within 30 days.

    .EXAMPLE
    Get-MgApplicationSAML -RunFromAzureAutomation -ExpirationThresholdDays 7 -NotificationRecipient 'admin@company.com' -NotificationSender 'automation@company.com'

    Gets all SAML applications using managed identity and sends email notification for certificates expiring within 7 days.

    .NOTES
    More information on: https://itpro-tips.com/get-azure-ad-saml-certificate-details/

    This function requires the Microsoft.Graph.Beta.Applications module to be installed.

    .NOTES
    Limitations:
    The information about the SAML applications clams is not available in the Microsoft Graph API v1 but in https://main.iam.ad.ext.azure.com/api/ApplicationSso/<service-principal-id>/FederatedSsoV2 so we don't get them

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationSAML
#>

function Get-MgApplicationSAML {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

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
        # At the date of writing (december 2023), PreferredTokenSigningKeyEndDateTime parameter is only on Beta profile
        Import-Module 'Microsoft.Graph.Beta.Applications' -ErrorAction Stop -ErrorVariable mgGraphAppsMissing
    }
    catch {
        if ($mgGraphAppsMissing) {
            Write-Warning "Please install the Microsoft.Graph.Beta.Applications module: $($mgGraphAppsMissing.Exception.Message)"
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
        Write-Verbose "You need to have the $($permissionsNeeded -join ',') permission in the current token, disconnect to force getting a new token with the right permissions"
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

    [System.Collections.Generic.List[PSCustomObject]]$samlApplicationsArray = @()
    $samlApplications = Get-MgBetaServicePrincipal -Filter "PreferredSingleSignOnMode eq 'saml'"

    foreach ($samlApp in $samlApplications) {
        # Reset to $null before each call: prevents previous iteration's value from bleeding through on silent errors
        $ownerObjects = $null
        $ownerString = $null
        $ownerObjects = Get-MgServicePrincipalOwner -ServicePrincipalId $samlApp.Id -ErrorAction SilentlyContinue

        # Build owners string: DisplayName for each owner, joined with '|'
        if ($ownerObjects) {
            $ownerString = ($ownerObjects | ForEach-Object {
                    $props = $_.AdditionalProperties
                    if ($props.ContainsKey('displayName') -and -not [string]::IsNullOrEmpty($props['displayName'])) { $props['displayName'] }
                    else { $_.Id }
                }) -join '|'
        }

        $object = [PSCustomObject][ordered]@{
            DisplayName                         = $samlApp.DisplayName
            Id                                  = $samlApp.Id
            AppId                               = $samlApp.AppId
            EntraUrl                            = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($samlApp.Id)/appId/$($samlApp.AppId)"
            LoginUrl                            = $samlApp.LoginUrl
            LogoutUrl                           = $samlApp.LogoutUrl
            NotificationEmailAddresses          = $samlApp.NotificationEmailAddresses -join '|'
            AppRoleAssignmentRequired           = $samlApp.AppRoleAssignmentRequired
            PreferredSingleSignOnMode           = $samlApp.PreferredSingleSignOnMode
            SamlSigningCertificateEndTime       = $samlApp.PreferredTokenSigningKeyEndDateTime
            # PreferredTokenSigningKeyEndDateTime is date time, compared to now and see it is valid
            SamlSigningCertificateValid         = $samlApp.PreferredTokenSigningKeyEndDateTime -gt (Get-Date)
            SamlSigningCertificateExpiresInDays = if ($samlApp.PreferredTokenSigningKeyEndDateTime) { [int](New-TimeSpan -Start (Get-Date) -End $samlApp.PreferredTokenSigningKeyEndDateTime).TotalDays } else { $null }
            ReplyUrls                           = $samlApp.ReplyUrls -join '|'
            SignInAudience                      = $samlApp.SignInAudience
            Owners                              = $ownerString
        }

        $samlApplicationsArray.Add($object)
    }

    # Check for expiring certificates and send notification if enabled
    if ($RunFromAzureAutomation.IsPresent) {
        $expiringCertificates = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and
            $_.SamlSigningCertificateExpiresInDays -le $ExpirationThresholdDays
        }

        # Calculate statistics for different expiration categories
        $expiredCertificates = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and $_.SamlSigningCertificateExpiresInDays -le 0
        }
        $certificatesExpiring15Days = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and $_.SamlSigningCertificateExpiresInDays -le 15 -and $_.SamlSigningCertificateExpiresInDays -gt 0
        }
        $certificatesExpiring30Days = $samlApplicationsArray | Where-Object {
            $null -ne $_.SamlSigningCertificateExpiresInDays -and $_.SamlSigningCertificateExpiresInDays -le 30 -and $_.SamlSigningCertificateExpiresInDays -gt 0
        }

        Write-Verbose "Sending notification email. Found $($expiringCertificates.Count) SAML certificates expiring within $ExpirationThresholdDays days."

        $expiringCertificates = $expiringCertificates | Sort-Object SamlSigningCertificateExpiresInDays

        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
<title>Microsoft Entra ID SAML Application Certificates Expiration Alert</title>
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
        vertical-align: middle;
        color: #ffffff;
        background-color: #323130;
        padding: 3px 8px;
        text-align: left;
        font-family: "Segoe UI Semibold", SegoeUISemibold, "Segoe UI", SegoeUI, Roboto, "Helvetica Neue", Arial, sans-serif;
        font-weight: 600;
        font-size: 12px;
        line-height: 16px;
        word-wrap: break-word;
    }
    
    td { 
        vertical-align: middle;
        color: #11100f;
        padding: 3px 8px;
        border-bottom: solid 1px #c8c6c4;
        word-wrap: break-word;
        font-size: 12px;
        line-height: 16px;
    }
    
    .critical { background-color: #FFF0F0; color: #A80000; }
    .warning { background-color: #FDEFD0; color: #7A3A00; }
    .caution { background-color: #CCE4FF; color: #003882; }
    
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
    <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:12px;background:transparent;box-shadow:none;" role="presentation">
        <tr>
            <td width="33%" valign="top" style="width:33%;padding:4pt 3pt 4pt 5pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#FFF0F0;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#A80000;font-weight:600;">Expired certificates</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#A80000;font-weight:bold;">$($expiredCertificates.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#A80000;border-bottom:none;vertical-align:middle;">
                                        certificates already expired
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="34%" valign="top" style="width:34%;padding:4pt 3pt 4pt 3pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#FDEFD0;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#7A3A00;font-weight:600;">Expiring within 15 days</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#7A3A00;font-weight:bold;">$($certificatesExpiring15Days.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#7A3A00;border-bottom:none;vertical-align:middle;">
                                        certificates expire within 15 days
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
            <td width="33%" valign="top" style="width:33%;padding:4pt 5pt 4pt 3pt;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;background:#CCE4FF;border-collapse:collapse;margin-bottom:0;box-shadow:none;" role="presentation">
                    <tr>
                        <td valign="top" style="padding:6pt 8pt 6pt 8pt;border-bottom:none;">
                            <h4 align="center" style="margin:0 0 5pt 0;text-align:center;line-height:14pt;font-size:11pt;font-family:'Segoe UI Semibold',sans-serif;color:#003882;font-weight:600;">Expiring within 30 days</h4>
                            <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%;border-collapse:collapse;margin-bottom:0;background:transparent;box-shadow:none;" role="presentation">
                                <tr>
                                    <td width="50%" valign="top" style="width:50%;padding:2pt 0 2pt 0;text-align:right;border-bottom:none;">
                                        <span style="font-size:18pt;font-family:'Segoe UI',sans-serif;color:#003882;font-weight:bold;">$($certificatesExpiring30Days.Count)</span>
                                    </td>
                                    <td width="50%" valign="middle" style="width:50%;padding:2pt 0 2pt 6pt;font-size:9pt;font-family:'Segoe UI',sans-serif;color:#003882;border-bottom:none;vertical-align:middle;">
                                        certificates expire within 30 days
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    
    <h2>SAML Certificates requiring attention</h2>
    <table>
        <tr>
            <th>Application</th>
            <th>App ID</th>
            <th>Expires In (Days)</th>
            <th>Expiry Date</th>
            <th>Login URL</th>
            <th>Owners</th>
        </tr>
"@

        foreach ($cert in $expiringCertificates) {
            $rowClass = if ($cert.SamlSigningCertificateExpiresInDays -le 0) { 'critical' } elseif ($cert.SamlSigningCertificateExpiresInDays -le 14) { 'warning' } else { 'caution' }
            $expiresInDaysDisplay = if ($cert.SamlSigningCertificateExpiresInDays -lt 0) { "$($cert.SamlSigningCertificateExpiresInDays) (already expired)" } else { $cert.SamlSigningCertificateExpiresInDays }
            $appLink = "<strong style=`"color:#11100f;font-size:12px;line-height:16px;`">$($cert.DisplayName)</strong> <a href=`"$($cert.EntraUrl)`" style=`"text-decoration:none;font-size:14px;line-height:16px;`" title=`"Open in Entra`">&#x1F517;</a>"
            $ownersList = @($cert.Owners -split '\|' | Where-Object { $_ })
            $ownersHtml = if ($ownersList.Count -gt 1) { ($ownersList | ForEach-Object { "- $_" }) -join '<br>' } else { $ownersList -join '' }
            $emailBody += "<tr class=`"$rowClass`"><td>$appLink</td><td>$($cert.AppId)</td><td><strong>$expiresInDaysDisplay</strong></td><td>$($cert.SamlSigningCertificateEndTime)</td><td>$($cert.LoginUrl)</td><td>$ownersHtml</td></tr>"
        }

        if ($expiringCertificates.Count -eq 0) {
            $emailBody += '<tr><td colspan="6" style="text-align:center;padding:12px 8px;color:#605e5c;font-style:italic;">No certificates requiring attention - all SAML signing certificates are healthy.</td></tr>'
        }

        $emailFooter = if ($expiringCertificates.Count -gt 0) {
            '<p class="action-required">Action Required:</p><p>Please review and renew these SAML signing certificates before they expire to avoid authentication disruptions.</p>'
        }
        else {
            '<p style="color:#107C10;font-weight:600;">&#10003; All SAML signing certificates are healthy. No action required at this time.</p>'
        }

        $emailSubject = if ($expiringCertificates.Count -gt 0) {
            "Microsoft Entra ID SAML Certificates Expiring ($($expiringCertificates.Count) certificates)"
        }
        else {
            'Microsoft Entra ID SAML Certificates - All Healthy'
        }

        $emailBody += @"
    </table>

    <div class="footer">
        $emailFooter
        <hr style="border: none; border-top: 1px solid #d2d0ce; margin: 15px 0;">
        <p><em>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Get-MgApplicationSAML v0.71.0</em></p>
    </div>
</body>
</html>
"@

        try {
            $params = @{
                Message         = @{
                    Subject      = $emailSubject
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
            Write-Host -ForegroundColor Green "Expiration notification email sent successfully to $NotificationRecipient"
        }
        catch {
            Write-Warning "Failed to send notification email: $($_.Exception.Message)"
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationSAML.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting SAML applications to Excel file: $excelFilePath"
        $samlApplicationsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraSAMLApplications'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    elseif (-not $RunFromAzureAutomation.IsPresent) {
        return $samlApplicationsArray
    }
}
