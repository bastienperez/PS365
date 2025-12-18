<#
.SYNOPSIS
    Retrieves all Entra ID applications configured for SAML SSO.

.DESCRIPTION
    This function returns a list of all Entra ID applications configured for SAML Single Sign-On
    along with their SAML-related properties, including the PreferredTokenSigningKeyEndDateTime
    and its validity status.

.EXAMPLE
    Get-MgApplicationSAML

    Retrieves all Entra ID applications configured for SAML SSO.

.EXAMPLE
    Get-MgApplicationSAML -ForceNewToken

    Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

.EXAMPLE
    Get-MgApplicationSAML -ExportToExcel

    Gets all SAML applications and exports them to an Excel file.

.NOTES
    More information on: https://itpro-tips.com/get-azure-ad-saml-certificate-details/

    This function requires the Microsoft.Graph.Beta.Applications module to be installed.

    Author: Bastien Perez

    .LIMITATIONS
    The information about the SAML applications clams is not available in the Microsoft Graph API v1 but in https://main.iam.ad.ext.azure.com/api/ApplicationSso/<service-principal-id>/FederatedSsoV2 so we don't get them

    .CHANGELOG
    ## [1.2.0] - 2025-04-04
    ### Changed
    - Change Write-Warning message in the catch block to Import-Module

    ## [1.1.0] - 2025-02-26
    ### Changed
    - Transform the script into a function
    - Add `ForceNewToken` parameter
    - Test if already connected to Microsoft Graph and with the right permissions

    ## [1.0.0] - 2024-xx-xx
    ### Initial Release

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationSAML
#>

function Get-MgApplicationSAML {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )
    
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

    $permissionsNeeded = 'Application.Read.All'
    $permissionMissing = $permissionsNeeded -notin $scopes

    if ($permissionMissing) {
        Write-Verbose "You need to have the $permissionsNeeded permission in the current token, disconnect to force getting a new token with the right permissions"
    }

    if (-not $isConnected) {
        Write-Verbose "Connecting to Microsoft Graph. Scopes: $permissionsNeeded"
        $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
    }
    
    [System.Collections.Generic.List[PSCustomObject]]$samlApplicationsArray = @()
    $samlApplications = Get-MgBetaServicePrincipal -Filter "PreferredSingleSignOnMode eq 'saml'"

    foreach ($samlApp in $samlApplications) {
        $object = [PSCustomObject][ordered]@{
            DisplayName                         = $samlApp.DisplayName
            Id                                  = $samlApp.Id
            AppId                               = $samlApp.AppId
            LoginUrl                            = $samlApp.LoginUrl
            LogoutUrl                           = $samlApp.LogoutUrl
            NotificationEmailAddresses          = $samlApp.NotificationEmailAddresses -join '|'
            AppRoleAssignmentRequired           = $samlApp.AppRoleAssignmentRequired
            PreferredSingleSignOnMode           = $samlApp.PreferredSingleSignOnMode
            PreferredTokenSigningKeyEndDateTime = $samlApp.PreferredTokenSigningKeyEndDateTime
            # PreferredTokenSigningKeyEndDateTime is date time, compared to now and see it is valid
            PreferredTokenSigningKeyValid       = $samlApp.PreferredTokenSigningKeyEndDateTime -gt (Get-Date)
            ReplyUrls                           = $samlApp.ReplyUrls -join '|'
            SignInAudience                      = $samlApp.SignInAudience
        }

        $samlApplicationsArray.Add($object)
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationSAML.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting SAML applications to Excel file: $excelFilePath"
        $samlApplicationsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraSAMLApplications'
        Write-Host -ForegroundColor Green "Export completed successfully!"
    }
    else {
        return $samlApplicationsArray
    }
}