<#
    .SYNOPSIS
    Retrieves all Entra ID applications and their credentials (key and password).

    .DESCRIPTION
    This function returns a list of all Entra ID applications with their credentials information,
    including key credentials and password credentials, along with their validity status.
    The function also retrieves the owners of each application.

    .EXAMPLE
    Get-MgApplicationCredential
    Retrieves all Entra ID applications and their credentials.

    .EXAMPLE
    Get-MgApplicationCredential -ForceNewToken
    Forces the function to disconnect and reconnect to Microsoft Graph to obtain a new access token.

    .EXAMPLE
    Get-MgApplicationCredential -ExportToExcel
    Gets all application credentials and exports them to an Excel file.

    .NOTES
    Author: Bastien Perez

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationCredential

    .NOTES
#>

function Get-MgApplicationCredential {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    try {
        # for Get-MgApplication/Get-MgApplicationOwner
        Import-Module 'Microsoft.Graph.Applications' -ErrorAction Stop -ErrorVariable mgGraphAppsMissing
    }
    catch {
        if ($mgGraphAppsMissing) {
            Write-Warning "Failed to import Microsoft.Graph.Applications module: $($mgGraphAppsMissing.Exception.Message)"
        }
        if ($mgGraphIdentitySignInsMissing) {
            Write-Warning "Failed to import Microsoft.Graph.Identity.SignIns module: $($mgGraphIdentitySignInsMissing.Exception.Message)"
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

    [System.Collections.Generic.List[PSCustomObject]]$credentialsArray = @()

    $mgApps = Get-MgApplication -All

    foreach ($mgApp in $mgApps) {
        $owner = Get-MgApplicationOwner -ApplicationId $mgApp.Id

        # if severral owners, join them with '|'

        foreach ($keyCredential in $mgApp.KeyCredentials) {
            $object = [PSCustomObject][ordered]@{
                DisplayName           = $mgApp.DisplayName
                CredentialType        = 'KeyCredentials'
                AppId                 = $mgApp.AppId
                CredentialDescription = $keyCredential.DisplayName
                CredentialStartDate   = $keyCredential.StartDateTime
                CredentialExpiryDate  = $keyCredential.EndDateTime
                # CredentialExpiryDate is date time, compared to now and see it is valid
                CredentialValid       = $keyCredential.EndDateTime -gt (Get-Date)
                Type                  = $keyCredential.Type
                Usage                 = $keyCredential.Usage
                Owners                = $owner.AdditionalProperties.userPrincipalName -join '|'
            }

            $credentialsArray.Add($object)
        }

        foreach ($passwordCredential in $mgApp.PasswordCredentials) {
            $object = [PSCustomObject][ordered]@{
                DisplayName           = $mgApp.DisplayName
                CredentialType        = 'PasswordCredentials'
                AppId                 = $mgApp.AppId
                CredentialDescription = $passwordCredential.DisplayName
                CredentialStartDate   = $passwordCredential.StartDateTime
                CredentialExpiryDate  = $passwordCredential.EndDateTime
                # CredentialExpiryDate is date time, compared to now and see it is valid
                CredentialValid       = $passwordCredential.EndDateTime -gt (Get-Date)
                Type                  = 'NA'
                Usage                 = 'NA'
                Owners                = $owner.AdditionalProperties.userPrincipalName -join '|'
            }

            $credentialsArray.Add($object)
        }
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgApplicationCredential.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting application credentials to Excel file: $excelFilePath"
        $credentialsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'EntraApplicationCredentials'
        Write-Host -ForegroundColor Green "Export completed successfully!"
    }
    else {
        return $credentialsArray
    }
}