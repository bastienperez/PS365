<#
    .SYNOPSIS
    Retrieves Entra ID applications with expiring credentials or SAML certificates.

    .DESCRIPTION
    This function identifies all Entra ID applications that have credentials (keys/passwords) or SAML signing certificates 
    that will expire within a specified number of days. It combines data from both regular application credentials 
    and SAML-specific certificates to provide a comprehensive view of expiring security components.

    .PARAMETER DaysUntilExpiry
    The number of days to check for expiring credentials. Applications with credentials expiring within this timeframe will be returned.
    Default is 30 days.

    .PARAMETER ExportToExcel
    Exports the results to an Excel file.

    .EXAMPLE
    MgApplicationExpiringCertAndKey

    Retrieves all applications with credentials or SAML certificates expiring within the next 30 days.

    .EXAMPLE
    MgApplicationExpiringCertAndKey -DaysUntilExpiry 7

    Retrieves all applications with credentials or SAML certificates expiring within the next 7 days.

    .EXAMPLE
    MgApplicationExpiringCertAndKey -DaysUntilExpiry 60 -ExportToExcel

    Gets all applications with credentials expiring within 60 days and exports to Excel.

    .NOTES
    This function requires both Get-MgApplicationCredential and Get-MgApplicationSAML functions to be available.
    
    Author: Bastien Perez

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgApplicationExpiringCertAndKey
#>

function Get-MgApplicationExpiringCertAndKey {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [int]$DaysUntilExpiry = 30,
        
        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    Write-Verbose "Checking for credentials expiring within $DaysUntilExpiry days"

    [System.Collections.Generic.List[PSCustomObject]]$expiringCredentialsArray = @()

    try {
        # Get application credentials
        Write-Verbose 'Retrieving application credentials...'
        $appCredentials = Get-MgApplicationCredential -ForceNewToken:$ForceNewToken

        # Filter credentials expiring within specified days
        $expiringAppCredentials = $appCredentials | Where-Object { 
            $null -ne $_.CredentialExpiresInDays -and 
            $_.CredentialExpiresInDays -le $DaysUntilExpiry 
        }

        foreach ($credential in $expiringAppCredentials) {
            $object = [PSCustomObject][ordered]@{
                DisplayName = $credential.DisplayName
                AppId       = $credential.AppId
                Type        = 'OAuth2'
                ExpiryDate  = $credential.CredentialExpiryDate
            }
            $expiringCredentialsArray.Add($object)
        }

        Write-Verbose "Found $($expiringAppCredentials.Count) expiring application credentials"

        # Get SAML applications
        Write-Verbose 'Retrieving SAML applications...'
        $samlApps = Get-MgApplicationSAML -ForceNewToken:$false  # Don't force token again

        # Filter SAML certificates expiring within specified days
        $expiringSamlCerts = $samlApps | Where-Object { 
            $null -ne $_.SamlSigningCertificateExpiresInDays -and 
            $_.SamlSigningCertificateExpiresInDays -le $DaysUntilExpiry 
        }

        foreach ($samlApp in $expiringSamlCerts) {
            $object = [PSCustomObject][ordered]@{
                DisplayName = $samlApp.DisplayName
                AppId       = $samlApp.AppId
                Type        = 'SAML'
                ExpiryDate  = $samlApp.SamlSigningCertificateEndTime
            }
            $expiringCredentialsArray.Add($object)
        }

        Write-Verbose "Found $($expiringSamlCerts.Count) expiring SAML certificates"

        # Sort by expiry date (most urgent first)
        $expiringCredentialsArray = $expiringCredentialsArray | Sort-Object ExpiryDate

        Write-Host -ForegroundColor Yellow "Found $($expiringCredentialsArray.Count) total expiring credentials/certificates within $DaysUntilExpiry days"

        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $excelFilePath = "$($env:userprofile)\$now-ExpiringCredentials_$($DaysUntilExpiry)days.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting expiring credentials to Excel file: $excelFilePath"
            $expiringCredentialsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExpiringCredentials'
            Write-Host -ForegroundColor Green 'Export completed successfully!'
        }
        else {
            return $expiringCredentialsArray
        }
    }
    catch {
        Write-Error "Error retrieving expiring credentials: $($_.Exception.Message)"
        throw
    }
}