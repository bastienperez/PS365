<#
    .SYNOPSIS
        Retrieves the Service Connection Point (SCP) for Entra ID Hybrid Join from Active Directory.

    .DESCRIPTION
        This function queries the Active Directory configuration naming context to retrieve
        the Service Connection Point (SCP) object for device registration, which contains
        Azure AD tenant information. If Active Directory is not accessible, it returns
        an object with error information and null values for the SCP data.

    .EXAMPLE
        Get-EntraIDHybridJoinSCP

        Returns the SCP object with keywords containing AzureADName and azureADId when AD is accessible.

    .EXAMPLE
        $result = Get-EntraIDHybridJoinSCP
        if (-not $result.ADAccessible) {
            Write-Warning "Cannot access Active Directory: $($result.ErrorMessage)"
        }

        Tests if Active Directory is accessible and handles the error case.

    .NOTES
        Requires access to Active Directory and the configuration naming context.
        Must be run on a domain-joined computer or with appropriate AD access.
        
        The returned object includes:
        - WhenCreated: Creation date of the SCP object
        - WhenChanged: Last modification date of the SCP object  
        - Keywords: Pipe-separated keywords containing AzureADName and azureADId
        - Path: LDAP path to the SCP object
        - ErrorMessage: Error details if AD is not accessible (null if successful)
        - ADAccessible: Boolean indicating if AD was accessible
#>

function Get-EntraIDHybridJoinSCP {

    try {
        # Test AD accessibility by trying to connect to RootDSE
        $rootDSE = New-Object System.DirectoryServices.DirectoryEntry('LDAP://RootDSE')
        $configNC = $rootDSE.configurationNamingContext
        
        # If configurationNamingContext is null or empty, AD is not accessible
        if ([string]::IsNullOrEmpty($configNC)) {
            throw 'Unable to retrieve configuration naming context from Active Directory'
        }
        
        $rootDSE.Dispose()
        
        # Get SCP object
        $scp = New-Object System.DirectoryServices.DirectoryEntry
        $scp.Path = "LDAP://CN=62a0ff2e-97b9-4513-943f-0d221bd30080,CN=Device Registration Configuration,CN=Services,$configNC"
        
        # Test if we can access the SCP object
        $keywords = $scp.Keywords
        if ($null -eq $keywords -and $scp.Name -eq $null) {
            throw 'Unable to access Service Connection Point object in Active Directory'
        }

        $object = [PSCustomObject][ordered]@{    
            WhenCreated  = $scp.WhenCreated
            WhenChanged  = $scp.WhenChanged
            # keywords is {AzureADName:xx.onmicrosoft.com, azureADId:xxx}
            Keywords     = $scp.Keywords -join '|'
            Path         = $scp.Path
            ErrorMessage = $null
            ADAccessible = $true
        }

        return $object
    }
    catch {
        # Return object indicating AD is not accessible
        $object = [PSCustomObject][ordered]@{    
            WhenCreated  = $null
            WhenChanged  = $null
            Keywords     = $null
            Path         = $null
            ErrorMessage = "Active Directory is not reachable: $($_.Exception.Message)"
            ADAccessible = $false
        }

        return $object
    }
}