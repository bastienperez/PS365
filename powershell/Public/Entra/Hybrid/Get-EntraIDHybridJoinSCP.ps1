<#
    .SYNOPSIS
        Retrieves the Service Connection Point (SCP) for Entra ID Hybrid Join from Active Directory.

    .DESCRIPTION
        This function queries the Active Directory configuration naming context to retrieve
        the Service Connection Point (SCP) object for device registration, which contains
        Azure AD tenant information.

    .EXAMPLE
        Get-EntraIDHybridJoinSCP

        Returns the SCP object with keywords containing AzureADName and azureADId.

    .NOTES
        Requires access to Active Directory and the configuration naming context.
        Must be run on a domain-joined computer or with appropriate AD access.
#>

function Get-EntraIDHybridJoinSCP {

    # Get configuration naming context without AD PowerShell module
    $rootDSE = New-Object System.DirectoryServices.DirectoryEntry('LDAP://RootDSE')
    $configNC = $rootDSE.configurationNamingContext
    $rootDSE.Dispose()
    
    $scp = New-Object System.DirectoryServices.DirectoryEntry
    $scp.Path = "LDAP://CN=62a0ff2e-97b9-4513-943f-0d221bd30080,CN=Device Registration Configuration,CN=Services,$configNC"
    $scp.Keywords

    $object = [PSCustomObject][ordered]@{    
        WhenCreated = $scp.WhenCreated
        WhenChanged = $scp.WhenChanged
        # keywords is {AzureADName:xx.onmicrosoft.com, azureADId:xxx}
        Keywords    = $scp.Keywords -join '|'
        Path        = $scp.Path
    }

    return $object
}