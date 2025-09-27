function Get-IntuneRole {
    <#
    .SYNOPSIS
        Retrieves Intune roles from Microsoft Graph using direct Graph API requests.
    
    .DESCRIPTION
        The Get-IntuneRole function retrieves all roles from Microsoft Intune via the Microsoft Graph API
        using Invoke-MgGraphRequest. It can filter results based on whether roles are built-in or custom.
    
    .PARAMETER RoleType
        Specifies which types of roles to retrieve:
        - All: Retrieves both built-in and custom roles (default)
        - CustomOnly: Retrieves only custom roles
        - BuiltinOnly: Retrieves only built-in roles
    
    .PARAMETER MgGraphScope
        Specifies the Microsoft Graph scope required for accessing Intune roles.
        Default is "DeviceManagementRBAC.Read.All".
    
    .EXAMPLE
        Get-IntuneRole
        Retrieves all Intune roles (both built-in and custom).
    
    .EXAMPLE
        Get-IntuneRole -Type CustomOnly
        Retrieves only custom Intune roles.
    
    .EXAMPLE
        Get-IntuneRole -Type BuiltinOnly
        Retrieves only built-in Intune roles.
    
    .NOTES
        Author: Bastien Perez
        Date: August 13, 2025
    #>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateSet('All', 'CustomOnly', 'BuiltinOnly')]
        [string]$Type = 'All'
    )
    
    begin {

        $scopes = 'DeviceManagementRBAC.Read.All'
        # Check if the required Microsoft.Graph.Authentication module is available
        if (-not (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)) {
            Write-Error 'The Microsoft.Graph.Authentication module is required but not installed. Please install it using: Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser'
            return
        }
        
        # Import the required module
        Import-Module -Name Microsoft.Graph.Authentication -ErrorAction Stop
        
        # Check if we're connected to Microsoft Graph with the required scope
        try {
            $graphConnection = Get-MgContext
            if (-not $graphConnection -or $graphConnection.Scopes -notcontains $scopes) {
                Write-Verbose "Not connected to Microsoft Graph, connecting with scope: '$scopes'"

                Connect-MgGraph -Scopes $scopes -NoWelcome
            }
            elseif ($graphConnection.Scopes -notcontains $scopes) {
                Write-Verbose "Connected to Microsoft Graph, but missing required scope: '$scopes'. Disconnecting and reconnecting with the correct scopes."
                Disconnect-MgGraph
                
                Connect-MgGraph -Scopes $scopes -NoWelcome
            }            
        }
        catch {
            Write-Error "Error checking Microsoft Graph connection: $_"
            return
        }

        [System.Collections.Generic.List[PSCustomObject]]$IntuneRolesArray = @()
    }
    
    process {
        try {
            # Initialize empty array for results
            $allRoles = [System.Collections.Generic.List[PSCustomObject]]::new()
            
            # API URL for Intune role definitions
            # beta contains scopeTagIds
            $apiUrl = 'https://graph.microsoft.com/beta/deviceManagement/roleDefinitions'
            $nextLink = $apiUrl
            
            # Retrieve all pages of results
            while ($null -ne $nextLink) {
                Write-Verbose "Requesting data from: $nextLink"
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink
                
                if ($response.value) {
                    foreach ($role in $response.value) {
                        $allRoles.Add($role)
                    }
                }
                
                # Check if there are more pages
                $nextLink = $response.'@odata.nextLink'
            }
            
            if ($allRoles.Count -eq 0) {
                Write-Warning 'No Intune roles found.'
                return
            }
            
            Write-Verbose "Retrieved $($allRoles.Count) roles from Intune."
            
            # Filter roles based on the RoleType parameter
            $resultRoles = switch ($Type) {
                'CustomOnly' {
                    Write-Verbose 'Filtering to show custom roles only'
                    $allRoles | Where-Object { -not $_.isBuiltIn }
                    break
                }
                'BuiltinOnly' {
                    Write-Verbose 'Filtering to show built-in roles only'
                    $allRoles | Where-Object { $_.isBuiltIn }
                    break
                }
                default {
                    Write-Verbose 'Returning all roles'
                    $allRoles
                    break
                }
            }
            
            foreach ($role in $resultRoles) {
                if ($role.isBuiltIn) {
                    $roleType = 'Built-in'
                }
                else {
                    $roleType = 'Custom'
                }

                $object = [PSCustomObject][ordered]@{
                    Id                        = $role.id
                    DisplayName               = $role.displayName
                    Description               = $role.description
                    IsBuiltIn                 = $role.isBuiltIn
                    RoleType                  = $roleType
                    ScopeTagIds               = $role.roleScopeTagIds -join '|'
                    # is hashtable, convert to string separated by |
                    AllowedResourceActions    = ($role.rolePermissions.resourceActions.allowedResourceActions -join '|')
                    NotAllowedResourceActions = ($role.rolePermissions.resourceActions.notAllowedResourceActions -join '|')
                }

                $IntuneRolesArray.Add($object)
            }
            
            # Return the filtered roles
            return $IntuneRolesArray
        }
        catch {
            Write-Error "An error occurred while retrieving Intune roles: $_"
        }
    }
    
    end {
        # Nothing to clean up
    }
}