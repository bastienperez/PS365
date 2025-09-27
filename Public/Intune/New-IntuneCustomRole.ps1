<#
.SYNOPSIS
    Creates custom role profiles in Microsoft Intune.

.DESCRIPTION
    This function creates predefined custom role profiles in Microsoft Intune, such as 
    Autopilot Operators, Device Administrators, etc. Each profile has a specific set of permissions.
    
.PARAMETER ProfileType
    The type of profile to create. Valid values are:
    - AutopilotOperator: Permissions to manage Autopilot devices and enrollment profiles
    - DeviceAdministrator: Permissions to manage device configuration and policies
    - ApplicationManager: Permissions to manage applications
    - Custom: Create a custom profile with specified permissions

.PARAMETER DisplayName
    Optional. The display name of the custom role. If not provided, a default name will be used based on the profile type.

.PARAMETER Description
    Optional. The description of the custom role. If not provided, a default description will be used based on the profile type.

.PARAMETER CustomPermissions
    Required when ProfileType is 'Custom'. An array of permission strings to assign to the custom role.

.PARAMETER ScopeTagIds
    Optional. An array of scope tag IDs to assign to the custom role.

.EXAMPLE
    New-IntuneCustomRole -ProfileType AutopilotOperator
    Creates an Intune custom role with Autopilot Operator permissions.

.EXAMPLE
    New-IntuneCustomRole -ProfileType Custom -DisplayName "My Custom Role" -Description "Custom permissions for my team" -CustomPermissions @("Microsoft.Intune_Audit_Read", "Microsoft.Intune_AppleDeviceSerialNumbers_Read")
    Creates a custom role with specific permissions.

.NOTES
    Requires Microsoft Graph PowerShell module with DeviceManagementRBAC.ReadWrite.All permissions.
    Author: Bastien Perez
    Date: August 13, 2025
#>
function New-IntuneCustomRole {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('AutopilotOperator', 'DeviceAdministrator', 'ApplicationManager', 'Custom')]
        [string]$ProfileType,
        
        [Parameter(Mandatory = $false)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $false)]
        [string[]]$CustomPermissions,
        
        [Parameter(Mandatory = $false)]
        [string[]]$ScopeTagIds = @()
    )
    
    begin {

        $scopes = 'DeviceManagementRBAC.ReadWrite.All'
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
        
        # Define permission sets for different profile types
        $permissionSets = @{
            AutopilotOperator = @(
                'Microsoft.Intune_AppleDeviceSerialNumbers_Delete',
                'Microsoft.Intune_AppleDeviceSerialNumbers_Update',
                'Microsoft.Intune_AppleDeviceSerialNumbers_Read',
                'Microsoft.Intune_AppleEnrollmentProfiles_Assign',
                'Microsoft.Intune_AppleDeviceSerialNumbers_Create',
                'Microsoft.Intune_AppleEnrollmentProfiles_Read',
                'Microsoft.Intune_Audit_Read',
                'Microsoft.Intune_EnrollmentProfiles_EnrollmentTimeMembershipAssign',
                'Microsoft.Intune_WindowsAutopilotSettings_Read', 
                'Microsoft.Intune_WindowsAutopilotSettings_Create', 
                'Microsoft.Intune_WindowsAutopilotSettings_Update', 
                'Microsoft.Intune_WindowsAutopilotSettings_Delete',
                'Microsoft.Intune_DeviceEnrollmentConfigurations_Read',
                'Microsoft.Intune_DeviceEnrollmentConfigurations_Create',
                'Microsoft.Intune_DeviceEnrollmentConfigurations_Update'
            )
            
            DeviceAdministrator = @(
                'Microsoft.Intune_Organization_Read',
                'Microsoft.Intune_DeviceConfigurations_Read',
                'Microsoft.Intune_DeviceConfigurations_Create',
                'Microsoft.Intune_DeviceConfigurations_Update',
                'Microsoft.Intune_DeviceConfigurations_Delete',
                'Microsoft.Intune_DeviceConfigurations_Assign',
                'Microsoft.Intune_DeviceComplianceScripts_Read',
                'Microsoft.Intune_DeviceComplianceScripts_Create',
                'Microsoft.Intune_DeviceComplianceScripts_Update',
                'Microsoft.Intune_DeviceComplianceScripts_Delete',
                'Microsoft.Intune_DeviceComplianceScripts_Assign',
                'Microsoft.Intune_DeviceCompliancePolicies_Read',
                'Microsoft.Intune_DeviceCompliancePolicies_Create',
                'Microsoft.Intune_DeviceCompliancePolicies_Update',
                'Microsoft.Intune_DeviceCompliancePolicies_Delete',
                'Microsoft.Intune_DeviceCompliancePolicies_Assign',
                'Microsoft.Intune_ManagedDevices_Read',
                'Microsoft.Intune_ManagedDevices_PrivilegedOperations'
            )
            
            ApplicationManager = @(
                'Microsoft.Intune_Organization_Read',
                'Microsoft.Intune_Apps_Read',
                'Microsoft.Intune_Apps_Create',
                'Microsoft.Intune_Apps_Update',
                'Microsoft.Intune_Apps_Delete',
                'Microsoft.Intune_Apps_Assign',
                'Microsoft.Intune_AppConfiguration_Read',
                'Microsoft.Intune_AppConfiguration_Create',
                'Microsoft.Intune_AppConfiguration_Update',
                'Microsoft.Intune_AppConfiguration_Delete',
                'Microsoft.Intune_AppConfiguration_Assign',
                'Microsoft.Intune_AppProtection_Read',
                'Microsoft.Intune_AppProtection_Create',
                'Microsoft.Intune_AppProtection_Update',
                'Microsoft.Intune_AppProtection_Delete',
                'Microsoft.Intune_AppProtection_Assign'
            )
        }
        
        # Define default display names and descriptions
        $defaultDisplayNames = @{
            AutopilotOperator = 'Intune Autopilot Operator - Custom Role'
            DeviceAdministrator = 'Intune Device Administrator - Custom Role'
            ApplicationManager = 'Intune Application Manager - Custom Role'
        }
        
        $defaultDescriptions = @{
            AutopilotOperator = 'Custom Intune role with permissions to add computers to Windows Autopilot'
            DeviceAdministrator = 'Custom Intune role with permissions to manage device configurations and compliance'
            ApplicationManager = 'Custom Intune role with permissions to manage applications and app policies'
        }
    }
    
    process {
        try {
            # Set permissions based on profile type
            if ($ProfileType -eq 'Custom') {
                if (-not $CustomPermissions -or $CustomPermissions.Count -eq 0) {
                    Write-Error "When using ProfileType 'Custom', you must provide permissions using the CustomPermissions parameter."
                    return
                }
                $permissions = $CustomPermissions
            }
            else {
                $permissions = $permissionSets[$ProfileType]
            }
            
            # Set display name and description
            if (-not $DisplayName) {
                $DisplayName = $defaultDisplayNames[$ProfileType]
                if (-not $DisplayName -and $ProfileType -eq 'Custom') {
                    $DisplayName = "Custom Intune Role - $(Get-Date -Format 'yyyy-MM-dd')"
                }
            }
            
            if (-not $Description) {
                $Description = $defaultDescriptions[$ProfileType]
                if (-not $Description -and $ProfileType -eq 'Custom') {
                    $Description = "Custom Intune role created on $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
                }
            }
            
            # Prepare the role definition request body
            $roleDefinition = @{
                displayName = $DisplayName
                description = $Description
                rolePermissions = @(
                    @{
                        resourceActions = @{
                            allowedResourceActions = $permissions
                            notAllowedResourceActions = @()
                        }
                    }
                )
                isBuiltIn = $false
                roleScopeTagIds = $ScopeTagIds
            }
            
            $roleDefinitionJson = ConvertTo-Json -InputObject $roleDefinition -Depth 10
            
            if ($PSCmdlet.ShouldProcess($DisplayName, "Create Intune Custom Role")) {
                Write-Verbose "Creating custom role: $DisplayName"
                
                $apiUrl = "https://graph.microsoft.com/beta/deviceManagement/roleDefinitions"
                $response = Invoke-MgGraphRequest -Method POST -Uri $apiUrl -Body $roleDefinitionJson -ContentType "application/json"
                
                Write-Verbose "Custom role created successfully with ID: $($response.id)"
                
                # Format the response
                $roleResponse = [PSCustomObject][ordered]@{
                    Id = $response.id
                    DisplayName = $response.displayName
                    Description = $response.description
                    IsBuiltIn = $response.isBuiltIn
                    Permissions = $response.rolePermissions.resourceActions.allowedResourceActions
                    ScopeTagIds = $response.roleScopeTagIds
                    CreatedDateTime = Get-Date
                }
                
                return $roleResponse
            }
        }
        catch {
            Write-Error "Error creating custom role: $_"
        }
    }
    
    end {
        # Nothing to do in the end block
    }
}