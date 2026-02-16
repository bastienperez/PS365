<#
.SYNOPSIS
    Retrieves the current MDM enrollment policy for Azure AD.

.DESCRIPTION
    This function retrieves the current state of automatic MDM enrollment during device 
    registration in Microsoft Entra ID.

.PARAMETER AsObject
    When specified, returns the complete policy object instead of just the state.

.EXAMPLE
    Get-IntuneAutoMDMEnrollmentPolicy
    
    Retrieves the current MDM enrollment policy state (Enabled/Disabled).

.EXAMPLE
    Get-IntuneAutoMDMEnrollmentPolicy -AsObject
    
    Retrieves the complete policy object from Microsoft Graph.

.NOTES
    Requires the following Microsoft Graph permissions:
    - Policy.Read.All
    
.OUTPUTS
    String - Returns 'Enabled' or 'Disabled' by default
    PSObject - Returns full policy object when -AsObject is used
#>
function Get-IntuneAutoMDMEnrollmentPolicy {
    [CmdletBinding()]
    param (
        [Parameter()]
        [switch]$AsObject
    )

    try {
        $policyId = '0000000a-0000-0000-c000-000000000000'
        
        Write-Verbose "Retrieving current MDM enrollment policy..."
        $policy = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies/$policyId"
        
        if ($AsObject) {
            return $policy
        }
        
        $currentValue = $policy.isMdmEnrollmentDuringRegistrationDisabled
        $currentState = if ($currentValue) { 'Disabled' } else { 'Enabled' }
        
        Write-Verbose "Current state: $currentState (isMdmEnrollmentDuringRegistrationDisabled: $currentValue)"
        
        return $currentState
    }
    catch {
        Write-Error "Error retrieving MDM enrollment policy: $($_.Exception.Message)"
        throw
    }
}
