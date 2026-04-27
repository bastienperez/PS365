<#
.SYNOPSIS
    Configures the MDM enrollment policy for Microsoft Entra ID.

.DESCRIPTION
    This function enables or disables automatic MDM enrollment during device 
    registration in Microsoft Entra ID.

.PARAMETER State
    Specifies whether MDM enrollment should be enabled or disabled during registration.
    - 'Enabled'  : Enables automatic MDM enrollment (sets isMdmEnrollmentDuringRegistrationDisabled to false)
    - 'Disabled' : Disables automatic MDM enrollment (sets isMdmEnrollmentDuringRegistrationDisabled to true)

.EXAMPLE
    Set-IntuneAutoMDMEnrollmentPolicy -State 'Disabled'
    
    Disables automatic MDM enrollment during device registration.

.EXAMPLE
    Set-IntuneAutoMDMEnrollmentPolicy -State 'Enabled'
    
    Enables automatic MDM enrollment during device registration.

.NOTES
    Requires the following Microsoft Graph permissions:
    - Policy.ReadWrite.MobilityManagement

    Connect-MgGraph -Scopes 'Policy.ReadWrite.MobilityManagement'

.LINK
    https://ps365.clidsys.com/docs/commands/Set-IntuneAutoMDMEnrollmentPolicy
#>
function Set-IntuneAutoMDMEnrollmentPolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet('Enabled', 'Disabled')]
        [string]$State
    )

    try {
        # Convert State to the correct boolean value for the API
        # Enabled means isMdmEnrollmentDuringRegistrationDisabled = false
        # Disabled means isMdmEnrollmentDuringRegistrationDisabled = true
        $disableEnrollment = ($State -eq 'Disabled')
        
        $policyId = '0000000a-0000-0000-c000-000000000000'
        
        Write-Verbose 'Retrieving current MDM enrollment policy...'
        $policyObject = Get-IntuneAutoMDMEnrollmentPolicy -AsObject
        $isDisabled = $policyObject.isMdmEnrollmentDuringRegistrationDisabled
        $currentState = if ($isDisabled) { 'Disabled' } else { 'Enabled' }

        Write-Verbose "Current state: $currentState (isMdmEnrollmentDuringRegistrationDisabled: $isDisabled)"
        Write-Verbose "Target state: $State (isMdmEnrollmentDuringRegistrationDisabled: $disableEnrollment)"

        # Check if change is needed
        if ($currentState -eq $State) {
            Write-Host "MDM enrollment policy is already set to $State - No change needed" -ForegroundColor Green
            return
        }
        
        $body = @{
            isMdmEnrollmentDuringRegistrationDisabled = $disableEnrollment
        } | ConvertTo-Json -Depth 10

        Write-Verbose "Updating MDM policy from $currentState to $State"
        Write-Verbose "PATCH body: $body"
        
        $result = Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/beta/policies/mobileDeviceManagementPolicies/$policyId" -Body $body -ContentType 'application/json'
        
        Write-Host "MDM policy updated successfully - Changed from $currentState to $State" -ForegroundColor Green
        return $result
    }
    catch {
        Write-Error "Error configuring MDM policy: $($_.Exception.Message)"
        throw
    }
}