<#
.SYNOPSIS
    Configures the MDM enrollment policy for Azure AD.

.DESCRIPTION
    This function enables or disables automatic MDM enrollment during device 
    registration in Azure AD using the Microsoft Graph API.

.PARAMETER Disabled
    Specifies whether MDM enrollment should be disabled during registration.
    - $true  : Disables automatic MDM enrollment
    - $false : Enables automatic MDM enrollment

.EXAMPLE
    Set-InMDMEnrollmentPolicy -Disabled $true
    Disables automatic MDM enrollment during device registration.

.EXAMPLE
    Set-InMDMEnrollmentPolicy -Disabled $false
    Enables automatic MDM enrollment during device registration.

.NOTES
    Requires the following Microsoft Graph permissions:
    - Policy.ReadWrite.MobileDeviceManagement
#>
function Set-InMDMEnrollmentPolicy {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory = $true)]
        [bool]$Disabled
    )

    try {
        $policyId = "0000000a-0000-0000-c000-000000000000"
        
        $body = @{
            requests = @(
                @{
                    id     = [guid]::NewGuid().ToString()
                    method = "PATCH"
                    url    = "/policies/mobileDeviceManagementPolicies/$policyId"
                    body   = @{
                        isMdmEnrollmentDuringRegistrationDisabled = $Disabled
                    }
                }
            )
        } | ConvertTo-Json -Depth 10

        if ($PSCmdlet.ShouldProcess("MDM Enrollment Policy", "Set isMdmEnrollmentDuringRegistrationDisabled to $Disabled")) {
            
            Write-Verbose "Configuring MDM policy with status: $($Disabled ? 'Disabled' : 'Enabled')"
            
            $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/`$batch" -Body $body -ContentType "application/json"
            
            if ($result.responses[0].status -eq 204 -or $result.responses[0].status -eq 200) {
                Write-Host "✓ MDM policy updated successfully" -ForegroundColor Green
                Write-Host "  Automatic MDM enrollment: $($Disabled ? 'Disabled' : 'Enabled')" -ForegroundColor Cyan
                return $result
            }
            else {
                Write-Warning "Request returned status: $($result.responses[0].status)"
                return $result
            }
        }
    }
    catch {
        Write-Error "Error configuring MDM policy: $($_.Exception.Message)"
        throw
    }
}
