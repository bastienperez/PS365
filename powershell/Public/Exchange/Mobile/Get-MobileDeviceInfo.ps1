<#
    .SYNOPSIS
    Retrieves detailed information about mobile devices associated with user mailboxes.

    .DESCRIPTION
    This script fetches mobile device details for specified user mailboxes or all mailboxes if none are specified.

    .PARAMETER UserPrincipalName
    An array of User Principal Names (UPNs) to retrieve mobile device details for.
    If not provided, details for all mailboxes will be retrieved.

    .EXAMPLE
    Get-MobileDeviceInfo

    Retrieves mobile device details for all user mailboxes.

    .EXAMPLE
    Get-MobileDeviceInfo -UserPrincipalName "<UserPrincipalName>"

    Retrieves mobile device details for the specified user.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MobileDeviceInfo

    .NOTES
    Author: Bastien Perez

    Works with both Exchange Online and on-premises Exchange environments and uses the appropriate cmdlets based on the environment.
#>

function Get-MobileDeviceInfo {

    [CmdletBinding()]
    param (

        [Parameter(ValueFromPipelineByPropertyName)]
        [string[]]
        $UserPrincipalName
    )
    begin {

        [System.Collections.Generic.List[PSCustomObject]]$mobileDetailsArray = @()

        if ($null -eq $UserPrincipalName) {
            $UserPrincipalName = (Get-EXOMailbox -ResultSize Unlimited).UserPrincipalName
        }

        $isExo = $false
        if (Get-Command -Name Get-EXOMobileDeviceStatistics -ErrorAction SilentlyContinue -WarningAction SilentlyContinue) {
            $isExo = $true
        }
    }
    process {
        foreach ($upn in $UserPrincipalName) {

            $mobiles = Get-MobileDevice -Mailbox $upn

            foreach ($mobile in $mobiles) {
                Write-Verbose "Getting info about mobile device(s) for $($mobile.Id)"
                
                if ($isexo) {
                    Write-Verbose 'Using EXO cmdlets to get mobile device statistics'
                    try {
                        $deviceStat = Get-EXOMobileDeviceStatistics -Id $mobile.Identity -ErrorAction Stop
                    }
                    catch {
                        Write-Host "Failed to get statistics for device $($mobile.DeviceId): $($_.Exception.Message)" -ForegroundColor Yellow
                        continue
                    }
                }
                else {
                    Write-Verbose 'Using on-premises cmdlets to get mobile device statistics'
                    try {
                        $deviceStat = Get-MobileDeviceStatistics -Identity $mobile.Identity -ErrorAction Stop
                    }
                    catch {
                        Write-Host "Failed to get statistics for device $($mobile.DeviceId): $($_.Exception.Message)" -ForegroundColor Yellow
                        continue
                    }
                }

                $object = [PSCustomObject][ordered]@{
                    UserUPN               = $upn
                    UserDisplayName       = $mobile.UserDisplayName
                    FriendlyName          = $mobile.FriendlyName
                    ID                    = $mobile.DeviceId
                    MobileOperator        = $mobile.DeviceMobileOperator
                    TelephoneNumber       = $mobile.DeviceTelephoneNumber
                    OS                    = $mobile.DeviceOS
                    OSLanguage            = $mobile.DeviceOSLanguage
                    FirstSyncTime         = $mobile.FirstSyncTime
                    LastSyncAttemptTime   = $deviceStat.LastSyncAttemptTime
                    LastSuccessSync       = $deviceStat.LastSuccessSync
                    ClientType            = $mobile.ClientType
                    DeviceModel           = $mobile.DeviceModel
                    DeviceType            = $mobile.DeviceType
                    ClientVersion         = $mobile.ClientVersion
                    DeviceId              = $mobile.DeviceId
                    DeviceUserAgent       = $mobile.DeviceUserAgent
                    Device                = $deviceStat.DeviceType
                    FoldersSynced         = $deviceStat.NumberOfFoldersSynced
                    Status                = $deviceStat.Status
                    IsRemoteWipeSupported = $deviceStat.IsRemoteWipeSupported
                    IsManaged             = $mobile.IsManaged
                    IsCompliant           = $mobile.IsCompliant
                    IsDisabled            = $mobile.IsDisabled
                    WhenCreatedUTC        = $mobile.WhenCreatedUTC
                    WhenChangedUTC        = $mobile.WhenChangedUTC
                }

                $mobileDetailsArray.Add($object)
            }
        }
    }
    end {
        return $mobileDetailsArray
    }
}