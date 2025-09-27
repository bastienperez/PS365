﻿
function Remove-GroupAddress {
    <#
    .SYNOPSIS
    Remove all mailbox addresses with one more more domains/words

    .DESCRIPTION
    Remove all mailbox addresses with one more more domains/words

    .PARAMETER Domains
    List of domains or words to find in the email addresses

    .EXAMPLE
    Remove-GroupAddress -Domains 'fabrikam.com' | Export-Csv c:\scripts\log.csv -NoTypeInformation

    .EXAMPLE
    Remove-GroupAddress -Domains 'wingtip.com|fabrikam.com|widget.com' | Export-Csv c:\scripts\log.csv -NoTypeInformation

    .NOTES
    Connect to Exchange Online Version 2
    Connect-CloudMFA -Tenant contoso -EXO2
    #>

    param (

        [Parameter(Mandatory)]
        $Domains
    )
    end {
        $EA = $ErrorActionPreference
        $ErrorActionPreference = 'Stop'
        $RemoveList = Get-DistributionGroup -ResultSize Unlimited | Select-Object @(
            'DisplayName'
            'PrimarySmtpAddress'
            'UserPrincipalName'
            @{
                Name       = 'EmailList'
                Expression = { @($_.emailaddresses) -match $Domains }
            }
            'ExchangeGuid'
            'Guid'
        )
        $RemoveList = $RemoveList | Where-Object { $_.EmailList }
        foreach ($Remove in $RemoveList) {
            try {
                Write-Host "$($Remove.DisplayName)" -ForegroundColor White
                Get-DistributionGroup -Identity $Remove.Guid.ToString() | Set-DistributionGroup -EmailAddresses @{Remove = @($Remove.EmailList) }
                Write-Host "$($Remove.DisplayName) Removed" -ForegroundColor Green
                [PSCustomObject][ordered]@{
                    Action             = "REMOVEEMAILS"
                    DisplayName        = $Remove.DisplayName
                    PrimarySmtpAddress = $Remove.PrimarySmtpAddress
                    UserPrincipalName  = $Remove.UserPrincipalName
                    Guid               = $Remove.Guid
                    Result             = "SUCCESS"
                    Log                = "SUCCESS"
                    Remove             = @($Remove.EmailList -ne '' -Join '|')
                }
            }
            catch {
                Write-Host "$($Remove.DisplayName) $($_.Exception.Message)" -ForegroundColor Red
                [PSCustomObject][ordered]@{
                    Action             = "REMOVEEMAILS"
                    DisplayName        = $Remove.DisplayName
                    PrimarySmtpAddress = $Remove.PrimarySmtpAddress
                    UserPrincipalName  = $Remove.UserPrincipalName
                    Guid               = $Remove.Guid
                    Result             = "FAILED"
                    Log                = $_.Exception.Message
                    Remove             = @($Remove.EmailList -ne '' -Join '|')
                }
            }
        }
        $ErrorActionPreference = $EA
    }
}
