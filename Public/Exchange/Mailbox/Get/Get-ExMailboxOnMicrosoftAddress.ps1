<#
.SYNOPSIS
Lists Exchange Online mailboxes with an @onmicrosoft.com email address

.DESCRIPTION
Retrieves Exchange Online mailboxes that have an email address ending with @onmicrosoft.com.
Returns an object for each mailbox with relevant details.

.EXAMPLE
Get-ExMailboxOnMicrosoftAddress
Returns objects containing mailbox details for those with @onmicrosoft.com addresses.

.EXAMPLE
Get-ExMailboxOnMicrosoftAddress -ByDomain "contoso.com"
Returns mailboxes with @onmicrosoft.com addresses filtered by the specified domain.

#>

function Get-ExMailboxOnMicrosoftAddress {
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [string]$ByDomain
    )

    [System.Collections.Generic.List[PSCustomObject]]$onMicrosoftMailboxes = @()

    if ($ByDomain) {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'" | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
    }
    elseif ($Identity) {
        [System.Collections.Generic.List[PSCustomObject]]$mailboxes = @()

        try {
            $mailbox = Get-EXOMailbox -Identity $Identity
            $mailboxes.Add($mailbox)
        }
        catch {
            Write-Warning "Mailbox not found: $Identity"
        }
    }
    else {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited
    }

    # search for @onmicrosoft.com addresses 
    Write-Host -ForegroundColor Cyan "Found $($mailboxes.Count) mailboxes. Searching for @onmicrosoft.com addresses..."
    foreach ($mailbox in $mailboxes) {
        foreach ($email in $mailbox.EmailAddresses) {
            if ($email -like '*.onmicrosoft.com') {
                $object = [PSCustomObject][ordered]@{ 
                    PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                    DisplayName        = $mailbox.DisplayName
                    ExchangeObjectId   = $mailbox.ExchangeObjectId
                    OnMicrosoftAddress = $email.replace('SMTP:', '').Replace('smtp:', '')
                }
                Write-Host -ForegroundColor Green "Found @onmicrosoft.com address: $($email.AddressString) for mailbox: $($mailbox.PrimarySmtpAddress)"
                $onMicrosoftMailboxes.Add($object)
            }
        }
    }
    return $onMicrosoftMailboxes
}