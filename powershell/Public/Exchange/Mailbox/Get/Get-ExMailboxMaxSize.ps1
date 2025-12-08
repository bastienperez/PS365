<#
    .SYNOPSIS
    Get Exchange Mailbox Max Send and Receive Size limits.

    .DESCRIPTION
    This function retrieves the maximum send and receive size limits for Exchange Online mailboxes.

    .PARAMETER Identity
    The identity of the mailbox to retrieve. This can be the primary SMTP address, alias, or GUID.

    .PARAMETER ByDomain
    The domain to filter mailboxes by their primary SMTP address.

    .EXAMPLE
    Get-ExMailboxMaxSize
        
    Retrieves the max send and receive size limits for all Exchange Online mailboxes.

    .EXAMPLE
    Get-ExMailboxMaxSize -ByDomain "contoso.com"
    
    Retrieves the max send and receive size limits for mailboxes in the specified domain.

    .EXAMPLE
    Get-ExMailboxMaxSize -Identity "user@contoso.com"
    
    Retrieves the max send and receive size limits for the specified mailbox.
#>
function Get-ExMailboxMaxSize {
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Identity,
        [Parameter(Mandatory = $false)]
        [string]$ByDomain
    )

    [System.Collections.Generic.List[PSCustomObject]]$exoMailboxesMaxSizeArray = @()

    if ($ByDomain) {
        $exoMailboxes = Get-Mailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'" | Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
    }
    elseif ($Identity) {
        [System.Collections.Generic.List[PSCustomObject]]$exoMailboxes = @()
        try {
            $mbx = Get-Mailbox -Identity $Identity -ErrorAction Stop
            $exoMailboxes.Add($mbx)
        }
        catch {
            Write-Warning "Mailbox not found: $Identity"
        }
    }
    else {
        $exoMailboxes = Get-Mailbox -ResultSize Unlimited
    }

    foreach ($mbx in $exoMailboxes) {
    
        $object = [PSCustomObject][ordered]@{ 
            PrimarySmtpAddress  = $mbx.PrimarySmtpAddress
            DisplayName         = $mbx.DisplayName
            ExchangeObjectId    = $mbx.ExchangeObjectId
            MaxReceiveSize      = $mbx.MaxReceiveSize
            MaxSendSize         = $mbx.MaxSendSize
            MailboxWhenCreated  = $mbx.WhenCreated
            MailboxWhenModified = $mbx.WhenChanged
        }

        $exoMailboxesMaxSizeArray.Add($object)
    }

    return $exoMailboxesMaxSizeArray
}