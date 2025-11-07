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
            PrimarySmtpAddress = $mbx.PrimarySmtpAddress
            DisplayName        = $mbx.DisplayName
            ExchangeObjectId   = $mbx.ExchangeObjectId
            MaxReceiveSize     = $mbx.MaxReceiveSize
            MaxSendSize        = $mbx.MaxSendSize
        }

        $exoMailboxesMaxSizeArray.Add($object)
    }

    return $exoMailboxesMaxSizeArray
}