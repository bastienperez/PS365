<#
    .SYNOPSIS
        Test if a given proxy address exists in the proxy addresses of a mailbox.

    .DESCRIPTION
        This cmdlet checks whether a specified proxy address is present in the proxy addresses of a given mailbox.
        It can process a single mailbox/proxy address pair or read multiple entries from a CSV file.

    .PARAMETER Mailbox
        The identity of the mailbox to check (e.g., email address, alias, or GUID).
        This parameter is used when processing a single mailbox/proxy address pair.

    .PARAMETER ProxyAddress
        The proxy address to check for in the mailbox's proxy addresses.
        This parameter is used when processing a single mailbox/proxy address pair.

    .PARAMETER CsvPath
        The path to a CSV file containing mailbox and proxy address pairs.
        This parameter is used when processing multiple entries from a CSV file.

    .PARAMETER MailboxColumn
        The name of the column in the CSV file that contains mailbox identities.
        Default is 'PrimarySmtpAddress'.

    .PARAMETER ProxyAddressColumn
        The name of the column in the CSV file that contains proxy addresses.
        Default is 'OnMicrosoftAddress'.

    .PARAMETER MatchOnly
        If specified, only returns entries where the proxy address matches.

    .PARAMETER NotMatchOnly
        If specified, only returns entries where the proxy address does not match.

    .EXAMPLE
        Test-ExMailboxProxyAddress -Mailbox "user@example.com" -ProxyAddress "alias@example.com"

        Tests if "alias@example.com" exists in the proxy addresses of the mailbox "user@example.com".

    .EXAMPLE
        Test-ExMailboxProxyAddress -CsvPath "C:\path\to\file.csv" -MailboxColumn "PrimarySmtpAddress" -ProxyAddressColumn "OnMicrosoftAddress"
        
        Reads mailbox and proxy address pairs from the specified CSV file and tests each pair.

    .EXAMPLE
        Test-ExMailboxProxyAddress -CsvPath "C:\path\to\file.csv" -MatchOnly

        Reads mailbox and proxy address pairs from the specified CSV file and returns only those that match.

    .EXAMPLE
        Test-ExMailboxProxyAddress -CsvPath "C:\path\to\file.csv" -NotMatchOnly 

        Reads mailbox and proxy address pairs from the specified CSV file and returns only those that do not match.
#>

function Test-ExMailboxProxyAddress {
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
        [string]$ProxyAddress,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
        [string]$CsvPath,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
        [string]$MailboxColumn = 'PrimarySmtpAddress',
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
        [string]$ProxyAddressColumn = 'OnMicrosoftAddress',
        
        [Parameter(Mandatory = $false)]
        [switch]$MatchOnly,
        
        [Parameter(Mandatory = $false)]
        [switch]$NotMatchOnly
    )

    [System.Collections.Generic.List[PSCustomObject]]$results = @()
    
    # Direct processing based on parameter type
    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        # If the column does not exist, throw an error
        
        $data = Import-Csv -Path $CsvPath -Delimiter ';'

        # If the columns do not exist, throw an error
        if (-not ($data | Get-Member -Name $MailboxColumn)) {
            throw "Column '$MailboxColumn' does not exist in the CSV file."
            return
        }
        if (-not ($data | Get-Member -Name $ProxyAddressColumn)) {
            throw "Column '$ProxyAddressColumn' does not exist in the CSV file."
            return
        }

        $itemsToProcess = $data | ForEach-Object { @{ Mailbox = $_.$MailboxColumn; ProxyAddress = $_.$ProxyAddressColumn } }
    }
    else {
        $itemsToProcess = @(@{ Mailbox = $Mailbox; ProxyAddress = $ProxyAddress })
    }
    
    # Process all items with the same logic
    foreach ($item in $itemsToProcess) {
        Write-Host -ForegroundColor Cyan "Checking mailbox: $($item.Mailbox) for email: $($item.ProxyAddress)"
        
        $statusValue = 'ERROR'
        
        try {
            $mb = Get-Mailbox -Identity $item.Mailbox -ErrorAction Stop
            
            if ($mb.EmailAddresses -contains "smtp:$($item.ProxyAddress)") {
                $statusValue = 'MATCH'
            }
            else {
                $statusValue = 'NOTMATCH'
            }
        }
        catch {
            $statusValue = 'ERROR'
        }
        
        $object = [PSCustomObject][ordered]@{
            Mailbox      = $item.Mailbox
            ProxyAddress = $item.ProxyAddress
            Status       = $statusValue
        }
        
        $results.Add($object)
    }

    if ($MatchOnly) {
        return $results | Where-Object { $_.Status -eq 'MATCH' }
    }
    elseif ($NotMatchOnly) {
        return $results | Where-Object { $_.Status -ne 'MATCH' }
    }
    else {
        return $results
    }
}