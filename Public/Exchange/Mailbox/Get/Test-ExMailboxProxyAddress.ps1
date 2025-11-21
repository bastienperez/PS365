<#
.SYNOPSIS
    Test if a given email address exists in the proxy addresses of a mailbox.

.DESCRIPTION
    This cmdlet checks whether a specified email address is present in the proxy addresses of a given mailbox.
    It can process a single mailbox/email pair or read multiple entries from a CSV file.

.EXAMPLE
    Test-ExMailboxProxyAddress -Mailbox "user@example.com" -EmailToCheck "alias@example.com"

    Tests if "alias@example.com" exists in the proxy addresses of the mailbox "user@example.com".

.EXAMPLE
    Test-ExMailboxProxyAddress -CsvPath "C:\path\to\file.csv" -MailboxColumn "PrimarySmtpAddress" -EmailColumn "OnMicrosoftAddress" -OnlyMatch

    Reads mailbox and email pairs from the specified CSV file and returns only those where the email address matches a proxy address in the mailbox.
#>
function Test-ExMailboxProxyAddress {
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
        [string]$Mailbox,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
        [string]$EmailToCheck,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
        [string]$CsvPath,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
        [string]$MailboxColumn = 'PrimarySmtpAddress',
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Csv')]
        [string]$EmailColumn = 'OnMicrosoftAddress',
        
        [Parameter(Mandatory = $false)]
        [switch]$OnlyMatch,
        
        [Parameter(Mandatory = $false)]
        [switch]$OnlyNotMatch
    )

    [System.Collections.Generic.List[PSCustomObject]]$results = @()

    if ($PSCmdlet.ParameterSetName -eq 'Csv') {
        $data = Import-Csv -Path $CsvPath -Delimiter ';'
        
        foreach ($row in $data) {
            Write-Host -ForegroundColor Cyan "Checking mailbox: $($row.$MailboxColumn) for email: $($row.$EmailColumn)"
            
            $mailboxValue = $row.$MailboxColumn
            $emailValue = $row.$EmailColumn
            $statusValue = 'ERROR'
            
            try {
                $mb = Get-Mailbox -Identity $mailboxValue -ErrorAction Stop
                
                if ($mb.EmailAddresses -contains "smtp:$emailValue") {
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
                Mailbox = $mailboxValue
                Email   = $emailValue
                Status  = $statusValue
            }
            
            $results.Add($object)
        }
    }
    else {
        $mailboxValue = $Mailbox
        $emailValue = $EmailToCheck
        $statusValue = 'ERROR'
        
        try {
            $mb = Get-Mailbox -Identity $mailboxValue -ErrorAction Stop
            
            if ($mb.EmailAddresses -contains "smtp:$emailValue") {
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
            Mailbox = $mailboxValue
            Email   = $emailValue
            Status  = $statusValue
        }
        $results.Add($object)
    }

    if ($OnlyMatch) {
        return $results | Where-Object { $_.Status -eq 'MATCH' }
    }
    elseif ($OnlyNotMatch) {
        return $results | Where-Object { $_.Status -ne 'MATCH' }
    }
    else {
        return $results
    }
}