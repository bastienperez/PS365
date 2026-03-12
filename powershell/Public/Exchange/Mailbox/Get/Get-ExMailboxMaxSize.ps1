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

    .PARAMETER ExportToExcel
    If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-ExMailboxMaxSize -ExportToExcel
    Exports results to an Excel file.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExMailboxMaxSize
#>
function Get-ExMailboxMaxSize {
    param (
        [Parameter(Mandatory = $false, position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [string]$ByDomain,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
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

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-ExMailboxMaxSize.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
        $exoMailboxesMaxSizeArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExMailboxMaxSize'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $exoMailboxesMaxSizeArray
    }
}