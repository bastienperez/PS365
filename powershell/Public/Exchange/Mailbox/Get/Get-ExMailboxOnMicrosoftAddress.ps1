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

    .EXAMPLE
    Get-ExMailboxOnMicrosoftAddress -Identity "user@contoso.com"

    Returns the @onmicrosoft.com address for the specified mailbox identity.

    .EXAMPLE
    $mailboxes | Get-ExMailboxOnMicrosoftAddress

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExMailboxOnMicrosoftAddress
#>

function Get-ExMailboxOnMicrosoftAddress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [string[]]$Identity,
        [Parameter(Mandatory = $false)]
        [string]$ByDomain
    )

    begin {
        Write-Verbose 'Starting Get-ExMailboxOnMicrosoftAddress'
        [System.Collections.Generic.List[PSCustomObject]]$mailboxes = @()
        [System.Collections.Generic.List[PSCustomObject]]$results = @()
        [System.Collections.Generic.List[string]]$orderedIdentities = @()
    }

    process {
        if ($Identity) {
            foreach ($id in $Identity) {
                $orderedIdentities.Add($id)
                try {
                    Write-Verbose "Getting mailbox: $id"
                    $mailbox = Get-EXOMailbox -Identity $id -ErrorAction Stop -Properties WhenCreated, WhenChanged
                    $mailboxes.Add($mailbox)
                }
                catch {
                    Write-Warning "Mailbox not found: $id"
                    # Add placeholder for non-existent mailbox
                    $mailboxes.Add([PSCustomObject]@{
                            PrimarySmtpAddress = $id
                            DisplayName        = $null
                            ExchangeObjectId   = $null
                            EmailAddresses     = @()
                            NotFound           = $true
                            WhenCreated        = $null
                            WhenChanged       = $null
                        })
                }
            }
        }
    }

    end {
        # Get mailboxes if not from pipeline
        if ($mailboxes.Count -eq 0) {
            if ($ByDomain) {
                Write-Verbose "Getting mailboxes for domain: $ByDomain"
                $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Filter "EmailAddresses -like '*@$ByDomain'"  -Properties WhenCreated, WhenChanged | 
                Where-Object { $_.PrimarySmtpAddress -like "*@$ByDomain" }
            }
            else {
                Write-Verbose 'Getting all mailboxes'
                $mailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties WhenCreated, WhenChanged
            }
        }

        Write-Host -ForegroundColor Cyan "Searching $($mailboxes.Count) mailbox(es) for @*.onmicrosoft.com addresses..."
        
        # If we have ordered identities, process in that order
        if ($orderedIdentities.Count -gt 0) {
            foreach ($id in $orderedIdentities) {
                $mailbox = $mailboxes | Where-Object { $_.PrimarySmtpAddress -eq $id }
                
                if ($mailbox.NotFound) {
                    # Mailbox doesn't exist
                    $object = [PSCustomObject][ordered]@{
                        PrimarySmtpAddress  = $id
                        DisplayName         = $null
                        ExchangeObjectId    = $null
                        WhenCreated         = $null
                        WhenChanged         = $null
                        OnMicrosoftAddress  = $null
                        MailboxWhenCreated  = $null
                        MailboxWhenModified = $null
                    }
                    $results.Add($object)
                }
                else {
                    $onMicrosoftAddresses = $mailbox.EmailAddresses | Where-Object { $_ -like '*.onmicrosoft.com' }
                    
                    if ($onMicrosoftAddresses) {
                        foreach ($email in $onMicrosoftAddresses) {
                            $cleanEmail = $email -replace '^(SMTP|smtp):', ''
                            
                            $object = [PSCustomObject][ordered]@{
                                PrimarySmtpAddress  = $mailbox.PrimarySmtpAddress
                                DisplayName         = $mailbox.DisplayName
                                ExchangeObjectId    = $mailbox.Id
                                OnMicrosoftAddress  = $cleanEmail
                                MailboxWhenCreated  = $mailbox.WhenCreated
                                MailboxWhenModified = $mailbox.WhenChanged
                            }
                            $results.Add($object)
                            
                            Write-Host -ForegroundColor Green "Found: $cleanEmail ($($mailbox.PrimarySmtpAddress))"
                        }
                    }
                    else {
                        # No onmicrosoft.com address found
                        $object = [PSCustomObject][ordered]@{
                            PrimarySmtpAddress  = $mailbox.PrimarySmtpAddress
                            DisplayName         = $mailbox.DisplayName
                            ExchangeObjectId    = $mailbox.Id
                            OnMicrosoftAddress  = $null
                            MailboxWhenCreated  = $mailbox.WhenCreated
                            MailboxWhenModified = $mailbox.WhenChanged
                        }
                        $results.Add($object)
                    }
                }
            }
        }
        else {
            # Process all mailboxes (no specific order required)
            foreach ($mailbox in $mailboxes) {
                $onMicrosoftAddresses = $mailbox.EmailAddresses | Where-Object { $_ -like '*.onmicrosoft.com' }
                
                if ($onMicrosoftAddresses) {
                    foreach ($email in $onMicrosoftAddresses) {
                        $cleanEmail = $email -replace '^(SMTP|smtp):', ''
                        
                        $object = [PSCustomObject][ordered]@{
                            PrimarySmtpAddress  = $mailbox.PrimarySmtpAddress
                            DisplayName         = $mailbox.DisplayName
                            ExchangeObjectId    = $mailbox.Id
                            OnMicrosoftAddress  = $cleanEmail
                            MailboxWhenCreated  = $mailbox.WhenCreated
                            MailboxWhenModified = $mailbox.WhenChanged
                        }
                        $results.Add($object)
                        
                        Write-Host -ForegroundColor Green "Found: $cleanEmail ($($mailbox.PrimarySmtpAddress))"
                    }
                }
                else {
                    # No onmicrosoft.com address found
                    $object = [PSCustomObject][ordered]@{
                        PrimarySmtpAddress  = $mailbox.PrimarySmtpAddress
                        DisplayName         = $mailbox.DisplayName
                        ExchangeObjectId    = $mailbox.Id
                        OnMicrosoftAddress  = $null
                        MailboxWhenCreated  = $mailbox.WhenCreated
                        MailboxWhenModified = $mailbox.WhenChanged
                    }
                    $results.Add($object)
                }
            }
        }
        
        return $results
    }
}