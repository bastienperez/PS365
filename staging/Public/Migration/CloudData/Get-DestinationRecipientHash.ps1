function Get-DestinationRecipientHash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('RemoteMailbox', 'MailContact')]
        $Type
    )

    $PS365Path = (Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath PS365)
    if (-not (Test-Path $PS365Path)) {
        $null = New-Item $PS365Path -type Directory -Force:$true -ErrorAction SilentlyContinue
    }
    if ($Type -eq 'RemoteMailbox') {
        $File = ('BACKUP Target Remote Mailboxes_{0}.xml' -f [DateTime]::Now.ToString('yyyy-MM-dd-hhmm'))
        $HashFile = 'TargetRemoteMailboxHash.xml'
        $RemoteXML = Join-Path -Path $PS365Path -ChildPath $File
        Get-RemoteMailbox -ResultSize Unlimited | Export-Clixml $RemoteXML
        Write-Host "Using the XML to create a hashtable . . . " -ForegroundColor White
        $RecipientList = Import-Clixml $RemoteXML
        $Hash = @{ }
        foreach ($Recipient in $RecipientList) {
            $Hash[$Recipient.PrimarySmtpAddress] = @{
                GUID                 = $Recipient.GUID
                RecipientTypeDetails = $Recipient.RecipientTypeDetails
                Identity             = $Recipient.Identity
                Alias                = $Recipient.Alias
                DisplayName          = $Recipient.DisplayName
                Name                 = $Recipient.Name
                EmailAddresses       = @($Recipient.EmailAddresses) -ne '' -join '|'
            }
        }
    }
    else {
        $File = ('BACKUP Target Contacts_{0}.xml' -f [DateTime]::Now.ToString('yyyy-MM-dd-hhmm'))
        $HashFile = 'TargetContactHash.xml'
        $RemoteXML = Join-Path -Path $PS365Path -ChildPath $File
        Get-MailContact -ResultSize Unlimited | Export-Clixml $RemoteXML
        Write-Host "Using the XML to create a hashtable . . . " -ForegroundColor White
        $RecipientList = Import-Clixml $RemoteXML
        $Hash = @{ }
        foreach ($Recipient in $RecipientList) {
            $Hash[($Recipient.ExternalEmailAddress).Split(':')[1]] = @{
                GUID                 = $Recipient.GUID
                RecipientTypeDetails = $Recipient.RecipientTypeDetails
                Identity             = $Recipient.Identity
                Alias                = $Recipient.Alias
                DisplayName          = $Recipient.DisplayName
                Name                 = $Recipient.Name
                EmailAddresses       = @($Recipient.EmailAddresses) -ne '' -join '|'
            }
        }
    }

    $OutputXml = Join-Path -Path $PS365Path -ChildPath $HashFile
    Write-Host "Hash has been exported to: " -ForegroundColor Cyan -NoNewline
    Write-Host "$OutputXml" -ForegroundColor Green
    $Hash | Export-Clixml $OutputXml
}
