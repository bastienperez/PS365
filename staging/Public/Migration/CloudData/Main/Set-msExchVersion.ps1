function Set-msExchVersion {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $DomainController,

        [Parameter()]
        [switch]
        $ShowVersionCountOnly,

        [Parameter()]
        [switch]
        $SkipConnection,

        [Parameter()]
        [switch]
        $ReuseADXML,

        [Parameter()]
        [switch]
        $DontViewEntireForest
    )
    if (-not (Get-Module ActiveDirectory -ListAvailable)) {
        Write-Host 'Active Directory PowerShell Module not found.  Halting Script.' -ForegroundColor Red
        continue
    }
    $PS365Path = (Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath PS365)
    if (-not (Test-Path $PS365Path)) {
        $null = New-Item $PS365Path -type Directory -Force:$true -ErrorAction SilentlyContinue
    }
    Import-Module ActiveDirectory -force
    $ADUserXML = Join-Path -Path $PS365Path -ChildPath 'ADUser_msExchVersion.xml'
    if ($ReuseADXML) {
        $UserList = Import-Clixml $ADUserXML
    }
    else {
        Write-Host "`r`nCreating XML of all Active Directory Users with the following values in the attribute, msExchRecipientTypeDetails:" -ForegroundColor Green
        Write-Host "2147483648 (RemoteMailbox), 8589934592 (RemoteRoomMailbox), 17179869184 (RemoteEquipmentMailbox), 34359738368 (RemoteSharedMailbox)" -BackgroundColor White -ForegroundColor Black
        Write-Host "`r`nBreakdown of msExchVersion via Active Directory of Remote Mailbox types will output shortly to Out Grid`r`n" -ForegroundColor Green
        Write-Host "`r`nPlease stand by . . .  `r`n" -ForegroundColor White

        $ADParams = @{
            LDAPFilter    = '(|(msExchRecipientTypeDetails=8589934592)(msExchRecipientTypeDetails=2147483648)(msExchRecipientTypeDetails=17179869184)(msExchRecipientTypeDetails=34359738368))'
            Properties    = 'msExchVersion', 'DisplayName', 'UserPrincipalName', 'ObjectGuid'
            ResultSetSize = $null
            Server        = $domainController
        }
        $UserList = Get-ADUser @ADParams | Select-Object *
        $UserList | Export-Clixml $ADUserXML
    }

    $UserHash = @{ }
    foreach ($User in $UserList) {
        $UserHash[$User.ObjectGuid.ToString()] = @{
            msExchVersion     = $User.msExchVersion
            DisplayName       = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
        }
    }

    $VersionList = $UserList | Group-Object msExchVersion | Sort-Object Count -Descending
    $ShowVersion = [System.Collections.Generic.List[PSCustomObject]]::New()
    foreach ($Version in $VersionList) {
        $ShowVersion.Add([PSCustomObject][ordered]@{
                'Count'   = $Version.Count
                'Version' = $Version.Name
            })
    }
    $ShowVersion | Out-GridView -Title 'Current breakdown of msExchVersion found in Remote Mailboxes'
    if ($ShowVersionCountOnly) { continue }
    if (-not $SkipConnection) {
        Get-PSSession | Remove-PSSession
        Connect-Exchange -DontViewEntireForest:$DontViewEntireForest -PromptConfirm
    }
    $RemoteMailboxXML = Join-Path -Path $PS365Path -ChildPath 'RemoteMailbox_msExchVersion.xml'
    Write-Host 'Fetching Remote Mailboxes...' -ForegroundColor Cyan

    Get-RemoteMailbox -DomainController $DomainController -ResultSize Unlimited | Select-Object * | Export-Clixml $RemoteMailboxXML
    $RemoteMailboxList = Import-Clixml $RemoteMailboxXML | Sort-Object DisplayName, OrganizationalUnit

    Write-Host " Remote Mailboxes found in Active Directory (via msExchRecipientTypeDetails). Count:  $($UserList.Count)  " -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host " Remote Mailboxes found in Exchange (via Get-RemoteMailbox). Count: $($RemoteMailboxList.Count)  " -ForegroundColor DarkBlue -BackgroundColor White
    Write-Host '  >> We will not modify any users in Active Directory unless a matching GUID is found in Exchange <<  ' -ForegroundColor DarkRed -BackgroundColor White

    $RMHash = Get-RemoteMailboxHash -Key Guid -RemoteMailboxList $RemoteMailboxList

    Write-Host 'Choose which Remote Mailboxes to modify msExchVersion - Prior to modification, you will choose which version' -ForegroundColor Black -BackgroundColor White
    Write-Host 'To select use Ctrl/Shift + click (Individual) or Ctrl + A (All)' -ForegroundColor Black -BackgroundColor White

    $Choice = Select-SetmsExchVersion -RemoteMailboxList $RemoteMailboxList -UserHash $UserHash |
    Out-GridView -OutputMode Multiple -Title 'Choose which Remote Mailboxes to modify msExchVersion'
    $ChoiceCSV = Join-Path -Path $PS365Path -ChildPath ('Before modify msExchVersion {0}.csv' -f [DateTime]::Now.ToString('yyyy-MM-dd-hhmm'))
    $Choice | Export-Csv $ChoiceCSV -NoTypeInformation -Encoding UTF8

    if ($Choice) { Get-DecisionbyOGV } else { Write-Host 'Halting as nothing was selected' ; continue }

    $VersionDecision = @(
        [PSCustomObject][ordered]@{
            Version       = 'Exchange2007'
            msExchVersion = '4535486012416'
        }
        [PSCustomObject][ordered]@{
            Version       = 'Exchange2010'
            msExchVersion = '44220983382016'
        }
        [PSCustomObject][ordered]@{
            Version       = 'Exchange2013'
            msExchVersion = '88218628259840'
        }
        [PSCustomObject][ordered]@{
            Version       = 'Exchange2016'
            msExchVersion = '1125899906842624'
        }) | Out-GridView -OutputMode Single -Title 'Choose the msExchVersion to apply to the mailboxes you just selected'
    if ($VersionDecision) { Get-DecisionbyOGV } else { Write-Host 'Halting as nothing was selected' ; continue }

    $Result = Invoke-SetmsExchVersion -DomainController $DomainController -Choice $Choice -RMHash $RMHash -UserHash $UserHash -VersionDecision $VersionDecision.msExchVersion

    $Result | Out-GridView -Title ('Results of modifying msExchVersion to version: {0}  [ Count: {1} ]' -f $VersionDecision.msExchVersion, $Result.Count)
    $ResultCSV = Join-Path -Path $PS365Path -ChildPath ('After modify msExchVersion {0}.csv' -f [DateTime]::Now.ToString('yyyy-MM-dd-hhmm'))
    $Result | Export-Csv $ResultCSV -NoTypeInformation
}
