using namespace System.Management.Automation.Host
function Select-CloudDataConnection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Mailboxes', 'MailUsers', 'AzureADUsers')]
        $Type,

        [Parameter(Mandatory)]
        [ValidateSet('Source', 'Target')]
        $TenantLocation,

        [Parameter()]
        [switch]
        $OnlyEXO
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Get-PSSession | Remove-PSSession
    if ($Type -eq 'AzureADUsers' -or $TenantLocation -eq 'Target' -and -not $OnlyEXO) {
        try { Disconnect-AzureAD -ErrorAction Stop } catch { }
        if (-not ($null = Get-Module -Name 'AzureAD', 'AzureADPreview' -ListAvailable)) {
            Install-Module -Name AzureAD -Scope CurrentUser -Force -AllowClobber
        }
        Write-Host "`r`nEnter credentials for $TenantLocation Microsoft Entra ID`r`n" -ForegroundColor Cyan
        $null = Connect-AzureAD
        $InitialDomain = ((Get-AzureADDomain).where{ $_.IsInitial }).Name
        Write-Host "`r`nConnected to Microsoft Entra ID Tenant: $InitialDomain`r`n" -ForegroundColor Green
    }
    if ($Type -match 'Mailboxes|MailUsers' -or ($Type -eq 'AzureADUsers' -and $TenantLocation -eq 'Source')) {
        $Script:RestartConsole = $null
        Connect-CloudModuleImport -EXO2
        if ($RestartConsole) { return }

        Write-Host "`r`nEnter credentials for $TenantLocation Tenant Exchange Online`r`n" -ForegroundColor Cyan
        Connect-ExchangeOnline
        $InitialDomain = ((Get-AcceptedDomain).where{ $_.InitialDomain }).DomainName
        Write-Host "`r`nConnected to Exchange Online Tenant: $InitialDomain`r`n" -ForegroundColor Green
    }
    if ($InitialDomain) {
        $Yes = [ChoiceDescription]::new('&Yes', 'Domain: Yes')
        $No = [ChoiceDescription]::new('&No', 'Domain: No')
        $Title = 'Please make a selection'
        $Question = "Is this the $TenantLocation tenant: $InitialDomain"
        $Options = [ChoiceDescription[]]($Yes, $No)
        $Menu = $host.ui.PromptForChoice($Title, $Question, $Options, 1)
        switch ($Menu) {
            0 { $InitialDomain }
            1 { return }
        }
    }
    else {
        Write-Host 'Halting script: Not connected' -ForegroundColor Red
        return
    }
}