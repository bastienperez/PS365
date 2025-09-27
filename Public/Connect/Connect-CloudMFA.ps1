function Connect-CloudMFA {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param
    (
        [Parameter(Mandatory)]
        [string]
        $Tenant,

        [Parameter()]
        [switch]
        $EXO2,

        [Parameter()]
        [switch]
        $ExchangeOnline,

        [Parameter()]
        [switch]
        $MSOnline,

        [Parameter()]
        [switch]
        $AzureAD,

        [Parameter()]
        [switch]
        $Compliance,

        [Parameter()]
        [switch]
        $Teams,

        [Parameter()]
        [switch]
        $SharePoint,

        [Parameter()]
        [switch]
        $DeleteCredential
    )
    end {
        if ($Tenant -match 'onmicrosoft') { $Tenant = $Tenant.Split(".")[0] }
        $host.ui.RawUI.WindowTitle = "Tenant: $($Tenant.ToUpper())"
        $PoshPath = Join-Path $Env:USERPROFILE '.PS365'
        $TenantPath = Join-Path $PoshPath $Tenant
        $CredPath = Join-Path $TenantPath 'Credentials'
        $CredFile = Join-Path $CredPath CC.xml
        $LogPath = Join-Path $TenantPath 'Logs'

        if (-not ($null = Test-Path $CredFile)) {
            $ItemSplat = @{
                Type        = 'Directory'
                Force       = $true
                ErrorAction = 'SilentlyContinue'
            }
            $null = New-Item $PoshPath @ItemSplat
            $null = New-Item $CredPath @ItemSplat
            $null = New-Item $LogPath @ItemSplat
        }

        switch ($true) {
            $DeleteCredential {
                Write-Host "Credential is being deleted now" -ForegroundColor White
                Connect-CloudDeleteCredential -CredFile $CredFile
                break
            }
            { $EXO2 -or $ExchangeOnline -or $MSOnline -or $AzureAD -or $Compliance -or
                $SharePoint -or $Teams -or $PSBoundParameters.Count -eq 1 } {
                if ($null = Test-Path $CredFile) {
                    Connect-CloudMFAClip -CredFile $CredFile
                    [System.Management.Automation.PSCredential]$Credential = Import-Clixml -Path $CredFile
                }
                else {
                    [System.Management.Automation.PSCredential]$Credential = Get-Credential -Message 'Enter Office 365 username and password'
                    [System.Management.Automation.PSCredential]$Credential | Export-Clixml -Path $CredFile
                    [System.Management.Automation.PSCredential]$Credential = Import-Clixml -Path $CredFile
                }
            }
            $MSOnline {
                Connect-CloudModuleImport -MSOnline
                Connect-MsolService
                Write-Host "Connected to Microsoft Online" -ForegroundColor Green
            }
            $AzureAD {
                Connect-CloudModuleImport -AzureAD
                $ConnectAz = Connect-AzureAD
                Write-Host ("Connected to Microsoft Entra ID ({0})" -f $ConnectAz.TenantDomain) -ForegroundColor Green
            }
            $SharePoint {
                Connect-CloudModuleImport -SharePoint
                $SharePointAdminSite = 'https://' + $Tenant + '-admin.sharepoint.com'
                Connect-SPOService -Url $SharePointAdminSite
                Write-Host "Connected to SharePoint Online" -ForegroundColor Green
            }
            $Teams {
                Connect-CloudModuleImport -Teams
                Connect-MicrosoftTeams
            }
            { $EXO2 -or $ExchangeOnline -or $Compliance } {
                $Script:RestartConsole = $null
                Connect-CloudModuleImport -EXO2
                if ($RestartConsole) { return }
                if ($EXO2 -or $ExchangeOnline) {
                    Connect-ExchangeOnline -UserPrincipalName $Credential.UserName -ShowBanner:$false
                    Write-Host "Connected to Exchange Online" -foregroundcolor Green
                }
                if ($Compliance) {
                    Get-PSSession | Remove-PSSession
                    Connect-ExchangeOnline -ConnectionUri 'https://ps.compliance.protection.outlook.com/powershell-liveid' @Splat
                    Write-Host "Connected to Security & Compliance Center" -foregroundcolor Green
                }
            }
            default { }
        }
        Get-RSJob -State Completed | Remove-RSJob -ErrorAction SilentlyContinue
    }
}
