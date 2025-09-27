function Connect-MigrationWiz {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param
    (
        [Parameter()]
        [ValidateNotNull()]
        [mailaddress]
        $EmailAddress,

        [Parameter()]
        [switch]
        $SkipModuleCheck,

        [Parameter()]
        [switch]
        $DeleteCredential
    )
    end {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        if ( $Email ) {
            $EmailAddress = $Email
        }
        if (-not $SkipModuleCheck ) {
            if (-not (Get-Module -Name BitTitanManagement -ListAvailable).version.build -eq 85) {
                Install-Module -Name BitTitanManagement -RequiredVersion 0.0.85 -Force -Scope CurrentUser
                Import-Module -Name BitTitanManagement -Version 0.0.85 -Force
            }
            else {
                Import-Module -Name BitTitanManagement -Version 0.0.85 -Force
            }
        }
        $host.ui.RawUI.WindowTitle = "BitTitan Tenant: $($EmailAddress)"
        $PoshPath = Join-Path $Env:USERPROFILE '.PS365'
        $TenantPath = Join-Path $PoshPath $EmailAddress
        $CredPath = Join-Path $TenantPath 'Credentials'
        $CredFile = Join-Path $CredPath BitTitan.xml

        if (-not ($null = Test-Path $CredFile)) {
            $ItemSplat = @{
                Type        = 'Directory'
                Force       = $true
                ErrorAction = 'SilentlyContinue'
            }
            $null = New-Item $PoshPath @ItemSplat
            $null = New-Item $CredPath @ItemSplat
        }

        if (-not ($null = Test-Path $CredFile)) {
            [System.Management.Automation.PSCredential]$Credential = Get-Credential -Message 'Enter your BitTitan email address and password' -UserName $EmailAddress
            [System.Management.Automation.PSCredential]$Credential | Export-Clixml -Path $CredFile
        }
        switch ($true) {
            { $EmailAddress } {
                $Script:Email = $EmailAddress.Address
            }
            { $DeleteCredential } {
                Write-Host "Credential is being deleted now" -ForegroundColor White
                Connect-CloudDeleteCredential -CredFile $CredFile
                return
            }
            default { }
        }

        Write-Host "Obtaining MigrationWiz Ticket" -ForegroundColor White
        try {
            Get-MWTicket -CredFile $CredFile -ErrorAction Stop
            Write-Host "Successfully obtained MigrationWiz Ticket" -ForegroundColor Green
        }
        catch {
            Write-Host "Could not obtain MigrationWiz Ticket" -ForegroundColor Red
            $_.Exception.Message
        }
        $Script:StarColor = @{
            ';tag-1;' = 'RED'
            ';tag-2;' = 'GREEN'
            ';tag-3;' = 'BLUE'
            ';tag-4;' = 'ORANGE'
            ';tag-5;' = 'PURPLE'
            ';tag-6;' = 'MINT'
        }
        Enter-MWProject
        Write-Host "To switch projects type:" -NoNewline -ForegroundColor Yellow
        Write-Host " Enter-MWProject " -ForegroundColor Green -NoNewline
        Write-Host "Then hit enter and select the project from the menu that appears." -ForegroundColor Yellow
    }
}
