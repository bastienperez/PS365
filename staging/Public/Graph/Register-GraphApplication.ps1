function Register-GraphApplication {
    <#
    .SYNOPSIS
    Register Apps with preset permissions for quick access to graph endpoints

    .DESCRIPTION
    Register Apps with preset permissions for quick access to graph endpoints
    Use those permissions with the connection script, Connect-PoshGraph
    Please check the Microsoft Entra ID app that this app creates to understand the permissions you have prior to running any commands.

    Make sure you that clearly understand and inspect any script before you run them!!!
    I am not responsible for any data in your tenant.  Please test thoroughly.

    If you want to add or remove permissions you can find your app here:
    https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps

    Please see examples!

    .PARAMETER Tenant
    Use this to uniquely identify the tenant and permissions.
    You will use this to connect to graph with "Connect-PoshGraph"

    Please see examples!

    .PARAMETER App
    Currently just Intune and Teams to choose from, but more to follow.

    Note: The name of the app in Microsoft Entra ID will be named Intune + the date/time it was added (but you won't need this information to connect)

    .PARAMETER AddDelegateCredentials
    A GUI will appear, type username and password and click "Export Tenant Credentials"

    .EXAMPLE

    Register-GraphApplication -Tenant Contoso -App Intune

    # The registration is a one-time thing.
    # Once it is complete, use the below command each time to connect to Graph
    Connect-PoshGraph -Tenant Contoso


    .EXAMPLE

    Register-GraphApplication -Tenant ContosoIntune -App Intune


    # The registration is a one-time thing.
    # Once it is complete, use the below command each time to connect to Graph
    Connect-PoshGraph -Tenant ContosoIntune


    .NOTES
    General notes

    #>
    [CmdletBinding(DefaultParameterSetName = 'PlaceHolder')]
    param (
        [Parameter(Mandatory)]
        $Tenant,

        [Parameter(Mandatory)]
        [ValidateSet('EXO', 'Intune', 'Teams', 'IntunePrivileged')]
        $App,

        [Parameter()]
        [switch]
        $AddDelegateCredentials,

        [Parameter()]
        [switch]
        $ReturnAppObject,

        [Parameter(ParameterSetName = 'ExchangeCBA')]
        [switch]
        $AlsoCreateGraphConnection,

        [Parameter()]
        [switch]
        $GCCHigh
    )

    $PoshPath = Join-Path -Path $Env:USERPROFILE -ChildPath '.PS365/Credentials/Graph'
    $ItemSplat = @{
        Type        = 'Directory'
        Force       = $true
        ErrorAction = 'SilentlyContinue'
    }
    if (-not (Test-Path $PoshPath)) { $null = New-Item $PoshPath @ItemSplat }
    $TenantPath = Join-Path -Path $PoshPath -ChildPath $Tenant

    if (Test-Path $TenantPath) {
        if ($AlsoCreateGraphConnection -or $PSCmdlet.ParameterSetName -notcontains 'ExchangeCBA') {
            Write-Host "Connect-PoshGraph already has a connection named, $TenantPath" -ForegroundColor Yellow -NoNewline
            $UsePath = Read-Host ". Type 'YES' to overwrite"
            if ($UsePath -ne 'YES') {
                Write-Host "Please rerun your command and choose another name to represent your connection" -ForegroundColor Green
                Write-Host "Perhaps, try appending the the app's function to the company name" -ForegroundColor Green
                Write-Host "For example, Contoso-Intune" -ForegroundColor Green
                continue
            }
        }
    }
    if (-not (Test-Path $TenantPath)) { $null = New-Item $TenantPath @ItemSplat }

    if ($AlsoCreateGraphConnection -or $PSCmdlet.ParameterSetName -notcontains 'ExchangeCBA') {
        Write-Host "`r`nWe will create an Microsoft Entra ID Application with the " -ForegroundColor Cyan -NoNewline
        Write-Host "$App" -ForegroundColor Green -NoNewline

        Write-Host " API permission set. Credentials will be encrypted to $TenantPath. Once complete, connect to Graph with: `r`n" -ForegroundColor Cyan
        Write-Host "Connect-PoshGraph " -ForegroundColor Yellow -NoNewline
        Write-Host "-Tenant " -ForegroundColor White -NoNewline
        Write-Host "$Tenant`r`n`r`n" -ForegroundColor Green

    }

    if (-not ($null = Get-Module -Name 'AzureAD', 'AzureADPreview' -ListAvailable)) {
        Write-Host "Installing AzureAD module" -ForegroundColor Cyan
        Install-Module -Name AzureAD -Scope CurrentUser -Force -AllowClobber
        Import-Module -Name AzureAD -Force
    }

    $Latest = Find-Module CloneApp -Repository PSGallery
    $Installed = Get-Module -Name CloneApp -ListAvailable | Sort-Object version -Descending | Select-Object -First 1
    if ($Latest.Version -ne $Installed.Version) {
        Write-Host "Installing CloneApp module" -ForegroundColor Cyan
        Install-Module -Name CloneApp -Scope CurrentUser -Force -AllowClobber
        Import-Module -Name CloneApp -Force
    }

    Write-Host "Disconnecting any possible connections to Microsoft Entra ID" -ForegroundColor White

    try { $null = Disconnect-AzureAD -ErrorAction Stop } catch { }
    try {
        Write-Host "In the window that appears, please enter your credentials to login to Microsoft Entra ID . . . " -ForegroundColor White -NoNewline
        $AzureAD = Connect-AzureAD -ErrorAction Stop
        Write-Host "Connected" -ForegroundColor Green
        Write-Host "Microsoft Entra ID Tenant: " -ForegroundColor Green -NoNewline
        Write-Host "$($AzureAD.TenantId)" -ForegroundColor White
        Write-Host "Microsoft Entra ID Account: " -ForegroundColor Green -NoNewline
        Write-Host "$($AzureAD.Account)" -ForegroundColor White
    }
    catch {
        Write-Host "Not connected to Microsoft Entra ID. " -ForegroundColor Yellow -NoNewline
        Write-Host "Please run the same command again and connect to Microsoft Entra ID." -ForegroundColor Cyan
        return
    }

    $Params = @{
        Name                = $App
        ConsentAction       = 'Both'
        GithubUsername      = 'KevinBlumenfeld'
        GistFilename        = '{0}.xml' -f $App
        SecretDurationYears = 10
        Owner               = ($AzureAD.Account).toString()
        GCCHIGH             = $GCCHigh
    }
    $NewApp = Import-TemplateApp @Params

    $ConfigObject = [PSCustomObject][ordered]@{
        TenantClientID = $NewApp.ApplicationId
        TenantTenantID = $NewApp.TenantId
        TenantObjectID = $NewApp.ObjectId
        TenantSecret   = $NewApp.Secret | ConvertTo-SecureString -AsPlainText -Force
    }

    if ($AlsoCreateGraphConnection -or $PSCmdlet.ParameterSetName -notcontains 'ExchangeCBA') {

        $TenantConfig = Join-Path -Path $TenantPath -ChildPath ('{0}Config.xml' -f $Tenant)
        [PSCustomObject][ordered]@{
            Cred     = [PSCredential]::new($ConfigObject.TenantTenantID, $ConfigObject.TenantSecret)
            ClientId = $ConfigObject.TenantClientID
        } | Export-Clixml -Path $TenantConfig

        Write-Host ('{0}Tenant configuration encrypted to: {1}{0}' -f [Environment]::NewLine, $TenantConfig)
    }

    if ($AddDelegateCredentials -or $App -notmatch 'Intune|EXO') {
        Write-Host "`r`nAlso, the Microsoft Graph Credential Export Tool will now appear. " -ForegroundColor Green
        Write-Host "In the fields labeled, Username & Password, enter: $($AzureAD.Account) and password."  -ForegroundColor Black -BackgroundColor White
        Write-Host "Then click the button, Export Tenant Credentials"  -ForegroundColor Black -BackgroundColor White
        Export-GraphConfig -Tenant $Tenant
    }

    if ($ReturnAppObject) {
        $ConfigObject
    }
}
