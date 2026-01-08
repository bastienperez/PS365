<#
    .SYNOPSIS
    Sets the status of an enterprise application (service principal) in Microsoft Graph by setting its 'AccountEnabled' property.

    .DESCRIPTION
    This function takes an Application ID or DisplayName as input and enables or disables the service principal (enterprise app) 
    by using Invoke-MgGraphRequest to set AccountEnabled to true or false.

    .PARAMETER ApplicationID
    The App ID of the service principal to modify.

    .PARAMETER DisplayName
    The display name of the service principal to modify.

    .PARAMETER Status
    The status to set for the enterprise application. Valid values are 'Enabled' or 'Disabled'.

    .PARAMETER GenerateCmdlets
    If specified, the function will generate the cmdlets and save them to a file instead of executing them.

    .EXAMPLE
    Set-MgEnterpriseAppStatus -ApplicationID "12345678-1234-1234-1234-123456789012" -Status "Disabled"

    This command disables the enterprise application with the specified App ID.

    .EXAMPLE
    Set-MgEnterpriseAppStatus -DisplayName "MyApp" -Status "Enabled"

    This command enables the enterprise application with the specified display name.

    .LINK
    https://ps365.clidsys.com/docs/commands/Set-MgEnterpriseAppStatus

    .NOTES  
    To modify a registered application, use the `Set-MgRegisteredAppStatus` function instead.
#>

function Set-MgEnterpriseAppStatus {
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'ByApplicationID', Position = 0)]
        [string]$ApplicationID,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'ByDisplayName', Position = 0)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidateSet('Enabled', 'Disabled')]
        [string]$Status,

        [Parameter(Mandatory = $false)]
        [switch]$GenerateCmdlets
    )

    if ($PSCmdlet.ParameterSetName -eq 'ByApplicationID') {
        $uri = "/v1.0/servicePrincipals?`$filter=appId eq '$ApplicationID'"
        $identifier = "ApplicationID: $ApplicationID"
    }
    else {
        $uri = "/v1.0/servicePrincipals?`$filter=displayName eq '$DisplayName'"
        $identifier = "DisplayName: $DisplayName"
    }

    # Get service principal
    try {
        $result = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to retrieve service principal with $identifier $_"
        return
    }
    
    $servicePrincipals = $result.value

    if (-not $servicePrincipals -or $servicePrincipals.Count -eq 0) {
        Write-Error "Service principal not found with $identifier"
        return
    }

    if ($servicePrincipals.Count -gt 1) {
        Write-Error "Multiple service principals found with $identifier. Please use -ApplicationID parameter instead for more precise targeting."
        Write-Host 'Found the following service principals:' -ForegroundColor Yellow
        foreach ($sp in $servicePrincipals) {
            Write-Host "  - DisplayName: $($sp.displayName), AppId: $($sp.appId)" -ForegroundColor Yellow
        }
        return
    }
    else {
        $servicePrincipal = $servicePrincipals[0]
    }

    $targetStatus = $Status -eq 'Enabled'
    $action = if ($targetStatus) { 'Enabling' } else { 'Disabling' }
    $targetStatusText = if ($targetStatus) { 'enabled' } else { 'disabled' }
    
    Write-Host -ForegroundColor Cyan "$action enterprise application: $($servicePrincipal.displayName) ($identifier)"

    # Check current status
    if ($servicePrincipal.accountEnabled -eq $targetStatus) {
        Write-Host "Enterprise application '$($servicePrincipal.displayName)' is already $targetStatusText." -ForegroundColor Green
        return
    }

    # Set service principal status
    $statusUri = "/v1.0/servicePrincipals/$($servicePrincipal.id)"
    $body = @{
        'accountEnabled' = $targetStatus
    }

    if ($GenerateCmdlets) {
        $commands = @()
        $bodyJson = ConvertTo-Json -InputObject $body -Compress
        $command = "Invoke-MgGraphRequest -Uri `"$statusUri`" -Method PATCH -Body '$bodyJson'"
        $commands += $command

        return $commands
    }

    try {
        Invoke-MgGraphRequest -Uri $statusUri -Method PATCH -Body $body -ErrorAction Stop
        Write-Host "Enterprise application '$($servicePrincipal.displayName)' has been $targetStatusText." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to set status for enterprise application '$($servicePrincipal.displayName)': $_"
        return
    }
}