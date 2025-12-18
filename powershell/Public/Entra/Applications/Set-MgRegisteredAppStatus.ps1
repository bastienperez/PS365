<#
    .SYNOPSIS
    Sets the status of a registered application in Microsoft Graph by setting its 'isDisabled' property.

    .DESCRIPTION
    This function takes an Application ID or DisplayName as input and enables or disables the registered application
    by using Invoke-MgGraphRequest to set isDisabled to true or false/null.

    .PARAMETER ApplicationID
    The unique identifier of the registered application to modify.

    .PARAMETER DisplayName
    The display name of the registered application to modify.

    .PARAMETER Status
    The status to set for the registered application. Valid values are 'Enabled' or 'Disabled'.
    
    .PARAMETER GenerateCmdlets
    If specified, the function will generate the cmdlets and save them to a file instead of executing them.

    .EXAMPLE
    Set-MgRegisteredAppStatus -ApplicationID "your-application-id" -Status "Disabled"

    This command disables the registered application with the specified Application ID.

    .EXAMPLE
    Set-MgRegisteredAppStatus -DisplayName "MyApp" -Status "Enabled"

    This command enables the registered application with the specified display name.

    .LINK
    https://ps365.clidsys.com/docs/commands/Set-MgRegisteredAppStatus

    .NOTES
    To modify a service principal/enterprise app, use the `Set-MgEnterpriseAppStatus` function instead.
#>

function Set-MgRegisteredAppStatus {
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
        $uri = "/beta/applications?`$filter=id eq '$ApplicationID'"
        $identifier = "ApplicationID: $ApplicationID"
    }
    else {
        $uri = "/beta/applications?`$filter=displayName eq '$DisplayName'"
        $identifier = "DisplayName: $DisplayName"
    }

    # Get application(s)
    try {
        $result = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to retrieve application with $identifier $_"
        return
    }
    
    $applications = $result.value

    if (-not $applications -or $applications.Count -eq 0) {
        Write-Error "Application not found with $identifier"
        return
    }

    if ($applications.Count -gt 1) {
        Write-Error "Multiple applications found with $identifier. Please use -ApplicationID parameter instead for more precise targeting."
        Write-Host 'Found the following applications:' -ForegroundColor Yellow
        foreach ($appItem in $applications) {
            Write-Host "  - DisplayName: $($appItem.displayName), ApplicationID: $($appItem.id)" -ForegroundColor Yellow
        }
        return
    }

    $app = $applications[0]

    # For registered applications: isDisabled = $false = Enabled, isDisabled = $true = Disabled
    $targetIsDisabled = if ($Status -eq 'Disabled') { $true } else { $false }
    $action = if ($Status -eq 'Disabled') { 'Disabling' } else { 'Enabling' }
    $targetStatusText = if ($Status -eq 'Disabled') { 'disabled' } else { 'enabled' }
    
    Write-Host -ForegroundColor Cyan "$action application: $($app.displayName) ($identifier)"
    
    # Check current status (isDisabled = $null or $false = Enabled, isDisabled = $true = Disabled)
    $currentlyDisabled = $app.isDisabled -eq $true
    $currentlyEnabled = -not $currentlyDisabled
    
    if (($Status -eq 'Disabled' -and $currentlyDisabled) -or ($Status -eq 'Enabled' -and $currentlyEnabled)) {
        Write-Host "Application '$($app.displayName)' is already $targetStatusText." -ForegroundColor Green
        return
    }

    # Change the application status
    $statusUri = "/beta/applications/$($app.id)"
    $body = @{
        'isDisabled' = $targetIsDisabled
    }

    if ($GenerateCmdlets) {
        $commands = @()
        $bodyJson = ConvertTo-Json -InputObject $body -Compress
        $command = "Invoke-MgGraphRequest -Uri `"$statusUri`" -Method PATCH -Body '$bodyJson'"
        $commands += $command

        return $commands
    }
    else {
        try {
            Invoke-MgGraphRequest -Uri $statusUri -Method PATCH -Body $body -ErrorAction Stop
            Write-Host "Application '$($app.displayName)' has been $targetStatusText." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to set status for application '$($app.displayName)': $_"
            return
        }
    }
}