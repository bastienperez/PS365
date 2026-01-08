<#
    .SYNOPSIS
    Get registered application in Microsoft Graph

    .DESCRIPTION
    Retrieves registered applications from Microsoft Graph. If no ApplicationID is provided, returns all applications with selected properties.

    .PARAMETER ApplicationID
    The unique identifier of the registered application to retrieve. If not provided, all applications are returned.

    .EXAMPLE
    Get-MgRegisteredApp -ApplicationID "your-application-id"

    This command retrieves the registered application with the specified Application ID.

    .EXAMPLE
    Get-MgRegisteredApp

    This command retrieves all registered applications.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgRegisteredApp

    .NOTES
    Same as the function `Get-MgApplication` but with a simplified output.
#>

function Get-MgRegisteredApp {
    param (
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$ApplicationID
    )

    [System.Collections.Generic.List[PSCustomObject]]$registeredAppsArray = @()

    if ($ApplicationID) {
        # Get specific application
        $uri = "/beta/applications/$($ApplicationID)"
        $apps = @(Invoke-MgGraphRequest -Uri $uri -Method GET)
    }
    else {
        # Get all applications with selected properties
        $uri = "/beta/applications?`$select=uniqueName,id,createdByAppId,displayName,signInAudience,disabledByMicrosoftStatus,isDisabled,appId"
        $apps = Invoke-MgGraphRequest -Uri $uri | Select-Object -ExpandProperty Value
    }

    foreach ($app in $apps) {
        $customApp = [PSCustomObject]@{
            AppId                     = $app.appId
            DisplayName               = $app.displayName
            Id                        = $app.id
            UniqueName                = $app.uniqueName
            CreatedByAppId            = $app.createdByAppId
            SignInAudience            = $app.signInAudience
            DisabledByMicrosoftStatus = $app.disabledByMicrosoftStatus
            IsDisabled                = $app.isDisabled
        }
        $registeredAppsArray.Add($customApp)
    }

    return $registeredAppsArray
}