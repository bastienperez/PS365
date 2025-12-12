<#
.SYNOPSIS
Disables a registered application in Microsoft Graph by setting its 'isDisabled' property to true.

.DESCRIPTION
This function takes an Application ID as input and sends a PATCH request to the Microsoft Graph API to
disable the specified registered application.

.PARAMETER ApplicationID
The unique identifier of the registered application to be disabled.

.EXAMPLE
Disable-MgRegisteredApp -ApplicationID "your-application-id"

This command disables the registered application with the specified Application ID.
#>
function Disable-MgRegisteredApp {
    param (
        [string]$ApplicationID
    )

    $uri = "/beta/applications/$($ApplicationID)"
    $body = @{
        'isDisabled' = $true
    }

    # test if app e
    try {
        $app = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        $displayName = $app.displayName
        Write-Host -ForegroundColor Cyan "Disabling application: $DisplayName (ID: $ApplicationID)"
    }
    catch {
        Write-Error "Application with ID '$ApplicationID' not found. Error: $_"
        return
    }
    
    # Send the PATCH request to disable the application
    try {
        Invoke-MgGraphRequest -Uri $uri2 -Method PATCH -Body $body -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to disable application with ID '$ApplicationID'. Error: $_"
        return
    }

    Write-Output "Application '$DisplayName' has been disabled."
}