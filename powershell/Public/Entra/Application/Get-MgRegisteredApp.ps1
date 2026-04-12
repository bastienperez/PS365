<#
    .SYNOPSIS
    Get registered application in Microsoft Graph

    .DESCRIPTION
    Retrieves registered applications from Microsoft Graph. If no ApplicationID, ObjectID, or DisplayName is provided, returns all applications with selected properties.

    .PARAMETER ApplicationID
    The unique identifier (AppId) of the registered application to retrieve. If not provided, all applications are returned.

    .PARAMETER ObjectID
    (Optional) Retrieves the application for a specific application by its ObjectID.

    .PARAMETER DisplayName
    (Optional) Retrieves the application for a specific application by its DisplayName.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .PARAMETER NoPermissionCheck
    (Optional) Skip the Microsoft Graph scope verification performed against the current Get-MgContext token.

    .EXAMPLE
    Get-MgRegisteredApp -ApplicationID "your-application-id"

    This command retrieves the registered application with the specified Application ID (AppId).

    .EXAMPLE
    Get-MgRegisteredApp -ObjectID "xxx-xxx-xxx"

    This command retrieves the registered application with the specified ObjectID.

    .EXAMPLE
    Get-MgRegisteredApp -DisplayName "My Application"

    This command retrieves the registered application with the specified DisplayName.

    .EXAMPLE
    Get-MgRegisteredApp

    This command retrieves all registered applications.

    .EXAMPLE
    Get-MgRegisteredApp -ExportToExcel

    Exports the list of registered applications to an Excel file.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgRegisteredApp

    .NOTES
    Same as the Microsoft Graph built-in function `Get-MgApplication` but with a simplified output.
#>

function Get-MgRegisteredApp {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'ByAppId')]
        [string]$ApplicationID,

        [Parameter(Mandatory = $false, ParameterSetName = 'ByObjectId')]
        [string]$ObjectID,

        [Parameter(Mandatory = $false, ParameterSetName = 'ByDisplayName')]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false)]
        [switch]$NoPermissionCheck
    )

    if (-not $NoPermissionCheck.IsPresent) {
        $requiredScopes = @('Application.Read.All')
        if (-not (Test-MgGraphPermission -RequiredScopes $requiredScopes -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }
    }

    [System.Collections.Generic.List[PSCustomObject]]$registeredAppsArray = @()

    # Determine how to search for the Application(s): by ApplicationID (AppId), by ObjectID (GUID), by DisplayName, or all
    if ($ApplicationID) {
        # Get specific application by AppId
        $uri = "/beta/applications?`$filter=appId eq '$ApplicationID'&`$select=uniqueName,id,createdByAppId,displayName,signInAudience,disabledByMicrosoftStatus,isDisabled,appId"
        $result = Invoke-MgGraphRequest -Uri $uri -Method GET
        $apps = if ($result.Value) { $result.Value } else { @() }
    }
    elseif ($ObjectID) {
        # Get specific application by ObjectID
        $uri = "/beta/applications/$ObjectID?`$select=uniqueName,id,createdByAppId,displayName,signInAudience,disabledByMicrosoftStatus,isDisabled,appId"
        $apps = @(Invoke-MgGraphRequest -Uri $uri -Method GET)
    }
    elseif ($DisplayName) {
        # Get application by DisplayName
        $escaped = $DisplayName -replace "'", "''"
        $uri = "/beta/applications?`$filter=displayName eq '$escaped'&`$select=uniqueName,id,createdByAppId,displayName,signInAudience,disabledByMicrosoftStatus,isDisabled,appId"
        Write-Verbose "Filtering applications with: displayName eq '$escaped'"
        $result = Invoke-MgGraphRequest -Uri $uri -Method GET
        $apps = if ($result.Value) { $result.Value } else { @() }
        
        # If no exact match found, try to find apps where trimmed DisplayName matches
        if (-not $apps -or $apps.Count -eq 0) {
            Write-Verbose "No exact match found. Searching for apps with trimmed DisplayName matching '$DisplayName'..."
            $uri = "/beta/applications?`$filter=startswith(displayName,'$escaped')&`$select=uniqueName,id,createdByAppId,displayName,signInAudience,disabledByMicrosoftStatus,isDisabled,appId"
            $result = Invoke-MgGraphRequest -Uri $uri -Method GET
            $candidateApps = if ($result.Value) { $result.Value } else { @() }
            
            # Filter in PowerShell to find apps where trimmed name matches
            $apps = $candidateApps | Where-Object { $_.displayName.Trim() -eq $DisplayName }
            
            if ($apps) {
                Write-Verbose "Found $($apps.Count) application(s) with trimmed DisplayName matching '$DisplayName'"
            }
        }
    }
    else {
        # Get all applications with selected properties
        $uri = "/beta/applications?`$select=uniqueName,id,createdByAppId,displayName,signInAudience,disabledByMicrosoftStatus,isDisabled,appId"
        $result = Invoke-MgGraphRequest -Uri $uri -Method GET
        $apps = if ($result.Value) { $result.Value } else { @() }
    }

    if (-not $apps -or $apps.Count -eq 0) {
        Write-Host 'No applications found' -ForegroundColor Yellow
        return
    }

    Write-Host "$($apps.Count) application(s) found" -ForegroundColor Green

    foreach ($app in $apps) {
        # Check for leading/trailing spaces in DisplayName
        $recommendation = $null
        if ($app.displayName -ne $app.displayName.Trim()) {
            $recommendation = 'DisplayName contains leading or trailing spaces - consider renaming'
            Write-Warning "Application '$($app.displayName)' has leading or trailing spaces in the displayName"
        }

        $customApp = [PSCustomObject][ordered]@{
            AppId                     = $app.appId
            DisplayName               = $app.displayName
            Recommendation            = $recommendation
            Id                        = $app.id
            UniqueName                = $app.uniqueName
            CreatedByAppId            = $app.createdByAppId
            SignInAudience            = $app.signInAudience
            DisabledByMicrosoftStatus = $app.disabledByMicrosoftStatus
            IsDisabled                = $app.isDisabled
        }
        $registeredAppsArray.Add($customApp)
    }

    if ($ExportToExcel.IsPresent) {
        $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
        $excelFilePath = "$($env:userprofile)\$now-MgRegisteredApp.xlsx"
        Write-Host -ForegroundColor Cyan "Exporting registered applications to Excel file: $excelFilePath"
        $registeredAppsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'MgRegisteredApp'
        Write-Host -ForegroundColor Green 'Export completed successfully!'
    }
    else {
        return $registeredAppsArray
    }
}