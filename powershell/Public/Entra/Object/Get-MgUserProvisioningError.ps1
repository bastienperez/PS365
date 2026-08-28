<#
    .SYNOPSIS
    Lists the users carrying provisioning errors in Microsoft Entra ID.

    .DESCRIPTION
    Scans the tenant (or a single user) for the two kinds of provisioning errors exposed
    on the user object by Microsoft Graph:

    - serviceProvisioningErrors: errors raised by a downstream Microsoft 365 service
      (typically Exchange Online) when provisioning the user. The error detail is an XML
      payload, parsed to surface the service name, the error code and the description.
    - onPremisesProvisioningErrors: synchronization errors coming from Entra Connect,
      typically an attribute conflict (AttributeValueMustBeUnique on proxyAddresses or
      userPrincipalName).

    Microsoft Graph does not support server-side filtering on these properties, so the
    whole user list is retrieved and filtered client-side. One result row per error: a
    user carrying several errors produces several rows.

    Resolved service provisioning errors are excluded by default (Graph keeps them with
    isResolved = true); use -IncludeResolved to see them.

    .PARAMETER Identity
    (Optional) UserPrincipalName or object id of a single user to check. When omitted,
    the whole tenant is scanned.

    .PARAMETER ErrorSource
    (Optional) Restricts the results to one error source. Valid values: All,
    ServiceProvisioning, OnPremisesSync. Default is All.

    .PARAMETER IncludeResolved
    (Optional) Also returns the service provisioning errors flagged as resolved by
    Microsoft Graph. Ignored for on-premises errors, which Graph only keeps while they
    are active.

    .PARAMETER ForceNewToken
    Switch parameter to force getting a new token from Microsoft Graph.

    .PARAMETER ExportToExcel
    (Optional) If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-MgUserProvisioningError

    Returns every active provisioning error of the tenant, one row per error.

    .EXAMPLE
    Get-MgUserProvisioningError -Identity 'user@contoso.com'

    Returns the provisioning errors of a single user.

    .EXAMPLE
    Get-MgUserProvisioningError -ErrorSource OnPremisesSync

    Returns only the Entra Connect synchronization errors (attribute conflicts).

    .EXAMPLE
    Get-MgUserProvisioningError -IncludeResolved

    Also returns the service provisioning errors already resolved.

    .EXAMPLE
    Get-MgUserProvisioningError -ExportToExcel

    Exports the provisioning errors report to an Excel file in the user's profile directory.

    .OUTPUTS
    System.Collections.Generic.List[PSCustomObject]

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-MgUserProvisioningError

    .NOTES
    OUTPUT PROPERTIES
    UserPrincipalName     : the user carrying the error
    DisplayName           : display name of the user
    ErrorSource           : ServiceProvisioning (downstream service) or OnPremisesSync (Entra Connect)
    Service               : service instance for service errors (e.g. exchange), empty for sync errors
    Category              : error code for service errors, Graph category for sync errors
                            (PropertyConflict for AttributeValueMustBeUnique)
    Property              : attribute causing a sync error (e.g. ProxyAddresses), empty for service errors
    Description           : parsed error description for service errors, conflicting value for sync errors
    IsResolved            : True/False for service errors, empty for sync errors
    OccurredDateTime      : when the error was recorded
    OnPremisesSyncEnabled : whether the user is synchronized from on-premises
    Id                    : object id of the user

    Required Microsoft Graph permissions:
        - User.Read.All

    Version history:
        1.0 - Creation. Covers serviceProvisioningErrors (errorDetail XML parsed) and
              onPremisesProvisioningErrors, client-side filtering (no server-side filter
              support on these properties), one row per error.
#>
function Get-MgUserProvisioningError {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[PSCustomObject]])]
    param (
        [Parameter(Mandatory = $false, Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $false)]
        [ValidateSet('All', 'ServiceProvisioning', 'OnPremisesSync')]
        [string]$ErrorSource = 'All',

        [Parameter(Mandatory = $false)]
        [switch]$IncludeResolved,

        [Parameter(Mandatory = $false)]
        [switch]$ForceNewToken,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel
    )

    begin {
        [System.Collections.Generic.List[PSCustomObject]]$provisioningErrorsArray = @()

        $permissionsNeeded = @('User.Read.All')

        $isConnected = $null -ne (Get-MgContext -ErrorAction SilentlyContinue)
        if ($ForceNewToken.IsPresent) {
            $null = Disconnect-MgGraph -ErrorAction SilentlyContinue
            $isConnected = $false
        }
        if (-not $isConnected) {
            $null = Connect-MgGraph -Scopes $permissionsNeeded -NoWelcome
        }

        if (-not (Test-MgGraphPermission -RequiredScopes $permissionsNeeded -CallerName $MyInvocation.MyCommand.Name)) {
            return
        }

        $selectClause = 'id,userPrincipalName,displayName,onPremisesSyncEnabled,serviceProvisioningErrors,onPremisesProvisioningErrors'

        # errorDetail is an XML payload (serviceProvisioningXmlError). Parsed to surface
        # the readable description instead of the raw XML string.
        function ConvertFrom-ServiceErrorDetail {
            param ([string]$ErrorDetail)

            $parsed = [PSCustomObject]@{
                Service     = ''
                Category    = ''
                Description = $ErrorDetail
            }

            if ([string]::IsNullOrWhiteSpace($ErrorDetail)) {
                return $parsed
            }

            try {
                $xml = [xml]$ErrorDetail
                $serviceInstance = $xml.ServiceInstance

                if ($serviceInstance) {
                    $parsed.Service = [string]$serviceInstance.Name
                    $errorRecord = $serviceInstance.ObjectErrors.ErrorRecord

                    if ($errorRecord) {
                        $parsed.Category = ($errorRecord | ForEach-Object { [string]$_.ErrorCode } | Where-Object { $_ }) -join '; '
                        $description = ($errorRecord | ForEach-Object { [string]$_.ErrorDescription } | Where-Object { $_ }) -join '; '

                        if ($description) {
                            $parsed.Description = $description
                        }
                    }
                }
            }
            catch {
                # Not XML, or an unexpected shape: the raw string set above is kept
                Write-Verbose "Could not parse errorDetail as XML: $_"
            }

            return $parsed
        }

        function Add-UserProvisioningError {
            param ($User)

            if ($ErrorSource -ne 'OnPremisesSync') {
                foreach ($serviceError in @($User.serviceProvisioningErrors)) {
                    if ($null -eq $serviceError) {
                        continue
                    }

                    $isResolved = [bool]$serviceError.isResolved

                    if ($isResolved -and -not $IncludeResolved.IsPresent) {
                        continue
                    }

                    $parsed = ConvertFrom-ServiceErrorDetail -ErrorDetail ([string]$serviceError.errorDetail)

                    $provisioningErrorsArray.Add([PSCustomObject][ordered]@{
                            UserPrincipalName     = $User.userPrincipalName
                            DisplayName           = $User.displayName
                            ErrorSource           = 'ServiceProvisioning'
                            Service               = $parsed.Service
                            Category              = $parsed.Category
                            Property              = ''
                            Description           = $parsed.Description
                            IsResolved            = $isResolved
                            OccurredDateTime      = $serviceError.createdDateTime
                            OnPremisesSyncEnabled = $User.onPremisesSyncEnabled
                            Id                    = $User.id
                        })
                }
            }

            if ($ErrorSource -ne 'ServiceProvisioning') {
                foreach ($syncError in @($User.onPremisesProvisioningErrors)) {
                    if ($null -eq $syncError) {
                        continue
                    }

                    $provisioningErrorsArray.Add([PSCustomObject][ordered]@{
                            UserPrincipalName     = $User.userPrincipalName
                            DisplayName           = $User.displayName
                            ErrorSource           = 'OnPremisesSync'
                            Service               = ''
                            Category              = $syncError.category
                            Property              = $syncError.propertyCausingError
                            Description           = $syncError.value
                            IsResolved            = ''
                            OccurredDateTime      = $syncError.occurredDateTime
                            OnPremisesSyncEnabled = $User.onPremisesSyncEnabled
                            Id                    = $User.id
                        })
                }
            }
        }
    }

    process {
        if ($PSBoundParameters.ContainsKey('Identity')) {
            try {
                $singleUserUri = "https://graph.microsoft.com/v1.0/users/$Identity`?`$select=$selectClause"
                $user = Invoke-MgGraphRequest -Method GET -Uri $singleUserUri -ErrorAction Stop

                Add-UserProvisioningError -User $user
            }
            catch {
                Write-Warning "Unable to retrieve user '$Identity': $($_.Exception.Message)"
            }
        }
    }

    end {
        if (-not $PSBoundParameters.ContainsKey('Identity')) {
            Write-Host -ForegroundColor Cyan 'Retrieving all users with their provisioning error properties (client-side filtering, Graph does not support filtering on them)...'

            # Invoke-MgGraphRequest rather than Get-MgUser: the SDK model of older
            # Microsoft.Graph versions does not deserialize serviceProvisioningErrors
            $uri = "https://graph.microsoft.com/v1.0/users?`$select=$selectClause&`$top=999"
            $userCount = 0

            try {
                do {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                    $userCount = $userCount + @($response.value).Count

                    foreach ($user in @($response.value)) {
                        Add-UserProvisioningError -User $user
                    }

                    $uri = $response.'@odata.nextLink'
                } while ($uri)
            }
            catch {
                Write-Warning "Unable to retrieve the users: $($_.Exception.Message)"
                return
            }

            Write-Host -ForegroundColor Cyan "$userCount user(s) scanned."
        }

        $usersInError = @($provisioningErrorsArray | Select-Object -ExpandProperty UserPrincipalName -Unique)

        if ($provisioningErrorsArray.Count -gt 0) {
            Write-Host -ForegroundColor Yellow "$($provisioningErrorsArray.Count) provisioning error(s) found on $($usersInError.Count) user(s)."
        }
        else {
            Write-Host -ForegroundColor Green 'No provisioning error found.'
        }

        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $excelFilePath = "$($env:userprofile)\$now-MgUserProvisioningError.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
            $provisioningErrorsArray | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -TableStyle Light9 -WorksheetName 'Entra-ProvisioningErrors'
            Write-Host -ForegroundColor Green 'Export completed successfully!'
        }
        else {
            return $provisioningErrorsArray
        }
    }
}
