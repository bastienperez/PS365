<#
    .SYNOPSIS
    Reports the synchronization and export errors of the Microsoft Entra Connect Sync connector spaces.

    .DESCRIPTION
    Reads the connector space objects with Get-ADSyncCSObject on the Entra Connect Sync server and returns one row per
    object: the sync rules applied, the HasSyncError and HasExportError flags and the error texts.

    By default every object of every Active Directory connector is returned. Use -ErrorsOnly to keep only the objects
    in error. The password sync status is read from the object log (Get-ADSyncCSObjectLog, one call per object, slow
    on a large directory) and is collected only with -IncludePasswordSync.

    With -GroupByErrorType, the function returns a summary instead: the number of objects per error type, per
    connector and per error kind (Synchronization or Export).

    Must be run on the Entra Connect Sync server (ADSync module) with a member of the ADSyncAdmins group.

    .PARAMETER ConnectorName
    Name of one or more connectors to read. By default, all connectors of type AD.

    .PARAMETER DistinguishedName
    Distinguished name of a single object to read.

    .PARAMETER ObjectType
    Keeps only objects of these types, for example user, group, contact or computer.

    .PARAMETER ErrorsOnly
    Keeps only the objects with a synchronization error or an export error.

    .PARAMETER IncludePasswordSync
    Adds the last password sync status and timestamp of each object, and whether the password sync rule applies.
    One log query per object: combine with -ErrorsOnly or -DistinguishedName on a large directory.

    .PARAMETER GroupByErrorType
    Returns the error count per error type instead of one row per object. Implies -ErrorsOnly.

    .PARAMETER Top
    Stops after this number of returned objects.

    .PARAMETER ExportPath
    Path of a CSV file to write the result to. The result is still returned.

    .EXAMPLE
    Get-EntraConnectSyncErrorReport -ErrorsOnly

    Lists every object in synchronization or export error, on all AD connectors.

    .EXAMPLE
    Get-EntraConnectSyncErrorReport -ErrorsOnly -GroupByErrorType

    Shows how many objects fail for each error type.

    .EXAMPLE
    Get-EntraConnectSyncErrorReport -DistinguishedName 'CN=John Doe,OU=Users,DC=contoso,DC=com' -IncludePasswordSync

    Shows the rules applied to one user and its last password sync event.

    .EXAMPLE
    Get-EntraConnectSyncErrorReport -ErrorsOnly -ObjectType user -ExportPath 'C:\Temp\SyncErrors.csv'

    Exports the user objects in error to a CSV file.

    .NOTES
    Requires the ADSync module, installed with Microsoft Entra Connect Sync.
    The error text of the connector space is returned as is: the ...ErrorType columns hold its first line (for
    example permission-issue, attribute-value-must-be-unique), the ...Error columns the full text.
#>

function Get-EntraConnectSyncErrorReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string[]]$ConnectorName,

        [Parameter(Mandatory = $false)]
        [string]$DistinguishedName,

        [Parameter(Mandatory = $false)]
        [string[]]$ObjectType,

        [Parameter(Mandatory = $false)]
        [switch]$ErrorsOnly,

        [Parameter(Mandatory = $false)]
        [switch]$IncludePasswordSync,

        [Parameter(Mandatory = $false)]
        [switch]$GroupByErrorType,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Top,

        [Parameter(Mandatory = $false)]
        [string]$ExportPath
    )

    begin {
        if (-not (Get-Command -Name 'Get-ADSyncCSObject' -ErrorAction SilentlyContinue)) {
            throw 'The ADSync module is not available. Run this function on the Microsoft Entra Connect Sync server.'
        }

        if ($GroupByErrorType) {
            $ErrorsOnly = $true
        }

        [System.Collections.Generic.List[PSCustomObject]]$report = @()

        $connectors = @(Get-ADSyncConnector -ErrorAction Stop)

        if ($ConnectorName) {
            $connectors = @($connectors | Where-Object { $_.Name -in $ConnectorName })
            $unknownNames = @($ConnectorName | Where-Object { $_ -notin $connectors.Name })

            if ($unknownNames.Count -gt 0) {
                Write-Warning "[!] Connector not found: $($unknownNames -join ', ')"
            }
        }
        else {
            $connectors = @($connectors | Where-Object { $_.Type -eq 'AD' })
        }

        if ($connectors.Count -eq 0) {
            throw 'No connector to read.'
        }

        $passwordSyncRuleNames = @()
        $passwordHashSyncCloudEnabled = $null

        if ($IncludePasswordSync) {
            $passwordSyncRuleNames = @((Get-ADSyncRule | Where-Object { ($_.EnablePasswordSync) -and (-not $_.Disabled) -and ($_.Direction -eq 'Inbound') }).Name)
            $passwordHashSyncCloudEnabled = (Get-ADSyncAADCompanyFeature).PasswordHashSync
        }
    }

    process {
        $returned = 0

        foreach ($connector in $connectors) {
            if (($Top) -and ($returned -ge $Top)) {
                break
            }

            Write-Verbose "[i] Reading connector space of $($connector.Name)"

            $csParameters = @{
                ConnectorName = $connector.Name
            }

            if ($DistinguishedName) {
                $csParameters.DistinguishedName = $DistinguishedName
            }

            # Cast to array so that .Count is reliable with zero or one object
            $csObjects = @(Get-ADSyncCSObject @csParameters)

            $passwordHashSyncConnectorEnabled = $null
            if ($IncludePasswordSync) {
                $passwordHashSyncConnectorEnabled = (Get-ADSyncAADPasswordSyncConfiguration -SourceConnector $connector.Name).Enabled
            }

            $counter = 0

            foreach ($csObject in $csObjects) {
                $counter++

                if (($ObjectType) -and ($csObject.ObjectType -notin $ObjectType)) {
                    continue
                }

                $hasError = ($csObject.HasSyncError) -or ($csObject.HasExportError)

                if (($ErrorsOnly) -and (-not $hasError)) {
                    continue
                }

                if (($Top) -and ($returned -ge $Top)) {
                    break
                }

                if ($csObjects.Count -gt 100) {
                    Write-Progress -Activity "Reading $($connector.Name)" -Status "$counter / $($csObjects.Count)" -PercentComplete (($counter / $csObjects.Count) * 100)
                }

                $syncRules = @($csObject.Lineage.SyncRuleName | Where-Object { $_ })

                $syncErrorText = if ($csObject.SynchronizationError) { ([string]$csObject.SynchronizationError).Trim() } else { $null }
                $exportErrorText = if ($csObject.ExportError) { ([string]$csObject.ExportError).Trim() } else { $null }

                $row = [ordered]@{
                    ConnectorName            = $connector.Name
                    DistinguishedName        = $csObject.DistinguishedName
                    ObjectType               = $csObject.ObjectType
                    ObjectID                 = $csObject.ObjectID
                    ConnectedMVObjectId      = $csObject.ConnectedMVObjectId
                    HasSyncError             = [bool]$csObject.HasSyncError
                    SynchronizationErrorType = if ($syncErrorText) { ($syncErrorText -split '\r?\n')[0] } else { $null }
                    SynchronizationError     = $syncErrorText
                    HasExportError           = [bool]$csObject.HasExportError
                    ExportErrorType          = if ($exportErrorText) { ($exportErrorText -split '\r?\n')[0] } else { $null }
                    ExportError              = $exportErrorText
                    SyncRules                = if ($syncRules.Count -gt 0) { $syncRules -join ' -> ' } else { '-' }
                    IsTransient              = $csObject.IsTransient
                    IsPlaceHolder            = $csObject.IsPlaceHolder
                    IsConnector              = $csObject.IsConnector
                }

                if ($IncludePasswordSync) {
                    $passwordEvent = @(Get-ADSyncCSObjectLog -Identifier $csObject.ObjectID -Count 10000 | Where-Object { $_.Type -eq 'PasswordSync' }) | Select-Object -First 1

                    $row.PasswordSyncStatus = $passwordEvent.Status
                    $row.PasswordSyncTimestamp = $passwordEvent.TimeStamp
                    $row.HasPasswordSyncRule = ($syncRules | Where-Object { $_ -in $passwordSyncRuleNames }).Count -gt 0
                    $row.PasswordHashSyncConnectorEnabled = $passwordHashSyncConnectorEnabled
                    $row.PasswordHashSyncCloudEnabled = $passwordHashSyncCloudEnabled
                }

                $report.Add([PSCustomObject]$row)
                $returned++
            }

            Write-Progress -Activity "Reading $($connector.Name)" -Completed
        }
    }

    end {
        if ($GroupByErrorType) {
            [System.Collections.Generic.List[PSCustomObject]]$summary = @()

            foreach ($entry in $report) {
                if ($entry.HasSyncError) {
                    $summary.Add([PSCustomObject][ordered]@{
                            ConnectorName = $entry.ConnectorName
                            ErrorKind     = 'Synchronization'
                            ErrorType     = $entry.SynchronizationErrorType
                        })
                }
                if ($entry.HasExportError) {
                    $summary.Add([PSCustomObject][ordered]@{
                            ConnectorName = $entry.ConnectorName
                            ErrorKind     = 'Export'
                            ErrorType     = $entry.ExportErrorType
                        })
                }
            }

            $result = @($summary | Group-Object -Property ConnectorName, ErrorKind, ErrorType | ForEach-Object {
                    [PSCustomObject][ordered]@{
                        ConnectorName = $_.Group[0].ConnectorName
                        ErrorKind     = $_.Group[0].ErrorKind
                        ErrorType     = $_.Group[0].ErrorType
                        Count         = $_.Count
                    }
                } | Sort-Object -Property Count -Descending)
        }
        else {
            $result = $report
        }

        if (($ExportPath) -and ($result.Count -gt 0)) {
            # Object names and error texts come from the directory: neutralize values a spreadsheet would run as a formula
            $exportRows = foreach ($row in $result) {
                $safeRow = [ordered]@{}
                foreach ($property in $row.PSObject.Properties) {
                    $value = $property.Value
                    if (($value -is [string]) -and ($value -match '^[=+\-@\t\r]')) {
                        $value = "'" + $value
                    }
                    $safeRow[$property.Name] = $value
                }
                [PSCustomObject]$safeRow
            }
            $exportRows | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Verbose "[OK] Report exported to $ExportPath"
        }

        return $result
    }
}
