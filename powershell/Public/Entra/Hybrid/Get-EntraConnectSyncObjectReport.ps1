<#
    .SYNOPSIS
    Reports the objects of the Microsoft Entra Connect Sync connector spaces: sync rules applied, password sync status and sync or export errors.

    .DESCRIPTION
    Reads the connector space objects with Get-ADSyncCSObject on the Entra Connect Sync server and returns one row per
    object: the sync rules applied, the password sync status when asked, the HasSyncError and HasExportError flags and the error texts.

    By default every object of every Active Directory connector is returned. Use -ErrorsOnly to keep only the objects
    currently in error, on every connector: they are read with csexport.exe, which filters in the sync engine, so
    the whole connector space is not listed. The password sync status is read from the object log (Get-ADSyncCSObjectLog, one call per object, slow
    on a large directory) and is collected only with -IncludePasswordSync.

    With -GroupByErrorType, the function returns a summary instead: the number of objects per error type, per
    connector and per error kind (Synchronization or Export).

    For the health of the sync itself (service, scheduler, last runs, pending exports), use Get-EntraConnectSyncHealth.

    Must be run on the Entra Connect Sync server (ADSync module) with a member of the ADSyncAdmins group.

    .PARAMETER ConnectorName
    Name of one or more connectors to read. By default, all connectors of type AD (all connectors with -ErrorsOnly or -GroupByErrorType).

    .PARAMETER DistinguishedName
    Distinguished name of a single object to read.

    .PARAMETER ObjectType
    Keeps only objects of these types, for example user, group, contact or computer.

    .PARAMETER ErrorsOnly
    Keeps only the objects currently in import, synchronization or export error, on every connector (AD and Entra ID)
    unless -ConnectorName is set. Read with csexport.exe (fast), except with -DistinguishedName. Each row adds the
    on-premises DN, UPN, error dates, retry count and, for a duplicate attribute, the conflicting object (ErrorDetails).

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
    Get-EntraConnectSyncObjectReport -ErrorsOnly

    Lists every object in synchronization or export error, on all connectors (AD and Entra ID).

    .EXAMPLE
    Get-EntraConnectSyncObjectReport -ErrorsOnly -GroupByErrorType

    Shows how many objects fail for each error type.

    .EXAMPLE
    Get-EntraConnectSyncObjectReport -DistinguishedName 'CN=John Doe,OU=Users,DC=contoso,DC=com' -IncludePasswordSync

    Shows the rules applied to one user and its last password sync event.

    .EXAMPLE
    Get-EntraConnectSyncObjectReport -ErrorsOnly -ObjectType user -ExportPath 'C:\Temp\SyncErrors.csv'

    Exports the user objects in error to a CSV file.

    .NOTES
    Requires the ADSync module, installed with Microsoft Entra Connect Sync.
    Without -ErrorsOnly, the ...ErrorType columns hold the first line of the connector space error text (for
    example permission-issue), the ...Error columns the full text. With -ErrorsOnly, the ...ErrorType columns hold
    the error name given by csexport (for example AttributeValueMustBeUnique) and the ...Error columns its message.
#>

function Get-EntraConnectSyncObjectReport {
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
        elseif (-not $ErrorsOnly) {
            # Errors can sit on any connector (an export error to Entra ID is held by the Entra ID connector
            # space), so -ErrorsOnly reads every connector; the full object listing stays on AD connectors.
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

        if (($ErrorsOnly) -and (-not $DistinguishedName)) {
            # Get-ADSyncCSObject has no server-side error filter: listing a whole connector space to keep the
            # objects in error takes very long (the Entra ID connector holds every synced object). csexport.exe,
            # shipped with Entra Connect, filters in the sync engine and writes only the objects in error
            # (Private/Invoke-ADSyncCsExport).

            # First value of an attribute in the object holograms (the Entra ID connector space DN is a hash,
            # so the on-premises DN and the UPN are what identify the object to a reader)
            function Get-CsExportAttributeValue {
                param ($CsObject, [string]$Name)
                foreach ($section in 'synchronized-hologram', 'unapplied-export-hologram', 'pending-import-hologram') {
                    $node = $CsObject.SelectSingleNode("$section/entry/attr[@name='$Name']/value")
                    if (($null -ne $node) -and ($node.InnerText)) {
                        return $node.InnerText
                    }
                }
                return $null
            }

            function ConvertTo-CsExportDate {
                param ([string]$Value)
                $date = [datetime]::MinValue
                if (([string]::IsNullOrWhiteSpace($Value)) -or (-not [datetime]::TryParse($Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$date))) {
                    return $null
                }
                return $date
            }

            foreach ($connector in $connectors) {
                if (($Top) -and ($returned -ge $Top)) {
                    break
                }

                Write-Verbose "[i] Reading objects in error of $($connector.Name) with csexport"

                $seenObjectIds = [System.Collections.Generic.HashSet[string]]::new()

                # i = import and synchronization errors, e = export errors (one filter per csexport call)
                foreach ($filter in 'i', 'e') {
                    $xml = Invoke-ADSyncCsExport -ConnectorName $connector.Name -Filter $filter
                    if ($null -eq $xml) {
                        continue
                    }

                    foreach ($csObject in $xml.SelectNodes('/cs-objects/cs-object')) {
                        if (($Top) -and ($returned -ge $Top)) {
                            break
                        }

                        $objectId = $csObject.GetAttribute('id').Trim('{', '}')

                        # An object in import and export error comes back in both csexport calls
                        if (-not $seenObjectIds.Add($objectId)) {
                            continue
                        }

                        $objectTypeName = $csObject.GetAttribute('object-type')
                        if (($ObjectType) -and ($objectTypeName -notin $ObjectType)) {
                            continue
                        }

                        $mvLink = $csObject.SelectSingleNode('mv-link')

                        $row = [ordered]@{
                            ConnectorName            = $connector.Name
                            DistinguishedName        = $csObject.GetAttribute('cs-dn')
                            SourceDistinguishedName  = Get-CsExportAttributeValue -CsObject $csObject -Name 'onPremisesDistinguishedName'
                            UserPrincipalName        = Get-CsExportAttributeValue -CsObject $csObject -Name 'userPrincipalName'
                            DisplayName              = Get-CsExportAttributeValue -CsObject $csObject -Name 'displayName'
                            ObjectType               = $objectTypeName
                            ObjectID                 = $objectId
                            ConnectedMVObjectId      = if ($mvLink) { $mvLink.InnerText.Trim('{', '}') } else { $null }
                            HasSyncError             = $false
                            SynchronizationErrorType = $null
                            SynchronizationError     = $null
                            HasExportError           = $false
                            ExportErrorType          = $null
                            ExportError              = $null
                            ErrorDetails             = $null
                            ErrorFirstOccurred       = $null
                            ErrorLastOccurred        = $null
                            RetryCount               = $null
                        }

                        $extraDetails = [System.Collections.Generic.List[string]]::new()

                        # export-errordetail, import-errordetail... one element per kind of error
                        foreach ($detail in @($csObject.ChildNodes | Where-Object { $_.LocalName -like '*-errordetail' })) {
                            $errorName = $detail.SelectSingleNode('.//error-name')
                            $errorLiteral = $detail.SelectSingleNode('.//error-literal')
                            $errorTypeAttribute = $detail.GetAttribute('error-type')

                            $errorType = if (($errorName) -and ($errorName.InnerText)) { $errorName.InnerText } elseif ($errorTypeAttribute) { $errorTypeAttribute } else { $detail.LocalName }
                            # The engine wraps personal data in <pii> tags: drop them, keep the value
                            $errorText = if (($errorLiteral) -and ($errorLiteral.InnerText)) { $errorLiteral.InnerText.Trim() } else { $detail.InnerText.Trim() }
                            $errorText = $errorText -replace '</?pii>', ''

                            if ($detail.LocalName -like 'export-*') {
                                $row.HasExportError = $true
                                $row.ExportErrorType = $errorType
                                $row.ExportError = $errorText
                            }
                            else {
                                $row.HasSyncError = $true
                                $row.SynchronizationErrorType = $errorType
                                $row.SynchronizationError = $errorText
                            }

                            # Conflicting object and attribute of an AttributeValueMustBeUnique error, as JSON
                            foreach ($extra in $detail.SelectNodes('.//extra-error-details')) {
                                if ($extra.InnerText) {
                                    $extraDetails.Add($extra.InnerText.Trim())
                                }
                            }

                            $firstOccurred = ConvertTo-CsExportDate -Value $detail.GetAttribute('first-occurred')
                            $lastOccurred = ConvertTo-CsExportDate -Value $detail.GetAttribute('date-occurred')
                            if (($firstOccurred) -and ((-not $row.ErrorFirstOccurred) -or ($firstOccurred -lt $row.ErrorFirstOccurred))) {
                                $row.ErrorFirstOccurred = $firstOccurred
                            }
                            if (($lastOccurred) -and ((-not $row.ErrorLastOccurred) -or ($lastOccurred -gt $row.ErrorLastOccurred))) {
                                $row.ErrorLastOccurred = $lastOccurred
                            }

                            $retryCount = $detail.GetAttribute('retry-count')
                            if (($retryCount) -and ($null -ne ($retryCount -as [int]))) {
                                $row.RetryCount = [int]$row.RetryCount + [int]$retryCount
                            }
                        }

                        if ($extraDetails.Count -gt 0) {
                            $row.ErrorDetails = $extraDetails -join [Environment]::NewLine
                        }

                        if ($IncludePasswordSync) {
                            $passwordEvent = @(Get-ADSyncCSObjectLog -Identifier $objectId -Count 10000 | Where-Object { $_.Type -eq 'PasswordSync' }) | Select-Object -First 1

                            $row.PasswordSyncStatus = $passwordEvent.Status
                            $row.PasswordSyncTimestamp = $passwordEvent.TimeStamp
                            $row.PasswordHashSyncCloudEnabled = $passwordHashSyncCloudEnabled
                        }

                        $report.Add([PSCustomObject]$row)
                        $returned++
                    }
                }
            }

            return
        }

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
            if (($IncludePasswordSync) -and ($connector.Type -eq 'AD')) {
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
