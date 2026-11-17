function Invoke-ADSyncCsExport {
    <#
    .SYNOPSIS
    Runs csexport.exe (shipped with Microsoft Entra Connect Sync) on one connector and returns the objects matching a filter.

    .DESCRIPTION
    csexport filters in the sync engine: only the matching objects are written, so the whole connector space is
    never listed. Returns the XML document (<cs-objects>), or the number of objects with -CountOnly (streamed, for
    filters that can match many objects such as pending exports). Returns $null when csexport fails, after a warning.

    .PARAMETER ConnectorName
    Name of the connector (management agent).

    .PARAMETER Filter
    csexport filter letter: i (import and synchronization errors), e (export errors), m (pending imports),
    x (pending exports), s (disconnectors).

    .PARAMETER CountOnly
    Returns the number of objects instead of the XML document.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ConnectorName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('i', 'e', 'm', 'x', 's')]
        [string]$Filter,

        [Parameter(Mandatory = $false)]
        [switch]$CountOnly
    )

    $candidates = @()
    $adSyncModule = @(Get-Module -Name ADSync) + @(Get-Module -Name ADSync -ListAvailable) | Select-Object -First 1
    if ($adSyncModule) {
        $candidates += Join-Path -Path (Split-Path -Path $adSyncModule.ModuleBase -Parent) -ChildPath 'csexport.exe'
    }
    $candidates += Join-Path -Path $env:ProgramFiles -ChildPath 'Microsoft Azure AD Sync\Bin\csexport.exe'
    $csExportPath = $candidates | Where-Object { Test-Path -Path $_ } | Select-Object -First 1

    if (-not $csExportPath) {
        throw 'csexport.exe not found (expected in the Bin folder of Microsoft Azure AD Sync).'
    }

    # csexport refuses to overwrite an existing file: always a new one
    $tempFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "csexport-$([guid]::NewGuid()).xml"

    try {
        $csExportOutput = (& $csExportPath $ConnectorName $tempFile "/f:$Filter" 2>&1 | Out-String)

        if (-not (Test-Path -Path $tempFile)) {
            if ($csExportOutput -match 'Successfully') {
                if ($CountOnly) { return 0 }
                $emptyXml = [System.Xml.XmlDocument]::new()
                $emptyXml.LoadXml('<cs-objects />')
                return , $emptyXml
            }
            Write-Warning "[!] csexport failed on $ConnectorName (/f:$Filter): $($csExportOutput.Trim())"
            return $null
        }

        if ($CountOnly) {
            $count = 0
            $reader = [System.Xml.XmlReader]::Create($tempFile)
            try {
                while ($reader.ReadToFollowing('cs-object')) {
                    $count++
                }
            }
            finally {
                $reader.Dispose()
            }
            return $count
        }

        $xml = [System.Xml.XmlDocument]::new()
        $xml.Load($tempFile)
        # XmlDocument is enumerable: the unary comma stops PowerShell from unrolling it into its child nodes
        return , $xml
    }
    catch {
        Write-Warning "[!] csexport output of $ConnectorName (/f:$Filter) not readable: $($_.Exception.Message)"
        return $null
    }
    finally {
        Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
    }
}
