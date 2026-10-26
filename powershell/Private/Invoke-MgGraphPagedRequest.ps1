function Invoke-MgGraphPagedRequest {
    <#
    .SYNOPSIS
    GETs a Microsoft Graph resource through Invoke-MgGraphRequestWithRetry and follows @odata.nextLink.

    .DESCRIPTION
    Replacement for the Microsoft Graph SDK data cmdlets (users, groups... read cmdlets), so the module
    only depends on Microsoft.Graph.Authentication. The items are written to the pipeline as they
    arrive (PSCustomObject, -OutputType PSObject), like the SDK cmdlets: a collection response emits
    each item of 'value', a single resource response emits the resource itself.

    Errors behave like an SDK cmdlet: a failed request writes a non-terminating error, so the caller
    decides with -ErrorAction (Stop to catch it, SilentlyContinue to ignore it). Throttling (429) is
    retried by Invoke-MgGraphRequestWithRetry.

    .PARAMETER Uri
    The Graph request URI, relative ('/v1.0/users?$select=id') or absolute.

    .PARAMETER Headers
    Optional request headers (e.g. ConsistencyLevel = 'eventual' for advanced queries).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [hashtable]$Headers
    )

    $nextLink = $Uri
    while ($nextLink) {
        $requestParams = @{
            Uri        = $nextLink
            Method     = 'GET'
            OutputType = 'PSObject'
        }
        if ($Headers) {
            $requestParams['Headers'] = $Headers
        }

        try {
            $page = Invoke-MgGraphRequestWithRetry @requestParams
        }
        catch {
            # Same message as the SDK cmdlets ('[Request_ResourceNotFound] : Resource ... does not exist'):
            # the Graph error code and message are in the JSON body at the end of ErrorDetails
            $errorRecord = $_
            if ([string]$_.ErrorDetails.Message -match '(?s)(\{"error".*\})\s*$') {
                try {
                    $graphError = ($Matches[1] | ConvertFrom-Json -ErrorAction Stop).error
                    $exception = [System.Exception]::new("[$($graphError.code)] : $($graphError.message)", $_.Exception)
                    $errorRecord = [System.Management.Automation.ErrorRecord]::new($exception, [string]$graphError.code, [System.Management.Automation.ErrorCategory]::InvalidOperation, $Uri)
                }
                catch {
                    Write-Verbose "Graph error body not readable: $($_.Exception.Message)"
                }
            }
            $PSCmdlet.WriteError($errorRecord)
            return
        }

        if ($null -eq $page) {
            return
        }

        # A collection answers with 'value'; a single entity says so in its @odata.context ('.../$entity')
        $isEntity = [string]$page.'@odata.context' -match '/\$entity$'
        if ($null -ne $page.PSObject.Properties['value'] -and -not $isEntity) {
            foreach ($item in $page.value) {
                $item
            }
            $nextLink = $page.'@odata.nextLink'
        }
        else {
            $page
            $nextLink = $null
        }
    }
}
