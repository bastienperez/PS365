function Invoke-MgGraphRequestWithRetry {
    <#
    .SYNOPSIS
    Wraps Invoke-MgGraphRequest with automatic retry on HTTP 429 (throttling).

    .DESCRIPTION
    Manual pagination loops over large tenants can be throttled by Graph. Without
    a retry, the catch block typically logs a warning and continues, silently
    returning incomplete results. This helper retries on 429, honoring the
    Retry-After header when present and otherwise backing off exponentially
    (capped at 60s). Any non-429 error is rethrown unchanged.

    .PARAMETER Uri
    The Graph request URI.

    .PARAMETER Method
    HTTP method (default GET).

    .PARAMETER Body
    Optional request body.

    .PARAMETER MaxRetries
    Maximum number of retry attempts on 429 (default 5).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [string]$Method = 'GET',

        [Parameter()]
        $Body,

        [Parameter()]
        [int]$MaxRetries = 5
    )

    $attempt = 0
    while ($true) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; ErrorAction = 'Stop' }
            if ($PSBoundParameters.ContainsKey('Body')) {
                $params['Body'] = $Body
            }
            return Invoke-MgGraphRequest @params
        }
        catch {
            # Best-effort status-code extraction across SDK/exception shapes.
            $statusCode = $null
            foreach ($candidate in @(
                    { [int]$_.Exception.Response.StatusCode },
                    { [int]$_.Exception.ResponseStatusCode },
                    { [int]$_.Exception.StatusCode }
                )) {
                try {
                    $value = & $candidate
                    if ($value) { $statusCode = $value; break }
                }
                catch { }
            }

            if ($statusCode -eq 429 -and $attempt -lt $MaxRetries) {
                $attempt++
                $retryAfter = 0
                try { $retryAfter = [int]$_.Exception.Response.Headers.RetryAfter.Delta.TotalSeconds } catch { }
                if (-not $retryAfter -or $retryAfter -lt 1) {
                    $retryAfter = [int][math]::Min([math]::Pow(2, $attempt), 60)
                }
                Write-Warning "Microsoft Graph throttled the request (429). Retry $attempt/$MaxRetries after $retryAfter second(s)."
                Start-Sleep -Seconds $retryAfter
                continue
            }

            throw
        }
    }
}
