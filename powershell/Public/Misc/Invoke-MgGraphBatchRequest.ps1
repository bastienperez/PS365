<#
.SYNOPSIS
Sends multiple Microsoft Graph requests in a single HTTP call using JSON batching.

.DESCRIPTION
Wraps the Microsoft Graph $batch endpoint (https://learn.microsoft.com/en-us/graph/json-batching) to combine
up to 20 requests per HTTP call, drastically reducing round-trips for bulk read scenarios (e.g. one call per user).
Requests are automatically chunked into batches of 20. Sub-requests throttled by Graph (HTTP 429) are retried
automatically, honoring the Retry-After header announced by Graph (exponential backoff when absent).
Returns a hashtable of responses indexed by the request id, so callers can match each response back to its request.

Each request is a hashtable with the keys expected by the $batch endpoint: id (unique string), method (GET, POST...),
url (relative to the Graph version, e.g. /users/<id>/authentication/methods) and optionally body and headers.

Note: this function performs no permission check because the required scopes depend entirely on the URLs of the
requests passed by the caller. The caller is responsible for connecting with the appropriate scopes beforehand.

.PARAMETER Requests
List of request hashtables to send. Each hashtable must contain at least id, method and url keys.

.PARAMETER GraphVersion
Graph API version used for the $batch endpoint: beta (default) or v1.0.

.PARAMETER MaxRetries
Maximum number of retry rounds for throttled sub-requests (default 5).

.PARAMETER Activity
Activity name displayed by Write-Progress while batches are being processed.

.EXAMPLE
[System.Collections.Generic.List[hashtable]]$requests = @()
foreach ($user in $mgUsers) {
    $requests.Add(@{ id = "$($user.Id)"; method = 'GET'; url = "/users/$($user.Id)/authentication/methods" })
}
$responses = Invoke-MgGraphBatchRequest -Requests $requests -Activity 'Getting authentication methods'

Retrieves the authentication methods of every user in $mgUsers with 20x fewer HTTP calls, then reads each
result with $responses["$($user.Id)"].body.value.

.EXAMPLE
$responses = Invoke-MgGraphBatchRequest -Requests $requests -GraphVersion 'v1.0'

Sends the requests against the v1.0 endpoint instead of beta.

.LINK
https://ps365.clidsys.com/docs/commands/Invoke-MgGraphBatchRequest

.NOTES
[0.1.0] - 2026-08-05
# Added
- Initial version, promoted from the internal helper of Get-MgAuthNMethodInfo (PS365.Clidsys)
#>
function Invoke-MgGraphBatchRequest {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[hashtable]]$Requests,

        [Parameter(Mandatory = $false)]
        [ValidateSet('beta', 'v1.0')]
        [string]$GraphVersion = 'beta',

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 5,

        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Processing Graph batch requests'
    )

    $responsesById = @{}
    $batchSize = 20
    $totalBatches = [Math]::Ceiling($Requests.Count / $batchSize)

    for ($i = 0; $i -lt $Requests.Count; $i += $batchSize) {
        $batchNumber = [Math]::Floor($i / $batchSize) + 1
        $percentComplete = ($batchNumber / $totalBatches) * 100
        Write-Progress -Activity $Activity -Status "Batch $batchNumber / $totalBatches" -PercentComplete $percentComplete

        $endIndex = [Math]::Min($i + $batchSize - 1, $Requests.Count - 1)
        [System.Collections.Generic.List[hashtable]]$pendingRequests = @($Requests[$i..$endIndex])
        $retryCount = 0

        while ($pendingRequests.Count -gt 0) {
            $body = @{ requests = @($pendingRequests) } | ConvertTo-Json -Depth 5

            try {
                $batchResult = Invoke-MgGraphRequest -Method POST -Uri "/$GraphVersion/`$batch" -Body $body -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
            }
            catch {
                Write-Warning "Batch request failed. $($_.Exception.Message)"
                break
            }

            [System.Collections.Generic.List[hashtable]]$throttledRequests = @()
            $maxRetryAfter = 0

            foreach ($response in $batchResult.responses) {
                if ($response.status -eq 429) {
                    # Sub-request throttled, retry it after the delay announced by Graph
                    $retryAfter = 0
                    if ($response.headers.'Retry-After') {
                        $retryAfter = [int]$response.headers.'Retry-After'
                    }

                    if ($retryAfter -gt $maxRetryAfter) {
                        $maxRetryAfter = $retryAfter
                    }

                    $throttledRequest = $pendingRequests | Where-Object { $_.id -eq $response.id }
                    if ($throttledRequest) {
                        $throttledRequests.Add($throttledRequest)
                    }
                }
                else {
                    $responsesById[$response.id] = $response
                }
            }

            if ($throttledRequests.Count -eq 0) {
                break
            }

            $retryCount++
            if ($retryCount -gt $MaxRetries) {
                Write-Warning "$($throttledRequests.Count) request(s) still throttled after $MaxRetries retries, skipping them"
                break
            }

            if ($maxRetryAfter -le 0) {
                $maxRetryAfter = [Math]::Pow(2, $retryCount)  # 2s, 4s, 8s...
            }

            Write-Host -ForegroundColor Yellow "Rate limit reached, waiting $maxRetryAfter seconds before retrying $($throttledRequests.Count) request(s)..."
            Start-Sleep -Seconds $maxRetryAfter
            $pendingRequests = $throttledRequests
        }
    }

    Write-Progress -Activity $Activity -Completed

    return $responsesById
}
