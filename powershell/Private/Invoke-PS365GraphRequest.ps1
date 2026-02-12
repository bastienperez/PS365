<#
    .SYNOPSIS
    Enhanced version of Invoke-MgGraphRequest that supports caching and automatic pagination.

    .DESCRIPTION
    Wrapper around Invoke-MgGraphRequest that caches GET responses, handles automatic pagination
    via @odata.nextLink, and provides convenient parameters for common OData query options.

    .PARAMETER Uri
    The Microsoft Graph API URI to call.

    .PARAMETER Method
    The HTTP method to use. Defaults to 'GET'.

    .PARAMETER OutputType
    The output type for the response (e.g., 'PSObject').

    .PARAMETER Headers
    Optional headers to include in the request.

    .PARAMETER DisableCache
    If specified, bypasses the cache and makes a direct Graph API call.

    .PARAMETER Body
    The request body for POST/PATCH/PUT requests.

    .PARAMETER ContentType
    The content type of the request body.

    .PARAMETER ErrorAction
    The error action preference for the underlying Invoke-MgGraphRequest call.

    .PARAMETER All
    If specified, automatically follows @odata.nextLink to retrieve all pages of results.

    .PARAMETER Select
    OData $select query parameter. Comma-separated list of properties to return.

    .PARAMETER Filter
    OData $filter query parameter.

    .PARAMETER Expand
    OData $expand query parameter. Comma-separated list of relationships to expand.

    .PARAMETER Top
    OData $top query parameter. Maximum number of results to return.

    .PARAMETER ConsistencyLevel
    Sets the ConsistencyLevel header (e.g., 'eventual') required for some advanced queries.

    .PARAMETER ApiVersion
    The Graph API version to use. Defaults to 'v1.0'. Use 'beta' for beta endpoints.
#>
function Invoke-PS365GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $Uri,

        [Parameter(Mandatory = $false)]
        [string] $Method = 'GET',

        [Parameter(Mandatory = $false)]
        [string] $OutputType,

        [Parameter(Mandatory = $false)]
        [System.Collections.IDictionary] $Headers,

        [Parameter(Mandatory = $false)]
        [switch] $DisableCache,

        [Parameter(Mandatory = $false)]
        $Body,

        [Parameter(Mandatory = $false)]
        [string] $ContentType,

        [Parameter(Mandatory = $false)]
        [System.Management.Automation.ActionPreference] $ErrorAction,

        [Parameter(Mandatory = $false)]
        [switch] $All,

        [Parameter(Mandatory = $false)]
        [string] $Select,

        [Parameter(Mandatory = $false)]
        [string] $Filter,

        [Parameter(Mandatory = $false)]
        [string] $Expand,

        [Parameter(Mandatory = $false)]
        [int] $Top,

        [Parameter(Mandatory = $false)]
        [string] $ConsistencyLevel,

        [Parameter(Mandatory = $false)]
        [ValidateSet('v1.0', 'beta')]
        [string] $ApiVersion
    )

    # Initialize cache if not already done
    if (-not $script:GraphCache) {
        $script:GraphCache = @{}
    }

    # Build URI with OData query parameters if provided
    $builtUri = $Uri

    # If Uri is a relative path without version prefix, add ApiVersion
    if ($ApiVersion -and $builtUri -notmatch '^https?://' -and $builtUri -notmatch '^/(v1\.0|beta)/') {
        $builtUri = "/$ApiVersion$builtUri"
    }

    # Build OData query string
    $queryParams = @()
    if ($Select) { $queryParams += "`$select=$Select" }
    if ($Filter) { $queryParams += "`$filter=$Filter" }
    if ($Expand) { $queryParams += "`$expand=$Expand" }
    if ($Top) { $queryParams += "`$top=$Top" }

    if ($queryParams.Count -gt 0) {
        $separator = if ($builtUri -match '\?') { '&' } else { '?' }
        $builtUri = $builtUri + $separator + ($queryParams -join '&')
    }

    # Add ConsistencyLevel header if specified
    if ($ConsistencyLevel) {
        if (-not $Headers) { $Headers = @{} }
        $Headers['ConsistencyLevel'] = $ConsistencyLevel
    }

    $results = $null

    if ($Method -eq 'GET') {
        $cacheKey = $builtUri
        $isCacheable = $true
    } elseif ($Method -eq 'POST' -and $builtUri.EndsWith('security/runHuntingQuery')) {
        $cacheKey = $builtUri + "_" + ($Body -replace '\s', '')
        $isCacheable = $true
    } else {
        $cacheKey = $null
        $isCacheable = $false
    }

    $isBatch = $builtUri.EndsWith('$batch')
    $isInCache = $cacheKey -and $script:GraphCache.ContainsKey($cacheKey)

    if (!$DisableCache -and !$isBatch -and $isInCache -and $isCacheable) {
        Write-Verbose "Using graph cache: $cacheKey"
        $results = $script:GraphCache[$cacheKey]
    }

    if (!$results) {
        Write-Verbose "Invoking Graph: $builtUri"

        # Build splat parameters
        $params = @{
            Method = $Method
            Uri    = $builtUri
        }

        if ($Headers) { $params['Headers'] = $Headers }
        if ($OutputType) { $params['OutputType'] = $OutputType }
        if ($ContentType) { $params['ContentType'] = $ContentType }
        if ($ErrorAction) { $params['ErrorAction'] = $ErrorAction }

        # Cannot use Body with GET in PS 5.1
        if ($Method -ne 'GET' -and $Body) {
            $params['Body'] = $Body
        }

        $response = Invoke-MgGraphRequest @params

        # Handle pagination with -All switch
        if ($All -and $Method -eq 'GET' -and $response -is [System.Collections.IDictionary] -and $response.ContainsKey('value')) {
            $allResults = [System.Collections.Generic.List[object]]::new()
            foreach ($item in $response.value) {
                $allResults.Add($item)
            }

            $nextLink = if ($response.ContainsKey('@odata.nextLink')) { $response['@odata.nextLink'] } else { $null }

            while ($nextLink) {
                Write-Verbose "Following @odata.nextLink: $nextLink"
                $nextParams = @{
                    Method = 'GET'
                    Uri    = $nextLink
                }
                if ($Headers) { $nextParams['Headers'] = $Headers }
                if ($OutputType) { $nextParams['OutputType'] = $OutputType }
                if ($ErrorAction) { $nextParams['ErrorAction'] = $ErrorAction }

                $nextResponse = Invoke-MgGraphRequest @nextParams

                if ($nextResponse -is [System.Collections.IDictionary] -and $nextResponse.ContainsKey('value')) {
                    foreach ($item in $nextResponse.value) {
                        $allResults.Add($item)
                    }
                }

                $nextLink = if ($nextResponse -is [System.Collections.IDictionary] -and $nextResponse.ContainsKey('@odata.nextLink')) { $nextResponse['@odata.nextLink'] } else { $null }
            }

            $results = $allResults
        } else {
            $results = $response
        }

        # Update cache for cacheable requests
        if (!$isBatch -and $isCacheable -and $cacheKey) {
            $script:GraphCache[$cacheKey] = $results
        }
    }

    return $results
}
