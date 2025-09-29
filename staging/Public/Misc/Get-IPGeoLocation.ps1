<# IP-API.com
Completely free for non-commercial use
No API key required
Limit of 45 requests per minute
#>

<#FreeIPAPI
Free for commercial and non-commercial use
Limit of 60 requests per minute
No account required2
#>

<#GeoJS
Completely free and open-source
No limits mentioned
Provides essential data (country, city)
#>

function Convert-CIDRToIPRange {
    param (
        [string] $cidrNotation
    )

    $addr, $maskLength = $cidrNotation -split '/'
    [int]$maskLen = 0
    if (-not [int32]::TryParse($maskLength, [ref] $maskLen)) {
        Write-Warning 'No mask, setting to /32'
        $masklen = 32
    }
    if (0 -gt $maskLen -or $maskLen -gt 32) {
        throw 'CIDR mask length must be between 0 and 32'
    }
    $ipAddr = [Net.IPAddress]::Parse($addr)
    if ($null -eq $ipAddr) {
        throw "Cannot parse IP address: $addr"
    }
    if ($ipAddr.AddressFamily -ne [Net.Sockets.AddressFamily]::InterNetwork) {
        throw 'Can only process CIDR for IPv4'
    }

    $shiftCnt = 32 - $maskLen
    $mask = -bnot ((1 -shl $shiftCnt) - 1)
    $ipNum = [Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($ipAddr.GetAddressBytes(), 0))
    $ipStart = ($ipNum -band $mask)
    $ipEnd = ($ipNum -bor (-bnot $mask))

    # return
    return [PSCustomObject][ordered]@{
        Start = ([BitConverter]::GetBytes([Net.IPAddress]::HostToNetworkOrder($ipStart)) | ForEach-Object { $_ } ) -join '.'
        End   = ([BitConverter]::GetBytes([Net.IPAddress]::HostToNetworkOrder($ipEnd)) | ForEach-Object { $_ } ) -join '.'
    }
}


function Get-IPGeolocation {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$IPAddress,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('IP-API.com', 'FreeIPAPI', 'GeoJS')]
        [string]$APIProvider
    )

    # Function to expand CIDR notation to IP range
    
    # If input is CIDR, take only first IP from range
    $expandedIPs = foreach ($ip in $IPAddress) {
        if ($ip -like '*/*') {
            # CIDR
            $range = Convert-CIDRToIPRange $ip
            # we take only the first one
            $ip = $range.Start
        }
        elseif ($ip -like '*-*') {
            # get the first IP before '-'
            $ip = ($ip -split '-')[0]
        }
        # return IP to store in $expandedIPs
        $ip
    }


    $IPAddress = $expandedIPs

    # 13.36.124.14,15.236.248.50,15.236.252.198,13.36.124.14,15.236.252.198,15.236.248.50,13.36.124.14,15.236.252.198,15.236.248.50,20.199.100.40,20.199.100.40,13.36.124.14,15.236.252.198,15.236.248.50,81.200.41.0/24,81.200.33.21,185.189.236.11-185.189.236.52,87.253.233.100-87.253.233.250,87.253.234.100-87.253.234.250,87.253.238.217-87.253.238.224,212.234.180.194,93.93.188.175,31.172.232.198,85.31.193.42,85.31.192.42

    # Define API URLs
    $apis = @{
        'IP-API.com' = 'http://ip-api.com/json/'
        'FreeIPAPI' = 'https://freeipapi.com/api/json/'
        'GeoJS' = 'https://get.geojs.io/v1/ip/geo/'
    }

    foreach ($ip in $IPAddress) {
        try {
            $uri = "$($apis[$APIProvider])$ip"
            if ($APIProvider -eq 'geojs') { $uri += '.json' }
            
            $results = Invoke-RestMethod -Method Get -Uri $uri

            # Create custom object with mapped properties based on the API provider
            switch ($APIProvider) {
                'IP-API.com' {
                    [PSCustomObject][ordered]@{
                        IPAddress = $ip
                        Provider  = $APIProvider
                        Country   = $results.country
                        Region    = $results.regionName
                        City      = $results.city
                        ZipCode   = $results.zip
                        Latitude  = $results.lat
                        Longitude = $results.lon
                        ISP       = $results.isp
                    }
                }
                'FreeIPAPI' {
                    [PSCustomObject][ordered]@{
                        IPAddress = $ip
                        Provider  = $APIProvider
                        Country   = $results.countryName
                        Region    = $results.regionName
                        City      = $results.cityName
                        ZipCode   = $results.zipCode
                        Latitude  = $results.latitude
                        Longitude = $results.longitude
                    }
                }
                'GeoJS' {
                    [PSCustomObject][ordered]@{
                        IPAddress    = $ip
                        Provider     = $APIProvider
                        Country      = $results.country
                        Region       = $results.region
                        City         = $results.city
                        Latitude     = $results.latitude
                        Longitude    = $results.longitude
                        Organization = $results.organization
                    }
                }
            }
        }
        catch {
            Write-Error "Error retrieving data for IP $ip : $_"
        }
    }
}
