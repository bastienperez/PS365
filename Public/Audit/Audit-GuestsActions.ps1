function Audit-GuestsActions {
    [System.Collections.Generic.List[Object]]$extUsers = @()

    $uri = 'https://graph.microsoft.com/v1.0/users'

    $response = Invoke-MgGraphRequest -Uri $uri -Method GET

    $response.value | ForEach-Object {
        $_ | Where-Object { $_.UserPrincipalName -like '*#EXT#*' } | ForEach-Object {
            $extUsers.Add($_)
        }
    }

    while ($response.'@odata.nextLink') {
        $response = Invoke-MgGraphRequest -Uri $response.'@odata.nextLink' -Method GET
        $response.value | ForEach-Object {
            $_ | Where-Object { $_.UserPrincipalName -like '*#EXT#*' } | ForEach-Object {
                $extUsers.Add($_)
            }
        }

        $response = Invoke-MgGraphRequest -Uri $response.'@odata.nextLink' -Method GET
    }

    $extUsers | ForEach-Object {
        $auditEventsForUser = Search-UnifiedAuditLog -EndDate $((Get-Date)) -StartDate $((Get-Date).AddDays(-365)) -UserIds $_.UserPrincipalName
        Write-Host 'Events for' $_.DisplayName 'created at' $_.WhenCreated
        $auditEventsForUser | Format-Table
    }
    return $extUsers
}