function Convert-SyncedUsersToCloud {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String[]]$UserPrincipalName
    )

    Connect-MgGraph -Scopes 'User.ReadWrite.All'

    # The user(s) must be restored first : Restore-MgDirectoryDeletedItem -DirectoryObjectId $directoryObjectId
    # Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/Users/$($userObj.id)" -Body @{OnPremisesImmutableId = $null} -ErrorAction Stop
    foreach ($upn in $UserPrincipalName) {
        Write-Host -ForegroundColor Cyan "$upn - setting ImmutableId to $null"

        try {
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/Users/$upn" -Body @{OnPremisesImmutableId = $null } -ErrorAction Stop
        }
        catch {
            Write-Host -ForegroundColor Red "$upn = Unable to set ImmuableID = `$null. Error:$($_.Exception.Message)"
        }
    }
}