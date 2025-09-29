function Convert-SyncedUsersToCloud {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String[]]$UserPrincipalName
    )

    Connect-MgGraph -Scopes 'User.ReadWrite.All'

    Write-Host -ForegroundColor Yellow 'Before running this script, make sure that the user(s) are already deleted then restored from the Microsoft 365/Microsoft Entra ID recycle bin.'
    Write-Host -ForegroundColor Yellow 'If you want to do this by CMDLET, you can use `Restore-MgDirectoryDeletedItem -DirectoryObjectId $directoryObjectId`'
    
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