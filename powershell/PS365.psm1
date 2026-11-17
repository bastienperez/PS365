# Windows PowerShell 5.1 (Desktop edition): only the Entra Connect Sync functions are loaded. They run on the
# sync server, which often has no PowerShell 7, and are written to be 5.1-compatible. Every other function
# requires PowerShell 7 and is not loaded.
if ($PSVersionTable.PSEdition -eq 'Desktop') {
    $filesToLoad = @(
        Get-ChildItem -Path "$PSScriptRoot/Public" -File -Recurse -Filter 'Get-EntraConnectSync*.ps1'
        Get-Item -Path "$PSScriptRoot/Private/Invoke-ADSyncCsExport.ps1"
    )
    Write-Warning 'PS365: Windows PowerShell 5.1 detected, only the Entra Connect Sync functions are loaded (Get-EntraConnectSync*). Use PowerShell 7 for the full module.'
}
else {
    $filesToLoad = Get-ChildItem -Path "$PSScriptRoot/Public", "$PSScriptRoot/Private" -File -Recurse *.ps1
}

$filesToLoad | ForEach-Object {
    $file = $_.FullName
    try {
        . $file
    }
    catch {
        Write-Error "PS365: failed to load '$file': $($_.Exception.Message)"
    }
}
