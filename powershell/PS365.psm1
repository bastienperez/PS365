Get-ChildItem -Path "$PSScriptRoot/Public", "$PSScriptRoot/Private" -File -Recurse *.ps1 | ForEach-Object {
    $file = $_.FullName
    try {
        . $file
    }
    catch {
        Write-Error "PS365: failed to load '$file': $($_.Exception.Message)"
    }
}
