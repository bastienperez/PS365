function Set-MailboxFlag {
    param (
        [Parameter()]
        $ELCMailboxFlags = 24
    )

    if (-not ($null = Get-Module ActiveDirectory -ListAvailable)) {
        Write-Host "ActiveDirectory module for PowerShell not found! Please run from a computer with the ActiveDirectory module"
        return
    }
    Import-Module ActiveDirectory -Force

    $PS365Path = (Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath PS365)
    if (-not (Test-Path $PS365Path)) {
        $null = New-Item $PS365Path -type Directory -Force:$true -ErrorAction SilentlyContinue
    }
    $Result = Invoke-SetMailboxFlag -ELCMailboxFlags $ELCMailboxFlags
    $Result | Out-GridView -Title 'Results of Setting flag'
    $Result | Export-Csv (Join-Path $PS365Path 'RESULTS_SetMailboxFlag.csv') -NoTypeInformation -Append
}