function Select-ADConnectServer {
    param ()
    $RootPath = $env:USERPROFILE + "\ps\"
    $User = $env:USERNAME

    if (-not(Test-Path $RootPath)) {
        try {
            New-Item -ItemType Directory -Path $RootPath -ErrorAction STOP | Out-Null
        }
        catch {
            throw $_.Exception.Message
        }           
    }
    $PDCEmulator = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().PdcRoleOwner.name
    while (-not $ADConnect) {
        $ADConnect = Invoke-Command -ComputerName $PDCEmulator -ScriptBlock { ( Get-ADComputer -Filter { ( OperatingSystem -Like 'Windows*' ) -AND ( OperatingSystem -Like '*Server*' ) } ).DNSHostName } |
        Sort-Object | Out-GridView -OutputMode Single -Title "SELECT THE AD CONNECT SERVER AND CLICK OK"
    }
    $ADConnect |  Out-File ($RootPath + "$($user).ADConnectServer") -Force
}