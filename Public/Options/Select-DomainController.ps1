function Select-DomainController {
    param ()
    $RootPath = $env:USERPROFILE + "\ps\"
    $User = $env:USERNAME
    $DomainController = $null

    if (-not(Test-Path $RootPath)) {
        try {
            New-Item -ItemType Directory -Path $RootPath -ErrorAction STOP | Out-Null
        }
        catch {
            throw $_.Exception.Message
        }           
    }

    while (-not $DomainController) {
        try {
            $DomainController = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().domains.DomainControllers.Name |  
            Out-GridView -OutputMode Single -Title "SELECT A DOMAIN CONTROLLER AND CLICK OK"
        }
        catch {
            $DomainController = (([system.directoryservices.activedirectory.domain]::GetComputerDomain()).domaincontrollers).name |  
            Out-GridView -OutputMode Single -Title "SELECT A DOMAIN CONTROLLER AND CLICK OK"
        }
    }
    $DomainController |  Out-File ($RootPath + "$($user).DomainController") -Force
}
    