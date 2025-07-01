function Invoke-EntraConnectSync {
    [Alias('Sync-AADConnect')]
    Param (
        [Parameter(Mandatory = $False)]    
        [switch] $Initial,
        [Parameter(Mandatory = $False)]
        [int]$Sleep = 10
    )
    <#
    .SYNOPSIS
    Forces Sync of Microsoft Entra ID Connect Sync (Synchronizes on premises Active Directory with Microsoft Entra ID/Office 366)

    .EXAMPLE
    # Delta Sync - This should be used unless there are extenuating circumstances

    Invoke-EntraConnectSync

    # Initial Sync
    # Typically used if an OU is added or removed from list of OUs to be synced
    # Can also be used when a normal Delta sync is not syncing a particular change
    
    Invoke-EntraConnectSync -Initial

    # Use the -Sleep switch to add a longer delay to sync if AD Connect is already in the midst of syncing
    # This only comes into play if the initial sync errors out...
    # ...at that point, the script will wait the number of seconds you specify and attempt sync again
    
    Invoke-EntraConnectSync -Sleep 240

    Invoke-EntraConnectSync -Initial -Sleep 240

    #>
    $RootPath = $env:USERPROFILE + "\ps\"
    $User = $env:USERNAME
    
    while (-not(Test-Path ($RootPath + "$($user).ADConnectServer"))) {
        Select-ADConnectServer
    }
    
    $aadComputer = Get-Content ($RootPath + "$($user).ADConnectServer")

    if ($initial) {
        Start-Job -Name ADConnectSync -ScriptBlock {
            $aadcomputer = $args[0]
            $Sleep = $args[1]
            Start-Sleep -Seconds 10
            $session = New-PSSession -ComputerName $aadComputer
            Invoke-Command -Session $session -ScriptBlock {
                $Sleep = $args[0]
                $Synced = $False
                while (-not$Synced) {
                    try {
                        Start-ADSyncSyncCycle -PolicyType Initial -erroraction Stop
                        $Synced = $True
                    }
                    catch {
                        while (Get-ADSyncConnectorRunStatus) {
                            Start-Sleep -Seconds $Sleep
                        }
                    }
                }
            } -ArgumentList $Sleep
            Remove-PSSession $session
        } -ArgumentList $aadComputer, $Sleep | Out-Null
    }
    else {
        Start-Job -Name ADConnectSync -ScriptBlock {
            $aadcomputer = $args[0]
            $Sleep = $args[1]
            Start-Sleep -Seconds 10
            $session = New-PSSession -ComputerName $aadComputer
            Invoke-Command -Session $session -ScriptBlock {
                $Sleep = $args[0]
                $Synced = $False
                while (-not$Synced) {
                    try {
                        Start-ADSyncSyncCycle -PolicyType Delta -erroraction Stop
                        $Synced = $True
                    }
                    catch {
                        while (Get-ADSyncConnectorRunStatus) {
                            Start-Sleep -Seconds $Sleep
                        }
                    }
                }
            } -ArgumentList $Sleep
            Remove-PSSession $session
        } -ArgumentList $aadComputer, $Sleep | Out-Null
    }
}

