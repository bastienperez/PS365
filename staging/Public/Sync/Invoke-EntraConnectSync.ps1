<#
    .SYNOPSIS
    Forces Sync of Microsoft Entra ID Connect Sync (Synchronizes on premises Active Directory with Microsoft Entra ID/Office 366)

    .DESCRIPTION
    This function initiates a synchronization cycle for Microsoft Entra ID Connect. It can perform either a delta sync or an initial sync based on the parameters provided.

    .PARAMETER Initial
    Indicates whether to perform an initial sync (as opposed to a delta sync).

    .EXAMPLE
    Invoke-EntraConnectSync

    Run a delta synchronization.

    .EXAMPLE
    Invoke-EntraConnectSync -Initial

    Run an initial synchronization. Typically used if an OU is added or removed from list of OUs to be synced or if rules are changed.
    
    Invoke-EntraConnectSync -Initial
    #>

function Invoke-EntraConnectSync {
    param (
        [Parameter(Mandatory = $False)]    
        [switch] $Initial
    )

    $RootPath = $env:USERPROFILE + '\ps\'
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
                while (-not $Synced) {
                    try {
                        Start-ADSyncSyncCycle -PolicyType Initial -ErrorAction Stop
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
                while (-not $Synced) {
                    try {
                        Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
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

