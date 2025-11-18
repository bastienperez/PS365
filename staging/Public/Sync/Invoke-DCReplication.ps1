<#
    .SYNOPSIS
    Force Replication on each Domain Controller in the Forest 
    
    .DESCRIPTION
    This function initiates replication between Domain Controllers in an Active Directory forest. It can perform either a directed replication between two specified Domain Controllers or a forest-wide replication to synchronize all Domain Controllers.
    It uses PowerShell remoting to execute the replication commands on the Domain Controllers.

    .PARAMETER SyncAll
    If specified, the function will perform a forest-wide replication on all Domain Controllers.
    Default behavior if no parameters are provided.

    .PARAMETER From
    The source Domain Controller from which to replicate data.
    If you specify this parameter, you must also specify the `To` parameter.

    .PARAMETER To
    The target Domain Controller to which data will be replicated.
    If you specify this parameter, you must also specify the `From` parameter.

    .EXAMPLE
    Invoke-DCReplication
    Initiates forest-wide replication on all Domain Controllers.
    Same as using the -SyncAll parameter.

    .EXAMPLE
    Invoke-DCReplication -Sync All
    Initiates forest-wide replication on all Domain Controllers.
    Same as calling the function without parameters.

    .EXAMPLE
    Invoke-DCReplication -From "DC1" -To "DC2"
    Initiates directed replication from Domain Controller "DC1" to Domain Controller "DC2".

    .NOTES
    This function requires the Active Directory PowerShell module.
    Ensure you have the necessary permissions to perform replication operations.

    #>

function Invoke-DCReplication {
    param(
        [Parameter(ParameterSetName = 'SyncAll')]
        [switch]$SyncAll,
        
        [Parameter(ParameterSetName = 'DirectedSync', Mandatory = $true)]
        [string]$From,
        
        [Parameter(ParameterSetName = 'DirectedSync', Mandatory = $true)]
        [string]$To
    )

    if ($PSCmdlet.ParameterSetName -eq 'DirectedSync') {
        ### Force Directed Replication between specific Domain Controllers ###
        Write-Host "Initiating directed replication from $From to $To..." -ForegroundColor Yellow
        
        $session = New-PSSession -ComputerName ($env:LOGONSERVER).Split('\')[2]
        Invoke-Command -Session $session -ScriptBlock {
            param($FromDC, $ToDC)
            
            try {
                # Get all naming contexts for comprehensive replication
                $namingContexts = (Get-ADRootDSE -Server $FromDC).namingContexts
                
                foreach ($nc in $namingContexts) {
                    Write-Host "Replicating naming context: $nc" -ForegroundColor Cyan
                    $result = repadmin /replicate $ToDC $FromDC $nc
                    
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "Successfully replicated $nc from $FromDC to $ToDC" -ForegroundColor Green
                    }
                    else {
                        Write-Warning "Failed to replicate $nc from $FromDC to $ToDC"
                        Write-Host $result -ForegroundColor Red
                    }
                }
            }
            catch {
                Write-Error "Error during directed replication: $_"
            }
        } -ArgumentList $From, $To
        Remove-PSSession $session
    }
    else {
        ### Force Replication on each Domain Controller in the Forest ###
        Write-Host 'Initiating forest-wide replication...' -ForegroundColor Yellow
        
        $session = New-PSSession -ComputerName ($env:LOGONSERVER).Split('\')[2]
        Invoke-Command -Session $session -ScriptBlock { 
            ((Get-ADForest).Domains | ForEach-Object { 
                Get-ADDomainController -Filter * -Server $_ 
            }).hostname | ForEach-Object { 
                repadmin /syncall /APeqd $_ 
            } 
        }
        
        Remove-PSSession $session
        
        Write-Host 'Forest-wide replication completed.' -ForegroundColor Green
    }
}