<#
    .SYNOPSIS
    Force Replication on each Domain Controller in the Forest 
    
    .EXAMPLE
    Invoke-DCReplication
  
    .EXAMPLE
    Invoke-DCReplication -Sync All

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