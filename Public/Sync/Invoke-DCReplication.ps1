function Invoke-DCReplication {
    [Alias('Sync-AD')]

    param ()
    <#
    .SYNOPSIS
    Force Replication on each Domain Controller in the Forest 
    
    .EXAMPLE
    Invoke-DCReplication
  
    .EXAMPLE
    Sync-AD

    #>
   
    ### Force Replication on each Domain Controller in the Forest ###
    $session = New-PSSession -ComputerName ($env:LOGONSERVER).Split('\')[2]
    Invoke-Command -Session $session -ScriptBlock { ((Get-ADForest).Domains | ForEach-Object { Get-ADDomainController -Filter * -Server $_ }).hostname | ForEach-Object { repadmin /syncall /APeqd $_ } }
    Remove-PSSession $session
}