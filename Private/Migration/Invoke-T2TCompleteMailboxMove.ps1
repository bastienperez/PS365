function Invoke-T2TCompleteMailboxMove {
    [cmdletbinding()]
    param (
        [Parameter(ValueFromPipeline, Mandatory)]
        [ValidateNotNullOrEmpty()]
        $UserList
    )
    process {
        foreach ($User in $UserList) {
            try {
                $Param = @{
                    Identity      = $User.Guid
                    BatchName     = $User.BatchName
                    Confirm       = $False
                    CompleteAfter = $null
                    ErrorAction   = 'Stop'
                }
                Set-MoveRequest @Param
                [PSCustomObject][ordered]@{
                    DisplayName   = $User.DisplayName
                    CompleteAfter = 'NULL'
                    Action        = 'SET'
                    Result        = 'Success'
                    Message       = ''
                }
            }
            catch {
                [PSCustomObject][ordered]@{
                    DisplayName      = $User.DisplayName
                    CompleteAfter    = $LocalTime
                    CompleteAfterUTC = $UTCTime
                    Action           = 'SET'
                    Result           = 'Failed'
                    Message          = $_.Exception.Message
                }
            }
        }
    }
}
