function Invoke-CompleteMailboxMove {

    param (

        [Parameter(ValueFromPipeline, Mandatory)]
        [ValidateNotNullOrEmpty()]
        $UserList,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        $CompleteAfter

    )
    begin {
        if ($CompleteAfter) {
            $LocalTime = $CompleteAfter.ToLocalTime()
            $UTCTime = $CompleteAfter
        }
        else {
            $Yesterday = (Get-Date).AddDays(-1)
            $LocalTime = $Yesterday
            $UTCTime = $Yesterday.ToUniversalTime()
        }
    }
    process {
        foreach ($User in $UserList) {
            try {
                $Param = @{
                    Identity                   = $User.Guid
                    BatchName                  = $User.BatchName
                    SuspendWhenReadyToComplete = $False
                    Confirm                    = $False
                    CompleteAfter              = $LocalTime
                    ErrorAction                = 'Stop'
                }
                Set-MoveRequest @Param
                [PSCustomObject][ordered]@{
                    DisplayName      = $User.DisplayName
                    CompleteAfter    = $LocalTime
                    CompleteAfterUTC = $UTCTime
                    Action           = "SET"
                    Result           = "Success"
                    Message          = ""
                }
            }
            catch {
                [PSCustomObject][ordered]@{
                    DisplayName      = $User.DisplayName
                    CompleteAfter    = $LocalTime
                    CompleteAfterUTC = $UTCTime
                    Action           = "SET"
                    Result           = "Failed"
                    Message          = $_.Exception.Message
                }
            }
            try {
                Resume-MoveRequest $User.Guid
                [PSCustomObject][ordered]@{
                    DisplayName      = $User.DisplayName
                    CompleteAfter    = ""
                    CompleteAfterUTC = ""
                    Action           = "RESUME"
                    Result           = "Success"
                    Message          = ""
                }
            }
            catch {
                [PSCustomObject][ordered]@{
                    DisplayName      = $User.DisplayName
                    CompleteAfter    = ""
                    CompleteAfterUTC = ""
                    Action           = "RESUME"
                    Result           = "Failed"
                    Message          = $_.Exception.Message
                }
            }
        }
    }
}
