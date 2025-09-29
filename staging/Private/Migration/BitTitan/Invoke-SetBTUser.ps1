function Invoke-SetBTUser {
    param (
        [Parameter(ValueFromPipeline, Mandatory)]
        [ValidateNotNullOrEmpty()]
        $UserList
    )
    begin {

    }
    process {
        foreach ($User in $UserList) {
            $Get = @{
                Ticket              = $BitTic
                PrimaryEmailAddress = $User.PrimarySmtpAddress
                WarningAction       = 'SilentlyContinue'
                ErrorAction         = 'Stop'
                RetrieveAll         = $true
            }
            $Set = @{
                WarningAction = 'SilentlyContinue'
                ErrorAction   = 'Stop'
            }
            switch ($User) {
                { $_.FirstName } { $Set.Add('FirstName', $User.FirstName) }
                { $_.LastName } { $Set.Add('LastName', $User.LastName) }
                { $_.DisplayName } { $Set.Add('DisplayName', $User.DisplayName) }
                # { $_.UserPrincipalName } { $Set.Add('UserPrincipalName', $User.UserPrincipalName) }
                default { }
            }
            if ($Get.PrimaryEmailAddress) {
                try {
                    $GetResult = Get-BT_CustomerEndUser @Get
                    Write-Host "User found`t: $($GetResult.PrimaryEmailAddress)" -ForegroundColor White
                    $Result = $GetResult | Set-BT_CustomerEndUser -Ticket $BitTic @Set
                    Write-Host "User set `t: $($GetResult.PrimaryEmailAddress)" -ForegroundColor Green
                    [PSCustomObject][ordered]@{
                        'DisplayName'        = '{0} {1}' -f $User.FirstName, $User.LastName
                        'PrimarySmtpAddress' = $Result.PrimaryEmailAddress
                        'UserPrincipalName'  = $Result.UserPrincipalName
                        'FirstName'          = $Result.FirstName
                        'LastName'           = $Result.LastName
                        'Result'             = 'SUCCESS'
                        'Log'                = 'SUCCESS'
                        'Action'             = 'SET'
                        'Updated'            = $Result.Updated.ToLocalTime()
                        'Id'                 = $Result.Id
                    }
                }
                catch {
                    [PSCustomObject][ordered]@{
                        'DisplayName'        = '{0} {1}' -f $User.FirstName, $User.LastName
                        'PrimarySmtpAddress' = $User.PrimarySmtpAddress
                        'UserPrincipalName'  = $User.UserPrincipalName
                        'FirstName'          = $User.FirstName
                        'LastName'           = $User.LastName
                        'Result'             = 'FAILED'
                        'Log'                = $_.Exception.Message
                        'Action'             = 'SET'
                        'Updated'            = ''
                        'Id'                 = ''
                    }
                }
            }
        }
    }
}