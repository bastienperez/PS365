function Invoke-SetMailboxMoveForward {
    [CmdletBinding(DefaultParameterSetName = 'SharePoint')]
    param (
        [Parameter(Mandatory, ParameterSetName = 'SharePoint')]
        [ValidateNotNullOrEmpty()]
        [string]
        $SharePointURL,

        [Parameter(Mandatory, ParameterSetName = 'SharePoint')]
        [ValidateNotNullOrEmpty()]
        [string]
        $ExcelFile,

        [Parameter(Mandatory, ParameterSetName = 'CSV')]
        [ValidateNotNullOrEmpty()]
        [string]
        $MailboxCSV
    )
    end {
        switch ($PSCmdlet.ParameterSetName) {
            'SharePoint' {
                $SharePointSplat = @{
                    SharePointURL = $SharePointURL
                    ExcelFile     = $ExcelFile
                }
                $UserChoice = Import-SharePointExcelDecision @SharePointSplat | Where-Object { $_.ForwardingAddress -or $_.ForwardingSmtpAddress }
            }
            'CSV' {
                $CSVSplat = @{
                    MailboxCSV = $MailboxCSV
                }
                $UserChoice = Import-MailboxCsvDecision @CSVSplat | Where-Object { $_.ForwardingAddress -or $_.ForwardingSmtpAddress }
            }
        }
        if ($UserChoice -ne 'Quit' ) {
            foreach ($User in $UserChoice) {
                $SetSplat = @{
                    warningaction = 'silentlycontinue'
                    ErrorAction   = 'Stop'
                    Identity      = $User.ExchangeGuid.toString()
                    Confirm       = $false
                    Force         = $true
                }
                switch ($User) {
                    { $_.ForwardingAddress } { $SetSplat.Add('ForwardingAddress', $_.ForwardingAddress) }
                    { $_.ForwardingSmtpAddress } { $SetSplat.Add('ForwardingSmtpAddress', $_.ForwardingSmtpAddress) }
                    { $_.DeliverToMailboxAndForward } { $SetSplat.Add('DeliverToMailboxAndForward', $_.DeliverToMailboxAndForward -as [bool]) }
                    Default { }
                }
                try {
                    Set-Mailbox @SetSplat
                    [PSCustomObject][ordered]@{
                        DisplayName  = $User.DisplayName
                        Result       = 'SUCCESS'
                        Identity     = $User.UserPrincipalName
                        ExchangeGuid = $User.ExchangeGuid.toString()
                        Forward      = @($User.ForwardingAddress, $User.ForwardingSmtpAddress).where{ $_ } -join '|'
                        Log          = 'SUCCESS'
                        Action       = 'SET'
                    }
                }
                catch {
                    [PSCustomObject][ordered]@{
                        DisplayName  = $User.DisplayName
                        Result       = 'FAILED'
                        Identity     = $User.UserPrincipalName
                        ExchangeGuid = $User.ExchangeGuid.toString()
                        Forward      = @($User.ForwardingAddress, $User.ForwardingSmtpAddress).where{ $_ } -join '|'
                        Log          = $_.Exception.Message
                        Action       = 'SET'
                    }
                }
            }
        }
    }
}
