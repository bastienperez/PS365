function Invoke-SetMailboxMoveConvertToShared {
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
                $UserChoice = Import-SharePointExcelDecision @SharePointSplat | Where-Object { $_.ConvertToShared -eq 'True' }
            }
            'CSV' {
                $CSVSplat = @{
                    MailboxCSV = $MailboxCSV
                }
                $UserChoice = Import-MailboxCsvDecision @CSVSplat | Where-Object { $_.ConvertToShared -eq 'True' }
            }
        }
        if ($UserChoice -ne 'Quit' ) {
            foreach ($User in $UserChoice) {
                $SetSplat = @{
                    warningaction = 'silentlycontinue'
                    ErrorAction   = 'Stop'
                    Identity      = $User.ExchangeGuid.toString()
                    Type          = 'Shared'
                }
                try {
                    Set-Mailbox @SetSplat
                    [PSCustomObject][ordered]@{
                        DisplayName  = $User.DisplayName
                        Result       = 'SUCCESS'
                        Identity     = $User.UserPrincipalName
                        ExchangeGuid = $User.ExchangeGuid.toString()
                        Log          = 'SUCCESS'
                        Action       = 'CONVERTTOSHARED'
                    }
                }
                catch {
                    [PSCustomObject][ordered]@{
                        DisplayName  = $User.DisplayName
                        Result       = 'FAILED'
                        Identity     = $User.UserPrincipalName
                        ExchangeGuid = $User.ExchangeGuid.toString()
                        Log          = $_.Exception.Message
                        Action       = 'CONVERTTOSHARED'
                    }
                }
            }
        }
    }
}
