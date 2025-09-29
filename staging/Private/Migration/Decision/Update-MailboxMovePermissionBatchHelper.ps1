function Update-MailboxMovePermissionBatchHelper {
    [CmdletBinding(DefaultParameterSetName = 'SharePoint')]
    param (
        [Parameter(ParameterSetName = 'SharePoint')]
        [ValidateNotNullOrEmpty()]
        [string]
        $SharePointURL,

        [Parameter(ParameterSetName = 'SharePoint')]
        [ValidateNotNullOrEmpty()]
        [string]
        $ExcelFile,

        [Parameter()]
        $BatchLink,

        [Parameter()]
        [PSObject]
        $UserChoice,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $UserInputBatch
    )
    end {
        $BatchLink | Select-Object -ExcludeProperty BatchName -Property @(
            @{
                Name       = "BatchName"
                Expression = { if ($_.PrimarySmtpAddress -in $UserChoice.PrimarySmtpAddress) {
                        $UserInputBatch
                    }
                    else { $_.BatchName }
                }
            }
            '*'
        )
    }
}
