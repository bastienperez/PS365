function Get-MailboxMoveReportData {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $WantsDetailOnTheseMoveRequests
    )
    Write-Host "Please stand by... exporting data to PS365 directory on your Desktop..." -ForegroundColor Cyan
    foreach ($Wants in $WantsDetailOnTheseMoveRequests) {
        Get-MailboxMoveReportDataHelper -Wants $Wants | Sort-Object -Property CreationTime -Descending
    }
}
