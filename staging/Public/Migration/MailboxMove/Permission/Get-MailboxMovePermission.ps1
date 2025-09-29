function Get-MailboxMovePermission {
    <#
    .SYNOPSIS
    Get permissions for on-premises mailboxes.
    The permissions that that mailbox has and those with permission to that mailbox

    .DESCRIPTION
    Get permissions for on-premises mailboxes.
    The permissions that that mailbox has and those with permission to that mailbox

    .PARAMETER SharePointURL
    Sharepoint url ex. https://fabrikam.sharepoint.com/sites/Contoso

    .PARAMETER ExcelFile
    Excel file found in "Shared Documents" of SharePoint site specified in SharePointURL
    ex. "Batches.xlsx"
    Minimum headers required are: BatchName, UserPrincipalName

    .PARAMETER MailboxCSV
    Path to csv of mailboxes. Minimum headers required are: BatchName, UserPrincipalName

    .EXAMPLE
    Get-MailboxMovePermission -MailboxCSV c:\scripts\batches.csv

    .EXAMPLE
    Get-MailboxMovePermission -SharePointURL 'https://fabrikam.sharepoint.com/sites/Contoso' -ExcelFile 'Batches.xlsx'

    .NOTES
    General notes
    #>

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

        [Parameter()]
        [switch]
        $IncludeMigrated,

        [Parameter()]
        [switch]
        $SkipBatchesLookup,

        [Parameter()]
        [switch]
        $UseApplyFunction,

        [Parameter()]
        [switch]
        $Remove,

        [Parameter()]
        [switch]
        $PassThru
    )
    end {
        switch ($PSCmdlet.ParameterSetName) {
            'SharePoint' {
                $SharePointSplat = @{
                    SharePointURL  = $SharePointURL
                    ExcelFile      = $ExcelFile
                    NoBatch        = $true
                    NoConfirmation = $true
                }
                $UserChoice = Import-SharePointExcelDecision @SharePointSplat
                $BatchHash = @{ }
                if (-not $SkipBatchesLookup) {
                    Import-SharePointExcel -SharePointURL $SharePointURL -ExcelFile $ExcelFile | ForEach-Object {
                        if (-not $BatchHash.ContainsKey($_.PrimarySMTPAddress)) {
                            $BatchHash.Add($_.PrimarySMTPAddress, @{
                                    BatchName  = $_.BatchName
                                    IsMigrated = $_.IsMigrated
                                })
                        }
                    }
                }
            }
            'CSV' {
                $UserChoice = Import-MailboxCsvDecision -MailboxCSV $MailboxCSV
            }
        }
        $UserChoiceRegex = '^(?:{0})$' -f (($UserChoice.PrimarySMTPAddress | ForEach-Object { [Regex]::Escape($_) }) -join '|')
        $PermissionChoice = Get-PermissionDecision
        $DirectionChoice = Get-PermissionDirectionDecision

        $PermissionResult = @{
            SharePointURL    = $SharePointURL
            ExcelFile        = $ExcelFile
            UserChoiceRegex  = $UserChoiceRegex
            PermissionChoice = $PermissionChoice
            DirectionChoice  = $DirectionChoice
        }
        switch ($true) {
            { $BatchHash } { $PermissionResult.Add('BatchHash', $BatchHash) }
            $Remove { $PermissionResult.Add('Remove', $true) }
            $IncludeMigrated { $PermissionResult.Add('IncludeMigrated', $IncludeMigrated) }
            $UseApplyfunction{
                Get-MailboxMoveApplyPermissionResult @PermissionResult | Out-GridView -Title "Choose which permissions to apply" -OutputMode Multiple
                return
            }
            $PassThru { Get-MailboxMovePermissionResult @PermissionResult | Out-GridView -Title "Permission Results" -OutputMode Multiple }
            { -not $PassThru -and -not $Link -and -not $UseApplyfunction} {
                Get-MailboxMovePermissionResult @PermissionResult | Out-GridView -Title "Permission Results"
            }
            Default { }
        }
    }
}
