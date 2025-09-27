function Get-MailboxMovePermissionResult {
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

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $UserChoiceRegex,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $PermissionChoice,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $DirectionChoice,

        [Parameter()]
        $BatchHash,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [switch]
        $Remove
    )
    end {
        if ($PermissionChoice -match "FullAccess|SendAs|SendOnBehalf") {
            $SPPermMailbox = @{
                SharePointURL  = $SharePointURL
                ExcelFile      = 'Permissions.xlsx'
                WorksheetName  = 'Mailbox'
                NoBatch        = $true
                NoConfirmation = $true
            }
            $MailboxPermission = Import-SharePointExcel @SPPermMailbox
            $DelegateResult = @{
                PermissionChoice  = $PermissionChoice
                DirectionChoice   = $DirectionChoice
                MailboxPermission = $MailboxPermission
            }
            Get-MailboxMoveDelegateResult @DelegateResult | Where-Object { $_.PrimarySmtpAddress -and $_.GrantedSMTP -and
                $BatchHash[$_.PrimarySmtpAddress].isMigrated -ne (-not $IncludeMigrated) -and $BatchHash[$_.GrantedSMTP].isMigrated -ne (-not $IncludeMigrated)
            } | Select-Object @(
                @{
                    Name       = 'BatchName'
                    Expression = { $BatchHash[$_.PrimarySmtpAddress].BatchName }
                }
                @{
                    Name       = 'Location'
                    Expression = { 'Mailbox' }
                }
                'Object'
                'PrimarySmtpAddress'
                'Granted'
                'GrantedSMTP'
                'Permission'
                @{
                    Name       = 'Type'
                    Expression = 'DisplayType'
                }
                'Checking'
                @{
                    Name       = 'GrantedBatch'
                    Expression = { $BatchHash[$_.GrantedSMTP].BatchName }
                }
            )
        }
        if ($PermissionChoice -match "Folder") {
            $SPPermMailbox = @{
                SharePointURL  = $SharePointURL
                ExcelFile      = 'Permissions.xlsx'
                WorksheetName  = 'Folder'
                NoBatch        = $true
                NoConfirmation = $true
            }
            $FolderPermission = Import-SharePointExcel @SPPermMailbox
            $FolderResult = @{
                DirectionChoice  = $DirectionChoice
                FolderPermission = $FolderPermission
                UserChoiceRege   = $UserChoiceRegex
            }
            if ($Remove) {
                $FolderOrRights = 'Folder'
            }
            else {
                $FolderOrRights = 'AccessRights'
            }
            Get-MailboxMoveFolderResult @FolderResult | Where-Object { $_.PrimarySmtpAddress -and $_.GrantedSMTP -and
                $BatchHash[$_.PrimarySmtpAddress].isMigrated -ne (-not $IncludeMigrated) -and $BatchHash[$_.GrantedSMTP].isMigrated -ne (-not $IncludeMigrated)
            } | Select-Object @(
                @{
                    Name       = 'BatchName'
                    Expression = { $BatchHash[$_.PrimarySmtpAddress].BatchName }
                }
                @{
                    Name       = 'Location'
                    Expression = 'Folder'
                }
                'Object'
                'PrimarySmtpAddress'
                'Granted'
                'GrantedSMTP'
                @{
                    Name       = 'Permission'
                    Expression = $FolderOrRights
                }
                @{
                    Name       = 'Type'
                    Expression = 'DisplayType'
                }
                'Checking'
                @{
                    Name       = 'GrantedBatch'
                    Expression = { $BatchHash[$_.GrantedSMTP].BatchName }
                }
            )
        }
    }
}
