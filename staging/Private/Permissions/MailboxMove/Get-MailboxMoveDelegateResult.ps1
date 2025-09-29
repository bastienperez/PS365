
function Get-MailboxMoveDelegateResult {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $PermissionChoice,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $DirectionChoice,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $MailboxPermission,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $UserChoiceRegex
    )
    end {
        $MailboxPermissionRegex = (($PermissionChoice | Where-Object { $_.Options -match "FullAccess|SendAs|SendOnBehalf" }) |
            ForEach-Object { [Regex]::Escape($_.Options) }) -join '|'

        $OrElements = foreach ($Direction in $DirectionChoice.Options) {
            if ($Direction -match 'delegates') {
                '$_.PrimarySMTPAddress -match $UserChoiceRegex'
            }

            if ($Direction -match 'delegated') {
                '$_.GrantedSMTP -match $UserChoiceRegex'
            }
        }
        $AndElements = '$_.Permission -match $MailboxPermissionRegex'
        $Filter = [ScriptBlock]::Create((($OrElements -join ' -or '), $AndElements -join " -and "))
        foreach ($Permission in $MailboxPermission) {
            $Permission | Where-Object $Filter
        }
    }
}
