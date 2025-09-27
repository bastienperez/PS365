function Import-MailboxCsvDecisionDomainChoice {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $MailboxCsv,

        [Parameter()]
        [switch]
        $ChooseDomain
    )
    end {
        $UserChoiceSplat = @{
            DecisionObject = Import-Csv -Path $MailboxCSV
            ChooseDomain   = $ChooseDomain
        }
        $UserChoice = Get-UserDecision @UserChoiceSplat
        $UserChoice
    }
}