function Get-MWMailboxConnector {
    [CmdletBinding()]
    param
    (

    )
    end {
        Get-MW_MailboxConnector -Ticket $MigWizTicket -RetrieveAll:$true
    }
}