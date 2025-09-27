function Invoke-GetBTUser {
    [CmdletBinding()]
    param
    (

    )
    end {
        Get-BT_CustomerEndUser -Ticket $BitTic -IsDeleted:$false -RetrieveAll:$true
    }
}