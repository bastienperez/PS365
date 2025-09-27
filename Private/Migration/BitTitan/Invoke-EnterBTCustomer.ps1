function Invoke-EnterBTCustomer {
    [CmdletBinding()]
    param
    (

    )
    end {
        Get-BT_Customer -RetrieveAll:$true -IsArchived:$False -SortBy_Updated_Descending
    }
}