function GetRecipientNameHash {
    <#
    .SYNOPSIS

    .EXAMPLE

    #>
    param (

    )
    Begin {
        $RecipientNameHash = @{ }
    }

    Process {
        $RecipientNameHash[$_.Name] = $_.PrimarySMTPAddress
    }
    End {
        $RecipientNameHash
    }
}
