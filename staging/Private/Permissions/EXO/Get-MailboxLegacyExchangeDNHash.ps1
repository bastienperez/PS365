function GetMailboxLegacyExchangeDNHash {
    <#
    .SYNOPSIS

    .EXAMPLE

    #>
    param (

    )
    Begin {
        $MailboxLegacyExchangeDNHash = @{ }
    }

    Process {
        $MailboxLegacyExchangeDNHash[$_.LegacyExchangeDN] = $_.PrimarySMTPAddress
    }
    End {
        $MailboxLegacyExchangeDNHash
    }
}
