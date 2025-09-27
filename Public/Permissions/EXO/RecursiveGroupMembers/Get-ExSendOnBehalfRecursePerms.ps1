function Get-ExSendOnBehalfRecursePerms {
    <#
    .SYNOPSIS
    Outputs Send On Behalf permissions for each mailbox that has permissions assigned.
    This is for Office 365

    .EXAMPLE

    (Get-Mailbox -ResultSize unlimited | Select-Object -expandproperty distinguishedname) | Get-ExSendOnBehalfRecursePerms | Export-csv .\SendOB.csv -NoTypeInformation

    #>
    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        $DistinguishedName,

        [parameter()]
        [hashtable] $RecipientMailHash,

        [parameter()]
        [hashtable] $RecipientHash,

        [parameter()]
        [hashtable] $RecipientDNHash,

        [parameter()]
        [hashtable] $GroupMemberHash
    )
    Begin {


    }
    Process {
        $Mailbox = $_
        $listGroupMembers = [System.Collections.Generic.HashSet[string]]::new()
        (Get-Mailbox $Mailbox -erroraction silentlycontinue).GrantSendOnBehalfTo | where-object { $_ -ne $null } |
        ForEach-Object {
            $CurGranted = $_
            if ($GroupMemberHash.ContainsKey($CurGranted) -and $GroupMemberHash[$CurGranted]) {
                $GroupMemberHash[$CurGranted] | ForEach-Object {
                    [void]$listGroupMembers.Add($_)
                }
            }
            elseif (-not($GroupMemberHash.ContainsKey($CurGranted))) {
                if ($RecipientMailHash.ContainsKey($CurGranted)) {
                    $CurGranted = $RecipientMailHash["$CurGranted"].Name
                    $Type = $RecipientMailHash["$CurGranted"].RecipientTypeDetails
                }
                $Email = $CurGranted
                if ($RecipientHash.ContainsKey($CurGranted)) {
                    $Email = $RecipientHash["$CurGranted"].PrimarySMTPAddress
                    $Type = $RecipientHash["$CurGranted"].RecipientTypeDetails
                }
                [PSCustomObject][ordered]@{
                    Mailbox              = $RecipientDNHash["$Mailbox"].Name
                    MailboxPrimarySMTP   = $RecipientDNHash["$Mailbox"].PrimarySMTPAddress
                    Granted              = $CurGranted
                    GrantedSMTP          = $Email
                    RecipientTypeDetails = $Type
                    Permission           = "SendOnBehalf"
                }
            }
        }
        if ($listGroupMembers.Count -gt 0) {
            foreach ($CurlistGroupMember in $listGroupMembers) {
                [PSCustomObject][ordered]@{
                    Mailbox              = $RecipientDNHash["$Mailbox"].Name
                    MailboxPrimarySMTP   = $RecipientDNHash["$Mailbox"].PrimarySMTPAddress
                    Granted              = $RecipientDNHash["$CurlistGroupMember"].Name
                    GrantedSMTP          = $RecipientDNHash["$CurlistGroupMember"].PrimarySMTPAddress
                    RecipientTypeDetails = $Type
                    Permission           = "SendOnBehalf"
                }
            }
        }
    }
    END {

    }
}
