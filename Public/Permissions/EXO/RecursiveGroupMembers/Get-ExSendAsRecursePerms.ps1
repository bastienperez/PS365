function Get-ExSendAsRecursePerms {
    <#
    .SYNOPSIS
    Outputs Send As permissions for each mailbox that has permissions assigned.
    This is for Office 365

    .EXAMPLE

    (Get-Mailbox -ResultSize unlimited | Select-Object -expandproperty distinguishedname) | Get-ExSendAsRecursePerms | Export-csv .\SA.csv -NoTypeInformation

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
        $listGroupMembers = [System.Collections.Generic.HashSet[string]]::new()
        Get-RecipientPermission $_ |
        Where-Object {
            $_.AccessRights -like "*SendAs*" -and
            !$_.IsInherited -and !$_.identity.tostring().startswith('S-1-5-21-') -and
            !$_.trustee.tostring().startswith('NT AUTHORITY\SELF') -and !$_.trustee.tostring().startswith('NULL SID')
        } | ForEach-Object {
            $Identity = $_.Identity
            $Trustee = $_.Trustee
            if ($GroupMemberHash.ContainsKey($Trustee) -and $GroupMemberHash[$Trustee]) {
                $GroupMemberHash[$Trustee] | ForEach-Object {
                    [void]$listGroupMembers.Add($_)
                }
            }
            elseif (-not($GroupMemberHash.ContainsKey($Trustee))) {
                if ($RecipientMailHash.ContainsKey($Trustee)) {
                    $Trustee = $RecipientMailHash["$Trustee"].Name
                    $Type = $RecipientMailHash["$Trustee"].RecipientTypeDetails
                }
                $Email = $Trustee
                if ($RecipientHash.ContainsKey($Trustee)) {
                    $Email = $RecipientHash["$Trustee"].PrimarySMTPAddress
                    $Type = $RecipientHash["$Trustee"].RecipientTypeDetails
                }
                [PSCustomObject][ordered]@{
                    Mailbox              = $_.Identity
                    MailboxPrimarySMTP   = $RecipientHash["$($_.Identity)"].PrimarySMTPAddress
                    Granted              = $Trustee
                    GrantedSMTP          = $Email
                    RecipientTypeDetails = $Type
                    Permission           = "SendAs"
                }
            }
        }
        if ($listGroupMembers.Count -gt 0) {
            foreach ($CurlistGroupMember in $listGroupMembers) {
                [PSCustomObject][ordered]@{
                    Mailbox              = $Identity
                    MailboxPrimarySMTP   = $RecipientHash["$($Identity)"].PrimarySMTPAddress
                    Granted              = $RecipientDNHash["$CurlistGroupMember"].Name
                    GrantedSMTP          = $RecipientDNHash["$CurlistGroupMember"].PrimarySMTPAddress
                    RecipientTypeDetails = $Type
                    Permission           = "SendAs"
                }
            }
        }
    }
    END {

    }
}
