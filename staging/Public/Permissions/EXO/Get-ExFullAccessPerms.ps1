function Get-ExFullAccessPerms {
    <#
    .SYNOPSIS
    Outputs Full Access permissions for each object that has permissions assigned.
    This is for On-Premises Exchange 2010, 2013, 2016+

    .EXAMPLE

    (Get-Mailbox -ResultSize unlimited | Select-Object -expandproperty distinguishedname) | Get-ExFullAccessPerms | Export-csv .\FA.csv -NoTypeInformation

    If not running from Exchange Management Shell (EMS), run this first:

    Connect-Exchange2 -NoPrefix

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
        [hashtable] $RecipientLiveIDHash

    )
    Begin {


    }
    Process {
        Write-Verbose "Inspecting: `t $_"
        Get-EXOMailboxPermission -Identity $_ -Verbose:$false |
        Where-Object {
            $_.AccessRights -like "*FullAccess*" -and
            !$_.IsInherited -and !$_.user.startswith('S-1-5-21-') -and
            !$_.user.startswith('NT AUTHORITY\SELF') -and
            !$_.user.contains('\') -and
            !$_.Deny
        } | ForEach-Object {
            $Type = $null
            $User = $_.User
            Write-Verbose "Has Full Access: `t $User"
            if ($RecipientMailHash.ContainsKey($_.User)) {
                $User = $RecipientMailHash["$($_.User)"].Name
                $Type = $RecipientMailHash["$($_.User)"].RecipientTypeDetails
            }
            $Email = $_.User
            if ($RecipientHash.ContainsKey($_.User)) {
                $Email = $RecipientHash["$($_.User)"].PrimarySMTPAddress
                $Type = $RecipientHash["$($_.User)"].RecipientTypeDetails
            }
            if ($RecipientLiveIDHash.ContainsKey($_.User)) {
                $User = $RecipientLiveIDHash["$($_.User)"].Name
                $Email = $RecipientLiveIDHash["$($_.User)"].PrimarySMTPAddress
                $Type = $RecipientLiveIDHash["$($_.User)"].RecipientTypeDetails
            }
            [PSCustomObject][ordered]@{
                Object               = $_.Identity
                PrimarySmtpAddress   = $RecipientHash["$($_.Identity)"].PrimarySMTPAddress
                Granted              = $User
                GrantedSMTP          = $Email
                RecipientTypeDetails = $Type
                Permission           = "FullAccess"
            }
        }
    }
    END {

    }
}
