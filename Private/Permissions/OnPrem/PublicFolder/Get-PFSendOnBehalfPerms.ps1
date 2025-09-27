function Get-PFSendOnBehalfPerms {
    <#
    .SYNOPSIS
    Outputs SendOnBehalf permissions for each mailbox that has permissions assigned.
    This is for On-Premises Exchange 2010, 2013, 2016+

    .EXAMPLE

	(Get-Mailbox -ResultSize unlimited | Select-Object -expandproperty distinguishedname) | Get-PFSendOnBehalfPerms | Export-csv .\SOB.csv -NoTypeInformation

    If not running from Exchange Management Shell (EMS), run this first:

    Connect-Exchange2 -NoPrefix

    #>
    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline = $true)]
        $DistinguishedName,

        [parameter()]
        [hashtable] $ADHashDN,

        [parameter()]
        [hashtable] $ADHashCN
    )
    Begin {
        try {
            import-module activedirectory -ErrorAction Stop -Verbose:$false
        }
        catch {
            Write-Host "This module depends on the ActiveDirectory module."
            Write-Host "Please download and install from https://www.microsoft.com/en-us/download/details.aspx?id=45520"
            throw
        }
    }
    Process {
        ForEach ($curDN in $DistinguishedName) {
            $mailbox = $curDN
            Write-Verbose "Inspecting: `t $Mailbox"
            (Get-MailPublicFolder $curDN -erroraction silentlycontinue).GrantSendOnBehalfTo |
            where-object { $_ -ne $null } | ForEach-Object {
                $CN = $_
                $DisplayName = $ADHashCN["$CN"].DisplayName
                Write-Verbose "Has Send On Behalf DN: `t $DisplayName"
                Write-Verbose "                   CN: `t $CN"
                try {
                    Get-ADGroupMember "$DisplayName" -Recursive -ErrorAction stop |
                    ForEach-Object {
                        New-Object -TypeName PSObject -property @{
                            Object             = $ADHashDN["$mailbox"].DisplayName
                            UserPrincipalName  = $ADHashDN["$mailbox"].UserPrincipalName
                            PrimarySMTPAddress = $ADHashDN["$mailbox"].PrimarySMTPAddress
                            Granted            = $ADHashDN["$($_.distinguishedname)"].DisplayName
                            GrantedUPN         = $ADHashDN["$($_.distinguishedname)"].UserPrincipalName
                            GrantedSMTP        = $ADHashDN["$($_.distinguishedname)"].PrimarySMTPAddress
                            Checking           = $CN
                            GroupMember        = $($_.distinguishedname)
                            Type               = "GroupMember"
                            Permission         = "SendOnBehalf"
                        }
                    }
                }
                catch {
                    New-Object -TypeName PSObject -property @{
                        Object             = $ADHashDN["$mailbox"].DisplayName
                        UserPrincipalName  = $ADHashDN["$mailbox"].UserPrincipalName
                        PrimarySMTPAddress = $ADHashDN["$mailbox"].PrimarySMTPAddress
                        Granted            = $ADHashCN["$CN"].DisplayName
                        GrantedUPN         = $ADHashCN["$CN"].UserPrincipalName
                        GrantedSMTP        = $ADHashCN["$CN"].PrimarySMTPAddress
                        Checking           = $CN
                        GroupMember        = ""
                        Type               = "User"
                        Permission         = "SendOnBehalf"
                    }
                }
            }
        }
    }
    END {

    }
}
