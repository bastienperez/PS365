function Get-FullAccessPermsRecursive {

    [CmdletBinding()]
    Param (
        [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        $DistinguishedName,

        [parameter()]
        [hashtable] $ADHashDN,

        [parameter()]
        [hashtable] $ADHash
    )
    begin {

    }
    process {
        foreach ($curDN in $DistinguishedName) {
            $mailbox = $curDN
            Write-Verbose "Inspecting: `t $mailbox"
            Get-MailboxPermission $curDN |
            Where-Object {
                $_.AccessRights -like "*FullAccess*" -and
                !$_.IsInherited -and !$_.user.tostring().startswith('S-1-5-21-') -and
                !$_.user.tostring().startswith('NT AUTHORITY\SELF') -and
                !$_.Deny
            } | ForEach-Object {
                $User = $_.User
                Write-Verbose "Has Full Access: `t $User"
                try {
                    Get-ADGroupMember ($_.user -split "\\")[1] -Recursive -ErrorAction stop |
                    ForEach-Object {
                        New-Object -TypeName PSObject -property @{
                            Object             = $ADHashDN["$mailbox"].DisplayName
                            UserPrincipalName  = $ADHashDN["$mailbox"].UserPrincipalName
                            PrimarySMTPAddress = $ADHashDN["$mailbox"].PrimarySMTPAddress
                            Granted            = $ADHashDN["$($_.distinguishedname)"].DisplayName
                            GrantedUPN         = $ADHashDN["$($_.distinguishedname)"].UserPrincipalName
                            GrantedSMTP        = $ADHashDN["$($_.distinguishedname)"].PrimarySMTPAddress
                            Checking           = $User
                            GroupMember        = $($_.distinguishedname)
                            Type               = "GroupMember"
                            Permission         = "FullAccess"
                        }
                    }
                }
                catch {
                    New-Object -TypeName PSObject -property @{
                        Object             = $ADHashDN["$mailbox"].DisplayName
                        UserPrincipalName  = $ADHashDN["$mailbox"].UserPrincipalName
                        PrimarySMTPAddress = $ADHashDN["$mailbox"].PrimarySMTPAddress
                        Granted            = $ADHash["$User"].DisplayName
                        GrantedUPN         = $ADHash["$User"].UserPrincipalName
                        GrantedSMTP        = $ADHash["$User"].PrimarySMTPAddress
                        Checking           = $User
                        GroupMember        = ""
                        Type               = "User"
                        Permission         = "FullAccess"
                    }
                }
            }
        }
    }
    end {

    }
}
