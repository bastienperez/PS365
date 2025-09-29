function Get-PSGUser {
    [CmdletBinding()]
    Param
    (

    )
    $UserList = Get-GSUser -Filter *

    foreach ($User in $UserList) {
        $Forward = Get-GSGmailAutoForwardingSettings -User $User.User
        [PSCustomObject][ordered]@{
            Name                       = $User.Name.FullName
            GivenName                  = $User.Name.GivenName
            FamilyName                 = $User.Name.FamilyName
            PrimaryEmail               = $User.PrimaryEmail
            Emails                     = @($User.Emails.Address) -ne $User.PrimaryEmail -join '|'
            Addresses                  = @(@($User.Addresses).ForEach{
                    if ($_.Type) { '{0}:{1}' -f $_.Type, $_.Formatted }
                    else { $_.Formatted }
                }) -ne '' -join '|'
            EmailType                  = @(@($User.Emails).ForEach{ if ($_.Type) { '{0}:{1}' -f $_.Type, $_.Address }
                    else { $_.Address }
                }) -ne '' -join '|'
            User                       = $User.User
            Department                 = $User.Department
            Title                      = $User.Organizations.Title
            PhonePrimary               = @(@($User.Phones).ForEach{ if ($_.Primary) {
                        if ($_.Type) { '{0}:{1}' -f $_.Type, $_.Value }
                        else { $_.Value }
                    } }) -ne '' -join '|'
            PhoneType                  = @(@($User.Phones).ForEach{ if ($_.Type) { '{0}:{1}' -f $_.Type, $_.Value }
                    else { $_.Value }
                }) -ne '' -join '|'
            OrganizationPrimary        = @(@($User.Organization).ForEach{ if ($_.Primary) {
                        if ($_.CustomType) { '{0}:{1}' -f $_.CustomType, $_.Department }
                        else { $_.Department }
                    }
                }) -ne '' -join '|'
            OrgUnitPath                = $User.OrgUnitPath
            Forward                    = @(@($Forward).ForEach{ if ($_.Enabled) { '{0}:{1}' -f $_.Disposition, $_.EmailAddress }
                    elseif ($_.EmailAddress) { 'DISABLED::{0}:{1}' -f $_.Disposition, $_.EmailAddress }
                }) -ne '' -join '|'
            RecoveryPhone              = $User.RecoveryPhone
            RecoveryEmail              = $User.RecoveryEmail
            Suspended                  = $User.Suspended
            SuspensionReason           = $User.SuspensionReason
            ChangePasswordAtNextLogin  = $User.changePasswordAtNextLogin
            CreationTime               = $User.CreationTime
            IncludeInGlobalAddressList = $User.IncludeInGlobalAddressList
            IpWhitelisted              = $User.IpWhitelisted
            IsAdmin                    = $User.IsAdmin
            IsDelegatedAdmin           = $User.IsDelegatedAdmin
            IsEnforcedIn2Sv            = $User.IsEnforcedIn2Sv
            IsEnrolledIn2Sv            = $User.IsEnrolledIn2Sv
            IsMailboxSetup             = $User.IsMailboxSetup
            NonEditableAliases         = @($User.NonEditableAliases) -ne '' -join '|'
            ThumbnailPhotoEtag         = $User.ThumbnailPhotoEtag
            ThumbnailPhotoUrl          = $User.ThumbnailPhotoUrl
        }
    }
}
