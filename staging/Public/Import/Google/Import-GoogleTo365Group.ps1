
function Import-GoogleTo365Group {
    <#
    .SYNOPSIS
    NOT READY!! JUST FRAMEWORK - DONT USE!
    Import CSV of Google Groups into Office 365 as Office 365 Groups (Unified Groups)

    .DESCRIPTION
    Import CSV of Google Groups into Office 365 as Office 365 Groups (Unified Groups)

    .PARAMETER LogPath
    The full path and file name of the log ex. c:\scripts\AddGroupsLog.csv (use csv for best results)

    .PARAMETER Group
    Google Group(s) and respective attributes

    .PARAMETER DontAddOwnersToManagedBy
    Google Group "Owners" will not be added Office 365's "ManagedBy"

    .PARAMETER DontAddManagersToManagedBy
    Google Group "Managers" will not be added Office 365's "ManagedBy"

    .PARAMETER DontCopyManagedByToMember
    Unless you use this switch, all in "ManagedBy" will become members of the Distribution Group

    .PARAMETER INVITED_CAN_JOIN_TranslatesTo
    If Google Group's "whoCanJoin" attribute contains the option INVITED_CAN_JOIN,
    the default behavior sets, "MemberJoinRestriction" to 'ApprovalRequired'.

    Use this parameter to override with either 'Open' or 'Closed'

    .PARAMETER CAN_REQUEST_TO_JOIN_TranslatesTo
    If Google Group's "whoCanJoin" attribute contains the option CAN_REQUEST_TO_JOIN,
    the default behavior sets, "MemberJoinRestriction" to 'ApprovalRequired'.

    Use this parameter to override with either 'Open' or 'Closed'

    .EXAMPLE
    Import-Csv C:\scripts\GoogleGroups.csv | Import-GoogleTo365Group -LogPath C:\scripts\365GroupResults.csv

    .NOTES
    Choosing both -DontAddOwnersToManagedBy & -DontAddManagersToManagedBy results in
    the ManagedBy field being populated with the account that runs this script.

    The same is true if the Google Group has both no managers and no owners

    #>

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory)]
        $LogPath,

        [Parameter(Mandatory, ValueFromPipeline)]
        $Group,

        [Parameter()]
        [switch] $DontAddOwnersToManagedBy,

        [Parameter()]
        [switch] $DontAddManagersToManagedBy,

        [Parameter()]
        [switch] $DontCopyManagedByToMember,

        [Parameter()]
        [ValidateSet('Open', 'Closed')]
        [string] $INVITED_CAN_JOIN_TranslatesTo,

        [Parameter()]
        [ValidateSet('Open', 'Closed')]
        [string] $CAN_REQUEST_TO_JOIN_TranslatesTo
    )
    Begin {

    }
    Process {
        ForEach ($CurGroup in $Group) {

            $Alias = ($CurGroup.Email -split "@")[0]

            # Managers and Owners
            $ManagedBy = [System.Collections.Generic.HashSet[string]]::new()

            if (-not $DontAddManagersToManagedBy -and -not [string]::IsNullOrWhiteSpace($CurGroup.Managers)) {
                $CurGroup.Managers -split "`r`n" | ForEach-Object {
                    $ManagedBy.Add($_) > $null
                }
            }
            if (-not $DontAddOwnersToManagedBy -and -not [string]::IsNullOrWhiteSpace($CurGroup.Owners)) {
                $CurGroup.Owners -split "`r`n" | ForEach-Object {
                    $ManagedBy.Add($_) > $null
                }
            }

            # whoCanJoin
            $MemberJoinRestriction = switch ($CurGroup.whoCanJoin) {

                'ALL_IN_DOMAIN_CAN_JOIN' { 'Open' }
                'CAN_REQUEST_TO_JOIN' {
                    if ($CAN_REQUEST_TO_JOIN_TranslatesTo) {
                        $CAN_REQUEST_TO_JOIN_TranslatesTo
                    }
                    else {
                        'ApprovalRequired'
                    }
                }
                'INVITED_CAN_JOIN' {
                    if ($INVITED_CAN_JOIN_TranslatesTo) {
                        $INVITED_CAN_JOIN_TranslatesTo
                    }
                    else {
                        'ApprovalRequired'
                    }
                }

            }

            # whoCanLeave
            $MemberDepartRestriction = switch ($CurGroup.whoCanLeaveGroup) {

                'ALL_MEMBERS_CAN_LEAVE' { 'Open' }
                Default { 'Open' }

            }

            $NewHash = @{
                Name                    = $CurGroup.Name
                DisplayName             = $CurGroup.Name
                Alias                   = $Alias
                ManagedBy               = $ManagedBy
                PrimarySmtpAddress      = $CurGroup.Email
                MemberJoinRestriction   = $MemberJoinRestriction
                MemberDepartRestriction = $MemberDepartRestriction
                Notes                   = $CurGroup.Description
            }

            # Are Owners and/or Managers copied to the Group's Membership?
            if ($DontCopyManagedByToMember) {
                $NewHash['CopyOwnerToMember'] = $false
            }
            else {
                $NewHash['CopyOwnerToMember'] = $true
            }


            $SetHash = @{
                Identity                      = $CurGroup.Email
                HiddenFromAddressListsEnabled = -not [bool]::Parse($CurGroup.includeInGlobalAddressList)
            }

            # messageModerationLevel (A moderator approves messages sent to recipient before delivered)
            switch ($CurGroup.messageModerationLevel) {

                'MODERATE_NONE' {$SetHash['ModerationEnabled'] = $false}
                'MODERATE_ALL_MESSAGES' {$SetHash['ModerationEnabled'] = $true}
                'MODERATE_NON_MEMBERS' {
                    $SetHash['ModerationEnabled'] = $true
                    $SetHash['BypassModerationFromSendersOrMembers'] = $CurGroup.Email
                }

            }

            # whoCanPostMessage (who can email the DL)
            switch ($CurGroup.whoCanPostMessage) {

                'ANYONE_CAN_POST' { $SetHash['RequireSenderAuthenticationEnabled'] = $false }
                'ALL_IN_DOMAIN_CAN_POST' { $SetHash['RequireSenderAuthenticationEnabled'] = $true }
                'ALL_MANAGERS_CAN_POST' {
                    $SetHash['RequireSenderAuthenticationEnabled'] = $true
                    $SetHash['AcceptMessagesOnlyFromSendersOrMembers'] = $ManagedBy
                }
                'ALL_MEMBERS_CAN_POST' {
                    $SetHash['RequireSenderAuthenticationEnabled'] = $true
                    $SetHash['AcceptMessagesOnlyFromSendersOrMembers'] = $CurGroup.Email
                }

            }

            # Create a splat with only parameters with values for New-DistributionGroup
            $NewSplat = @{}
            ForEach ($Key in $NewHash.keys) {
                if ($($NewHash.item($Key))) {
                    $NewSplat.add($Key, $($NewHash.item($Key)))
                }
            }

            # Create a splat with only parameters with values for Set-DistributionGroup
            $SetSplat = @{}
            ForEach ($Key in $SetHash.keys) {
                if ($($SetHash.item($Key))) {
                    $SetSplat.add($Key, $($SetHash.item($Key)))
                }
            }

            try {
                # $NewDL = New-DistributionGroup @NewSplat -ErrorAction Stop
                [PSCustomObject][ordered]@{
                    Time            = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
                    Result          = 'SUCCESS'
                    Action          = 'CREATING'
                    Object          = 'GROUP'
                    Name            = $CurGroup.Name
                    Email           = $CurGroup.Email
                    Message         = 'SUCCESS'
                    ExtendedMessage = 'SUCCESS'

                } | Export-Csv -Path $LogPath -NoTypeInformation -Append
                Write-HostLog -Message "Creating`t$($NewDL.Name)`t$($NewDL.PrimarySmtpAddress)" -Status Success
                try {
                    # Set-DistributionGroup @SetSplat -ErrorAction Stop -WarningAction SilentlyContinue
                    [PSCustomObject][ordered]@{
                        Time            = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
                        Result          = 'SUCCESS'
                        Action          = 'SETTING'
                        Object          = 'GROUP'
                        Name            = $CurGroup.Name
                        Email           = $CurGroup.Email
                        Message         = 'SUCCESS'
                        ExtendedMessage = 'SUCCESS'

                    } | Export-Csv -Path $LogPath -NoTypeInformation -Append
                    Write-HostLog -Message "Setting`t$($NewDL.Name)`t$($NewDL.PrimarySmtpAddress)" -Status Success
                }
                catch {
                    $Failure = $_.CategoryInfo.Reason
                    [PSCustomObject][ordered]@{
                        Time            = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
                        Result          = 'FAILURE'
                        Action          = 'SETTING'
                        Object          = 'GROUP'
                        Name            = $CurGroup.Name
                        Email           = $CurGroup.Email
                        Message         = $Failure
                        ExtendedMessage = $_.Exception.Message

                    } | Export-Csv -Path $LogPath -NoTypeInformation -Append
                }
            }
            catch {
                $Failure = $_.CategoryInfo.Reason
                if ($_ -match 'The email address') {
                    $Failure = "The email address $($CurGroup.Email) isn't correct"
                }

                if ($_ -match 'is already managed by recipient') {
                    $Failure = 'DL already managed by recipient'
                }
                [PSCustomObject][ordered]@{
                    Time            = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
                    Result          = 'FAILURE'
                    Action          = 'CREATING'
                    Object          = 'GROUP'
                    Name            = $CurGroup.Name
                    Email           = $CurGroup.Email
                    Message         = $Failure
                    ExtendedMessage = $_.Exception.Message

                } | Export-Csv -Path $LogPath -NoTypeInformation -Append
                Write-HostLog -Message "Creating`t$($CurGroup.Name)`t$Failure" -Status Failed
            }
        }
    }
    End {

    }
}
