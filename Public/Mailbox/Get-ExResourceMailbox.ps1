<#
    .SYNOPSIS
    Exports Exchange Online resource mailboxes and their calendar processing settings

    .DESCRIPTION
    This function retrieves Exchange Online resource mailboxes (rooms and equipment)
    resource mailboxes (rooms and equipment) and their calendar processing parameters. It also includes a
    It also includes a "splatting" feature that deploys distribution groups to display all individual members.

    PARAMETER Filter
    Specifies filters to target specific mailboxes.
    If no filter is provided, all resource mailboxes will be returned.
    Example: '{emailaddresses -like "*contoso.com"}''

    .PARAMETER ResourceMailboxes
    Collection of resource mailboxes to be processed.
    If not specified, the function will automatically retrieve the mailboxes according to the filters.

    .PARAMETER UseExchangeDNHash
    Uses a hash of Exchange Legacy DNs to optimize performance when processing
    delegates and reservation strategies.

    .PARAMETER IgnoreExpandGroups
    Disables the expansion of distribution groups to list individual members
    in ResourceDelegates, BookInPolicy, RequestInPolicy and RequestOutOfPolicy properties.

    .EXAMPLE
    Get-ExResourceMailbox | Export-Csv c:\scripts\All365Mailboxes.csv -NoTypeInformation -Encoding UTF8
    
    Exports all resource mailboxes to a CSV file.

    .EXAMPLE
    Get-ExResourceMailbox -Filter '{emailaddresses -like "*contoso.com"}' -ExpandGroups | Export-Csv c:\scripts\365ResourceMailboxes.csv -NoTypeInformation -Encoding UTF8

    Exports resource mailboxes from the contoso.com domain with group expansion.

    #>
    
function Get-ExResourceMailbox {

    [CmdletBinding()]
    param (
        [Parameter()]
        [string[]] $Filter,

        [Parameter()]
        $ResourceMailboxes,

        [Parameter()]
        [switch] $UseExchangeDNHash
    )

    begin {

        if ($UseExchangeDNHash) {
            $MailboxLegacyExchangeDNHash = Get-Mailbox -ResultSize Unlimited | Get-MailboxLegacyExchangeDNHash
        }
        
        if (-not $ResourceMailboxes) {
            
            [System.Collections.Generic.List[PSCustomObject]]$ResourceMailboxes = @()

            if ($Filter) {
                foreach ($CurFilter in $Filter) {
                    $mbxes = Get-Mailbox -Filter $CurFilter -RecipientTypeDetails RoomMailbox, EquipmentMailbox -ResultSize Unlimited
                }
            }
            else {
                Write-Verbose 'No filter provided, retrieving all resource mailboxes'
                $mbxes = Get-Mailbox -RecipientTypeDetails RoomMailbox, EquipmentMailbox -ResultSize Unlimited
            }

            $mbxes | ForEach-Object {
                $ResourceMailboxes.Add($_)
            }
        }

        foreach ($resource in $ResourceMailboxes) {
            Write-Verbose "Processing resource mailbox: $($resource.Identity)"
            $calProcArray = Get-CalendarProcessing -Identity $resource.Guid.ToString()

            foreach ($calProc in $calProcArray) {

                $parameters = @('BookInPolicy', 'RequestInPolicy', 'RequestOutOfPolicy', 'ResourceDelegates')

                foreach ($param in $parameters) {
                    $applyTo = @()

                    $applyToSplatted = @()

                    if ($calProc.$param) {
                        foreach ($object in $calProc.$param) {
                            if ($UseExchangeDNHash) {
                                $applyTo += $MailboxLegacyExchangeDNHash[$object]
                            }
                            else {
                                $recipient = Get-Recipient $object
                                $applyTo += $recipient.PrimarySmtpAddress
                            }
                            
                            if ($recipient.RecipientType -like '*Group*') {
                                $applyToSplatted += (Get-DistributionGroupMember -Identity $object | Select-Object @{ Name = 'PrimarySmtpAddress'; Expression = { $_.PrimarySmtpAddress.ToString() -join '|' } }).PrimarySmtpAddress
                            }
                            else {
                                if ($UseExchangeDNHash) {
                                    $applyToSplatted += $MailboxLegacyExchangeDNHash[$object]
                                }
                                else {
                                    $applyToSplatted += $recipient.PrimarySmtpAddress
                                }
                            }
                        }
                    }

                    $calProc | Add-Member -MemberType NoteProperty -Name $param-applyTo -Value $($applyTo -join '|')
                    $calProc | Add-Member -MemberType NoteProperty -Name $param-applyToSplatted -Value $($applyToSplatted -join '|')
                }
            }

            [PSCustomObject][ordered]@{
                DisplayName                         = $resource.DisplayName
                Office                              = $resource.Office
                RecipientTypeDetails                = $resource.RecipientTypeDetails
                Identity                            = $resource.Identity
                PrimarySmtpAddress                  = $resource.PrimarySmtpAddress
                Alias                               = $resource.Alias
                AutomateProcessing                  = $calProc.AutomateProcessing
                ResourceDelegates                   = @($calProc.ResourceDelegates) -ne '' -join '|'
                ResourceDelegatesResolved           = $calProc.'ResourceDelegates-applyTo'
                ResourceDelegatesResolvedSplatted   = $calProc.'ResourceDelegates-apply'
                AllBookInPolicy                     = $calProc.AllBookInPolicy
                AllRequestOutOfPolicy               = $calProc.AllRequestOutOfPolicy
                AllRequestInPolicy                  = $calProc.AllRequestInPolicy
                BookInPolicy                        = @($calProc.BookInPolicy) -ne '' -join '| '
                BookInPolicyResolved                = $calProc.'BookInPolicy-applyTo'
                BookInPolicyResolvedSplatted        = $calProc.'BookInPolicy-applyToSplatted'
                RequestInPolicy                     = @($calProc.RequestInPolicy) -ne '' -join '| '
                RequestInPolicyResolved             = $calProc.'RequestInPolicy-applyTo'
                RequestInPolicyResolvedSplatted     = $calProc.'RequestInPolicy-applyToSplatted'
                RequestOutOfPolicy                  = @($calProc.RequestOutOfPolicy) -ne '' -join '| '
                RequestOutOfPolicyResolved          = $calProc.'RequestOutOfPolicy-applyTo'
                RequestOutOfPolicyResolvedSplatted  = $calProc.'RequestOutOfPolicy-applyToSplatted'
                MaximumDurationInMinutes            = $calProc.MaximumDurationInMinutes
                BookingWindowInDays                 = $calProc.BookingWindowInDays
                ConflictPercentageAllowed           = $calProc.ConflictPercentageAllowed
                MaximumConflictInstances            = $calProc.MaximumConflictInstances
                AdditionalResponse                  = $calProc.AdditionalResponse
                AddAdditionalResponse               = $calProc.AddAdditionalResponse
                AddNewRequestsTentatively           = $calProc.AddNewRequestsTentatively
                ForwardRequestsToDelegates          = $calProc.ForwardRequestsToDelegates
                TentativePendingApproval            = $calProc.TentativePendingApproval
                AddOrganizerToSubject               = $calProc.AddOrganizerToSubject
                AllowConflicts                      = $calProc.AllowConflicts
                AllowRecurringMeetings              = $calProc.AllowRecurringMeetings
                DeleteAttachments                   = $calProc.DeleteAttachments
                DeleteComments                      = $calProc.DeleteComments
                DeleteNonCalendarItems              = $calProc.DeleteNonCalendarItems
                DeleteSubject                       = $calProc.DeleteSubject
                EnableResponseDetails               = $calProc.EnableResponseDetails
                EnforceSchedulingHorizon            = $calProc.EnforceSchedulingHorizon
                IsValid                             = $calProc.IsValid
                OrganizerInfo                       = $calProc.OrganizerInfo
                ProcessExternalMeetingMessages      = $calProc.ProcessExternalMeetingMessages
                RemoveForwardedMeetingNotifications = $calProc.RemoveForwardedMeetingNotifications
                RemoveOldMeetingMessages            = $calProc.RemoveOldMeetingMessages
                RemovePrivateProperty               = $calProc.RemovePrivateProperty
                ScheduleOnlyDuringWorkHours         = $calProc.ScheduleOnlyDuringWorkHours
                ObjectState                         = $calProc.ObjectState
                MailboxOwnerId                      = $calProc.MailboxOwnerId
            }
        }
    }
}