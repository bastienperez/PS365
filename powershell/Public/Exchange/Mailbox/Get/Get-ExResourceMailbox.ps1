<#
    .SYNOPSIS
    Exports Exchange Online resource mailboxes and their calendar processing settings

    .DESCRIPTION
    This function retrieves Exchange Online resource mailboxes (rooms and equipment)
    and their calendar processing parameters. It also includes a
    It also includes a "splatting" feature that deploys distribution groups to display all individual members.

    PARAMETER Filter
    Specifies filters to target specific mailboxes.
    If no filter is provided, all resource mailboxes will be returned.
    Example: '{emailaddresses -like "*contoso.com"}''

    .PARAMETER PrimarySmtpAddress
    Specifies an array of primary SMTP addresses to retrieve specific mailboxes.
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

    .PARAMETER ExportToExcel
    If specified, exports the results to an Excel file in the user's profile directory.

    .EXAMPLE
    Get-ExResourceMailbox -ExportToExcel

    Exports results to an Excel file.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExResourceMailbox
    #>

function Get-ExResourceMailbox {

    [CmdletBinding()]
    param (

        [Parameter(Mandatory = $false, position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String[]] $PrimarySmtpAddress,

        [Parameter()]
        [String[]] $Filter,

        [Parameter()]
        [switch] $UseExchangeDNHash,

        [Parameter()]
        [switch] $ExportToExcel
    )

    begin { $pipelineIdentities = [System.Collections.Generic.List[string]]::new() }
    process {
        foreach ($pipelineIdentity in $PrimarySmtpAddress) { $pipelineIdentities.Add($pipelineIdentity) }
    }
    end {
        # Do not assign an empty array back to the validated parameter when no address was supplied.
        $requestedPrimarySmtpAddresses = $pipelineIdentities.ToArray()
        $Mailboxes = @()
        $ResourceMailboxes = [System.Collections.Generic.List[PSCustomObject]]::new()
        $resourceMailboxesResults = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($UseExchangeDNHash) {
            $MailboxLegacyExchangeDNHash = Get-Mailbox -ResultSize Unlimited | Get-MailboxLegacyExchangeDNHash
        }
        if ($requestedPrimarySmtpAddresses.Count -eq 0) {

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
        else {
            foreach ($smtpAddress in $requestedPrimarySmtpAddresses) {
                $mbx = Get-Mailbox -Identity $smtpAddress -ErrorAction SilentlyContinue
                if ($mbx) {
                    $ResourceMailboxes.Add($mbx)
                }
                else {
                    Write-Warning "Mailbox $smtpAddress not found"
                }
            }
        }
        foreach ($resource in $ResourceMailboxes) {
            Write-Verbose "Processing resource mailbox: $($resource.Identity)"
            $calProcArray = Get-CalendarProcessing -Identity $resource.Guid.ToString()

            foreach ($calProc in $calProcArray) {

                $parameters = @('BookInPolicy', 'RequestInPolicy', 'RequestOutOfPolicy', 'ResourceDelegates')

                foreach ($param in $parameters) {
                    $applyToArray = [System.Collections.Generic.List[string]]::new()
                    $applyToSplatted = [System.Collections.Generic.List[string]]::new()

                    foreach ($object in @($calProc.$param)) {
                        if ($null -eq $object -or "$object" -eq '') { continue }
                        $objectName = "$object"

                        # Resolve once per entry (never reuse the previous iteration value)
                        $recipientObject = $null
                        $resolvedSmtp = $null
                        if ($UseExchangeDNHash -and $MailboxLegacyExchangeDNHash.ContainsKey($objectName)) {
                            $resolvedSmtp = "$($MailboxLegacyExchangeDNHash[$objectName])"
                        }
                        if (-not $resolvedSmtp) {
                            try {
                                # ADObjectId: prefer the GUID/DN (a bare name like "Accueil pyramides" is not always resolvable)
                                $lookupId = if ($object.ObjectGuid -and $object.ObjectGuid -ne [guid]::Empty) { $object.ObjectGuid.ToString() }
                                elseif ($object.DistinguishedName) { $object.DistinguishedName }
                                else { $objectName }
                                $recipientObject = Get-Recipient -Identity $lookupId -ErrorAction SilentlyContinue
                                if (-not $recipientObject) {
                                    $escapedName = $objectName.Replace("'", "''")
                                    $recipientObject = Get-Recipient -Filter "DisplayName -eq '$escapedName' -or Name -eq '$escapedName'" -ResultSize 2 -ErrorAction Stop | Select-Object -First 1
                                }
                                if (-not $recipientObject) { throw "Recipient $objectName not found" }
                                $resolvedSmtp = if ($recipientObject.PrimarySmtpAddress) { "$($recipientObject.PrimarySmtpAddress)" } else { "$($recipientObject.Name)" }
                            }
                            catch {
                                Write-Warning "Resource '$($resource.PrimarySmtpAddress)': recipient '$objectName' not found in property '$param' (deleted or orphaned object). Reported as '$objectName(NotFound)'."
                                $resolvedSmtp = "$objectName(NotFound)"
                            }
                        }
                        $applyToArray.Add($resolvedSmtp)

                        $isGroup = $recipientObject -and ($recipientObject.RecipientTypeDetails -match 'Group|RoomList')
                        if ($isGroup) {
                            $groupIdentity = $recipientObject.Guid.ToString()
                            $memberAddresses = @()
                            try {
                                switch -Regex ($recipientObject.RecipientTypeDetails) {
                                    'GroupMailbox' {
                                        $memberAddresses = @(Get-UnifiedGroupLinks -Identity $groupIdentity -LinkType Members -ResultSize Unlimited -ErrorAction Stop).PrimarySmtpAddress
                                    }
                                    'Dynamic' {
                                        $dynamicGroup = Get-DynamicDistributionGroup -Identity $groupIdentity -ErrorAction Stop
                                        $memberAddresses = @(Get-Recipient -RecipientPreviewFilter $dynamicGroup.RecipientFilter -OrganizationalUnit $dynamicGroup.RecipientContainer -ResultSize Unlimited -ErrorAction Stop).PrimarySmtpAddress
                                    }
                                    default {
                                        $memberAddresses = @(Get-DistributionGroupMember -Identity $groupIdentity -ResultSize Unlimited -ErrorAction Stop).PrimarySmtpAddress
                                    }
                                }
                            }
                            catch {
                                Write-Warning "Resource '$($resource.PrimarySmtpAddress)': cannot expand group '$objectName' (type $($recipientObject.RecipientTypeDetails)) in property '$param'. Error: $($_.Exception.Message)"
                            }
                            foreach ($memberAddress in $memberAddresses) {
                                if ($memberAddress) { $applyToSplatted.Add("$memberAddress") }
                            }
                        }
                        else {
                            $applyToSplatted.Add($resolvedSmtp)
                        }
                    }

                    $calProc | Add-Member -MemberType NoteProperty -Name "$param-applyTo" -Value ($applyToArray -join '|') -Force
                    $calProc | Add-Member -MemberType NoteProperty -Name "$param-applyToArraySplatted" -Value ($applyToSplatted -join '|') -Force
                }
            }

            $resourceMailboxesResults.Add([PSCustomObject][ordered]@{
                    DisplayName                         = $resource.DisplayName
                    Office                              = $resource.Office
                    RecipientTypeDetails                = $resource.RecipientTypeDetails
                    Identity                            = $resource.Identity
                    PrimarySmtpAddress                  = $resource.PrimarySmtpAddress
                    Alias                               = $resource.Alias
                    AutomateProcessing                  = $calProc.AutomateProcessing
                    ResourceDelegates                   = @($calProc.ResourceDelegates) -ne '' -join '|'
                    ResourceDelegatesResolved           = $calProc.'ResourceDelegates-applyTo'
                    ResourceDelegatesSplatted           = $calProc.'ResourceDelegates-applyToArraySplatted'
                    AllBookInPolicy                     = $calProc.AllBookInPolicy
                    AllRequestOutOfPolicy               = $calProc.AllRequestOutOfPolicy
                    AllRequestInPolicy                  = $calProc.AllRequestInPolicy
                    BookInPolicy                        = @($calProc.BookInPolicy) -ne '' -join '| '
                    BookInPolicyResolved                = $calProc.'BookInPolicy-applyTo'
                    BookInPolicySplatted                = $calProc.'BookInPolicy-applyToArraySplatted'
                    RequestInPolicy                     = @($calProc.RequestInPolicy) -ne '' -join '| '
                    RequestInPolicyResolved             = $calProc.'RequestInPolicy-applyTo'
                    RequestInPolicySplatted             = $calProc.'RequestInPolicy-applyToArraySplatted'
                    RequestOutOfPolicy                  = @($calProc.RequestOutOfPolicy) -ne '' -join '| '
                    RequestOutOfPolicyResolved          = $calProc.'RequestOutOfPolicy-applyTo'
                    RequestOutOfPolicySplatted          = $calProc.'RequestOutOfPolicy-applyToArraySplatted'
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
                    MailboxWhenCreated                  = $resource.WhenCreated
                    MailboxWhenModified                 = $resource.WhenChanged
                })
        }
        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $excelFilePath = "$($env:userprofile)\$now-ExResourceMailbox.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
            $resourceMailboxesResults | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExResourceMailbox' -TableStyle Light9
            Write-Host -ForegroundColor Green 'Export completed successfully!'
        }
        else {
            return $resourceMailboxesResults
        }
    }
}
