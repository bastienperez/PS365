
<#
	.SYNOPSIS
	Get mailbox forwarding configuration

	.DESCRIPTION
	Get mailbox forwarding configuration from ForwardingAddress, ForwardingSMTPAddress and Inbox Rules

	.EXAMPLE
	Get-ExMailboxForwarding

	Returns all mailboxes with forwarding configuration.

	.EXAMPLE
	Get-ExMailboxForwarding -Identity "user1@example.com","user2@example.com"

	Returns forwarding configuration for specified mailboxes.

	.EXAMPLE
	Get-ExMailboxForwarding -ForwardingAndForwardingSMTPOnly

	Returns only mailboxes with ForwardingAddress or ForwardingSMTPAddress set.

	.EXAMPLE
	Get-ExMailboxForwarding -InboxRulesOnly

	Returns only mailboxes with inbox rules that forward emails.

	.PARAMETER ExportToExcel
	If specified, exports the results to an Excel file in the user's profile directory.

	.EXAMPLE
	Get-ExMailboxForwarding -ExportToExcel

	Exports results to an Excel file.

	.LINK
	https://ps365.clidsys.com/docs/commands/Get-ExMailboxForwarding

	.NOTES
	Priority 1: ForwardingAddress
	Priority 2: ForwardingSMTPAddress
	Priority 3: Inbox Rules

	ForwardingWorks uses the outbound spam policy that applies to each mailbox: the first enabled rule that matches, by priority, else the default policy.
#>

function Get-ExMailboxForwarding {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false, position=0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Identity,

        [Parameter(Mandatory = $false)]
        [switch]$ForwardingAndForwardingSMTPOnly,

        [Parameter(Mandatory = $false)]
        [switch]$InboxRulesOnly,

        [Parameter(Mandatory = $false)]
        [switch]$ExportToExcel,

        [Parameter(Mandatory = $false, HelpMessage = 'Optional output directory for the Excel export (defaults to the user profile).')]
        [string]$ExportPath,

        [Parameter(Mandatory = $false)]
        [switch]$ExchangeOnPremise
    )

    begin { $pipelineIdentities = [System.Collections.Generic.List[string]]::new() }
    process {
        foreach ($pipelineIdentity in $Identity) { $pipelineIdentities.Add($pipelineIdentity) }
    }
    end {
        # Do not assign an empty array back to the validated parameter when no identity was supplied.
        $requestedIdentities = $pipelineIdentities.ToArray()
        $Mailboxes = @()
        function ConvertFrom-Recipient {
            param (
                [Parameter(Mandatory = $true)]
                [string]$Recipient
            )

            # "name" [EX:/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=xxxx] if recipient is an object in the organization (mailbox, mail contact, etc.)
            # "name" [SMTP: is not in the same organization

            if ($Recipient -like '*`[SMTP:*@*') {
                # need to escape the [
                $recipientConverted = ($Recipient -split 'SMTP:')[1].TrimEnd(']')
            }
            # we use the LegacyExchangeDN after the EX: to get the recipient domain
            elseif ($Recipient -like '*`[EX:*') {
                # remove the last character (])
                $temp = ($Recipient -split 'EX:')[1].TrimEnd(']')
                $recipientConverted = $hashRecipients[$temp]
            }
            else {
                $recipientConverted = 'unknown format'
            }

            return $recipientConverted
        }
        function Get-OutboundRuleGroupMember {
            # Identifiers of the members of a group named in an outbound spam rule, nested groups included.
            # Returns $null when the group can't be read.
            param (
                [Parameter(Mandatory = $true)]
                [string]$GroupIdentity
            )

            if ($outboundGroupCache.ContainsKey($GroupIdentity)) {
                return , $outboundGroupCache[$GroupIdentity]
            }

            $identifiers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $pendingGroups = [System.Collections.Generic.Queue[string]]::new()
            $visitedGroups = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $pendingGroups.Enqueue($GroupIdentity)
            $failed = $false

            while (($pendingGroups.Count -gt 0) -and (-not $failed)) {
                $currentGroup = $pendingGroups.Dequeue()
                if (-not $visitedGroups.Add($currentGroup)) {
                    continue
                }

                try {
                    $groupRecipient = Get-Recipient -Identity $currentGroup -ErrorAction Stop
                    if ($groupRecipient.RecipientTypeDetails -eq 'GroupMailbox') {
                        $members = @(Get-UnifiedGroupLinks -Identity $currentGroup -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
                    }
                    else {
                        $members = @(Get-DistributionGroupMember -Identity $currentGroup -ResultSize Unlimited -ErrorAction Stop)
                    }
                }
                catch {
                    Write-Warning "Cannot read group '$currentGroup' used by an outbound spam rule: $($_.Exception.Message)"
                    $failed = $true
                    continue
                }

                foreach ($member in $members) {
                    if ([string]$member.RecipientType -like '*Group*') {
                        $pendingGroups.Enqueue([string]$member.DistinguishedName)
                        continue
                    }

                    foreach ($value in @($member.PrimarySmtpAddress, $member.DistinguishedName, $member.Name, $member.Alias, $member.ExternalDirectoryObjectId)) {
                        if ($value) {
                            [void]$identifiers.Add([string]$value)
                        }
                    }
                }
            }

            if ($failed) {
                $identifiers = $null
            }

            $outboundGroupCache[$GroupIdentity] = $identifiers
            return , $identifiers
        }

        function Test-OutboundRuleCondition {
            # Tests one condition or exception of an outbound spam rule.
            # Returns $true, $false, or $null when a group can't be read. Values of one condition use OR logic.
            param (
                [string[]]$Users,
                [string[]]$Groups,
                [string[]]$Domains,
                [string[]]$MailboxIdentifiers,
                [string]$MailboxDomain
            )

            foreach ($identifier in $MailboxIdentifiers) {
                if ($Users -contains $identifier) {
                    return $true
                }
            }

            foreach ($domain in $Domains) {
                # Subdomains are included, per the outbound spam policy documentation
                if (($MailboxDomain -eq $domain) -or ($MailboxDomain -like "*.$domain")) {
                    return $true
                }
            }

            $unknownGroup = $false
            foreach ($group in $Groups) {
                $groupMembers = Get-OutboundRuleGroupMember -GroupIdentity $group
                if ($null -eq $groupMembers) {
                    $unknownGroup = $true
                    continue
                }

                foreach ($identifier in $MailboxIdentifiers) {
                    if ($groupMembers.Contains($identifier)) {
                        return $true
                    }
                }
            }

            if ($unknownGroup) {
                return $null
            }

            return $false
        }

        function Resolve-OutboundSpamPolicy {
            # Returns the outbound spam policy that applies to a mailbox: the first enabled rule that matches,
            # by priority, else the default policy. PolicyName and Mode are $null when it can't be determined.
            param (
                [Parameter(Mandatory = $true)]
                $Mailbox
            )

            $mailboxIdentifiers = @(
                foreach ($value in @($Mailbox.PrimarySmtpAddress, $Mailbox.DistinguishedName, $Mailbox.Name, $Mailbox.Alias, $Mailbox.ExternalDirectoryObjectId)) {
                    if ($value) { [string]$value }
                }
            )
            $mailboxDomain = ([string]$Mailbox.PrimarySmtpAddress).Split('@')[-1]

            foreach ($rule in $enabledOutboundRules) {
                $conditionResults = @()
                # Different types of conditions use AND logic
                $conditionSets = @(
                    @{ Users = @($rule.From); Groups = @(); Domains = @() }
                    @{ Users = @(); Groups = @($rule.FromMemberOf); Domains = @() }
                    @{ Users = @(); Groups = @(); Domains = @($rule.SenderDomainIs) }
                )
                foreach ($conditionSet in $conditionSets) {
                    if (($conditionSet.Users + $conditionSet.Groups + $conditionSet.Domains | Where-Object { $_ }).Count -gt 0) {
                        $conditionResults += , (Test-OutboundRuleCondition @conditionSet -MailboxIdentifiers $mailboxIdentifiers -MailboxDomain $mailboxDomain)
                    }
                }

                if ($conditionResults -contains $false) {
                    continue
                }

                # Any exception that matches excludes the mailbox from this rule
                $exceptionSet = @{
                    Users   = @($rule.ExceptIfFrom | Where-Object { $_ })
                    Groups  = @($rule.ExceptIfFromMemberOf | Where-Object { $_ })
                    Domains = @($rule.ExceptIfSenderDomainIs | Where-Object { $_ })
                }
                $exceptionResult = Test-OutboundRuleCondition @exceptionSet -MailboxIdentifiers $mailboxIdentifiers -MailboxDomain $mailboxDomain
                if ($exceptionResult -eq $true) {
                    continue
                }

                if (($conditionResults -contains $null) -or ($null -eq $exceptionResult)) {
                    return [PSCustomObject]@{ PolicyName = $null; Mode = $null }
                }

                $rulePolicy = $outboundSpamPolicies | Where-Object { $_.Name -eq $rule.HostedOutboundSpamFilterPolicy } | Select-Object -First 1
                if ($rulePolicy) {
                    return [PSCustomObject]@{ PolicyName = $rulePolicy.Name; Mode = $rulePolicy.AutoForwardingMode }
                }

                return [PSCustomObject]@{ PolicyName = $null; Mode = $null }
            }

            if ($defaultOutboundPolicy) {
                return [PSCustomObject]@{ PolicyName = $defaultOutboundPolicy.Name; Mode = $defaultOutboundPolicy.AutoForwardingMode }
            }

            return [PSCustomObject]@{ PolicyName = $null; Mode = $null }
        }

        $mailboxesList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $forwardList = [System.Collections.Generic.List[PSCustomObject]]::new()
        $forwardListMailboxSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $inboxForwardList = [System.Collections.Generic.List[PSCustomObject]]::new()
        Assert-PS365ExchangeOnlineConnection
        Write-Host -ForegroundColor cyan 'Get Accepted Domain in Exchange Online to identify internal/external forward'
        $internalDomains = (Get-AcceptedDomain).DomainName
        $remoteDomains = Get-RemoteDomain
        foreach ($remoteDomain in $remoteDomains) {
            Write-Host "Remote Domain '$remotedomain' AutoForwardEnabled: $($remoteDomain.AutoForwardEnabled)" -ForegroundColor Cyan
        }
        if (-not $ExchangeOnPremise) {
            $outboundSpamPolicies = [array](Get-HostedOutboundSpamFilterPolicy)
            $outboundSpamRules = @(Get-HostedOutboundSpamFilterRule)
            $enabledOutboundRules = @($outboundSpamRules | Where-Object { $_.State -eq 'Enabled' } | Sort-Object -Property Priority)
            $defaultOutboundPolicy = $outboundSpamPolicies | Where-Object { $_.IsDefault -or ($_.Name -eq 'Default') } | Select-Object -First 1
            $outboundGroupCache = @{}

            Write-Host -ForegroundColor Cyan "You have $($outboundSpamPolicies.Count) OutboundSpamPolicies":
            foreach ($outboundSpamPolicy in $outboundSpamPolicies) {

                if ($outboundSpamPolicy.IsDefault -or ($outboundSpamPolicy.Name -eq 'Default')) {
                    # The default policy has no rule and applies to every sender not matched by another policy
                    $state = 'Enabled'
                }
                else {
                    $state = ($outboundSpamRules | Where-Object { $_.HostedOutboundSpamFilterPolicy -eq $outboundSpamPolicy.Name }).State
                }

                if ($state -eq 'Enabled') {
                    $prefix = ''
                    $color = 'Cyan'
                }
                else {
                    $prefix = '[NOT ENABLED] '
                    $color = 'Gray'
                }

                Write-Host "- $prefix`OutboundSpamPolicy '$($outboundSpamPolicy.Name)' - AutoForwardingMode: $($outboundSpamPolicy.AutoForwardingMode)" -ForegroundColor $color

                if ($outboundSpamPolicy.AutoForwardingMode -eq 'Automatic' -and $state -eq 'Enabled') {
                    Write-Host "Careful, the value 'Automatic is now the same as AutoForwardEnable=Off, means autoForward is disabled even if the Remote domain(s) are configured with AutoForwardEnable = `$true
		Sources:
		https://office365itpros.com/2020/11/12/microsoft-clamps-down-mail-forwarding-exchange-online/
		http://blog.icewolf.ch/archive/2020/10/06/how-to-control-the-many-ways-of-email-forwarding-in.aspx
		https://learn.microsoft.com/en-us/microsoft-365/security/office-365-security/external-email-forwarding?view=o365-worldwide
		https://techcommunity.microsoft.com/t5/exchange-team-blog/all-you-need-to-know-about-automatic-email-forwarding-in/ba-p/2074888

		RoadMap ID: MC221113" -ForeGround Yellow
                }
            }
        }
        $hashRecipients = @{}
        Write-Host -ForegroundColor Cyan 'Get all Exchange recipients'
        if (-not $ExchangeOnPremise) {
            $recipients = Get-EXORecipient -ResultSize Unlimited
        }
        else {
            $recipients = Get-Recipient -ResultSize Unlimited
        }
        $recipients | ForEach-Object {
            $hashRecipients.Add($_.Name, $_.PrimarySmtpAddress)
        }
        $properties = @('Identity', 'Name', 'DistinguishedName', 'PrimarySmtpAddress', 'ForwardingAddress', 'ForwardingSmtpAddress', 'DeliverToMailboxAndForward', 'LegacyExchangeDN', 'UserPrincipalName', 'DisplayName', 'Alias', 'ExternalDirectoryObjectId')
        if (-not $ExchangeOnPremise) {

            # Get-EOMailbox doesn't contain all the properties we need by default, so we need to specify them
            Get-EXOMailbox -ResultSize Unlimited -Properties $properties | Select-Object $properties | ForEach-Object {
                $mailboxesList.Add($_)
                $hashRecipients.Add($_.LegacyExchangeDN, $_.PrimarySmtpAddress)
            }

            Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
                $hashRecipients.Add($_.LegacyExchangeDN, $_.PrimarySmtpAddress)
            }

            Get-UnifiedGroup -ResultSize Unlimited | ForEach-Object {
                $hashRecipients.Add($_.LegacyExchangeDN, $_.PrimarySmtpAddress)
            }
        }
        else {
            Get-Mailbox -ResultSize Unlimited | Select-Object $properties | ForEach-Object {
                $mailboxesList.Add($_)
                $hashRecipients.Add($_.LegacyExchangeDN, $_.PrimarySmtpAddress)
            }

            Get-DistributionGroup -ResultSize Unlimited | ForEach-Object {
                $hashRecipients.Add($_.LegacyExchangeDN, $_.PrimarySmtpAddress)
            }
        }
        if ($requestedIdentities.Count -gt 0) {
            # Build a hashtable keyed on PrimarySMTPAddress for O(1) lookups instead of O(N*M)
            # linear scans when the caller passes multiple identities against a large mailbox list.
            $mailboxByPrimarySmtp = @{}
            foreach ($mbxObj in $mailboxesList) {
                $mailboxByPrimarySmtp[[string]$mbxObj.PrimarySMTPAddress] = $mbxObj
            }

            $tempMailboxesList = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($mbx in $requestedIdentities) {
                $mailbox = $mailboxByPrimarySmtp[$mbx]
                if ($mailbox) {
                    $tempMailboxesList.Add($mailbox)
                }
                else {
                    Write-Warning "$mbx mailbox not found."
                }
            }

            $mailboxesList = $tempMailboxesList
        }
        # else get all mailboxes
        else {
            # all mailboxes are in $mailboxesList
            # nothing to do here
        }
        if (-not($InboxRulesOnly)) {
            foreach ($mailbox in $mailboxesList) {
                Write-Host -ForegroundColor cyan "Processing ForwardingAddress|ForwardingSMTPAddress - $($mailbox.Name) - $($mailbox.PrimarySMTPAddress)"

                $autoForwardMode = $null
                $autoForwardLabel = 'Autoforward mode ='
                if ((-not $ExchangeOnPremise) -and (($null -ne $mailbox.ForwardingAddress) -or ($null -ne $mailbox.ForwardingSmtpAddress))) {
                    $outboundPolicy = Resolve-OutboundSpamPolicy -Mailbox $mailbox
                    $autoForwardMode = $outboundPolicy.Mode
                    $autoForwardLabel = "Autoforward mode = $($outboundPolicy.Mode), policy = $($outboundPolicy.PolicyName)"
                }
                #$forward = $mailbox | Where-Object { ($null -ne $_.ForwardingSMTPAddress) -or ($null -ne $_.ForwardingAddress) } | Select-Object Name, PrimarySmtpAddress, ForwardingSMTPAddress, ForwardingAddress, @{Name = 'ForwardingAddressConvertSMTP'; Expression = { if ($null -ne $_.ForwardingAddress) { $hashRecipients[$_.ForwardingAddress] } } }, DeliverToMailboxAndForward

                <# --- Forwarding Address part ---
			ForwardingAddress is a RecipientIdParameter and used when you want to forward emails to a mail-enabled object.
			The target object should exists in your ActiveDirectory | Exchange Online as a mail-enabled object like MailUser, Contact or RemoteMailUser.
			If you do not have a mail-enabled object for your forwarding address then this will not work.
			ForwardingAddress can be set by using the -ForwardingAddress parameter in the command set-mailbox.
			#>

                if ($ExchangeOnPremise) {
                    $forwardingAddress = $mailbox | Where-Object { ($null -ne $_.ForwardingAddress) } | Select-Object Name, PrimarySmtpAddress, ForwardingAddress, @{Name = 'ForwardingAddressConverted'; Expression = { if ($null -ne $_.ForwardingAddress) { $hashRecipients[$_.ForwardingAddress.Name].Address } } }, DeliverToMailboxAndForward
                }
                else {
                    $forwardingAddress = $mailbox | Where-Object { ($null -ne $_.ForwardingAddress) } | Select-Object Name, PrimarySmtpAddress, ForwardingAddress, @{Name = 'ForwardingAddressConverted'; Expression = { if ($null -ne $_.ForwardingAddress) { $hashRecipients[$_.ForwardingAddress] } } }, DeliverToMailboxAndForward
                }

                #$forward = $mailbox | Where-Object { ($null -ne $_.ForwardingSMTPAddress -and -not($internalDomains -contains $_.ForwardingSMTPAddress.split('@')[1])) -or ($null -ne $_.ForwardingAddress -and -not($internalDomains -contains $hashRecipients[$_.ForwardingSMTPAddress].split('@')[1]) ) } | Select-Object Name, PrimarySmtpAddress, ForwardingSMTPAddress, ForwardingAddress, @{Name = 'ForwardingAddressConvertSMTP'; Expression = { if ($null -ne $_.ForwardingAddress) { $hashRecipients[$_.ForwardingAddress] } } }, DeliverToMailboxAndForward

                if ($null -ne $forwardingAddress) {
                    Write-Host -ForegroundColor yellow "$($mailbox.Name) - $($mailbox.PrimarySMTPAddress) - 1 forwardingAddress parameter found"

                    # The target may be missing from the recipients list: don't reuse the domain of the previous mailbox
                    $recipientDomain = $null
                    if ($forwardingAddress.ForwardingAddressConverted) {
                        $recipientDomain = $forwardingAddress.ForwardingAddressConverted.Split('@')[1]
                    }

                    if (-not $recipientDomain) {
                        $forwardingWorks = 'Unknown (forwarding target not found in the recipients list)'
                    }
                    elseif ($internalDomains -contains $recipientDomain) {
                        $forwardingWorks = "True ($recipientDomain = internalDomain)"
                    }
                    elseif ((-not $ExchangeOnPremise) -and ($null -eq $autoForwardMode)) {
                        $forwardingWorks = 'Unknown (cannot determine the outbound spam policy of this mailbox)'
                    }
                    elseif ($autoForwardMode -eq 'Automatic' -or $autoForwardMode -eq 'Off') {
                        $forwardingWorks = "False ($autoForwardLabel)"
                    }
                    else {
                        $forwardingWorks = "Yes ($autoForwardLabel) and address used is an internal object (contact or mailbox)"
                    }

                    $object = [PSCustomObject][ordered]@{
                        Identity                               = $mailbox.Identity
                        Name                                   = $mailbox.Name
                        DisplayName                            = $mailbox.DisplayName
                        PrimarySmtpAddress                     = $mailbox.PrimarySmtpAddress
                        UserPrincipalName                      = $mailbox.UserPrincipalName
                        ForwardingAddressConverted             = $forwardingAddress.ForwardingAddressConverted
                        ForwardType                            = 'ForwardingAddress'
                        ForwardScope                           = ''
                        Precedence                             = '-'
                        ForwardingAddress                      = $forwardingAddress.ForwardingAddress
                        ForwardingSMTPAddress                  = '-'
                        ForwardingWorks                        = $forwardingWorks
                        DeliverToMailboxAndForward             = $forwardingAddress.DeliverToMailboxAndForward
                        InboxRulePriority                      = '-'
                        InboxRuleEnabled                       = '-'
                        InboxRuleForwardAddressConverted       = '-'
                        InboxRuleRedirectTo                    = '-'
                        InboxRuleForwardTo                     = '-'
                        InboxRuleForwardAsAttachmentTo         = '-'
                        InboxRuleSendTextMessageNotificationTo = '-'
                        InboxRuleDescription                   = '-'
                    }

                    #Add object to an array
                    $forwardList.Add($object)
                    [void]$forwardListMailboxSet.Add([string]$mailbox.PrimarySmtpAddress)
                }

                <# --- Forwarding SMTP Address part ---
			On the other hand, ForwardingSMTPAddress, it is a ProxyAddresses Value and has lower priority than ForwardingAddress.
			You can set this attribute with a remote SMTP address even if there is no mail-enabled Object exists in your ActiveDirectory | Exchange Online
			User can set ForwardingSMTPAddress in OWA.
			The ForwardingSMTPAddress has a higher priority than InboxRule :
			'This is expected behavior. Forwarding on a mailbox overrides an inbox redirection rule. To enable the redirection rule, remove forwarding on the mailbox.'
			(https://support.microsoft.com/en-us/help/3069075/inbox-rule-to-redirect-messages-doesn-t-work-if-forwarding-is-set-up-o
			)

			#>
                $forwardingSMTPAddress = $mailbox | Where-Object { ($null -ne $_.ForwardingSMTPAddress) }
                #$forward = $mailbox | Where-Object { ($null -ne $_.ForwardingSMTPAddress -and -not($internalDomains -contains $_.ForwardingSMTPAddress.split('@')[1])) -or ($null -ne $_.ForwardingAddress -and -not($internalDomains -contains $hashRecipients[$_.ForwardingSMTPAddress].split('@')[1]) ) } | Select-Object Name, PrimarySmtpAddress, ForwardingSMTPAddress, ForwardingAddress, @{Name = 'ForwardingAddressConvertSMTP'; Expression = { if ($null -ne $_.ForwardingAddress) { $hashRecipients[$_.ForwardingAddress] } } }, DeliverToMailboxAndForward

                if ($null -ne $forwardingSMTPAddress) {
                    Write-Host -ForegroundColor yellow "$($mailbox.Name) - $($mailbox.PrimarySMTPAddress) - 1 forwardingSMTPAddress parameter found"

                    # O(1) HashSet lookup instead of an O(N) scan through $forwardList.PrimarySMTPAddress
                    if ($forwardListMailboxSet.Contains([string]$mailbox.PrimarySmtpAddress)) {
                        $precedence = 'ForwardingAddress is already set for this mailbox. ForwardingAddress has a higher priority than the ForwardingSMTPAddress. This ForwardingSMTPAddress is ignored'
                    }
                    else {
                        $precedence = '-'
                    }

                    if ($ExchangeOnPremise) {
                        # in exchange on premise, the ForwardingSMTPAddress is a ProxyAddress and is stored in the ProxyAddressesString attribute (the value is smtp:xxx)
                        $recipientDomain = $forwardingSMTPAddress.ForwardingSmtpAddress.ProxyAddressString.Split('@')[1]
                        $forwardingAddressConverted = $forwardingSMTPAddress.ForwardingSmtpAddress.ProxyAddressString.replace('smtp:', '')
                    }
                    else {
                        $recipientDomain = $forwardingSMTPAddress.forwardingSMTPAddress.Split('@')[1]
                        $forwardingAddressConverted = $forwardingSMTPAddress.forwardingSMTPAddress.replace('smtp:', '')
                    }


                    if ($internalDomains -contains $recipientDomain) {
                        $forwardingWorks = "True ($recipientDomain = internalDomain)"
                    }
                    elseif ((-not $ExchangeOnPremise) -and ($null -eq $autoForwardMode)) {
                        $forwardingWorks = 'Unknown (cannot determine the outbound spam policy of this mailbox)'
                    }
                    elseif ($autoForwardMode -eq 'Automatic' -or $autoForwardMode -eq 'Off') {
                        $forwardingWorks = "False ($autoForwardLabel)"
                    }
                    else {
                        $forwardingWorks = "Maybe ($autoForwardLabel), check if RemoteDomain(s) allows external forwarding and check if TransportRule(s) exist to prevent external forwarding"
                    }

                    $object = [PSCustomObject][ordered]@{
                        Identity                               = $mailbox.Identity
                        Name                                   = $mailbox.Name
                        DisplayName                            = $mailbox.DisplayName
                        PrimarySmtpAddress                     = $mailbox.PrimarySmtpAddress
                        UserPrincipalName                      = $mailbox.UserPrincipalName
                        ForwardingAddressConverted             = $forwardingAddressConverted
                        ForwardType                            = 'ForwardingSMTPAddress'
                        ForwardScope                           = ''
                        Precedence                             = $precedence
                        ForwardingAddress                      = '-'
                        ForwardingSMTPAddress                  = $forwardingSMTPAddress.ForwardingSMTPAddress
                        ForwardingWorks                        = $forwardingWorks
                        DeliverToMailboxAndForward             = $forwardingSMTPAddress.DeliverToMailboxAndForward
                        InboxRulePriority                      = '-'
                        InboxRuleEnabled                       = '-'
                        InboxRuleForwardAddressConverted       = '-'
                        InboxRuleRedirectTo                    = '-'
                        InboxRuleForwardTo                     = '-'
                        InboxRuleForwardAsAttachmentTo         = '-'
                        InboxRuleSendTextMessageNotificationTo = '-'
                        InboxRuleDescription                   = '-'
                    }

                    #Add object to an array
                    $forwardList.Add($object)
                    [void]$forwardListMailboxSet.Add([string]$mailbox.PrimarySmtpAddress)
                }
            }
        }
        $i = 0
        if (-not($ForwardingAndForwardingSMTPOnly)) {
            foreach ($mailbox in $mailboxesList) {
                $i++
                Write-Host -ForegroundColor cyan "Processing Inbox rules - $($mailbox.Name) - $($mailbox.PrimarySMTPAddress) [$i/$($mailboxesList.count)]"

                $inboxForwardRules = @(Get-InboxRule -Mailbox "$($mailbox.DistinguishedName)" | Where-Object { ($null -ne $_.ForwardTo) -or ($null -ne $_.ForwardAsAttachmentTo) -or ($null -ne $_.RedirectTo) -or ($_.SendTextMessageNotificationTo.count -gt 0) }) | Select-Object Identity, Enabled, ForwardTo, ForwardAsAttachmentTo, RedirectTo, SendTextMessageNotificationTo, Description, Priority

                if ($inboxForwardRules.count -gt 0) {
                    Write-Host -ForegroundColor yellow "$($mailbox.Name) - $($mailbox.PrimarySMTPAddress) - $(($inboxForwardRules ).count) forward rule(s) found"
                }

                foreach ($inboxForwardRule in $inboxForwardRules) {
                    # Read the mailbox itself: -InboxRulesOnly doesn't fill the forwarding list
                    if (($null -ne $mailbox.ForwardingAddress) -or ($null -ne $mailbox.ForwardingSmtpAddress)) {
                        $precedence = 'ForwardingAddress | ForwardingSMTPAddress is already set for this mailbox. They have a higher priority than inbox rules. This inbox rule will be ignored unless DeliverToMailboxAndForward is set to $true'
                    }
                    else {
                        $precedence = '-'
                    }

                    # ForwardTo, ForwardAsAttachmentTo, RedirectTo are in the following format:
                    # "name" [EX:/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=xxxx] if recipient is an object in the organization (mailbox, mail contact, etc.)
                    # "name" [SMTP: is not in the same organization
                    # ForwardTo, ForwardAsAttachmentTo, RedirectTo can be a list of recipients
                    # SendTextMessageNotificationTo is a list of phone numbers

                    foreach ($forwardTo in $inboxForwardRule.ForwardTo) {

                        if ($ExchangeOnPremise) {
                            $inboxForwardRuleDescription = $inboxForwardRule.description.ToString().replace("`t", '') # delete line breaks and tabs
                        }
                        else {
                            $inboxForwardRuleDescription = $inboxForwardRule.description.replace("`r`n", ' ').replace("`t", '') # delete line breaks and tabs
                        }

                        $recipientConverted = ConvertFrom-Recipient -Recipient $forwardTo

                        $object = [PSCustomObject][ordered]@{
                            Identity                               = $mailbox.Identity
                            Name                                   = $mailbox.Name
                            DisplayName                            = $mailbox.DisplayName
                            PrimarySmtpAddress                     = $mailbox.PrimarySmtpAddress
                            UserPrincipalName                      = $mailbox.UserPrincipalName
                            WhenCreated                            = $mailbox.WhenCreated
                            WhenChanged                            = $mailbox.WhenChanged
                            ForwardingAddressConverted             = $recipientConverted
                            ForwardType                            = 'InboxRule'
                            ForwardScope                           = $forwardingScope
                            Precedence                             = $precedence
                            ForwardingAddress                      = '-'
                            ForwardingSMTPAddress                  = '-'
                            ForwardingWorks                        = 'Not evaluated yet (check precedence and InboxRuleEnabled and forward address)'
                            DeliverToMailboxAndForward             = '-'
                            InboxRulePriority                      = $inboxForwardRule.Priority
                            InboxRuleEnabled                       = $inboxForwardRule.Enabled
                            InboxRuleForwardAddressConverted       = $recipientConverted
                            InboxRuleRedirectTo                    = '-'
                            InboxRuleForwardTo                     = $forwardTo
                            InboxRuleForwardAsAttachmentTo         = '-'
                            InboxRuleSendTextMessageNotificationTo = '-'
                            InboxRuleDescription                   = $inboxForwardRuleDescription
                        }

                        #Add object to an array
                        # $forwardList.Add($object)
                        $inboxForwardList.Add($object)
                    }

                    foreach ($forwardAsAttachmentTo in $inboxForwardRule.ForwardAsAttachmentTo) {
                        if ($ExchangeOnPremise) {
                            $inboxForwardRuleDescription = $inboxForwardRule.description.ToString().("`r`n", ' ').replace("`t", '') # delete line breaks and tabs
                        }
                        else {
                            $inboxForwardRuleDescription = $inboxForwardRule.description.replace("`r`n", ' ').replace("`t", '') # delete line breaks and tabs
                        }

                        $recipientConverted = ConvertFrom-Recipient -Recipient $forwardAsAttachmentTo

                        $object = [PSCustomObject][ordered]@{
                            Identity                               = $mailbox.Identity
                            Name                                   = $mailbox.Name
                            DisplayName                            = $mailbox.DisplayName
                            PrimarySmtpAddress                     = $mailbox.PrimarySmtpAddress
                            UserPrincipalName                      = $mailbox.UserPrincipalName
                            ForwardingAddressConverted             = $recipientConverted
                            ForwardType                            = 'InboxRule'
                            ForwardScope                           = $forwardingScope
                            Precedence                             = $precedence
                            ForwardingAddress                      = '-'
                            ForwardingSMTPAddress                  = '-'
                            ForwardingWorks                        = 'Not evaluated yet (check precedence and InboxRuleEnabled and forward address)'
                            DeliverToMailboxAndForward             = '-'
                            InboxRulePriority                      = $inboxForwardRule.Priority
                            InboxRuleEnabled                       = $inboxForwardRule.Enabled
                            InboxRuleForwardAddressConverted       = $recipientConverted
                            InboxRuleRedirectTo                    = '-'
                            InboxRuleForwardTo                     = '-'
                            InboxRuleForwardAsAttachmentTo         = $forwardAsAttachmentTo
                            InboxRuleSendTextMessageNotificationTo = '-'
                            InboxRuleDescription                   = $inboxForwardRuleDescription
                            MailboxWhenCreated                     = $mailbox.WhenCreated
                            MailboxWhenModified                    = $mailbox.WhenChanged
                        }

                        #Add object to an array
                        #$forwardList.Add($object)
                        $inboxForwardList.Add($object)
                    }

                    foreach ($redirectTo in $inboxForwardRule.RedirectTo) {

                        if ($ExchangeOnPremise) {
                            $inboxForwardRuleDescription = $inboxForwardRule.description.ToString().replace("`r`n", ' ').replace("`t", '') # delete line breaks and tabs
                        }
                        else {
                            $inboxForwardRuleDescription = $inboxForwardRule.description.replace("`r`n", ' ').replace("`t", '') # delete line breaks and tabs
                        }

                        $recipientConverted = ConvertFrom-Recipient -Recipient $redirectTo

                        $object = [PSCustomObject][ordered]@{
                            Identity                               = $mailbox.Identity
                            Name                                   = $mailbox.Name
                            DisplayName                            = $mailbox.DisplayName
                            PrimarySmtpAddress                     = $mailbox.PrimarySmtpAddress
                            UserPrincipalName                      = $mailbox.UserPrincipalName
                            ForwardingAddressConverted             = $recipientConverted
                            ForwardType                            = 'InboxRule'
                            ForwardScope                           = $forwardingScope
                            Precedence                             = $precedence
                            ForwardingAddress                      = '-'
                            ForwardingSMTPAddress                  = '-'
                            ForwardingWorks                        = 'Not evaluated yet (check precedence and InboxRuleEnabled and forward address)'
                            DeliverToMailboxAndForward             = '-'
                            InboxRulePriority                      = $inboxForwardRule.Priority
                            InboxRuleEnabled                       = $inboxForwardRule.Enabled
                            InboxRuleForwardAddressConverted       = $recipientConverted
                            InboxRuleRedirectTo                    = $redirectTo
                            InboxRuleForwardTo                     = '-'
                            InboxRuleForwardAsAttachmentTo         = '-'
                            InboxRuleSendTextMessageNotificationTo = '-'
                            InboxRuleDescription                   = $inboxForwardRuleDescription
                            MailboxWhenCreated                     = $mailbox.WhenCreated
                            MailboxWhenModified                    = $mailbox.WhenChanged
                        }

                        #Add object to an array
                        #$forwardList.Add($object)
                        $inboxForwardList.Add($object)
                    }

                    foreach ($sendTextMessageNotificationTo in $inboxForwardRule.SendTextMessageNotificationTo) {

                        if ($ExchangeOnPremise) {
                            $sendTextMessageNotificationToDescription = $sendTextMessageNotificationTo.Description.ToString().("`r`n", ' ').replace("`t", '') # delete line breaks and tabs
                        }
                        else {
                            $sendTextMessageNotificationToDescription = $sendTextMessageNotificationTo.Description.replace("`r`n", ' ').replace("`t", '') # delete line breaks and tabs
                        }

                        $forwardingScope = 'External'

                        $object = [PSCustomObject][ordered]@{
                            Identity                               = $mailbox.Identity
                            Name                                   = $mailbox.Name
                            DisplayName                            = $mailbox.DisplayName
                            PrimarySmtpAddress                     = $mailbox.PrimarySmtpAddress
                            UserPrincipalName                      = $mailbox.UserPrincipalName
                            ForwardingAddressConverted             = $sendTextMessageNotificationTo
                            ForwardType                            = 'InboxRule'
                            ForwardScope                           = $forwardingScope
                            InboxRulePriority                      = $inboxForwardRule.Priority
                            InboxRuleEnabled                       = $inboxForwardRule.Enabled
                            InboxRuleForwardAddressConverted       = $temp
                            InboxRuleRedirectTo                    = '-'
                            InboxRuleForwardTo                     = '-'
                            InboxRuleForwardAsAttachmentTo         = '-'
                            InboxRuleSendTextMessageNotificationTo = $sendTextMessageNotificationTo
                            InboxRuleDescription                   = $sendTextMessageNotificationToDescription
                            MailboxWhenCreated                     = $mailbox.WhenCreated
                            MailboxWhenModified                    = $mailbox.WhenChanged
                        }

                        #Add object to an array
                        #$forwardList.Add($object)
                        $inboxForwardList.Add($object)
                    }
                }
            }
        }
        $inboxForwardList | ForEach-Object {
            $forwardList.Add($_)
        }
        Write-Host -ForegroundColor cyan "$($forwardList.count) forward(s) found"
        $forwardList | ForEach-Object {
            if ((($_.ForwardingAddressConverted -like '*@*') -and -not($internalDomains -contains $_.ForwardingAddressConverted.split('@')[1])) -or (($_.ForwardingSMTPAddress -like '*@*') -and -not($internalDomains -contains $_.ForwardingSMTPAddress.split('@')[1])) -or (($_.InboxRuleForwardAddressConverted -like '*@*') -and -not($internalDomains -contains $_.InboxRuleForwardAddressConverted.split('@')[1]))) {
                $_.ForwardScope = 'External'
            }
            elseif (($_.ForwardType -eq 'ForwardingAddress') -and (-not $_.ForwardingAddressConverted)) {
                $_.ForwardScope = 'Unknown'
            }
            else {
                $_.ForwardScope = 'Internal'
            }
        }
        if ($ExportToExcel.IsPresent) {
            $now = Get-Date -Format 'yyyy-MM-dd_HHmmss'
            $excelFilePath = "$(if ($ExportPath) { $ExportPath } else { $env:userprofile })\$now-ExMailboxForwarding.xlsx"
            Write-Host -ForegroundColor Cyan "Exporting to Excel file: $excelFilePath"
            $forwardList | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -WorksheetName 'ExMailboxForwarding' -TableStyle Light9
            Write-Host -ForegroundColor Green 'Export completed successfully!'
        }
        else {
            return $forwardList
        }
    }
}
