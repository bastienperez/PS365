<#
    .SYNOPSIS
    Get mailboxes based on a custom attribute and compare specified attributes.

    .DESCRIPTION
    This function retrieves mailboxes based on a custom attribute and compares specified attributes. It returns a list of mailboxes along with their primary SMTP address, the custom attribute value, and whether the specified attributes match.
    It's more or less a `Get-ExoMailbox with specific filter

    .PARAMETER Attribute
    The custom attribute to retrieve from the mailboxes.

    .PARAMETER CheckAttributes
    Optional. An array of two attributes to compare. If specified, the function checks if the values of these attributes match.
    Can be useful to compare two attributes for examples

    .EXAMPLE
    Get-ExMailboxFromAttribute -Attribute "CustomAttribute" -CheckAttributes @("Attribute1", "Attribute2")
    This example retrieves mailboxes with the custom attribute "CustomAttribute" and compares the values of "Attribute1" and "Attribute2".

    .EXAMPLE
    Get-ExMailboxFromAttribute -Attribute "CustomAttribute"
    This example retrieves mailboxes with the custom attribute "CustomAttribute" without comparing any attributes.

    .LINK
    https://ps365.clidsys.com/docs/commands/Get-ExMailboxFromAttribute

    .NOTES
    You can also use `Get-AttributeMatching` to compare two attributes directly in AD/Exchange/Entra ID

#>

function Get-ExMailboxFromAttribute {
    param (
        [Parameter(Mandatory)]
        [string]$Attribute,
        [Parameter(Mandatory = $false)]
        [string[]]$CheckAttributes
    )

    [System.Collections.Generic.List[PSCustomObject]] $mailboxesFound = @()

    # Attribute allows to get PrimarySMTPAddress and this specific attribute (for example forwardingSMTPAddress) and compare
    try {
        $allmbx = Get-EXOMailbox -ResultSize unlimited -Properties $Attribute -ErrorAction Stop -Properties WhenCreated, WhenModified | Where-Object { $null -ne $_.$Attribute }
    }
    catch {
        Write-Warning $_.Exception.Message
        return
    }
    
    $allmbx | ForEach-Object {
        $object = [PSCustomObject][ordered]@{
            Name               = $_.Name
            PrimarySmtpAddress = $_.PrimarySmtpAddress
            $Attribute         = $_.$Attribute
        }

        if ($CheckAttributes -and $CheckAttributes.Count -eq 2) {
            $firstAttribute = $_.($CheckAttributes[0])
            $secondAttribute = $_.($CheckAttributes[1])
            $attributesMatch = $firstAttribute -eq $secondAttribute
        }

        $object | Add-Member -MemberType NoteProperty -Name 'Match' -Value $attributesMatch
        $object | Add-Member -MemberType NoteProperty -Name 'MailboxWhenCreated' -Value $_.WhenCreated
        $object | Add-Member -MemberType NoteProperty -Name 'MailboxWhenModified' -Value $_.WhenChanged

        $mailboxesFound.Add($object)
    }

    return $mailboxesFound
}