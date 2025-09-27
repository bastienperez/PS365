<#
    .SYNOPSIS
    Retrieves Active Directory objects based on a custom attribute value and optionally compares two attributes.

    .DESCRIPTION
    This function searches for Active Directory objects that have a specific attribute populated or that match a specific attribute value.
    The function returns detailed information about each object including Name, DistinguishedName, UserPrincipalName, Mail, 
    the specified custom attribute value, and a Match property indicating whether the compared attributes are equal.
    It can optionally compare two attributes on each found object to determine if their values match.
    This is an effective method for identifying inconsistencies in attributes such as email addresses and user principal names, for example.

    .PARAMETER Attribute
    The name of the custom attribute to search for or retrieve from AD objects.

    .PARAMETER AttributeValue
    Optional. The specific value to search for in the custom attribute.
    If not specified, returns all objects with the attribute populated.

    .PARAMETER CheckAttributes
    Optional. An array of exactly two attribute names to compare on each found object. 
    The function will add a 'Match' property indicating whether these two attributes have the same value.

    .EXAMPLE
    Get-ADObjectFromCustomAttribute -Attribute "extensionAttribute1" -CheckAttributes @("mail", "userPrincipalName")
    Retrieves all AD objects that have extensionAttribute1 populated and compares the 'mail' and 'userPrincipalName' attributes.

    .EXAMPLE
    Get-ADObjectFromCustomAttribute -Attribute "department" -AttributeValue "IT"
    Retrieves all AD objects where the department attribute equals "IT".

    .EXAMPLE
    Get-ADObjectFromCustomAttribute -Attribute "extensionAttribute5"
    Retrieves all AD objects that have extensionAttribute5 populated, without performing any attribute comparison.

#>

function Get-ADObjectFromAttribute {
    param (
        [Parameter(Mandatory)]
        [string]$Attribute,
        [Parameter(Mandatory = $false)]
        [string]$AttributeValue,
        [Parameter(Mandatory = $false)]
        [string[]]$CheckAttributes
    )

    [System.Collections.Generic.List[PSCustomObject]] $objectsFound = @()

    if ($AttributeValue) {
        $ldapFilter = "($Attribute=$AttributeValue)"
    }
    else {
        $ldapFilter = '(CN=*)'
    }

    try {
        $allObjects = Get-ADObject -LDAPFilter $ldapFilter -Properties $Attribute -ErrorAction Stop
    }
    catch {
        Write-Warning $_.Exception.Message
        return
    }
    
    $allObjects | ForEach-Object {
        $object = [PSCustomObject][ordered]@{
            Name              = $_.Name
            DistinguishedName = $_.DistinguishedName
            UserPrincipalName = $_.UserPrincipalName
            Mail              = $_.Mail
            $Attribute        = $_.$Attribute
        }

        if ($CheckAttributes -and $CheckAttributes.Count -eq 2) {
            $firstAttribute = $_.($CheckAttributes[0])
            $secondAttribute = $_.($CheckAttributes[1])
            $attributesMatch = $firstAttribute -eq $secondAttribute
        }

        $object | Add-Member -MemberType NoteProperty -Name 'Match' -Value $attributesMatch

        $objectsFound.Add($object)
    }

    return $objectsFound
}