<#
.SYNOPSIS
Get AD OBject based on a custom attribute and compare specified attributes.

.DESCRIPTION
This function retrieves AD Objects based on a custom attribute and compares specified attributes.
It returns a list of AD Objects along with their primary SMTP address, the custom attribute value, and whether the specified attributes match.
It's more or less a Get-ADObject with specific filter

.PARAMETER Attribute
The custom attribute to retrieve from the AD Objects.

.PARAMETER CheckAttributes
Optional. An array of two attributes to compare. If specified, the function checks if the values of these attributes match.
Can be useful to compare two attributes for examples

.EXAMPLE
Get-ADObjectFromAttribute -Attribute "CustomAttribute" -CheckAttributes @("Attribute1", "Attribute2")
This example retrieves AD Objects with the custom attribute "CustomAttribute" and compares the values of "Attribute1" and "Attribute2".

.EXAMPLE
Get-ADObjectFromAttribute -Attribute "CustomAttribute"
This example retrieves AD Objects with the custom attribute "CustomAttribute" without comparing any attributes.

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

    [System.Collections.Generic.List[PSObject]] $objectsFound = @()

    if ($AttributeValue) {
        #$filter = "$Attribute -eq '$AttributeValue'"
        $ldapFilter = "($Attribute=$AttributeValue)"
    }
    else {
        #$filter = "*"
        $ldapFilter = '(CN=*)'
    }

    try {
        #$allObjects = Get-ADObject -Filter * -Properties $Attribute | Where-Object { $_.$Attribute -ne $null } -ErrorAction Stop
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