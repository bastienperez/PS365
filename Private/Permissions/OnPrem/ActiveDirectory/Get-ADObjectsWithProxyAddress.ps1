﻿function GetADObjectsWithProxyAddress {

    param (
        [Parameter()]
        [hashtable] $DomainNameHash
    )

    # Find writable Global Catalog
    $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Forest')
    $dc = ([System.DirectoryServices.ActiveDirectory.GlobalCatalog]::FindOne($context, [System.DirectoryServices.ActiveDirectory.LocatorOptions]'ForceRediscovery, WriteableRequired')).name

    $Properties = @(
        'UserPrincipalName', 'samaccountname', 'displayname', 'canonicalname', 'proxyaddresses'
        'mail', 'msExchRecipientDisplayType', 'msExchRecipientTypeDetails')

    $Selectproperties = @(
        'UserPrincipalName', 'distinguishedname', 'canonicalname', 'displayname', 'mail'
        'Objectguid', 'msExchRecipientDisplayType', 'msExchRecipientTypeDetails'
    )
    $CalculatedProps = @(
        @{n = "logon"; e = { ($DomainNameHash.($_.distinguishedname -replace '^.+?DC=' -replace ',DC=', '.')) + "\" + $_.samaccountname } },
        @{n = "PrimarySMTPAddress" ; e = { ( $_.proxyAddresses | Where-Object { $_ -cmatch "SMTP:" }).Substring(5) } }
    )
    Get-ADObject -filter 'proxyaddresses -ne "$null"' -server ($dc + ":3268") -SearchBase (Get-ADRootDSE).rootdomainnamingcontext -SearchScope Subtree -Properties $Properties |
    Select-Object ($Selectproperties + $CalculatedProps)
}
