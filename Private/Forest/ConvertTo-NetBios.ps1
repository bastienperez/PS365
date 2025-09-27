function ConvertTo-NetBios {

    param(
        $domainName
    )
    
    $root = [adsi] "LDAP://$domainname/RootDSE"
    $configContext = $root.Properties['configurationNamingContext'][0]
    $searchr = [adsi] "LDAP://cn=Partitions,$configContext"
    
    $search = New-Object System.DirectoryServices.DirectorySearcher
    $search.SearchRoot = $searchr
    $search.SearchScope = [System.DirectoryServices.SearchScope] 'OneLevel'
    $search.filter = "(&(objectcategory=Crossref)(dnsRoot=$domainName)(netBIOSName=*))"
    
    $result = $search.Findone() 
    
    if ($result) {
    
        $result.Properties['netbiosname']
    }
    
}