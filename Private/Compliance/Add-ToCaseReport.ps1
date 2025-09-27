
function Add-ToCaseReport {
    param(
        [string]$casename,
        
        [String]$casestatus,
        
        [datetime]$casecreatedtime,
        
        [string]$casemembers,
        
        [datetime]$caseClosedDateTime,
        
        [string]$caseclosedby,
        
        [string]$holdname,
        
        [String]$Holdenabled,
        
        [string]$holdcreatedby,
        
        [string]$holdlastmodifiedby,
        
        [string]$ExchangeLocation,
        
        [string]$sharePointlocation,
        
        [string]$ContentMatchQuery,
        
        [datetime]$holdcreatedtime,
        
        [datetime]$holdchangedtime,

        [string]$outputpath
    )
    $object = [PSCustomObject][ordered]@{
        'Case name'               = $casename
        'Case status'             = $casestatus
        'Hold name'               = $holdname
        'Hold enabled'            = $Holdenabled

        'Case members'            = $casemembers
        
        'Case created time'       = $casecreatedtime
        'Case closed time'        = $caseClosedDateTime
        'Case closed by'          = $caseclosedby
        'Exchange locations'      = $ExchangeLocation
        'SharePoint locations'    = $sharePointlocation
        'SharePoint locations'    = $sharePointlocation
        'Hold query'              = $ContentMatchQuery
        'Hold created by'         = $holdcreatedby
        'Hold created time (UTC)' = $holdcreatedtime

        'Hold last changed by'    = $holdlastmodifiedby
        'Hold changed time (UTC)' = $holdchangedtime
    }
    
    $object | Export-Csv -Path $outputPath -NoTypeInformation -Append -Encoding ascii 
}
