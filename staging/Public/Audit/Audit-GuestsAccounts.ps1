function Audit-GuestsAccounts {
  #https://petri.com/knowing-guest-accounts-office-365
  $EndDate = (Get-Date).AddDays(1); $StartDate = (Get-Date).AddDays(-10)
  $Records = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -Operations 'Add Member to Group' -ResultSize 2000 -Formatted)
  if ($Records.Count -eq 0) {
    Write-Host 'No Group Add Member records found.' 
  }
  else {
    Write-Host 'Processing' $Records.Count 'audit records...'
    $Report = [System.Collections.Generic.List[Object]]::new()
    foreach ($Rec in $Records) {
      $AuditData = ConvertFrom-Json $Rec.Auditdata
      # Only process the additions of guest users to groups
      if ($AuditData.ObjectId -like '*#EXT#*') {
        $TimeStamp = Get-Date $Rec.CreationDate -Format g
        # Try and find the timestamp when the Guest account was created in AAD
        try { $AADCheck = (Get-Date(Get-AzureADUser -ObjectId $AuditData.ObjectId).RefreshTokensValidFromDateTime -Format g) }
        catch { Write-Host 'Azure Active Directory record for' $AuditData.ObjectId 'no longer exists' }
        if ($TimeStamp -eq $AADCheck) {
          # It's a new record, so let's write it out
          $NewGuests++
          $ReportLine = [PSCustomObject][ordered]@{
            TimeStamp = $TimeStamp
            User      = $AuditData.UserId
            Action    = $AuditData.Operation
            GroupName = $AuditData.modifiedproperties.newvalue[1]
            Guest     = $AuditData.ObjectId 
          }      
          $Report.Add($ReportLine) 
        }
      }
    }
  }
  Write-Host $NewGuests 'new guest records found...'
  $Report | Sort-Object GroupName, Timestamp | Get-Unique -AsString | Format-Table Timestamp, Groupname, Guest
  return $Report
}