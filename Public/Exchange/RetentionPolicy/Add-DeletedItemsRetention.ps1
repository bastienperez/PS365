function Add-DeletedItemsRetention {
    New-RetentionPolicyTag -AgeLimitForRetention 30 -RetentionAction DeleteAndAllowRecovery -Name 'Deleted Items Remove After 30 days' -RetentionEnabled $True -MessageClass * -Type DeletedItems
    [array]$retentionTags = Get-RetentionPolicy -Identity 'Default MRM Policy' | Select-Object -ExpandProperty RetentionPolicyTagLinks

    $RetentionTags += 'Deleted Items Remove After 30 days'

    Set-RetentionPolicy -Identity 'Default MRM Policy' -RetentionPolicyTagLinks $RetentionTags
}