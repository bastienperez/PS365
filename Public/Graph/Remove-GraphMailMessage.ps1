function Remove-GraphMailMessage {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        $MessageList
    )
    process {
        foreach ($Message in $MessageList) {
            $RestSplat = @{
                Uri         = "https://graph.microsoft.com/v1.0/users/{0}/mailFolders/{1}/messages/{2}" -f $Message.UserPrincipalName, $Message.ParentFolderId, $Message.Id
                Headers     = @{ "Authorization" = "Bearer $Token" }
                Method      = 'Delete'
                ErrorAction = 'Stop'
            }
            try { $null = Invoke-RestMethod @RestSplat }
            catch { Write-Host "Error Deleting Message $($Message.UserPrincipalName): - $($_.Exception.Message)" -ForegroundColor Red }
        }
    }
}
