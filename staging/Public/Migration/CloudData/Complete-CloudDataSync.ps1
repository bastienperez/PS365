using namespace System.Management.Automation.Host
function Complete-CloudDataSync {
    [CmdletBinding()]
    param (
        [Parameter()]
        $ResultSize = 'Unlimited',

        [Parameter()]
        [ValidateScript( { Test-Path $_ } )]
        $AlternateCSVFilePath
    )
    $PS365Path = (Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath PS365)
    if (-not ($null = Test-Path $PS365Path)) {
        $null = New-Item $PS365Path -Type Directory -Force -ErrorAction SilentlyContinue
    }
    if ($AlternateCSVFilePath) {
        $ResultFile = $AlternateCSVFilePath
    }
    else {
        $ResultFile = Join-Path -Path $PS365Path -ChildPath 'SyncCloudData_Results.csv'
    }
    $ResultObject = Import-Csv $ResultFile
    $Converted = Select-CompleteCloudDataSync -ResultObject $ResultObject
    $ChoiceList = $Converted | Out-GridView -OutputMode Multiple -Title 'Choose which objects to modify at Target'
    if ($ChoiceList) {
        $ChoiceList | Export-Csv (Join-Path -Path $PS365Path -ChildPath 'ConvertCloudData_Converted.csv') -NoTypeInformation
        $InitialDomain = Select-CloudDataConnection -Type Mailboxes -TenantLocation Target -OnlyEXO
        while (-not $InitialDomain) {
            Write-Host "`r`nPlease connect to Target Tenant now." -ForegroundColor White -BackgroundColor DarkMagenta
            $InitialDomain = Select-CloudDataConnection -Type Mailboxes -TenantLocation Target -OnlyEXO
        }
        $WriteResult = Invoke-CompleteCloudDataSync -ChoiceList $ChoiceList
        $WriteResultFile = Join-Path -Path $PS365Path -ChildPath 'CompleteCloudData_Results.csv'
        $WriteResult | Out-GridView -Title $WriteResultFile
        $WriteResult | Export-Csv $WriteResultFile -NoTypeInformation -Append
    }
}
