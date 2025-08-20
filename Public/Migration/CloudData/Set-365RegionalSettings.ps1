function Set-365RegionalSettings {
    <#
    .SYNOPSIS
    Imports all mailboxes Regional Configuration from CSV created by Get-365RegionalSettings

    .DESCRIPTION
    Imports all mailboxes Regional Configuration from CSV created by Get-365RegionalSettings

    Results are exported as PSCustomObject which can be exported to CSV etc.

    .PARAMETER CSVFilePath
    Path to the CSVFile previously exported by Get-365RegionalSettings

    .EXAMPLE
    Set-365RegionalSettings -CSVFilePath .\RegionalSettings.csv | Export-Csv .\SetRegionalSettings_Results.csv -Append -notypeinformation

    .NOTES

    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $CSVFilePath,
        [Parameter(Mandatory = $false)]
        [switch]$Allmailboxes,
        [Parameter(Mandatory = $false)]
        [ValidateSet('en-us', 'en-uk', 'fr-fr', 'es-es', 'de-de', 'it-it', 'ja-jp', 'ko-kr', 'pt-br', 'zh-cn', 'zh-tw', 'nl-nl', 'ru-ru', 'pl-pl', 'tr-tr', 'sv-se', 'da-dk', 'fi-fi', 'nb-no', 'el-gr', 'cs-cz', 'hu-hu', 'ro-ro', 'hr-hr', 'sk-sk', 'sl-si', 'bg-bg', 'uk-ua', 'et-ee', 'lv-lv', 'lt-lt', 'ar-sa', 'he-il', 'th-th', 'id-id', 'ms-my', 'vi-vn', 'fil-ph', 'hi-in', 'bn-in', 'gu-in', 'ta-in', 'te-in', 'kn-in', 'mr-in', 'pa-in', 'ur-pk', 'fa-ir')]
        $CountryCode
        # (Get-ChildItem "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Time zones" | foreach {Get-ItemProperty $_.PSPath}).PSChildName
        #[ValidateSet('Afghanistan Standard Time','Alaskan Standard Time','Aleutian Standard Time','Altai Standard Time','Arab Standard Time','Arabian Standard Time','Arabic Standard Time','Argentina Standard Time','Astrakhan Standard Time','Atlantic Standard Time','AUS Central Standard Time','Aus Central W. Standard Time','AUS Eastern Standard Time','Azerbaijan Standard Time','Azores Standard Time','Bahia Standard Time','Bangladesh Standard Time','Belarus Standard Time','Bougainville Standard Time','Canada Central Standard Time','Cape Verde Standard Time','Caucasus Standard Time','Cen. Australia Standard Time','Central America Standard Time','Central Asia Standard Time','Central Brazilian Standard Time','Central Europe Standard Time','Central European Standard Time','Central Pacific Standard Time','Central Standard Time','Central Standard Time (Mexico)','Chatham Islands Standard Time','China Standard Time','Cuba Standard Time','Dateline Standard Time','E. Africa Standard Time','E. Australia Standard Time','E. Europe Standard Time','E. South America Standard Time','Easter Island Standard Time','Eastern Standard Time','Eastern Standard Time (Mexico)','Egypt Standard Time','Ekaterinburg Standard Time','Fiji Standard Time','FLE Standard Time','Georgian Standard Time','GMT Standard Time','Greenland Standard Time','Greenwich Standard Time','GTB Standard Time','Haiti Standard Time','Hawaiian Standard Time','India Standard Time','Iran Standard Time','Israel Standard Time','Jordan Standard Time','Kaliningrad Standard Time','Kamchatka Standard Time','Korea Standard Time','Libya Standard Time','Line Islands Standard Time','Lord Howe Standard Time','Magadan Standard Time','Magallanes Standard Time','Marquesas Standard Time','Mauritius Standard Time','Mid-Atlantic Standard Time','Middle East Standard Time','Montevideo Standard Time','Morocco Standard Time','Mountain Standard Time','Mountain Standard Time (Mexico)','Myanmar Standard Time','N. Central Asia Standard Time','Namibia Standard Time','Nepal Standard Time','New Zealand Standard Time','Newfoundland Standard Time','Norfolk Standard Time','North Asia East Standard Time','North Asia Standard Time','North Korea Standard Time','Omsk Standard Time','Pacific SA Standard Time','Pacific Standard Time','Pacific Standard Time (Mexico)','Pakistan Standard Time','Paraguay Standard Time','Qyzylorda Standard Time','Romance Standard Time','Russia Time Zone 10','Russia Time Zone 11','Russia Time Zone 3','Russian Standard Time','SA Eastern Standard Time','SA Pacific Standard Time','SA Western Standard Time','Saint Pierre Standard Time','Sakhalin Standard Time','Samoa Standard Time','Sao Tome Standard Time','Saratov Standard Time','SE Asia Standard Time','Singapore Standard Time','South Africa Standard Time','South Sudan Standard Time','Sri Lanka Standard Time','Sudan Standard Time','Syria Standard Time','Taipei Standard Time','Tasmania Standard Time','Tocantins Standard Time','Tokyo Standard Time','Tomsk Standard Time','Tonga Standard Time','Transbaikal Standard Time','Turkey Standard Time','Turks And Caicos Standard Time','Ulaanbaatar Standard Time','US Eastern Standard Time','US Mountain Standard Time','UTC','UTC+12','UTC+13','UTC-02','UTC-08','UTC-09','UTC-11','Venezuela Standard Time','Vladivostok Standard Time','Volgograd Standard Time','W. Australia Standard Time','W. Central Africa Standard Time','W. Europe Standard Time','W. Mongolia Standard Time','West Asia Standard Time','West Bank Standard Time','West Pacific Standard Time','Yakutsk Standard Time','Yukon Standard Time')]
        #$TimeZone
    )

    $hashTableCountryCodeTimeZone = @{
        'en-us' = 'Pacific Standard Time'
        'en-uk' = 'GMT Standard Time'
        'fr-fr' = 'Romance Standard Time'
        'es-es' = 'Romance Standard Time'
        'de-de' = 'W. Europe Standard Time'
        'it-it' = 'W. Europe Standard Time'
        'ja-jp' = 'Tokyo Standard Time'
        'ko-kr' = 'Korea Standard Time'
        'pt-br' = 'E. South America Standard Time'
        'zh-cn' = 'China Standard Time'
        'zh-tw' = 'Taipei Standard Time'
        'nl-nl' = 'W. Europe Standard Time'
        'ru-ru' = 'Russian Standard Time'
        'pl-pl' = 'Central European Standard Time'
        'tr-tr' = 'Turkey Standard Time'
        'sv-se' = 'W. Europe Standard Time'
        'da-dk' = 'Romance Standard Time'
        'fi-fi' = 'FLE Standard Time'
        'nb-no' = 'W. Europe Standard Time'
        'el-gr' = 'GTB Standard Time'
        'cs-cz' = 'Central Europe Standard Time'
        'hu-hu' = 'Central Europe Standard Time'
        'ro-ro' = 'GTB Standard Time'
        'hr-hr' = 'Central Europe Standard Time'
        'sk-sk' = 'Central Europe Standard Time'
        'sl-si' = 'Central Europe Standard Time'
        'bg-bg' = 'FLE Standard Time'
        'uk-ua' = 'FLE Standard Time'
        'et-ee' = 'FLE Standard Time'
        'lv-lv' = 'FLE Standard Time'
        'lt-lt' = 'FLE Standard Time'
        'ar-sa' = 'Arabian Standard Time'
    }

    if (Test-Path $CSVFilePath) {
        $mailboxList = Import-Csv $CSVFilePath | Where-Object { $_.Language -and $_.TimeZone }
    }
    else {
        if ($Allmailboxes) {
            if ([string]::IsNullOrWhitespace($CountryCode)) {
                Write-Warning "You need to provide a valid CountryCode"
                return
            }

            $mailboxList = Get-Mailbox -ResultSize Unlimited |  Where-Object RecipientTypeDetails -ne 'DiscoveryMailbox' | Select-Object DisplayName, PrimarySmtpAddress, ExchangeGuid, Language, TimeZone, DateFormat, TimeFormat

            $timeZone = $hashTableCountryCodeTimeZone[$CountryCode]

            if ($null -eq $timeZone) {
                Write-Warning "You need to provide a valid CountryCode"
                return
            }
            else {
                $mailboxList | Add-Member -MemberType NoteProperty -Name TimeZone -Value $timeZone -Force
            }

            $language = $CountryCode

            $mailboxList | Add-Member -MemberType NoteProperty -Name Language -Value $language -Force
        }
        else {
            Write-Warning "You need to provide a valid CSV file path or use -Allmailboxes switch to process all mailboxes"
            return
        }
    }

    foreach ($mailbox in $mailboxList) {
        
        try {
            Set-MailboxRegionalConfiguration -Identity $mailbox.PrimarySmtpAddress -Language $mailbox.Language -TimeZone $mailbox.TimeZone -LocalizeDefaultFolderName $true -ErrorAction Stop
            <#
            $NewConfig = Get-mailboxRegionalConfiguration -Identity $mailbox.ExchangeGuid
            [PSCustomObject][ordered]@{
                DisplayName        = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                ExchangeGuid       = $mailbox.ExchangeGuid
                NewLanguage        = $NewConfig.Language
                NewTimeZone        = $NewConfig.TimeZone
                SourceLanguage     = $mailbox.Language
                SourceTimeZone     = $mailbox.TimeZone
                DateFormat         = $mailbox.DateFormat
                TimeFormat         = $mailbox.TimeFormat
                Log                = 'SUCCESS'
            }
            #>
        }
        catch {
            [PSCustomObject][ordered]@{
                DisplayName        = $mailbox.DisplayName
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                ExchangeGuid       = $mailbox.ExchangeGuid
                NewLanguage        = 'FAILED'
                NewTimeZone        = 'FAILED'
                SourceLanguage     = $mailbox.Language
                SourceTimeZone     = $mailbox.TimeZone
                DateFormat         = $mailbox.DateFormat
                TimeFormat         = $mailbox.TimeFormat
                Log                = $_.Exception.Message
            }
        }
    }
}
