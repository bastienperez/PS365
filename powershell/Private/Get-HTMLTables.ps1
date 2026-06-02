function Get-HTMLTables {
    param(
        [Parameter(Mandatory)]
        [string]$URL,

        [Parameter(Mandatory = $false)]
        [int]$TableNumber,

        [Parameter(Mandatory = $false)]
        [bool]$LocalFile
    )

    [System.Collections.Generic.List[PSObject]]$tablesArray = @()

    if ($LocalFile) {
        $html = New-Object -ComObject 'HTMLFile'
        $source = Get-Content -Path $URL -Raw
        $html.IHTMLDocument2_write($source)
        $tables = @($html.getElementsByTagName('TABLE'))
    }
    else {
        $webRequest = Invoke-WebRequest $URL -UseBasicParsing
        $htmlContent = $webRequest.Content

        $tableMatches = [regex]::Matches(
            $htmlContent,
            '<table[^>]*>.*?</table>',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
        )

        $tables = @()
        foreach ($tableMatch in $tableMatches) {
            $tableHtml = $tableMatch.Value
            $mockTable = [PSCustomObject]@{
                InnerHtml = $tableHtml
                Rows      = @()
            }

            $rowMatches = [regex]::Matches(
                $tableHtml,
                '<tr[^>]*>.*?</tr>',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
            )

            foreach ($rowMatch in $rowMatches) {
                $rowHtml = $rowMatch.Value
                $mockRow = [PSCustomObject]@{
                    InnerHtml = $rowHtml
                    Cells     = @()
                }

                $cellMatches = [regex]::Matches(
                    $rowHtml,
                    '<(th|td)[^>]*>.*?</(th|td)>',
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Singleline
                )

                foreach ($cellMatch in $cellMatches) {
                    $cellHtml = $cellMatch.Value
                    $isHeader = $cellHtml -match '^<th'
                    $innerText = [regex]::Replace($cellHtml, '<[^>]+>', '') -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace '&quot;', '"'

                    $mockCell = [PSCustomObject]@{
                        tagName   = if ($isHeader) { 'TH' } else { 'TD' }
                        InnerText = $innerText.Trim()
                    }

                    $mockRow.Cells += $mockCell
                }

                $mockTable.Rows += $mockRow
            }

            $tables += $mockTable
        }
    }

    if ($PSBoundParameters.ContainsKey('TableNumber')) {
        $tables = @($tables[$TableNumber])
    }

    $tableNumberIndex = 0
    foreach ($table in $tables) {
        $titles = @()
        $rows = @($table.Rows)
        $tableNumberIndex++

        foreach ($row in $rows) {
            $cells = @($row.Cells)

            if ($cells.Count -eq 0) {
                continue
            }

            if ($cells[0].tagName -eq 'TH') {
                $titles = @($cells | ForEach-Object { ('' + $_.InnerText).Trim() })
                continue
            }

            if (-not $titles) {
                $titles = @(1..($cells.Count + 2) | ForEach-Object { "P$_" })
            }

            $resultObject = [ordered]@{
                TableNumber = $tableNumberIndex
            }

            for ($counter = 0; $counter -lt $cells.Count; $counter++) {
                $title = $titles[$counter]
                if (-not $title) {
                    continue
                }

                $resultObject[$title] = ('' + $cells[$counter].InnerText).Trim()
            }

            $tablesArray.Add([PSCustomObject]$resultObject)
        }
    }

    return $tablesArray
}