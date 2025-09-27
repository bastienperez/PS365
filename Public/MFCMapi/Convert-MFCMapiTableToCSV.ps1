function Convert-MFCMapiTableToCSV {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$InputFile
    )

    $content = Get-Content $InputFile

    # Initialize a list to store all CSV lines
    [System.Collections.Generic.List[PSCustomObject]]$csvList = @()

    # Initialize a new ordered custom object for the CSV line
    $csv = [PSCustomObject][ordered]@{}

    # Iterate over each line in the content
    foreach ($line in $content) {
        # If the line matches 'Row: (a number)'
        if ($line -match 'Row: (\d+)') {
            # If the object is not empty, add it to the list
            if ($csv.PSObject.Properties.Name) {
                $csvList.Add($csv)

                # Initialize a new ordered custom object for the CSV line
                $csv = [PSCustomObject][ordered]@{}
            }
        }
        # If the line matches '(something)::(something)'
        elseif ($line -match '([^:]+)::(.*)') {
            # Extract the header and value and add them to the CSV object
            $header = $matches[1]
            $value = $matches[2]
            $csv | Add-Member -NotePropertyName $header -NotePropertyValue $value
        }
    }

    # Add the last CSV line to the list if it's not empty
    if ($csv.PSObject.Properties.Name) {
        $csvList.Add($csv)
    }

    # output the CSV
    $csvList | Export-Csv -Path ($InputFile -replace '\.txt$', '.csv') -Delimiter ';' -NoTypeInformation
}