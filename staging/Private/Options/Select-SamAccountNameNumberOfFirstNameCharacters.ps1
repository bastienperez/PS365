function Select-SamAccountNameNumberOfFirstNameCharacters {
    param (
        [Parameter()]
        [int]$SamAccountNameCharacters,
        [Parameter()]
        [int]$SamAccountNameNumberOfLastNameCharacters
    )
    
    $RootPath = $env:USERPROFILE + "\ps\"
    $User = $env:USERNAME

    if (-not(Test-Path $RootPath)) {
        try {
            $null = New-Item -ItemType Directory -Path $RootPath -ErrorAction STOP
        }
        catch {
            throw $_.Exception.Message
        }           
    }
    
    if ($SamAccountNameNumberOfLastNameCharacters) {
        [array]$SamAccountNameNumberOfFirstNameCharacters = 1..($SamAccountNameCharacters - $SamAccountNameNumberOfLastNameCharacters)  | % {$_ -join ","}  | 
            Out-GridView -OutputMode Single -Title "Select the Maximum number of characters from the user's First Name that will make up the SamAccountName (Choose 1 and click OK)" 
        $SamAccountNameNumberOfFirstNameCharacters | Out-File ($RootPath + "$($user).SamAccountNameNumberOfFirstNameCharacters") -Force
    }
    else {
        [array]$SamAccountNameNumberOfFirstNameCharacters = 1..($SamAccountNameCharacters)  | % {$_ -join ","}  | 
            Out-GridView -OutputMode Single -Title "Select the Maximum number of characters from the user's First Name that will make up the SamAccountName (Choose 1 and click OK)"
        $SamAccountNameNumberOfFirstNameCharacters | Out-File ($RootPath + "$($user).SamAccountNameNumberOfFirstNameCharacters") -Force
    }
}  