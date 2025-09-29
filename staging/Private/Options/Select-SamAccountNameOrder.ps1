function Select-SamAccountNameOrder {
    param ()
    $RootPath = $env:USERPROFILE + "\ps\"
    $User = $env:USERNAME

    if (-not(Test-Path $RootPath)) {
        try {
            $null = New-Item -ItemType Directory -Path $RootPath -ErrorAction Stop
        }
        catch {
            throw $_.Exception.Message
        }           
    }

    [array]$SamAccountNameOrder = "First Name then Last Name (example: JSmith)", "Last Name then First Name (example: SmithJ)" | 
    Out-GridView -OutputMode Single -Title "The SamAccountName is represented by First Name and Last Name - In which order (Choose 1 and click OK)"

    if ($SamAccountNameOrder -eq "First Name then Last Name (example: JSmith)") {
        "SamFirstFirst" | Out-File ($RootPath + "$($user).SamAccountNameOrder") -Force
    }
    else {
        "SamLastFirst" | Out-File ($RootPath + "$($user).SamAccountNameOrder") -Force
    }

}  