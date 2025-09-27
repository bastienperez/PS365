function Install-ModuleOnServer {
    <#
    .SYNOPSIS
    Installs Module on a server that you have C$ access to
    By default installs PS365 unless specified with -Module parameter

    .DESCRIPTION
    Installs Module on a server that you have C$ access to.
    By default installs PS365 unless specified with -Module parameter
    Helpful when you want a module installed from the PowerShell Gallery
    but PowerShell 5 is not on the server.
    Use this command from any other computer that has PS 5

    .PARAMETER Server
    The server Where-Object you want to module installed from the PowerShell Gallery

    .PARAMETER Module
    Specify the module from the PowerShell Gallery that you want to install
    The default is PS365 and will be installed if this parameter is not used.

    .EXAMPLE
    Install-Module PS365 -Force -Scope CurrentUser
    Install-ModuleOnServer -Server DC01
    * Run from Win10 or 2016. Requires access to remote computer's C$ share

    .NOTES
    The computer from which you install this module must have PowerShell 5.
    The goal is to install the module on that computer and then use this function to
    install it on remote computers (that do not have PS 5)
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string] $Server,

        [Parameter()]
        [string] $Module = 'PS365'
    )
    End {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $Path = "\\{0}\c$\Program Files\WindowsPowerShell\Modules" -f $Server
        Write-Host "`nAttempting to install module here: $Path`n" -ForegroundColor White
        if (Test-Path $Path) {
            $SaveSplat = @{
                Name        = $Module
                Path        = $Path
                Force       = $true
                ErrorAction = 'Stop'
                Confirm     = $false
            }
            try {
                $ModulePath = Join-Path $Path $Module
                Save-Module @SaveSplat
                $Latest = (Get-ChildItem -Directory -Path $ModulePath -Force | Sort-Object LastWriteTime -Descending)[0].fullname
                Get-ChildItem -Path $Latest | Copy-Item -Destination $ModulePath -Recurse -Force -ErrorAction Stop
                Write-Host "Successfully installed $Module module`n" -ForegroundColor Green
                Remove-Item -Path $Latest -Recurse -Force -Confirm:$false

                if ($Module -eq 'PS365') {
                    $SaveExcel = @{
                        Name        = 'ImportExcel'
                        Path        = $Path
                        Force       = $true
                        ErrorAction = 'Stop'
                        Confirm     = $false
                    }
                    $ExcelPath = Join-Path $Path 'ImportExcel'
                    Save-Module @SaveExcel
                    $ExcelLatest = (Get-ChildItem -Directory -Path $ExcelPath -Force | Sort-Object LastWriteTime -Descending)[0].fullname
                    Get-ChildItem -Path $ExcelLatest | Copy-Item -Destination $ExcelPath -Recurse -Force -ErrorAction Stop
                    Write-Host "Successfully installed $Module module dependency, ImportExcel`n" -ForegroundColor Green
                    Remove-Item -Path $ExcelLatest -Recurse -Force -Confirm:$false

                    $SavePoshRS = @{
                        Name        = 'PoshRSJob'
                        Path        = $Path
                        Force       = $true
                        ErrorAction = 'Stop'
                        Confirm     = $false
                    }
                    $PoshRSPath = Join-Path $Path 'PoshRSJob'
                    Save-Module @SavePoshRS
                    $PoshRSLatest = (Get-ChildItem -Directory -Path $PoshRSPath -Force | Sort-Object LastWriteTime -Descending)[0].fullname
                    Get-ChildItem -Path $PoshRSLatest | Copy-Item -Destination $PoshRSPath -Recurse  -Force -ErrorAction Stop
                    Write-Host "Successfully installed $Module module dependency, PoshRSJob`n" -ForegroundColor Green
                    Remove-Item -Path $PoshRSLatest -Recurse -Force -Confirm:$false
                }
                Write-Host "Please restart PowerShell on $Server" -ForegroundColor Cyan
                Write-Host "Type the following command on $Server`:  " -NoNewline -ForegroundColor Cyan
                Write-Host "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" -ForegroundColor DarkYellow
                $Here = @'
$gci = @{
    Path    = "C:\Program Files\WindowsPowerShell\Modules\PS365"
    Filter  = '*.ps1'
    Recurse = $true
}
Get-ChildItem @gci | % {try{. $_.fullname}catch{}}
'@
                Write-Host "If $Server has PowerShell 2 (common with Exchange 2010)," -ForegroundColor Magenta
                Write-Host "copy and paste this code block in PowerShell on $Server" -ForegroundColor Magenta
                Write-Host "`n#####################################################`n" -ForegroundColor White
                Write-Host $Here -ForegroundColor Magenta
                Write-Host "`n#####################################################`n" -ForegroundColor White
            }
            catch {
                Write-Host "Failed to install $Module module." -BackgroundColor Yellow -ForegroundColor Black
                $_.Exception.Message
                Write-Host "Please close PowerShell on $Server Server and try again." -BackgroundColor Yellow -ForegroundColor Black
            }
        }
        else {
            Write-Host "Path $Path not found." -BackgroundColor Yellow -ForegroundColor Black
            Write-Host "Unable to install $Module module." -BackgroundColor Yellow -ForegroundColor Black
        }
    }
}
