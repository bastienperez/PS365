<#
    .SYNOPSIS
	Downloads the Microsoft 365 Apps binary using the Office Deployment Tool (ODT).

    .DESCRIPTION
	This script downloads the Office Deployment Tool from Microsoft's official site,
	extracts it, and uses a specified configuration XML file to download the Microsoft 365 Apps binary.

    .PARAMETER ConfigFilePath
	The path to the configuration XML file used for the download.
	You can generate this file using the Office Customization Tool (OCT) at https://config.office.com/.
	Can be a relative or absolute path.

    .EXAMPLE
	Invoke-M365AppsDownload -ConfigFilePath .\Configuration-XX.xml

	Runs the script to download the Microsoft 365 Apps binary.

    .NOTES
	Ensure that you have a valid configuration XML file in the ConfigFiles folder.
#>

function Invoke-M365AppsDownload {
	param(
		[Parameter(Mandatory = $true)]
		[string]$ConfigFilePath
	)

	if (-not (Test-Path $ConfigFilePath)) {
		Write-Warning "The configuration file $ConfigFilePath does not exist. Please create it and run the script again."
		return 1
	}
	else {
		#$configFileFullPath = (Get-ChildItem $ConfigFilePath).FullName
		$configFileFullPath = (Get-Item $ConfigFilePath).FullName
	}

	$odtFolder = "$env:SystemDrive\Custom-OfficeDeploymentTool"
    
	if (-not (Test-Path $odtFolder -ErrorAction SilentlyContinue)) {
		Write-Host -ForegroundColor Cyan "Creating the folder $odtFolder"
		try {
			$null = New-Item -ItemType Directory -Path $odtFolder -ErrorAction Stop
		}
		catch {

			Write-Warning "Unable to create the folder $odtFolder (maybe because user has no admin rights). Trying another location..."
			# try to create in the userfolder
			$odtFolder = "$env:USERPROFILE\Custom-OfficeDeploymentTool"
			try {
				$null = New-Item -ItemType Directory -Path $odtFolder -ErrorAction Stop
			}
			catch {
				Write-Warning "Unable to create the folder $odtFolder. Error: $($_.Exception.Message)"
				return
			}
		}
	}

	Write-Host -ForegroundColor Cyan "Downloading the Office Deployment Tool to $odtFolder"

	$url = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'

	$response = Invoke-RestMethod -Uri $url
	# content haslink `https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19029-20278.exe`

	$regex = 'https:\/\/download\.microsoft\.com\/download\/[a-z0-9\-\/]+\/officedeploymenttool_[0-9\-]+\.exe'
	$downloadLink = [regex]::Match($response, $regex).Value
    
	Write-Host -ForegroundColor Cyan "Downloading the Office Deployment Tool from $downloadLink"
	Invoke-WebRequest -Uri $downloadLink -OutFile "$odtFolder\officedeploymenttool.exe" -UseBasicParsing

	Write-Host -ForegroundColor Cyan "Extracting the Office Deployment Tool to $odtFolder"
	. $odtFolder\officedeploymenttool.exe /extract:$odtFolder /quiet

	# pause in case of UAC prompt, otherwise setup.exe may not be available yet
	Start-Sleep -Seconds 30
	# if UAC is enabled, LASTEXITCODE can be other than 0 even if extraction is successful
	# so we need to check if setup.exe is present
	if ($LASTEXITCODE -ne 0 -and -not (Test-Path "$odtFolder\setup.exe")) {
		Write-Warning "Extraction failed with exit code $LASTEXITCODE."
		return
	}

	# Wait for the extraction to complete
	# si setup.exe is not yet available, wait a bit
	while (-not (Test-Path "$odtFolder\setup.exe")) {
		Write-Host -ForegroundColor Yellow 'Waiting for the Office Deployment Tool extraction to complete...'
		Start-Sleep -Seconds 5
	}

	try {
		Write-Host -ForegroundColor Cyan "Downloading Microsoft 365 Apps binary using configuration file $ConfigFilePath to $odtFolder"
		Set-Location $odtFolder
		. .\setup.exe /download $configFileFullPath

		if ($LASTEXITCODE -eq 0) {
			Write-Host -ForegroundColor Green 'Download completed successfully.'
		}
		else {
			Write-Warning "Download failed with exit code $LASTEXITCODE."
		}
	}
	catch {
		Write-Warning $_.Exception.Message
	}
}