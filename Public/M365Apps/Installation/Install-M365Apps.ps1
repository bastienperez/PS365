<#
	.SYNOPSIS
		Installs Microsoft 365 Apps using the Office Deployment Tool.

	.DESCRIPTION
		Uses the Office Deployment Tool to install Microsoft 365 Apps with a specified configuration file.

	.PARAMETER ODTFolderPath
		The network-accessible location containing the Microsoft 365 Apps source files and the Office Deployment Tool `setup.exe`.
		This can be a UNC path (e.g., \\server\share\Office) or a local path (e.g., C:\OfficeSource).
		Can be a relative or absolute path.

	.PARAMETER ConfigFilePath
		The path to the configuration XML file used for the installation.
		You can generate this file using the Office Customization Tool (OCT) at https://config.office.com/.
		Can be a relative or absolute path.

	.EXAMPLE
		Install-M365Apps -ODTFolderPath '\\server\share\Office' -ConfigFilePath ".\Configuration-XX.xml"

	.EXAMPLE
		Install-M365Apps -ODTFolderPath 'OfficeSource' -ConfigFilePath 'C:\Configs\OfficeConfig.xml'

	.NOTES
		Ensure that the source path is accessible and that the configuration file is correctly formatted.
		The script requires administrative privileges to run.
		The Office Deployment Tool must be downloaded and extracted prior to running this script.
		The ODT folder must contains the ODT setup.exe file and the source files for installation (see `Invoke-M365AppsDownload.ps1`).
#>

function Install-M365Apps {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$ODTFolderPath ,
		[Parameter(Mandatory = $true)]
		[string]$ConfigFilePath
	)

	Write-Host -ForegroundColor cyan 'Start script'

	if (-not (Test-Path $ODTFolderPath)) {
		Write-Warning "The Office Deployment Tool folder $ODTFolderPath does not exist. Please create it and run the script again."
		return 1
	}
	else {
		$odtFolder = (Get-Item $ODTFolderPath).FullName
	}

	if (-not (Test-Path $ConfigFilePath)) {
		Write-Warning "The configuration file $ConfigFilePath does not exist. Please create it and run the script again."
		return 1
	}
	else {
		$configFileFullPath = (Get-Item $ConfigFilePath).FullName
	}
	
	try {
		Write-Host -ForegroundColor cyan 'Installing Microsoft 365 Apps...' -NoNewline

		. $odtFolder\setup.exe /configure $configFileFullPath

		#$process = Start-Process -FilePath "$ODTFolderFullPath\setup.exe" -ArgumentList "/Configure '$ConfigFileFullPath'" -Wait -PassThru -ErrorAction Stop

		if ($LASTEXITCODE -eq 0) {
			Write-Host -ForegroundColor Green 'Office setup started without error.'
		}
		else {
			Write-Warning "Installer failed with exit code $LASTEXITCODE."
		}
	}
	catch {
		Write-Warning $_.Exception.Message
	}
}