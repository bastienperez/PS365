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

	Installs Microsoft 365 Apps from the specified network share using the provided configuration file.

	.EXAMPLE
	Install-M365Apps -ODTFolderPath 'C:\Custom-OfficeDeploymentTool' -ConfigFilePath 'C:\Configs\OfficeConfig.xml'

	Installs Microsoft 365 Apps from the specified local folder using the provided configuration file.

	.LINK
	https://ps365.clidsys.com/docs/commands/Install-M365Apps

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
		[string]$ODTFolderPath,
		[Parameter(Mandatory = $true)]
		[string]$ConfigFilePath
	)

	# check run as admin
	if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
		Write-Warning 'You do not have Administrator rights to run this script! Please re-run this script as an Administrator!'
		return 1
	}
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
	
	$setupExePath = "$odtFolder\setup.exe"
	if (-not (Test-Path $setupExePath)) {
		Write-Warning "setup.exe was not found in $odtFolder. Run Invoke-M365AppsDownload first."
		return 1
	}

	# Verify setup.exe is validly signed by Microsoft before executing it.
	$setupSignature = Get-AuthenticodeSignature -FilePath $setupExePath
	if ($setupSignature.Status -ne 'Valid' -or $setupSignature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
		Write-Warning "Authenticode verification failed for setup.exe (Status: $($setupSignature.Status)). Aborting to avoid running an untrusted binary."
		return 1
	}

	try {
		Write-Host -ForegroundColor cyan 'Installing Microsoft 365 Apps...'
		Write-Verbose "Executing $setupExePath /configure $configFileFullPath"
		# Start-Process with an argument array avoids quoting/injection issues.
		$process = Start-Process -FilePath $setupExePath -ArgumentList '/configure', $configFileFullPath -Wait -PassThru -ErrorAction Stop

		if ($process.ExitCode -eq 0) {
			Write-Host -ForegroundColor Green 'Office setup started without error.'
		}
		else {
			Write-Warning "Installer failed with exit code $($process.ExitCode)."
		}
	}
	catch {
		Write-Warning $_.Exception.Message
	}
}