# Description: This script is used to generate the 'Command Reference' section of the PS365 docusaurus site
# * This command needs to be run from the root of the project. e.g. ./build/Build-CommandReference.ps1
# * If running the docusaurus site locally you will need to stop and start Docusaurus to clear the 'Module not found' errors after running this command


$powershellModuleFolder = './powershell'
$powershellModuleName = 'PS365.psm1'
$websiteFolder = './website/docs'
$githubURL = 'https://github.com/bastienperez/ps365'
if (-not (Get-Module Alt3.Docusaurus.Powershell -ListAvailable)) { Install-Module Alt3.Docusaurus.Powershell -Scope CurrentUser -Force -SkipPublisherCheck }
if (-not (Get-Module PlatyPS -ListAvailable)) { Install-Module PlatyPS -Scope CurrentUser -Force -SkipPublisherCheck }
if (-not (Get-Module Pester -ListAvailable)) { Install-Module Pester -Scope CurrentUser -Force -SkipPublisherCheck }

Import-Module Alt3.Docusaurus.Powershell
Import-Module PlatyPS

# Generate the command reference markdown
#$commandsIndexFile = "./website/docs/commands/readme.md"
#$readmeContent = Get-Content $commandsIndexFile  # Backup the readme.md since it will be deleted by New-DocusaurusHelp

# Get all the filenames in the ./powershell/internal folder without the extension
$internalCommands = Get-ChildItem @("./$powershellModuleFolder/Public") -Filter *.ps1 | ForEach-Object { $_.BaseName }
#New-DocusaurusHelp -Module "./$powershellModuleFolder/$powershellModuleName" -DocsFolder "./$websiteFolder" -NoPlaceHolderExamples -EditUrl "$githubURL/blob/main/powershell/public/" -Exclude $internalCommands
New-DocusaurusHelp -Module "./$powershellModuleFolder/$powershellModuleName" -DocsFolder "./$websiteFolder" -NoPlaceHolderExamples -Exclude $internalCommands

# Update the markdown to include the synopsis as description so it can be displayed correctly in the doc links.
$cmdMarkdownFiles = Get-ChildItem ./website/docs/commands
foreach ($file in $cmdMarkdownFiles) {
    # Read file as array of lines for indexing
    $contentLines = Get-Content -Path $file.FullName
    
    # Get the synopsis line
    $synopsisIndex = $contentLines.IndexOf('## SYNOPSIS')
    $synopsis = $null
    if ($synopsisIndex -ge 0 -and ($synopsisIndex + 2) -lt $contentLines.Count) {
        $synopsis = $contentLines[$synopsisIndex + 2]
    }
    
    # Add description to frontmatter if synopsis exists and description doesn't
    if (-not [string]::IsNullOrWhiteSpace($synopsis)) {
        $hasDescription = $contentLines | Where-Object { $_ -match '^description:' }
        if (-not $hasDescription) {
            $idLineIndex = $contentLines.IndexOf(($contentLines | Where-Object { $_ -match '^id:' } | Select-Object -First 1))
            if ($idLineIndex -ge 0) {
                $contentLines = @($contentLines[0..($idLineIndex - 1)]) + @("description: $synopsis") + @($contentLines[$idLineIndex..($contentLines.Count - 1)])
            }
        }
    }
    
    <# Custom for mintlify
    Remove sidebar_class_name:
    Remove hide_title:
    Remove hide_table_of_contents:    
    #>
    
    # Filter out unwanted lines
    $contentLines = $contentLines | Where-Object { 
        $_ -notmatch '^sidebar_class_name:' -and 
        $_ -notmatch '^hide_title:' -and 
        $_ -notmatch '^hide_table_of_contents:' 
    }
    
    # Remove the entire ProgressAction section
    $updatedContent = $contentLines -join "`n"
    $updatedContent = [regex]::Replace($updatedContent, '(?s)### -ProgressAction.*?(?=###|\z)', '')
    
    # Write back to file
    Set-Content -Path $file.FullName -Value $updatedContent -NoNewline

    # Remove the docusaurus.sidebar.js file if it exists
    $sidebarFile = Join-Path -Path $file.DirectoryName -ChildPath 'docusaurus.sidebar.js'
    if (Test-Path -Path $sidebarFile) {
        Remove-Item -Path $sidebarFile -Force
    }
}

#Set-Content $commandsIndexFile $readmeContent  # Restore the readme content