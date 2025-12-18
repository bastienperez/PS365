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
    $content = Get-Content $file
    $updatedContent = $content
    
    $synopsis = $content[($content.IndexOf('## SYNOPSIS') + 2)] # Get the synopsis
    if (-not [string]::IsNullOrWhiteSpace($synopsis)) {
        # Check if description already exists in the frontmatter
        if ($content -notmatch '^description:') {
            $updatedContent = $updatedContent.Replace('id:', "description: $($synopsis)`nid:")
        }
    }

    <# Custom for mintlify
    Remove sidebar_class_name:
    Remove hide_title:
    Remove hide_table_of_contents:    
    #>

    $updatedContent = $updatedContent -replace "sidebar_class_name:.*(\r?\n)", ''
    $updatedContent = $updatedContent -replace "hide_title:.*(\r?\n)", ''
    $updatedContent = $updatedContent -replace "hide_table_of_contents:.*(\r?\n)", ''
    
    # Remove the entire ProgressAction section
    $updatedContent = $updatedContent -replace "### -ProgressAction.*?(?=^###|\z)", '', 'Singleline'

    Set-Content $file $updatedContent

}

#Set-Content $commandsIndexFile $readmeContent  # Restore the readme content