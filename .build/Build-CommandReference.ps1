# Description: This script is used to generate the 'Command Reference' section of the PS365 docusaurus site
# * This command needs to be run from the root of the project. e.g. ./build/Build-CommandReference.ps1
# * If running the docusaurus site locally you will need to stop and start Docusaurus to clear the 'Module not found' errors after running this command

$powershellModuleFolder = './powershell'
$powershellModuleName = 'PS365.psm1'
$websiteFolder = './website/docs'
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
New-DocusaurusHelp -Module "./$powershellModuleFolder/$powershellModuleName" -DocsFolder "./$websiteFolder" -NoPlaceHolderExamples -Exclude $internalCommands -VendorAgnostic

# Update the markdown to include the synopsis as description so it can be displayed correctly in the doc links.
$cmdMarkdownFiles = Get-ChildItem ./website/docs/commands
foreach ($file in $cmdMarkdownFiles) {
    $content = Get-Content $file
    $synopsis = $content[($content.IndexOf('## SYNOPSIS') + 2)] # Get the synopsis
    if (![string]::IsNullOrWhiteSpace($synopsis)) {
        $content = $content.Replace('id:', "description: $($synopsis)`nid:")
        Set-Content $file $content
    }

    # Remove lines containing "external help file:", "schema:", or "online version:" only within the first 10 lines
    $first10Lines = $content | Select-Object -First 10 | Where-Object { $_ -notmatch '^(external help file:|schema:|online version:)' }
    $remainingLines = $content | Select-Object -Skip 10
    $content = $first10Lines + $remainingLines

    # Remove the -ProgressAction parameter section (common parameter not useful in documentation)
    $contentText = $content -join "`n"
    $contentText = $contentText -replace '(?s)\n*### -ProgressAction.*?Accept wildcard characters: False\s*```\s*\n*', "`n"
    $content = $contentText -split "`n"

    Set-Content $file $content
}

# Update docs.json navigation based on PowerShell module structure
$docsJsonPath = './website/docs.json'
$docsJson = Get-Content $docsJsonPath -Raw | ConvertFrom-Json

# Function to convert folder names to display names dynamically
function ConvertTo-DisplayName {
    param([string]$FolderName)
    
    # Split on capital letters and join with spaces
    $result = $FolderName -creplace '(?<!^)([A-Z])', ' $1'
    
    # Handle numbers followed by letters (like M365)
    $result = $result -replace '(\d+)([A-Z])', '$1 $2'
    
    return $result.Trim()
}

# Build navigation groups from PowerShell folder structure
[System.Collections.Generic.List[PSCustomObject]]$newGroups = @()

# Keep the "Getting started" group as is
$gettingStartedGroup = $docsJson.navigation.groups | Where-Object { $_.group -eq 'Getting started' }
if ($gettingStartedGroup) {
    $null = $newGroups.Add($gettingStartedGroup)
}

# Get all main folders in Public
$mainFolders = Get-ChildItem -Path "./$powershellModuleFolder/Public" -Directory | Sort-Object Name

foreach ($mainFolder in $mainFolders) {
    $groupName = ConvertTo-DisplayName -FolderName $mainFolder.Name

    # Check if folder has subfolders (like Entra, Exchange) or direct .ps1 files
    $subFolders = Get-ChildItem -Path $mainFolder.FullName -Directory
    $ps1Files = Get-ChildItem -Path $mainFolder.FullName -Filter '*.ps1' -File

    if ($subFolders.Count -gt 0) {
        # Has subfolders - create nested groups
        [System.Collections.Generic.List[PSCustomObject]]$subPages = @()

        foreach ($subFolder in ($subFolders | Sort-Object Name)) {
            $subGroupName = ConvertTo-DisplayName -FolderName $subFolder.Name

            # Get all .ps1 files in subfolder and map to mdx paths
            $subPs1Files = Get-ChildItem -Path $subFolder.FullName -Filter '*.ps1' -File | Sort-Object Name
            [System.Collections.Generic.List[string]]$subGroupPages = @()

            foreach ($ps1File in $subPs1Files) {
                $mdxPath = "docs/commands/$($ps1File.BaseName)"
                # Only add if the mdx file exists
                if (Test-Path "./website/$mdxPath.mdx") {
                    $null = $subGroupPages.Add($mdxPath)
                }
            }

            if ($subGroupPages.Count -gt 0) {
                $subGroup = [PSCustomObject]@{
                    group = $subGroupName
                    pages = $subGroupPages.ToArray()
                }
                $null = $subPages.Add($subGroup)
            }
        }

        if ($subPages.Count -gt 0) {
            $group = [PSCustomObject]@{
                group = $groupName
                pages = $subPages.ToArray()
            }
            $null = $newGroups.Add($group)
        }
    }
    elseif ($ps1Files.Count -gt 0) {
        # Direct .ps1 files - create flat group
        [System.Collections.Generic.List[string]]$pages = @()

        foreach ($ps1File in ($ps1Files | Sort-Object Name)) {
            $mdxPath = "docs/commands/$($ps1File.BaseName)"
            # Only add if the mdx file exists
            if (Test-Path "./website/$mdxPath.mdx") {
                $null = $pages.Add($mdxPath)
            }
        }

        if ($pages.Count -gt 0) {
            $group = [PSCustomObject]@{
                group = $groupName
                pages = $pages.ToArray()
            }
            $null = $newGroups.Add($group)
        }
    }
}

# Update the navigation groups in docs.json
$docsJson.navigation.groups = $newGroups.ToArray()

# Save the updated docs.json with proper formatting
$docsJson | ConvertTo-Json -Depth 10 | Set-Content $docsJsonPath -Encoding UTF8