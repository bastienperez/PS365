# Description: This script is used to generate the 'Command Reference' section of the PS365 Mintlify site
# * This command needs to be run from the root of the project. e.g. ./build/Build-CommandReference.ps1
#
# Uses Microsoft.PowerShell.PlatyPS directly (no Alt3.Docusaurus.Powershell): that wrapper targets
# Docusaurus, not Mintlify, which is why the previous version had to patch the generated markdown
# afterwards with regex (title/sidebarTitle injection, ProgressAction removal, description injection).
# PlatyPS's own -Metadata and -Locale parameters produce the right frontmatter directly.

$powershellModuleFolder = './powershell'
$powershellModuleName = 'PS365.psm1'
$commandsFolder = './website/docs/commands'
$siteBaseUrl = 'https://ps365.clidsys.com/docs/commands'

# if (-not (Get-Module Microsoft.PowerShell.PlatyPS -ListAvailable)) { Install-Module Microsoft.PowerShell.PlatyPS -Scope CurrentUser -Force -SkipPublisherCheck }
Import-Module Microsoft.PowerShell.PlatyPS -Force

# Best-effort: most section headers follow this, but "### EXAMPLE n" still comes back as
# "EXEMPLE" on an fr-FR OS regardless (see the post-processing regex below).
[System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'

# Get all the public command names (recursive: the Public folder is organized in nested subfolders)
$publicCommands = Get-ChildItem -Path "$powershellModuleFolder/Public" -Filter *.ps1 -File -Recurse | ForEach-Object { $_.BaseName }

$module = Import-Module "$powershellModuleFolder/$powershellModuleName" -PassThru -Force
$privateCommands = @($module.ExportedCommands.Keys | Where-Object { $_ -notin $publicCommands })

Write-Host "Documenting $($publicCommands.Count) public commands, excluding $($privateCommands.Count) private ones"

if (Test-Path $commandsFolder) { Remove-Item "$commandsFolder/*.mdx" -Force -ErrorAction SilentlyContinue }
else { New-Item -ItemType Directory -Path $commandsFolder -Force | Out-Null }

$tempFolder = Join-Path ([System.IO.Path]::GetTempPath()) "ps365-docs-$([guid]::NewGuid())"
New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null

foreach ($commandName in $publicCommands) {
    $cmd = Get-Command -Module $module.Name -Name $commandName -ErrorAction SilentlyContinue
    if ($null -eq $cmd) { Write-Warning "Command $commandName not found in module, skipped"; continue }

    $synopsis = (Get-Help $cmd.Name -ErrorAction SilentlyContinue).Synopsis
    $synopsis = if ([string]::IsNullOrWhiteSpace($synopsis)) { $null } else { $synopsis.Trim() }

    # "$commandName" (not the bare variable) forces a distinct string instance for each key.
    # PlatyPS's YAML metadata writer emits a full reflection dump instead of a plain scalar when
    # two hashtable values share the exact same string reference.
    $metadata = @{
        title        = "$commandName"
        sidebarTitle = "$commandName"
    }
    if ($null -ne $synopsis) { $metadata.description = $synopsis }

    New-MarkdownCommandHelp -CommandInfo $cmd -OutputFolder $tempFolder -Force -Locale en-US `
        -Metadata $metadata -HelpUri "$siteBaseUrl/$commandName" -ExcludeDontShow | Out-Null

    $generated = Join-Path $tempFolder "$($module.Name)/$commandName.md"
    if (-not (Test-Path $generated)) { Write-Warning "No markdown generated for $commandName"; continue }

    $content = Get-Content $generated -Raw

    # Drop the "external help file:"/"schema:" lines: internal PlatyPS bookkeeping, not useful in the docs
    $content = $content -replace '(?m)^(external help file|PlatyPS schema version):.*\r?\n', ''

    # Drop the ALIASES section when PlatyPS could not find any and left its literal placeholder
    $content = $content -replace '(?s)\r?\n## ALIASES\r?\n\r?\nThis cmdlet has the following aliases,\r?\n\s*\{\{Insert list of aliases\}\}\r?\n', "`n"

    # RELATED LINKS: PlatyPS emits "- [](url)" (empty link text) when comment-based help has no .LINK label
    $content = $content -replace '(?m)^- \[\]\((https?://[^\)]+)\)\r?$', '- [$1]($1)'

    # "EXAMPLE" heading comes back translated as "EXEMPLE" on an fr-FR OS: it's produced by
    # Get-Help's own comment-based-help parser (OS help-engine locale), not by PlatyPS's -Locale.
    $content = $content -replace '(?m)^### EXEMPLE (\d+)\r?$', '### EXAMPLE $1'

    Set-Content -Path (Join-Path $commandsFolder "$commandName.mdx") -Value $content -NoNewline
}

Remove-Module -ModuleInfo $module -Force
Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue

# Update docs.json navigation based on PowerShell module structure
$docsJsonPath = './website/docs.json'
$docsJson = Get-Content $docsJsonPath -Raw | ConvertFrom-Json

# Function to convert folder names to display names dynamically
function ConvertTo-DisplayName {
    param([string]$FolderName)

    # Insert space only between a lowercase/digit and an uppercase letter (real CamelCase boundary)
    $result = $FolderName -creplace '(?<=[a-z0-9])(?=[A-Z])', ' '

    # Insert space between an acronym and a following Word (e.g. "MSCommerce" -> "MS Commerce")
    $result = $result -creplace '(?<=[A-Z])(?=[A-Z][a-z])', ' '

    return $result.Trim()
}

# Recursively build the "pages" array for one Public folder, mixing nested groups (subfolders)
# and flat page entries (.ps1 files directly in that folder) at any depth - the Public tree is
# not always 2 levels deep (e.g. Exchange/Mailbox/Get, Exchange/Mailbox/Set).
function Get-NavigationPages {
    param([string]$FolderPath)

    [System.Collections.Generic.List[object]]$pages = @()

    $subFolders = Get-ChildItem -Path $FolderPath -Directory | Sort-Object Name
    foreach ($subFolder in $subFolders) {
        $subPages = Get-NavigationPages -FolderPath $subFolder.FullName
        if ($subPages.Count -gt 0) {
            $null = $pages.Add([PSCustomObject]@{
                    group = ConvertTo-DisplayName -FolderName $subFolder.Name
                    pages = $subPages.ToArray()
                })
        }
    }

    $ps1Files = Get-ChildItem -Path $FolderPath -Filter '*.ps1' -File | Sort-Object Name
    foreach ($ps1File in $ps1Files) {
        $mdxPath = "docs/commands/$($ps1File.BaseName)"
        if (Test-Path "./website/$mdxPath.mdx") {
            $null = $pages.Add($mdxPath)
        }
    }

    # The unary comma prevents PowerShell from unrolling the List onto the output stream when it
    # holds a single item, which would otherwise turn $subPages into a bare string/object upstream.
    return , $pages
}

# Build navigation groups from PowerShell folder structure
[System.Collections.Generic.List[PSCustomObject]]$newGroups = @()

# Keep the "Getting started" group as is
$gettingStartedGroup = $docsJson.navigation.groups | Where-Object { $_.group -eq 'Getting started' }
if ($null -ne $gettingStartedGroup) {
    $null = $newGroups.Add($gettingStartedGroup)
}

# Get all main folders in Public
$mainFolders = Get-ChildItem -Path "./$powershellModuleFolder/Public" -Directory | Sort-Object Name

foreach ($mainFolder in $mainFolders) {
    $pages = Get-NavigationPages -FolderPath $mainFolder.FullName
    if ($pages.Count -gt 0) {
        $null = $newGroups.Add([PSCustomObject]@{
                group = ConvertTo-DisplayName -FolderName $mainFolder.Name
                pages = $pages.ToArray()
            })
    }
}

<# Add Private Functions section if any private PS1 files exist
$privatePs1Files = Get-ChildItem -Path "./$powershellModuleFolder/Private" -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue | Sort-Object Name
if ($privatePs1Files.Count -gt 0) {
    [System.Collections.Generic.List[string]]$privatePages = @()
    foreach ($ps1File in $privatePs1Files) {
        $mdxPath = "docs/commands/$($ps1File.BaseName)"
        if (Test-Path "./website/$mdxPath.mdx") {
            $null = $privatePages.Add($mdxPath)
        }
    }
    if ($privatePages.Count -gt 0) {
        $privateGroup = [PSCustomObject]@{
            group = 'Private Functions'
            pages = $privatePages.ToArray()
        }
        $null = $newGroups.Add($privateGroup)
    }
}
#>

# Update the navigation groups in docs.json
$docsJson.navigation.groups = $newGroups.ToArray()

# Save the updated docs.json with proper formatting
$docsJson | ConvertTo-Json -Depth 10 | Set-Content $docsJsonPath -Encoding UTF8