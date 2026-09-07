# Description: This script is used to generate the 'Command Reference' section of the PS365 Mintlify site
# * This command needs to be run from the root of the project. e.g. ./build/Build-CommandReference.ps1
#
# Uses Microsoft.PowerShell.PlatyPS directly (no Alt3.Docusaurus.Powershell): that wrapper targets
# Docusaurus, not Mintlify, which is why the previous version had to patch the generated markdown
# afterwards with regex (title/sidebarTitle injection, ProgressAction removal, description injection).
# PlatyPS's own -Metadata and -Locale parameters produce the right frontmatter directly.

# Windows PowerShell 5.1 (Desktop edition) reads a BOM-less .ps1 with the system ANSI codepage,
# not UTF-8, so any non-ASCII character in comment-based help (e.g. "→") comes out double-encoded
# ("â†’") in the generated docs. PowerShell 7 does not have this problem. Re-exec under pwsh
# transparently rather than silently producing corrupted docs when launched from Windows PowerShell.
if ($PSVersionTable.PSEdition -eq 'Desktop') {
    if (-not (Get-Command pwsh -ErrorAction SilentlyContinue)) {
        throw "This script must run under PowerShell 7+ (pwsh) to avoid encoding corruption in the generated docs; pwsh was not found on PATH."
    }
    & pwsh -NoProfile -File $PSCommandPath @args
    exit $LASTEXITCODE
}

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

# Applies $Transform to the markdown outside of fenced ``` code blocks only. Mintlify's MDX
# parser treats "{...}" as a JS expression and bare "<word>" as an HTML/JSX tag; comment-based
# help routinely uses both as plain-text placeholders (e.g. "{{ Fill in the Description }}",
# "<service-principal-id>"), which breaks the page build. Code fences are already safe (MDX does
# not parse their content), so they must be left untouched.
function Repair-MdxProseText {
    param([string]$Content, [scriptblock]$Transform)

    $parts = [regex]::Split($Content, '(?s)(```.*?```)')
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($i % 2 -eq 0) { $parts[$i] = & $Transform $parts[$i] }
    }
    return ($parts -join '')
}

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

    # PlatyPS renders the example command as plain text, not inside a ```powershell fence (unlike
    # SYNTAX/PARAMETERS) - and it can span several lines (if/foreach blocks). Any curly brace in
    # that command (e.g. "Where-Object { $_.X }") is then read as raw prose and breaks Mintlify's
    # MDX parser. Fencing it is also more correct (syntax highlighting) and makes it immune to any
    # further prose-escaping below.
    #
    # The command/description split can't just be "stop at the first blank line": for a
    # multi-line command, PlatyPS itself inserts a spurious blank line after the command's first
    # line (a formatting quirk, not present in the original comment-based help), so the first
    # blank line can land INSIDE the command. The description, on the other hand, is never
    # observed with an internal blank line, so the LAST blank line before the next heading is the
    # actual command/description boundary.
    $fence = '```'
    $content = [regex]::Replace($content, '(?ms)^(### EXAMPLE \d+)\r?\n\r?\n(.*?)(?=\r?\n\r?\n#|\z)', {
            param($match)
            $heading = $match.Groups[1].Value
            $body = $match.Groups[2].Value
            $blankLines = [regex]::Matches($body, '\r?\n\r?\n')
            if ($blankLines.Count -eq 0) {
                $fencedBlock = @(($fence + 'powershell'), $body, $fence) -join "`n"
                return (@($heading, $fencedBlock) -join "`n`n")
            }
            $splitAt = $blankLines[$blankLines.Count - 1]
            $command = $body.Substring(0, $splitAt.Index)
            $description = $body.Substring($splitAt.Index + $splitAt.Length)
            $fencedBlock = @(($fence + 'powershell'), $command, $fence) -join "`n"
            return (@($heading, $fencedBlock, $description) -join "`n`n")
        })

    # INPUTS/OUTPUTS headings can be a raw .NET generic type name (e.g. "List`1[[...]]"): the
    # backtick is .NET's generic-arity notation, but in Markdown/MDX an unmatched backtick opens
    # an inline code span that is never closed, corrupting the rest of the page (Mintlify then
    # fails to render/link it). Escape only backticks inside headings; fenced code blocks
    # elsewhere in the file must keep their real triple backticks untouched.
    $content = [regex]::Replace($content, '(?m)^### .*$', { param($match) $match.Value -replace '`', '\`' })

    # "{{ Fill in the Description }}" (PlatyPS's literal placeholder when .INPUTS/.OUTPUTS has no
    # description) and bare "<placeholder>" tokens in comment-based help text both break
    # Mintlify's MDX parser (acorn expression / unclosed JSX tag). The placeholder carries no real
    # information, so it's replaced with plain text rather than escaped - wrapping it in backticks
    # would "fix" the crash but render it as a misleading inline-code snippet. Real placeholder
    # tokens like "<service-principal-id>" are informative, so those keep their text and are only
    # wrapped in backticks. Skips fenced code blocks, where "{"/"<" are legitimate (yaml
    # hashtables, generic types) and must stay untouched.
    $content = Repair-MdxProseText -Content $content -Transform {
        param($text)
        $text = $text -replace '\{\{\s*Fill[^{}]*\}\}', '_Not documented._'

        # Single alternation pass (double-brace tried first) so a "{{...}}" group is never
        # re-matched a second time as two nested "{...}" groups once the outer pair is handled.
        $text = [regex]::Replace($text, '\{\{[^{}]*\}\}|\{[^{}]*\}', { param($m) '`' + $m.Value + '`' })

        # HTML entities, not backticks: backtick-wrapping a "<placeholder>" glued directly to a
        # URL (e.g. ".../ApplicationSso/`<service-principal-id>`/...") still gets read as an
        # unclosed JSX tag by Mintlify's parser. Entities never trigger tag detection at all, and
        # render back to "<...>" in the final page. Real autolinks ("<https://...>") are excluded.
        $text = $text -replace '<(?!https?://)([A-Za-z][^<>\r\n]*)>', '&lt;$1&gt;'
        return $text
    }

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

# Save the updated docs.json with proper formatting. Windows PowerShell 5.1's "-Encoding UTF8"
# always writes a byte-order mark; Mintlify's JSON parser chokes on it ("Unexpected token '﻿'").
# [System.Text.UTF8Encoding]::new($false) writes UTF-8 without one.
$docsJsonText = $docsJson | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Resolve-Path $docsJsonPath).Path, $docsJsonText, [System.Text.UTF8Encoding]::new($false))