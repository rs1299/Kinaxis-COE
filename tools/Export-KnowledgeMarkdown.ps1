#Requires -Version 7.0
<#
.SYNOPSIS
    Collects all Markdown (.md) knowledge files in the Kinaxis-COE repository into a
    single distributable package suitable for uploading to a Microsoft 365 Copilot
    knowledge source (SharePoint library, Copilot Studio file knowledge, etc.).

.DESCRIPTION
    Produces three artifacts under the output folder:
      1. A mirrored copy of every .md file with its repo-relative folder structure
         preserved (good for a SharePoint document library).
      2. A single consolidated .md file with a table of contents and every document
         concatenated (good for a one-file upload when an agent only accepts a few files).
      3. A manifest.csv listing every file, its size, and last-write time.

    Markdown is not always treated as a first-class indexed type by M365 Copilot.
    Use -AlsoConvertToText to additionally emit a .txt copy of each file, which is a
    universally supported knowledge format.

.PARAMETER RepositoryRoot
    Root of the repository to scan. Defaults to the repository containing this script.

.PARAMETER OutputPath
    Destination folder for the package. Defaults to <RepositoryRoot>\build\knowledge-export.

.PARAMETER ExcludePathPattern
    One or more wildcard patterns; any file whose full path matches is skipped.

.PARAMETER AlsoConvertToText
    Also write a .txt copy of each .md file (broadest Copilot compatibility).

.EXAMPLE
    ./tools/Export-KnowledgeMarkdown.ps1

.EXAMPLE
    ./tools/Export-KnowledgeMarkdown.ps1 -AlsoConvertToText -Verbose
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string[]]$ExcludePathPattern = @('*\archive\*', '*\node_modules\*', '*\.venv\*', '*\build\knowledge-export\*'),

    [Parameter()]
    [switch]$AlsoConvertToText
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $OutputPath = Join-Path $RepositoryRoot 'build\knowledge-export'
}

$mirrorRoot = Join-Path $OutputPath 'files'
$consolidatedFile = Join-Path $OutputPath 'Kinaxis-COE-Knowledge-Consolidated.md'
$manifestFile = Join-Path $OutputPath 'manifest.csv'

Write-Verbose "Repository root : $RepositoryRoot"
Write-Verbose "Output path     : $OutputPath"

if (Test-Path $OutputPath) {
    Remove-Item $OutputPath -Recurse -Force
}
New-Item -ItemType Directory -Path $mirrorRoot -Force | Out-Null

# Discover candidate files.
$allMd = Get-ChildItem -Path $RepositoryRoot -Filter '*.md' -Recurse -File

$files = foreach ($f in $allMd) {
    $skip = $false
    foreach ($pattern in $ExcludePathPattern) {
        if ($f.FullName -like $pattern) { $skip = $true; break }
    }
    if (-not $skip) { $f }
}

$files = $files | Sort-Object FullName
Write-Host "Found $($files.Count) Markdown files to export." -ForegroundColor Cyan

$manifest = [System.Collections.Generic.List[object]]::new()
$tocBuilder = [System.Text.StringBuilder]::new()
$bodyBuilder = [System.Text.StringBuilder]::new()

[void]$tocBuilder.AppendLine('# Kinaxis-COE Consolidated Knowledge Base')
[void]$tocBuilder.AppendLine()
[void]$tocBuilder.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
[void]$tocBuilder.AppendLine()
[void]$tocBuilder.AppendLine('## Table of Contents')
[void]$tocBuilder.AppendLine()

$index = 0
foreach ($f in $files) {
    $index++
    $relative = [System.IO.Path]::GetRelativePath($RepositoryRoot, $f.FullName)

    # Mirror the file preserving folder structure.
    $destPath = Join-Path $mirrorRoot $relative
    $destDir = Split-Path -Parent $destPath
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }
    Copy-Item -Path $f.FullName -Destination $destPath -Force

    if ($AlsoConvertToText) {
        $txtPath = [System.IO.Path]::ChangeExtension($destPath, '.txt')
        Copy-Item -Path $f.FullName -Destination $txtPath -Force
    }

    # Consolidated document section.
    $anchor = "doc-$index"
    [void]$tocBuilder.AppendLine("$index. [$relative](#$anchor)")

    [void]$bodyBuilder.AppendLine()
    [void]$bodyBuilder.AppendLine('---')
    [void]$bodyBuilder.AppendLine()
    [void]$bodyBuilder.AppendLine("<a id=""$anchor""></a>")
    [void]$bodyBuilder.AppendLine("## [$index] $relative")
    [void]$bodyBuilder.AppendLine()
    [void]$bodyBuilder.AppendLine((Get-Content -Path $f.FullName -Raw))
    [void]$bodyBuilder.AppendLine()

    $manifest.Add([pscustomobject]@{
        Index        = $index
        RelativePath = $relative
        SizeBytes    = $f.Length
        LastModified = $f.LastWriteTime.ToString('s')
    })
}

# Write consolidated file (TOC + body).
Set-Content -Path $consolidatedFile -Value ($tocBuilder.ToString() + $bodyBuilder.ToString()) -Encoding utf8

# Write manifest.
$manifest | Export-Csv -Path $manifestFile -NoTypeInformation -Encoding utf8

Write-Host ""
Write-Host "Export complete:" -ForegroundColor Green
Write-Host "  Mirrored files : $mirrorRoot"
Write-Host "  Consolidated   : $consolidatedFile"
Write-Host "  Manifest       : $manifestFile"
if ($AlsoConvertToText) {
    Write-Host "  (.txt copies also written next to each .md)"
}
