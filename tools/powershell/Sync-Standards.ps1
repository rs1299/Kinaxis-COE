<#
.SYNOPSIS
    Creates or refreshes project-level standard reference files.

.DESCRIPTION
    This script does not duplicate the full standards into every project. Instead, it places
    lightweight STANDARD-REFERENCES.md files in each project so authors know which shared
    standards are authoritative. PowerShell 7 safe.

.EXAMPLE
    .\Sync-Standards.ps1 -RepositoryPath "C:\Repos\Kinaxis-COE"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RepositoryPath,
    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = Resolve-Path -LiteralPath $RepositoryPath
$productsPath = Join-Path $repo "products"

if (-not (Test-Path -LiteralPath $productsPath)) {
    throw "Products folder not found: $productsPath"
}

$standardReference = @'
# Standard References

This project uses the shared team standards from the repository root.

## Authoritative standards

- ../../standards/authoring/
- ../../standards/naming/
- ../../standards/ux/
- ../../standards/testing/
- ../../standards/data-model/
- ../../standards/performance/
- ../../standards/ai-authoring/

## Architecture references

- ../../architecture/patterns/
- ../../architecture/reference-architectures/
- ../../architecture/decision-records/

## Accelerator references

- ../../accelerators/templates/
- ../../accelerators/prompts/
- ../../accelerators/starter-kits/

## Rule

Do not fork or rewrite shared standards inside this project. If the project needs a standards change, submit a change request under governance/change-requests/ and update the shared standard after approval.
'@

$projects = Get-ChildItem -LiteralPath $productsPath -Directory
foreach ($project in $projects) {
    $target = Join-Path $project.FullName "STANDARD-REFERENCES.md"
    if ($Overwrite -or -not (Test-Path -LiteralPath $target)) {
        Set-Content -LiteralPath $target -Value $standardReference -Encoding UTF8
        Write-Host "Updated $target" -ForegroundColor Green
    }
}

Write-Host "Standards references synchronized." -ForegroundColor Green
