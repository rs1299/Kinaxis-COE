<#
.SYNOPSIS
    Creates a new solution project folder from the shared team structure.

.DESCRIPTION
    Adds a project under products/ with standard folders and markdown starter files.
    This version is PowerShell 7 safe and avoids backtick-u escape patterns.

.EXAMPLE
    .\New-SolutionProject.ps1 -RepositoryPath "C:\Repos\Kinaxis-COE" -Name "ProductTransitions" -DisplayName "Product Transitions"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RepositoryPath,
    [Parameter(Mandatory = $true)][ValidatePattern('^[A-Za-z0-9_-]+$')][string]$Name,
    [string]$DisplayName = $Name,
    [ValidateSet("solution", "tool", "standard")][string]$ProjectType = "solution",
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-DirectoryIfMissing {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Set-TextFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Content,
        [switch]$Overwrite
    )
    $parent = Split-Path -Parent $Path
    if ($parent) { New-DirectoryIfMissing -Path $parent }
    if ($Overwrite -or -not (Test-Path -LiteralPath $Path)) {
        Set-Content -LiteralPath $Path -Value $Content -Encoding UTF8
    }
}

$repo = Resolve-Path -LiteralPath $RepositoryPath
$projectRoot = Join-Path $repo "products/$Name"

if ((Test-Path -LiteralPath $projectRoot) -and -not $Force) {
    throw "Project already exists: $projectRoot. Use -Force to overwrite starter markdown files."
}

$dirs = @(
    "requirements",
    "architecture",
    "data-model",
    "resources/source",
    "resources/generated",
    "scripts/python",
    "scripts/powershell",
    "ux/wireframes",
    "ux/screenshots",
    "tests/acceptance",
    "tests/regression",
    "tests/performance",
    "tests/evidence",
    "decisions",
    "releases",
    "docs"
)

foreach ($dir in $dirs) {
    New-DirectoryIfMissing -Path (Join-Path $projectRoot $dir)
}

$projectReadme = @"
# $DisplayName

## Purpose

Describe the business purpose and intended outcomes of this solution project.

## Shared standards used

- ../../standards/authoring/
- ../../standards/naming/
- ../../standards/ux/
- ../../standards/testing/
- ../../standards/data-model/
- ../../standards/ai-authoring/

## Project structure

- requirements/ - Business and functional requirements.
- architecture/ - Solution architecture and dependency views.
- data-model/ - Logical and physical data model artifacts.
- resources/ - Maestro / RapidResponse source and generated resources.
- scripts/ - Project-specific automation.
- ux/ - Wireframes, screenshots, and UX notes.
- tests/ - Test plans, regression tests, performance tests, and evidence.
- decisions/ - Project-specific ADRs.
- releases/ - Release notes and deployment evidence.
- docs/ - Supplemental documentation.
"@
Set-TextFile -Path (Join-Path $projectRoot "README.md") -Overwrite:$Force -Content $projectReadme

$requirements = @"
# Requirements - $DisplayName

## Business goals

## In scope

## Out of scope

## Personas

## Functional requirements

## Non-functional requirements

## Assumptions

## Open questions
"@
Set-TextFile -Path (Join-Path $projectRoot "requirements/requirements.md") -Overwrite:$Force -Content $requirements

$architecture = @"
# Architecture - $DisplayName

## Architecture summary

## Shared patterns referenced

- ../../../architecture/patterns/

## Major components

## Data flow

## Resource dependencies

## Security and access considerations

## Performance considerations

## Open architecture decisions
"@
Set-TextFile -Path (Join-Path $projectRoot "architecture/architecture.md") -Overwrite:$Force -Content $architecture

$dataModel = @"
# Logical Data Model - $DisplayName

## Entities

## Relationships

## Key fields

## Semantic definitions

## Data quality rules

## Open modeling decisions
"@
Set-TextFile -Path (Join-Path $projectRoot "data-model/logical-model.md") -Overwrite:$Force -Content $dataModel

$ux = @"
# UX Design - $DisplayName

## Primary workflows

## User-facing worksheets or dashboards

## Actions and confirmations

## Exception handling

## Accessibility and usability notes

## Shared UX standards referenced

- ../../../standards/ux/
"@
Set-TextFile -Path (Join-Path $projectRoot "ux/ux-design.md") -Overwrite:$Force -Content $ux

$testPlan = @"
# Test Plan - $DisplayName

## Scope

## Requirement-to-test traceability

## Functional tests

## Negative and boundary tests

## Data-change tests

## Performance tests

## Security and access tests

## Evidence location

- tests/evidence/
"@
Set-TextFile -Path (Join-Path $projectRoot "tests/test-plan.md") -Overwrite:$Force -Content $testPlan

$adr = @"
# ADR-0001: Initial Project Structure

## Status

Accepted

## Context

The team needs a consistent structure for project-specific solution artifacts while reusing shared standards.

## Decision

Use the standard project structure generated by New-SolutionProject.ps1.

## Consequences

Project-specific artifacts remain under products/$Name, while standards and reusable patterns remain shared.
"@
Set-TextFile -Path (Join-Path $projectRoot "decisions/ADR-0001-initial-project-structure.md") -Overwrite:$Force -Content $adr

$releaseNotes = @"
# Release Notes - $DisplayName

## Unreleased

- Initial project folder created.
"@
Set-TextFile -Path (Join-Path $projectRoot "releases/release-notes.md") -Overwrite:$Force -Content $releaseNotes

Write-Host "Created solution project: $projectRoot" -ForegroundColor Green
