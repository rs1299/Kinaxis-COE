#Requires -Version 7.0
<#
.SYNOPSIS
    Publishes the Kinaxis-COE Markdown knowledge base into a Microsoft 365 SharePoint
    document library so it can be used as a knowledge source by an M365 Copilot
    declarative agent (e.g. the Supply Chain Architect agent).

.DESCRIPTION
    This is the "Bitbucket -> SharePoint" half of the recommended architecture:

        Bitbucket (system of record)
              |
              v
        Export-KnowledgeMarkdown.ps1   (mirrors *.md preserving folder structure)
              |
              v
        Publish-KnowledgeToSharePoint.ps1  (this script - uploads to SharePoint)
              |
              v
        M365 Copilot Supply Chain Architect agent (SharePoint = knowledge source)

    The script:
      1. Optionally runs Export-KnowledgeMarkdown.ps1 to produce a fresh mirror.
      2. Connects to a SharePoint site using PnP.PowerShell.
      3. Recreates the repo folder hierarchy under a target library folder.
      4. Uploads only files whose content hash differs from what is already in
         SharePoint (incremental sync), and removes orphaned files when -Mirror is set.

    Authentication uses an Entra ID app registration (client id + certificate or
    client secret) so it can run unattended in Azure Automation, a Bitbucket
    pipeline, or a scheduled task. Interactive login is also supported for testing.

.PARAMETER SiteUrl
    Full URL of the target SharePoint site, e.g. https://contoso.sharepoint.com/sites/SupplyChainArchitect

.PARAMETER LibraryName
    Server-relative document library (title). Defaults to 'Documents' (the default
    "Shared Documents" library).

.PARAMETER TargetFolder
    Folder path inside the library that becomes the root of the mirror,
    e.g. 'SupplyChainArchitect'. Created if it does not exist.

.PARAMETER SourcePath
    Folder containing the files to publish. Defaults to the 'files' mirror produced
    by Export-KnowledgeMarkdown.ps1 (<RepositoryRoot>\build\knowledge-export\files).

.PARAMETER RepositoryRoot
    Root of the repository. Defaults to the repository containing this script.

.PARAMETER RunExport
    Run Export-KnowledgeMarkdown.ps1 first to refresh the local mirror before uploading.

.PARAMETER Mirror
    Delete files/folders in the SharePoint target that no longer exist in the source
    (true mirror). Without this switch the sync is additive/update-only.

.PARAMETER ClientId
    Entra ID app registration (client) id for unattended auth.

.PARAMETER Tenant
    Tenant domain, e.g. contoso.onmicrosoft.com (required with app-only auth).

.PARAMETER CertificatePath
    Path to a .pfx certificate for app-only auth.

.PARAMETER CertificatePassword
    SecureString password for the .pfx (if protected).

.PARAMETER Interactive
    Use interactive browser sign-in instead of app-only auth (for local testing).

.PARAMETER WhatIf
    Show what would be uploaded/removed without making changes.

.EXAMPLE
    # Local test with interactive sign-in, refreshing the export first
    ./tools/Publish-KnowledgeToSharePoint.ps1 `
        -SiteUrl 'https://contoso.sharepoint.com/sites/SupplyChainArchitect' `
        -TargetFolder 'SupplyChainArchitect' `
        -RunExport -Interactive -Verbose

.EXAMPLE
    # Unattended app-only sync (Azure Automation / pipeline), full mirror
    ./tools/Publish-KnowledgeToSharePoint.ps1 `
        -SiteUrl 'https://contoso.sharepoint.com/sites/SupplyChainArchitect' `
        -TargetFolder 'SupplyChainArchitect' `
        -ClientId $env:SP_CLIENT_ID `
        -Tenant 'contoso.onmicrosoft.com' `
        -CertificatePath './cert.pfx' `
        -RunExport -Mirror

.NOTES
    Requires the PnP.PowerShell module:  Install-Module PnP.PowerShell -Scope CurrentUser
    App registration needs Sites.ReadWrite.All (application) SharePoint permission
    granted with admin consent for unattended runs.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$SiteUrl,

    [Parameter()]
    [string]$LibraryName = 'Documents',

    [Parameter(Mandatory)]
    [string]$TargetFolder,

    [Parameter()]
    [string]$SourcePath,

    [Parameter()]
    [string]$RepositoryRoot = (Split-Path -Parent $PSScriptRoot),

    [Parameter()]
    [switch]$RunExport,

    [Parameter()]
    [switch]$Mirror,

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$Tenant,

    [Parameter()]
    [string]$CertificatePath,

    [Parameter()]
    [securestring]$CertificatePassword,

    [Parameter()]
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'

# --- Preconditions ---------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop

if (-not $SourcePath) {
    $SourcePath = Join-Path $RepositoryRoot 'build\knowledge-export\files'
}

# --- Optionally refresh the local mirror -----------------------------------
if ($RunExport) {
    $exportScript = Join-Path $PSScriptRoot 'Export-KnowledgeMarkdown.ps1'
    if (-not (Test-Path $exportScript)) {
        throw "Cannot find Export-KnowledgeMarkdown.ps1 at $exportScript"
    }
    Write-Host "Running knowledge export..." -ForegroundColor Cyan
    & $exportScript -RepositoryRoot $RepositoryRoot
}

if (-not (Test-Path $SourcePath)) {
    throw "Source path does not exist: $SourcePath. Run the export first or pass -RunExport."
}

# --- Connect to SharePoint -------------------------------------------------
$connectParams = @{ Url = $SiteUrl }
if ($Interactive) {
    if (-not $ClientId) {
        throw "-Interactive still requires -ClientId (an Entra app registration with delegated SharePoint permissions)."
    }
    $connectParams['Interactive'] = $true
    $connectParams['ClientId'] = $ClientId
}
else {
    if (-not $ClientId -or -not $Tenant) {
        throw "App-only auth requires -ClientId and -Tenant (or use -Interactive)."
    }
    if (-not $CertificatePath) {
        throw "App-only auth requires -CertificatePath (a .pfx for the app registration)."
    }
    $connectParams['ClientId'] = $ClientId
    $connectParams['Tenant'] = $Tenant
    $connectParams['CertificatePath'] = $CertificatePath
    if ($CertificatePassword) {
        $connectParams['CertificatePassword'] = $CertificatePassword
    }
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline @connectParams
Write-Host "Connected." -ForegroundColor Green

try {
    # Resolve the library's server-relative root folder.
    $web = Get-PnPWeb
    $webUrl = $web.ServerRelativeUrl.TrimEnd('/')
    $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop
    $libRootRelative = $list.RootFolder.ServerRelativeUrl.TrimEnd('/')

    # Target root inside the library.
    $targetRootRelative = if ($TargetFolder) {
        "$libRootRelative/$($TargetFolder.Trim('/'))"
    } else {
        $libRootRelative
    }

    # --- Ensure a folder path exists (creates each segment) ---------------
    $ensuredFolders = [System.Collections.Generic.HashSet[string]]::new()
    function Confirm-SharePointFolder {
        param([string]$RelativeUrl)

        if ([string]::IsNullOrWhiteSpace($RelativeUrl)) { return }
        if ($ensuredFolders.Contains($RelativeUrl)) { return }

        # Build the path segment by segment starting below the library root.
        if (-not $RelativeUrl.StartsWith($libRootRelative)) { return }
        $sub = $RelativeUrl.Substring($libRootRelative.Length).Trim('/')
        if ([string]::IsNullOrWhiteSpace($sub)) { return }

        $current = $libRootRelative
        foreach ($segment in $sub -split '/') {
            $parent = $current
            $current = "$current/$segment"
            if ($ensuredFolders.Contains($current)) { continue }
            $existing = Get-PnPFolder -Url $current -ErrorAction SilentlyContinue
            if (-not $existing) {
                if ($PSCmdlet.ShouldProcess($current, 'Create folder')) {
                    Resolve-PnPFolder -SiteRelativePath ($current.Substring($webUrl.Length).Trim('/')) | Out-Null
                }
            }
            [void]$ensuredFolders.Add($current)
        }
    }

    Confirm-SharePointFolder -RelativeUrl $targetRootRelative

    # --- Enumerate local files -------------------------------------------
    $localFiles = Get-ChildItem -Path $SourcePath -Recurse -File
    Write-Host "Found $($localFiles.Count) local file(s) to sync." -ForegroundColor Cyan

    # Track what we upload so we can prune orphans in -Mirror mode.
    $expectedRelative = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase)

    $uploaded = 0
    $skipped = 0

    foreach ($file in $localFiles) {
        $relative = [System.IO.Path]::GetRelativePath($SourcePath, $file.FullName) -replace '\\', '/'
        [void]$expectedRelative.Add($relative)

        $destFileRelative = "$targetRootRelative/$relative"
        $destFolderRelative = ($destFileRelative -replace '/[^/]+$', '')

        Confirm-SharePointFolder -RelativeUrl $destFolderRelative

        # Incremental: compare local hash against the stored SharePoint file.
        $needsUpload = $true
        $existingFile = Get-PnPFile -Url $destFileRelative -AsFileObject -ErrorAction SilentlyContinue
        if ($existingFile) {
            $localHash = (Get-FileHash -Path $file.FullName -Algorithm SHA256).Hash
            $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString('N'))
            try {
                Get-PnPFile -Url $destFileRelative -Path (Split-Path $tempPath) `
                    -FileName (Split-Path $tempPath -Leaf) -AsFile -Force -ErrorAction Stop | Out-Null
                $remoteHash = (Get-FileHash -Path $tempPath -Algorithm SHA256).Hash
                if ($remoteHash -eq $localHash) { $needsUpload = $false }
            }
            finally {
                if (Test-Path $tempPath) { Remove-Item $tempPath -Force -ErrorAction SilentlyContinue }
            }
        }

        if (-not $needsUpload) {
            $skipped++
            Write-Verbose "Unchanged: $relative"
            continue
        }

        if ($PSCmdlet.ShouldProcess($destFileRelative, 'Upload file')) {
            Add-PnPFile -Path $file.FullName -Folder $destFolderRelative -Values @{} | Out-Null
        }
        $uploaded++
        Write-Host "  Uploaded: $relative" -ForegroundColor DarkGreen
    }

    # --- Prune orphaned files (mirror mode) ------------------------------
    $removed = 0
    if ($Mirror) {
        Write-Host "Mirror mode: checking for orphaned files under $targetRootRelative ..." -ForegroundColor Cyan
        $remoteItems = Get-PnPListItem -List $LibraryName -PageSize 500 |
            Where-Object {
                $_.FieldValues.FileRef -and
                $_.FieldValues.FileRef.StartsWith("$targetRootRelative/") -and
                $_.FileSystemObjectType -eq 'File'
            }

        foreach ($item in $remoteItems) {
            $fileRef = $item.FieldValues.FileRef
            $relative = $fileRef.Substring($targetRootRelative.Length).Trim('/')
            if (-not $expectedRelative.Contains($relative)) {
                if ($PSCmdlet.ShouldProcess($fileRef, 'Delete orphaned file')) {
                    Remove-PnPFile -ServerRelativeUrl $fileRef -Force
                }
                $removed++
                Write-Host "  Removed: $relative" -ForegroundColor DarkYellow
            }
        }
    }

    Write-Host ""
    Write-Host "Sync complete:" -ForegroundColor Green
    Write-Host "  Uploaded : $uploaded"
    Write-Host "  Unchanged: $skipped"
    if ($Mirror) { Write-Host "  Removed  : $removed" }
    Write-Host "  Target   : $SiteUrl -> $targetRootRelative"
    Write-Host ""
    Write-Host "Point your M365 Copilot agent's knowledge source at:" -ForegroundColor Cyan
    Write-Host "  $SiteUrl (library '$LibraryName', folder '$TargetFolder')"
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
