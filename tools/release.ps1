#Requires -Version 5.0

<#
.SYNOPSIS
  Publishes the PSFoundation module to the PowerShell Gallery.

.DESCRIPTION
  Builds source archives, generates SHA256 checksums, and publishes the module
  to a PowerShell repository (PSGallery by default). Designed to be called from
  CI after semantic-release determines the next version, though it can also be
  used interactively.

  The workflow is:
    1. Optionally rebuild dist/ archives via build.ps1.
    2. Generate SHA256 checksums for all dist/*.tar.gz and dist/*.zip files.
    3. Publish the module from ./src to the target repository.

  The script does NOT modify the module manifest version automatically. Use
  -Version to pass the resolved semantic version (which Publish-Module will use
  to override the manifest value in memory).

.PARAMETER SkipBuild
  Skip the initial build step. Use when archives are already present in dist/.

.PARAMETER Version
  Version string to publish (e.g. '1.2.3'). Passed through to Publish-Module
  -RequiredVersion.

.PARAMETER NuGetApiKey
  API key for the PowerShell repository. Falls back to the NUGET_API_KEY
  environment variable when not supplied.

.PARAMETER Gallery
  Target PSRepository name. Defaults to PSGallery.

.PARAMETER DryRun
  Report what WOULD be done without making changes or publishing.

.PARAMETER SkipPublish
  Build and generate checksums but skip the publish step. Useful for CI
  validation of the build artifacts.

.PARAMETER SkipChecksums
  Skip generation of the CHECKSUMS_SHA256.txt file.

.EXAMPLE
  PS> ./release.ps1 -Version 1.0.0 -NuGetApiKey $env:NUGET_API_KEY
  Builds, generates checksums, and publishes v1.0.0 to PSGallery.

.EXAMPLE
  PS> ./release.ps1 -Version 1.0.0 -DryRun
  Reports planned actions without publishing.

.EXAMPLE
  PS> ./release.ps1 -SkipBuild -SkipPublish
  Only generates checksums for existing dist/ artifacts.

.LINK
  https://github.com/adnoctem/PSFoundation

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [switch]$SkipBuild,

  [string]$Version,

  [string]$NuGetApiKey,

  [ValidateNotNullOrEmpty()]
  [string]$Gallery = 'PSGallery',

  [switch]$DryRun,

  [switch]$SkipPublish,

  [switch]$SkipChecksums
)

$ErrorActionPreference = 'Stop'

$repositoryRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Split-Path -Path $PSScriptRoot -Parent))
$distPath = Join-Path -Path $repositoryRoot -ChildPath 'dist'
$srcPath = Join-Path -Path $repositoryRoot -ChildPath 'src'
$buildScript = Join-Path -Path $PSScriptRoot -ChildPath 'build.ps1'
$checksumPath = Join-Path -Path $distPath -ChildPath 'CHECKSUMS_SHA256.txt'

# ---- Resolve API key --------------------------------------------------------
if (-not $NuGetApiKey) {
  $NuGetApiKey = $env:NUGET_API_KEY
}
if (-not $NuGetApiKey -and -not $DryRun -and -not $SkipPublish) {
  throw 'NuGetApiKey is required for publishing. Supply -NuGetApiKey or set the NUGET_API_KEY environment variable.'
}

# ---- Build ------------------------------------------------------------------
if (-not $SkipBuild) {
  if ($DryRun) {
    Write-Output "[DRY RUN] Would run build.ps1 to create dist/$((Get-Item -Path $repositoryRoot).Name).zip / .tar.gz"
  }
  else {
    Write-Output "Running build.ps1 ..."
    & $buildScript
    if ($LASTEXITCODE -ne 0) {
      throw "Build failed with exit code $LASTEXITCODE."
    }
  }
}

# ---- Generate checksums -----------------------------------------------------
if (-not $SkipChecksums) {
  if (-not (Test-Path -LiteralPath $distPath -PathType Container)) {
    Write-Warning "dist/ directory does not exist. Skipping checksum generation."
  }
  else {
    $archives = Get-ChildItem -LiteralPath $distPath -Include '*.tar.gz', '*.zip' -File |
      Sort-Object Name

    if ($archives.Count -eq 0) {
      Write-Warning 'No archive files found in dist/. Skipping checksum generation.'
    }
    else {
      if ($DryRun) {
        Write-Output "[DRY RUN] Would generate SHA256 checksums for $($archives.Count) archive(s) -> $checksumPath"
      }
      else {
        $checksums = foreach ($archive in $archives) {
          $hash = (Get-FileHash -Path $archive.FullName -Algorithm SHA256).Hash
          "$hash  $($archive.Name)"
        }

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllLines($checksumPath, $checksums, $utf8NoBom)
        Write-Output "Checksums written: $checksumPath"
      }
    }
  }
}

# ---- Validate module source -------------------------------------------------
if (-not $SkipPublish -and -not $DryRun) {
  if (-not (Test-Path -LiteralPath $srcPath -PathType Container)) {
    throw "Module source directory not found: $srcPath"
  }

  $manifestFiles = Get-ChildItem -LiteralPath $srcPath -Filter '*.psd1' -File
  if ($manifestFiles.Count -eq 0) {
    throw "No .psd1 module manifest found in: $srcPath"
  }
}

# ---- Publish ----------------------------------------------------------------
if ($SkipPublish) {
  Write-Output 'Publish skipped (SkipPublish is set).'
  exit 0
}

if ($DryRun) {
  $versionLabel = if ($Version) { $Version } else { '(from manifest)' }
  Write-Output "[DRY RUN] Would publish module to $Gallery (version: $versionLabel)"
  exit 0
}

$publishParams = @{
  Path = $srcPath
  Repository = $Gallery
  NuGetApiKey = $NuGetApiKey
  Force = $true
  ErrorAction = 'Stop'
}

if ($Version) {
  $publishParams.RequiredVersion = $Version
}

Write-Output "Publishing module to $Gallery ..."
Publish-Module @publishParams
Write-Output "Published successfully."
