#Requires -Version 5.0

<#
.SYNOPSIS
  Publishes the PSFoundation module to the PowerShell Gallery.

.DESCRIPTION
  Builds source archives, generates SHA256 checksums, and publishes the module
  to a PowerShell repository (PSGallery by default). Designed to be called from
  CI after semantic-release determines the next version, though it can also be
  used interactively.

  The script supports two modes:

  - Prepare mode (-Prepare): invoked by the @semantic-release/exec plugin during
    the semantic-release prepare phase. Writes the resolved next version into
    src/PSFoundation.psd1 (ModuleVersion plus, for prerelease suffixes such as
    '1.1.0-beta.1', the PSData.Prerelease key), rebuilds the dist/ archives and
    writes the CHECKSUMS file, then exits. The manifest change is committed by
    @semantic-release/git as part of the release commit, keeping module source,
    GitHub release bundles and PSGallery package versions permanently in sync.

  - Publish mode (default): optionally rebuilds dist/ archives via build.ps1,
    generates SHA256 checksums, and publishes the module from ./src to the
    target repository. The module manifest is the single source of truth for
    the published version: Publish-PSResource is preferred when available and
    Publish-Module is used as fallback. When -Version is supplied it must match
    the manifest version and is verified before publishing.

.PARAMETER SkipBuild
  Skip the initial build step. Use when archives are already present in dist/.

.PARAMETER Version
  Semantic version for the release (e.g. '1.2.3' or '1.2.3-beta.1'). Required
  with -Prepare, where it is written into the module manifest. In publish mode
  it is validated against the manifest version and must match it.

.PARAMETER Prepare
  Run the semantic-release prepare phase only: synchronize the module manifest
  to -Version, rebuild the dist/ archives, regenerate the CHECKSUMS file, then
  exit without publishing.

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
  PS> ./release.ps1 -Prepare -Version 1.0.0
  Synchronizes src/PSFoundation.psd1 to v1.0.0, rebuilds dist/, regenerates
  the CHECKSUMS file, and exits without publishing.

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
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [switch]$SkipBuild,

  [string]$Version,

  [switch]$Prepare,

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

function Split-ReleaseVersion {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Version
  )

  $parts = $Version -split '-', 2
  $coreVersion = $parts[0]

  if ($coreVersion -notmatch '^\d+\.\d+\.\d+(\.\d+)?$') {
    throw "Version '$Version' is not a valid semantic version. Expected 'X.Y.Z' or 'X.Y.Z-prerelease'."
  }

  $prerelease = if ($parts.Count -gt 1) { $parts[1] } else { $null }
  if ($prerelease -and $prerelease -notmatch '^[0-9A-Za-z-]+(\.[0-9A-Za-z-]+)*$') {
    throw "Prerelease label '$prerelease' in version '$Version' is invalid."
  }

  [pscustomobject]@{
    CoreVersion = $coreVersion
    Prerelease = $prerelease
  }
}

function Write-DistChecksum {
  [CmdletBinding()]
  param (
    [switch]$Skip,
    [switch]$DryRun
  )

  if ($Skip) { return }

  if (-not (Test-Path -LiteralPath $distPath -PathType Container)) {
    Write-Warning "dist/ directory does not exist. Skipping checksum generation."
    return
  }

  $archives = Get-ChildItem -LiteralPath $distPath -File |
    Where-Object { $_.Name -like '*.tar.gz' -or $_.Name -like '*.zip' } |
    Sort-Object Name

  if ($archives.Count -eq 0) {
    Write-Warning 'No archive files found in dist/. Skipping checksum generation.'
    return
  }

  if ($DryRun) {
    Write-Output "[DRY RUN] Would generate SHA256 checksums for $($archives.Count) archive(s) -> $checksumPath"
    return
  }

  $checksums = foreach ($archive in $archives) {
    $hash = (Get-FileHash -Path $archive.FullName -Algorithm SHA256).Hash
    "$hash  $($archive.Name)"
  }

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllLines($checksumPath, $checksums, $utf8NoBom)
  Write-Output "Checksums written: $checksumPath"
}

# ---- Prepare (semantic-release prepare phase) --------------------------------
# Invoked by the @semantic-release/exec plugin before the release commit is
# created. Synchronizes the module manifest to the resolved next version,
# rebuilds the dist/ archives and regenerates the CHECKSUMS file so the GitHub
# release assets match the release. Exits after this phase.
if ($Prepare) {
  if (-not $Version) {
    throw 'Version is required when using -Prepare. Pass -Version with the semantic-release next version (e.g. "1.0.0" or "1.1.0-beta.1").'
  }

  $versionInfo = Split-ReleaseVersion -Version $Version
  $manifestFiles = Get-ChildItem -LiteralPath $srcPath -Filter '*.psd1' -File
  if ($manifestFiles.Count -eq 0) {
    throw "No .psd1 module manifest found in: $srcPath"
  }
  $manifestFile = $manifestFiles[0].FullName

  $manifestContent = Get-Content -LiteralPath $manifestFile -Raw

  $rootIndentMatch = [regex]::Match($manifestContent, "(?m)^(\s*)RootModule\s*=\s*")
  if (-not $rootIndentMatch.Success) {
    throw "Could not determine module manifest indentation (RootModule key not found): $manifestFile"
  }
  $topLevelIndent = [regex]::Escape($rootIndentMatch.Groups[1].Value)

  $moduleVersionPattern = "(?m)^($topLevelIndent)(ModuleVersion)\s*=\s*'[^']*'"
  if (-not [regex]::IsMatch($manifestContent, $moduleVersionPattern)) {
    throw "Could not locate top-level ModuleVersion in module manifest: $manifestFile"
  }
  $manifestContent = [regex]::Replace($manifestContent, $moduleVersionPattern, "`$1`$2 = '$($versionInfo.CoreVersion)'")

  if ($versionInfo.Prerelease) {
    $prereleasePattern = "(?m)^(\s*)[# ]*Prerelease\s*=\s*'[^']*'"
    if (-not [regex]::IsMatch($manifestContent, $prereleasePattern)) {
      throw "Could not locate Prerelease key in module manifest: $manifestFile"
    }
    $manifestContent = [regex]::Replace($manifestContent, $prereleasePattern, "`$1Prerelease = '$($versionInfo.Prerelease)'")
  }
  else {
    $livePrereleasePattern = "(?m)^(\s*)Prerelease\s*=\s*'[^']*'(\r\n|\r|\n)"
    $livePrereleaseMatch = [regex]::Match($manifestContent, $livePrereleasePattern)
    if ($livePrereleaseMatch.Success) {
      $prereleaseIndent = $livePrereleaseMatch.Groups[1].Value
      $prereleaseEol = $livePrereleaseMatch.Groups[2].Value
      $manifestContent = [regex]::Replace($manifestContent, $livePrereleasePattern, "$prereleaseIndent# Prerelease = ''$prereleaseEol")
    }
  }

  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($manifestFile, $manifestContent, $utf8NoBom)
  Write-Output "Manifest version set to $Version ($manifestFile)"

  if (-not $SkipBuild) {
    Write-Output "Running build.ps1 ..."
    & $buildScript
    if ($LASTEXITCODE -ne 0) {
      throw "Build failed with exit code $LASTEXITCODE."
    }
  }

  Write-DistChecksum
  Write-Output 'Prepare phase complete.'
  exit 0
}

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
Write-DistChecksum -Skip:$SkipChecksums -DryRun:$DryRun

# ---- Validate module source -------------------------------------------------
if (-not $SkipPublish -and -not $DryRun) {
  if (-not (Test-Path -LiteralPath $srcPath -PathType Container)) {
    throw "Module source directory not found: $srcPath"
  }

  $manifestFiles = Get-ChildItem -LiteralPath $srcPath -Filter '*.psd1' -File
  if ($manifestFiles.Count -eq 0) {
    throw "No .psd1 module manifest found in: $srcPath"
  }
  $manifestFile = $manifestFiles[0].FullName

  $manifest = Test-ModuleManifest -Path $manifestFile

  $manifestPrerelease = $null
  if ($manifest.PrivateData -and $manifest.PrivateData.PSData) {
    $psData = $manifest.PrivateData.PSData
    if ($psData -is [System.Collections.IDictionary]) {
      if ($psData.Contains('Prerelease')) {
        $manifestPrerelease = [string]$psData['Prerelease']
      }
    }
    else {
      $prereleaseProperty = $psData.PSObject.Properties['Prerelease']
      if ($prereleaseProperty -and $prereleaseProperty.Value) {
        $manifestPrerelease = [string]$prereleaseProperty.Value
      }
    }
  }
  $manifestVersionString = if ($manifestPrerelease) {
    "$($manifest.Version)-$manifestPrerelease"
  }
  else {
    $manifest.Version.ToString()
  }

  if ($Version) {
    $versionInfo = Split-ReleaseVersion -Version $Version
    $requestedVersionString = if ($versionInfo.Prerelease) {
      "$($versionInfo.CoreVersion)-$($versionInfo.Prerelease)"
    }
    else {
      $versionInfo.CoreVersion
    }

    if ($requestedVersionString -ne $manifestVersionString) {
      throw "Version '$requestedVersionString' does not match the module manifest version '$manifestVersionString' ($manifestFile). Bump the manifest (e.g. via 'release.ps1 -Prepare -Version $requestedVersionString') before publishing."
    }
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

$publishPSResource = Get-Command -Name 'Publish-PSResource' -ErrorAction SilentlyContinue
if ($publishPSResource) {
  $publishParams = @{
    Path = $srcPath
    Repository = $Gallery
    ApiKey = $NuGetApiKey
    ErrorAction = 'Stop'
  }

  Write-Output "Publishing module to $Gallery (via Publish-PSResource) ..."
  Publish-PSResource @publishParams
}
else {
  $publishModule = Get-Command -Name 'Publish-Module' -ErrorAction SilentlyContinue
  if (-not $publishModule) {
    throw 'Neither Publish-PSResource nor Publish-Module is available. Install Microsoft.PowerShell.PSResourceGet or PowerShellGet and try again.'
  }

  if ($manifestPrerelease) {
    $psGet = Get-Module -Name PowerShellGet -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $psGet -or $psGet.Version.Major -lt 2) {
      throw "The manifest carries a prerelease label ('$manifestPrerelease'), which requires PowerShellGet 2.x or newer. Upgrade PowerShellGet or install Microsoft.PowerShell.PSResourceGet."
    }
  }

  # PowerShellGet requires the folder leaf passed to -Path to match the module
  # name, so stage a copy under a folder named after the module.
  $moduleName = [System.IO.Path]::GetFileNameWithoutExtension((Get-Item -LiteralPath $manifestFile).Name)
  $stagingRoot = Join-Path -Path $env:TEMP -ChildPath "$moduleName-Publish-$([guid]::NewGuid())"
  $stagingModulePath = Join-Path -Path $stagingRoot -ChildPath $moduleName
  New-Item -ItemType Directory -Path $stagingModulePath -Force | Out-Null
  Copy-Item -Path (Join-Path -Path $srcPath -ChildPath '*') -Destination $stagingModulePath -Recurse -Force

  try {
    $publishParams = @{
      Path = $stagingModulePath
      Repository = $Gallery
      NuGetApiKey = $NuGetApiKey
      Force = $true
      ErrorAction = 'Stop'
    }

    Write-Output "Publishing module to $Gallery (via Publish-Module) ..."
    Publish-Module @publishParams
  }
  finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
}
Write-Output "Published successfully."
