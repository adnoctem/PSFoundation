#Requires -Version 5.0

<#
.SYNOPSIS
  Installs the PSFoundation module to the local PowerShell module path.

.DESCRIPTION
  Copies or symlinks the src/ directory into the user's PowerShell Modules path
  so that the module can be imported with Import-Module PSFoundation from any
  session. By default, files are copied. Use -SymbolicLink to create a directory
  junction instead (instant updates, suitable for active development).

  The module version is read from src/PSFoundation.psd1 and used to create the
  versioned module folder. Use -Undo to remove a previously installed copy.

.PARAMETER SymbolicLink
  Create a directory junction instead of copying files. The junction points back
  to the src/ directory so changes take effect immediately without reinstalling.

.PARAMETER Undo
  Remove the installed module from the local module path. Works for both copy
  and symlink installations.

.PARAMETER Scope
  Install scope. Defaults to CurrentUser. AllUsers requires an elevated session.

.PARAMETER Force
  Skip confirmation prompts.

.EXAMPLE
  PS> ./install.ps1
  Copies src/ into $env:USERPROFILE\Documents\PowerShell\Modules\PSFoundation\<version>\.

.EXAMPLE
  PS> ./install.ps1 -SymbolicLink
  Creates a junction so that module changes are reflected immediately.

.EXAMPLE
  PS> ./install.ps1 -Undo
  Removes the locally installed PSFoundation module.

.LINK
  https://github.com/adnoctem/PSFoundation

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
  [switch]$SymbolicLink,

  [switch]$Undo,

  [ValidateSet('CurrentUser', 'AllUsers')]
  [string]$Scope = 'CurrentUser',

  [switch]$Force
)

$ErrorActionPreference = 'Stop'

$repositoryRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Split-Path -Path $PSScriptRoot -Parent))
$srcPath = Join-Path -Path $repositoryRoot -ChildPath 'src'
$manifestPath = Join-Path -Path $srcPath -ChildPath 'PSFoundation.psd1'

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
  throw "Module manifest not found: $manifestPath"
}

$manifest = Test-ModuleManifest -Path $manifestPath -ErrorAction Stop
$moduleVersion = $manifest.Version.ToString()
$moduleName = $manifest.Name

if ($Scope -eq 'AllUsers') {
  $modulesRoot = Join-Path -Path $env:ProgramFiles -ChildPath 'PowerShell\Modules'
}
else {
  $documentsPath = if ($IsWindows -or $env:OS -eq 'Windows_NT') {
    [Environment]::GetFolderPath('Personal')
  }
  else {
    Join-Path -Path $HOME -ChildPath 'Documents'
  }
  $modulesRoot = Join-Path -Path $documentsPath -ChildPath 'PowerShell\Modules'
}

$installPath = Join-Path -Path $modulesRoot -ChildPath "$moduleName\$moduleVersion"

# ---- Undo mode --------------------------------------------------------------
if ($Undo) {
  if (-not (Test-Path -LiteralPath $installPath)) {
    Write-Output "No installation found at: $installPath"
    exit 0
  }

  $item = Get-Item -LiteralPath $installPath -Force
  $isSymlink = ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq [System.IO.FileAttributes]::ReparsePoint

  if ($PSCmdlet.ShouldProcess($installPath, "Remove $($isSymlink ? 'symlink' : 'copy')")) {
    if ($isSymlink) {
      try {
        [System.IO.Directory]::Delete($installPath)
      }
      catch {
        Remove-Item -LiteralPath $installPath -Recurse -Force
      }
    }
    else {
      Remove-Item -LiteralPath $installPath -Recurse -Force
    }
    Write-Output "Removed: $installPath"
  }

  $parentPath = Split-Path -Path $installPath -Parent
  $remaining = Get-ChildItem -LiteralPath $parentPath -ErrorAction SilentlyContinue
  if ($null -eq $remaining -or @($remaining).Count -eq 0) {
    if ($PSCmdlet.ShouldProcess($parentPath, 'Remove empty module folder')) {
      Remove-Item -LiteralPath $parentPath -Recurse -Force
      Write-Output "Removed empty folder: $parentPath"
    }
  }

  exit 0
}

# ---- Install mode -----------------------------------------------------------
if (Test-Path -LiteralPath $installPath) {
  if (-not $Force -and -not $PSCmdlet.ShouldContinue("$installPath exists. Overwrite?", 'Install PSFoundation')) {
    Write-Warning 'Installation cancelled.'
    exit 0
  }

  if ($PSCmdlet.ShouldProcess($installPath, 'Remove existing installation')) {
    Remove-Item -LiteralPath $installPath -Recurse -Force -ErrorAction SilentlyContinue
  }
}

if (-not (Test-Path -LiteralPath $modulesRoot)) {
  New-Item -Path $modulesRoot -ItemType Directory -Force | Out-Null
}

$parentPath = Split-Path -Path $installPath -Parent
if (-not (Test-Path -LiteralPath $parentPath)) {
  New-Item -Path $parentPath -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $installPath)) {
  New-Item -Path $installPath -ItemType Directory -Force | Out-Null
}

if ($SymbolicLink) {
  if (-not $IsWindows -and $env:OS -ne 'Windows_NT') {
    throw 'SymbolicLink installation is only supported on Windows (directory junctions).'
  }

  if ($PSCmdlet.ShouldProcess($srcPath, "Create junction -> $installPath")) {
    New-Item -ItemType Junction -Path $installPath -Target $srcPath -Force | Out-Null
    Write-Output "Junction created: $installPath -> $srcPath"
  }
}
else {
  if ($PSCmdlet.ShouldProcess($srcPath, "Copy src/ -> $installPath")) {
    Copy-Item -Path "$srcPath\*" -Destination $installPath -Recurse -Force
    Write-Output "Module installed: $installPath"
  }
}
