<#
  Initial setup script used to download PowerShell module dependencies
  defined in the project manifest and set up the project for local use.

  .PARAMETER Force
    Reinstall all modules even if the required version is already present.
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'This script is intended for interactive use and Write-Host is appropriate for user feedback.')]

[CmdletBinding()]
param(
  [switch]$Force
)

# ---- Source the module maintenance functions (bypasses Test-ModuleManifest) --
$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$maintenancePath = Join-Path -Path $repoRoot -ChildPath 'src/maintenance.ps1'
if (Test-Path -LiteralPath $maintenancePath -PathType Leaf) {
  . $maintenancePath
}
else {
  Write-Error "Module source not found: $maintenancePath"
  exit 1
}

# ---- Ensure NuGet provider (required by PowerShellGet) ----------------------
$null = Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue

# ---- Configure module -------------------------------------------------------
$RepositoryRoot = Split-Path -Path $PSScriptRoot -Parent
$manifestPath = Join-Path -Path $RepositoryRoot -ChildPath 'src/PSFoundation.psd1'
$manifest = Test-ModuleManifest -Path $manifestPath -ErrorAction SilentlyContinue

Write-Host "Using manifest: $manifest"

foreach ($mod in $manifest.RequiredModules) {
  if ($mod -is [string]) {
    $name = $mod
    $minVer = $null
    $exactVer = $null
  }
  else {
    $name = $mod.Name
    $minVer = $mod.Version
    $exactVer = $mod.RequiredVersion
  }

  Write-Host "Ensuring module '$name' is installed.." -ForegroundColor Yellow

  $installed = Get-Module -ListAvailable -Name $name |
    Sort-Object Version -Descending |
    Select-Object -First 1

  # ---- Determine whether current install satisfies the requirement ----
  $satisfied = $false
  if ($installed -and -not $Force) {
    if ($exactVer) {
      $satisfied = $installed.Version -eq $exactVer
    }
    elseif ($minVer) {
      $satisfied = $installed.Version -ge $minVer
    }
  }

  if ($satisfied) {
    Write-Host "    -> OK (found $($installed.Version))" -ForegroundColor Green
    continue
  }

  # ---- Install ----

  if ($exactVer) {
    Add-PSModule -Name $name -Version $exactVer.ToString() -Scope CurrentUser -Force
  }
  elseif ($minVer) {
    Add-PSModule -Name $name -MinimumVersion $minVer.ToString() -Scope CurrentUser -Force
  }
  else {
    Add-PSModule -Name $name -Scope CurrentUser -Force
  }

  Write-Host "    -> Installed $name" -ForegroundColor Green
}

# ---------------------------------------------------------------
Write-Host "Successfully processed all RequiredModules!" -ForegroundColor Yellow
