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

function Ensure-ModuleInstalled {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Ensure- is a descriptive verb for a private helper that guarantees a module version is installed.')]

  param (
    [Parameter(Mandatory = $true)]
    [string]$Name,

    [version]$MinimumVersion,

    [version]$RequiredVersion,

    [switch]$Force
  )

  Write-Host "Ensuring module '$Name' is installed.." -ForegroundColor Yellow

  $installed = Get-Module -ListAvailable -Name $Name |
    Sort-Object Version -Descending |
    Select-Object -First 1

  # ---- Determine whether current install satisfies the requirement ----
  $satisfied = $false
  if ($installed -and -not $Force) {
    if ($RequiredVersion) {
      $satisfied = $installed.Version -eq $RequiredVersion
    }
    elseif ($MinimumVersion) {
      $satisfied = $installed.Version -ge $MinimumVersion
    }
  }

  if ($satisfied) {
    Write-Host "    -> OK (found $($installed.Version))" -ForegroundColor Green
    return
  }

  # ---- Install ----

  if ($RequiredVersion) {
    Add-PSModule -Name $Name -Version $RequiredVersion.ToString() -Scope CurrentUser -Force
  }
  elseif ($MinimumVersion) {
    Add-PSModule -Name $Name -MinimumVersion $MinimumVersion.ToString() -Scope CurrentUser -Force
  }
  else {
    Add-PSModule -Name $Name -Scope CurrentUser -Force
  }

  Write-Host "    -> Installed $Name" -ForegroundColor Green
}

foreach ($mod in $manifest.RequiredModules) {
  if ($mod -is [string]) {
    Ensure-ModuleInstalled -Name $mod -Force:$Force
  }
  else {
    Ensure-ModuleInstalled -Name $mod.Name -MinimumVersion $mod.Version -RequiredVersion $mod.RequiredVersion -Force:$Force
  }
}

# ---- Dev dependencies (test-only, not shipped with the module) --------------
$devDependencies = @(
  @{ Name = 'Pester'; MinimumVersion = '5.0.0' }
)

foreach ($dev in $devDependencies) {
  Ensure-ModuleInstalled -Name $dev.Name -MinimumVersion $dev.MinimumVersion -Force:$Force
}

# ---------------------------------------------------------------
Write-Host "Successfully processed all module dependencies!" -ForegroundColor Yellow
