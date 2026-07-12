#Requires -Version 5.0

<#
.SYNOPSIS
  Prunes (removes) installed PowerShell module versions.

.DESCRIPTION
  Thin wrapper around Remove-PSModule from the PSFoundation module source.
  Dot-sources the implementation directly, bypassing module-manifest validation.

  By default, keeps only the newest version of each module and removes older
  ones. Use -All to remove every version (prompts for confirmation unless
  -Force), or -LatestToKeep to control how many versions are retained.

  Usage:
    .\PSFoundation.ps1 prune [-Name <regex>] [-All] [-LatestToKeep <int>] [-Scope <scope>] [-Path <dir>] [-Force] [-WhatIf]

.EXAMPLE
  .\PSFoundation.ps1 prune -WhatIf
  Previews which old module versions would be removed.

.EXAMPLE
  .\PSFoundation.ps1 prune -All -Force
  Removes all CurrentUser module versions without prompting.

.EXAMPLE
  .\PSFoundation.ps1 prune -Name 'Pester' -LatestToKeep 2
  Keeps the two newest Pester versions, removes older ones.

.LINK
  https://github.com/adnoctem/PSFoundation

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$Name,

  [ValidateRange(1, [int]::MaxValue)]
  [int]$LatestToKeep = 1,

  [switch]$All,

  [ValidateSet('CurrentUser', 'AllUsers')]
  [string]$Scope = 'CurrentUser',

  [string]$Path,

  [switch]$Force
)

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$maintenancePath = Join-Path -Path $repoRoot -ChildPath 'src/maintenance.ps1'

if (-not (Test-Path -LiteralPath $maintenancePath -PathType Leaf)) {
  Write-Error "Module source not found: $maintenancePath"
  exit 1
}

. $maintenancePath
Remove-PSModule @PSBoundParameters
exit $LASTEXITCODE
