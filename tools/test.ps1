#Requires -Version 5.1

<#
.SYNOPSIS
  Runs Pester tests for the PSFoundation module.

.DESCRIPTION
  Invokes Pester against test files in the repository tests directory. By
  default all test files are executed. The script exits with the number of
  failed tests as its exit code, making it suitable for CI usage.

.PARAMETER Path
  Path to test files or directory. Defaults to the repository tests directory.

.EXAMPLE
  PS> ./test.ps1
  Runs all tests in the tests directory.

.EXAMPLE
  PS> ./test.ps1 -Path ./tests/user.Tests.ps1
  Runs only the user.Tests.ps1 test file.

.LINK
  https://github.com/adnoctem/PSFoundation

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[CmdletBinding()]
param (
  [string[]]$Path = @(Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'tests')
)

$ErrorActionPreference = 'Stop'

$pester = Get-Module -ListAvailable -Name Pester |
  Sort-Object Version -Descending |
  Select-Object -First 1

if (-not $pester -or $pester.Version -lt [version]'5.0.0') {
  Write-Warning 'Pester 5.0.0+ is required but was not found. Installing it now (CurrentUser scope)...'

  $maintenancePath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/maintenance.ps1'
  . $maintenancePath
  Add-PSModule -Name Pester -MinimumVersion '5.0.0' -Scope CurrentUser -Force

  $pester = Get-Module -ListAvailable -Name Pester |
    Sort-Object Version -Descending |
    Select-Object -First 1

  if (-not $pester -or $pester.Version -lt [version]'5.0.0') {
    Write-Error "Failed to install Pester 5.0.0+. Run '.\PSFoundation.ps1 init' and try again."
    exit 1
  }
}

Write-Verbose "Using Pester $($pester.Version)"
Import-Module Pester -MinimumVersion 5.0.0 -ErrorAction Stop

$config = [PesterConfiguration]@{
  Run = @{
    Path = $Path
    PassThru = $true
  }
  Output = @{
    Verbosity = 'Detailed'
  }
}

$result = Invoke-Pester -Configuration $config

exit $result.FailedCount
