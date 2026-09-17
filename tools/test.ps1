#Requires -Version 5.1

<#
.SYNOPSIS
  Runs Pester tests for the PSFoundation module.

.DESCRIPTION
  Invokes Pester against test files in the repository tests directory. By
  default all test files are executed except Integration-tagged tests. Exits
  with 1 for failed tests, failed containers or empty discovery, otherwise 0.

.PARAMETER Path
  Path to test files or directory. Defaults to the repository tests directory.
.PARAMETER Coverage
  Collect informational source coverage in JaCoCo format, with no percentage gate.
.PARAMETER OutputDirectory
  Write NUnit test results and optional coverage here. Coverage defaults this to
  build/test-results when no directory is supplied.
.PARAMETER IncludeIntegration
  Include tests tagged Integration. These may require privileges or change state.

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
  [string[]]$Path = @(Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'tests'),
  [switch]$Coverage,
  [string]$OutputDirectory,
  [switch]$IncludeIntegration
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
  Run    = @{
    Path     = $Path
    PassThru = $true
  }
  Output = @{
    Verbosity = 'Detailed'
  }
}

if (-not $IncludeIntegration) { $config.Filter.ExcludeTag = @('Integration') }
if ($Coverage -and -not $OutputDirectory) {
  $OutputDirectory = Join-Path (Split-Path $PSScriptRoot -Parent) 'build/test-results'
}
if ($OutputDirectory) {
  $OutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
  $null = [IO.Directory]::CreateDirectory($OutputDirectory)
  $config.TestResult.Enabled = $true
  $config.TestResult.OutputFormat = 'NUnitXml'
  $config.TestResult.OutputPath = Join-Path $OutputDirectory 'tests.xml'
}
if ($Coverage) {
  # Detailed verbosity dumps every uncovered command; the XML retains that
  # information while Normal keeps the CI log useful.
  $config.Output.Verbosity = 'Normal'
  $config.CodeCoverage.Enabled = $true
  $config.CodeCoverage.Path = @(Join-Path (Split-Path $PSScriptRoot -Parent) 'src/*.ps1')
  $config.CodeCoverage.OutputFormat = 'JaCoCo'
  $config.CodeCoverage.OutputPath = Join-Path $OutputDirectory 'coverage.xml'
  $config.CodeCoverage.CoveragePercentTarget = 0
}

$result = Invoke-Pester -Configuration $config

if ($result.FailedCount -gt 0 -or $result.FailedContainersCount -gt 0 -or $result.TotalCount -eq 0) { exit 1 }
exit 0
