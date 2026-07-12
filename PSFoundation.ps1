#Requires -Version 5.0

<#
.SYNOPSIS
  Central launcher for PSFoundation repository tool scripts.

.DESCRIPTION
  Routes a friendly command name to the matching PowerShell tool script under
  the tools directory, passing any remaining arguments through. Supports common
  aliases from different developer ecosystems so that `build`, `bundle`, `test`,
  `lint`, `check`, `format`, `fmt`, `init`, `setup`, `bootstrap`, `release`,
  and `publish` all work without needing to memorise exact verb/noun pairs.

.PARAMETER Command
  Tool to invoke. Accepts case-insensitive shorthand:
    init, initialize, setup, bootstrap  .ps1   initialize.ps1
    build, bundle, package              .ps1   build.ps1
    format, fmt, fix                    .ps1   format.ps1
    lint, check, analyze                .ps1   lint.ps1
    test, tests, pester                 .ps1   test.ps1
    release, publish                 .ps1   release.ps1
    install                           .ps1   install.ps1
    prune                             .ps1   prune.ps1

  All remaining positional and named arguments after Command are forwarded to
  the target script.

.EXAMPLE
  PS> .\PSFoundation.ps1 init
  Runs tools\initialize.ps1.

.EXAMPLE
  PS> .\PSFoundation.ps1 format -Check
  Runs tools\format.ps1 -Check.

.EXAMPLE
  PS> .\PSFoundation.ps1 build -Format Zip
  Runs tools\build.ps1 -Format Zip.

.EXAMPLE
  PS> .\PSFoundation.ps1 lint -Path ./src
  Runs tools\lint.ps1 -Path ./src.

.EXAMPLE
  PS> .\PSFoundation.ps1 test
  Runs tools\test.ps1 (all Pester tests).

.EXAMPLE
  PS> .\PSFoundation.ps1 release -DryRun
  Runs tools\release.ps1 -DryRun.

.EXAMPLE
  PS> .\PSFoundation.ps1 prune -WhatIf
  Runs tools\prune.ps1 -WhatIf.

.LINK
  https://github.com/adnoctem/PSFoundation

.NOTES
  Author: Maximilian Gindorfer <info@mvprowess.com>
  License: MIT
#>

# NOTE: deliberately NOT using [CmdletBinding()] here. This is a pass-through
# launcher: dropping advanced-function mode lets $args collect everything after
# the command name, which then splats to the target with named parameters intact.
param (
  [string]$Command
)

if (-not $Command) {
  Write-Error 'A command is required. Available: init(ialize), setup, bootstrap, build, bundle, package, format, fmt, fix, lint, check, analyze, test, pester, release, publish, install, prune'
  exit 1
}

$scriptMap = @{
  'init' = 'initialize'
  'initialize' = 'initialize'
  'setup' = 'initialize'
  'bootstrap' = 'initialize'
  'build' = 'build'
  'bundle' = 'build'
  'package' = 'build'
  'format' = 'format'
  'fmt' = 'format'
  'fix' = 'format'
  'lint' = 'lint'
  'check' = 'lint'
  'analyze' = 'lint'
  'test' = 'test'
  'tests' = 'test'
  'pester' = 'test'
  'release' = 'release'
  'publish' = 'release'
  'install' = 'install'
  'prune' = 'prune'
}

$commandKey = $Command.ToLowerInvariant()
$scriptName = $scriptMap[$commandKey]

if (-not $scriptName) {
  Write-Error "Unknown command '$Command'. Available: init(ialize), setup, bootstrap, build, bundle, package, format, fmt, fix, lint, check, analyze, test, pester, release, publish, install, prune"
  exit 1
}

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "tools/$scriptName.ps1"

if (-not (Test-Path -LiteralPath $scriptPath)) {
  Write-Error "Tool script not found: $scriptPath"
  exit 1
}

& $scriptPath @args
exit $LASTEXITCODE
