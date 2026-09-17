<p align="center">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/Powershell_256.png">
      <img src="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/Powershell_256.png" alt="PowerShell L" width="225">
    </picture>
    <h1 align="center">PSFoundation</h1>
</p>

[![License](https://img.shields.io/github/license/adnoctem/PSFoundation?label=License)][license]
[![Language](https://img.shields.io/github/languages/top/adnoctem/PSFoundation?label=PowerShell)][powershell]
[![PSGallery Version](https://img.shields.io/powershellgallery/v/PSFoundation)][psgallery_package]
[![CI Status](https://github.com/adnoctem/PSFoundation/actions/workflows/testing.yaml/badge.svg)][testing_workflow]
[![GitHub Release](https://img.shields.io/github/v/release/adnoctem/PSFoundation?label=Release)][github_releases]
[![GitHub Activity](https://img.shields.io/github/commit-activity/m/adnoctem/PSFoundation?label=Commits)][github_commits]
[![Semantic Release](https://img.shields.io/badge/Semantic_Release-enabled-brightgreen?logo=semanticrelease&logoColor=E5E4E7)][semantic_release]
[![Renovate](https://img.shields.io/badge/Renovate-enabled-brightgreen?logo=renovate&logoColor=1A1F6C)][renovate]
[![PreCommit](https://img.shields.io/badge/PreCommit-enabled-brightgreen?logo=precommit&logoColor=FAB040)][precommit]
[![Super-Linter](https://github.com/adnoctem/PSFoundation/actions/workflows/superlint.yaml/badge.svg)][superlinter_action]

`PSFoundation` is an open-source [MIT][license]-licensed [PowerShell][powershell] module library written and maintained by the [Ad Noctem
Collective][org] for Windows system administration, configuration management, and automation. The module targets both desktop Windows
installations and Windows Server environments and supports [PowerShell][powershell] 5.1 and above, including Windows PowerShell 5.1 as well
as newer PowerShell 7+ releases. It is published to the [PowerShell Gallery][psgallery_package] for easy discovery and installation.

The [`src`](src) directory contains the module source code — a collection of PowerShell functions organized by domain (registry, networking,
security, packages, system, etc.) — bundled together as a single importable module. The [`tools`](tools) directory contains the repository's
development tooling for building, formatting, linting, testing, and publishing the module.

### Module Coverage

PSFoundation provides functions across these domains:

| Module File        | Domain                                                                                                                         |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------ |
| `provisioning.ps1` | Offline domain join blob provisioning and djoin file creation                                                                  |
| `common.ps1`       | Operation result helpers and registry setting state management                                                                 |
| `data.ps1`         | Data transformation utilities (quote conversion, object merging)                                                               |
| `devices.ps1`      | Print and scan device enumeration and management                                                                               |
| `errors.ps1`       | Domain-specific deployment error guidance for AppX, WinGet, DISM, and MSI                                                      |
| `interop.ps1`      | COM interop and Outlook automation                                                                                             |
| `log.ps1`          | Console logging helpers                                                                                                        |
| `networking.ps1`   | IP validation, adapter resolution, address calculation, remote host reachability                                               |
| `packages.ps1`     | Win32 and AppX package lifecycle management                                                                                    |
| `permissions.ps1`  | Elevation, ownership takeover, and encrypted credential files                                                                  |
| `policies.ps1`     | LGPO integration and binary registry.pol reading and writing, including lossless raw round trips                               |
| `registry.ps1`     | Registry key and value CRUD with path resolution                                                                               |
| `security.ps1`     | Defender, firewall, event log analysis, certificate inventory, script signing                                                  |
| `settings.ps1`     | Default application associations                                                                                               |
| `system.ps1`       | OS version, .NET version, drive mapping, pending-reboot and file integrity checks, service account search, file lock diagnosis |
| `updates.ps1`      | Windows Update and Microsoft Store update management                                                                           |
| `user.ps1`         | User and SID information, AD credential validation, lockout source and FSMO reporting                                          |

## TL;DR

```pwsh
# Install from PowerShell Gallery
Install-Module -Name PSFoundation

# Or initialize the repository (download dependencies)
.\PSFoundation.ps1 init
# also: .\PSFoundation.ps1 initialize | setup | bootstrap

# format all PowerShell source files
.\PSFoundation.ps1 format
# also: .\PSFoundation.ps1 fmt | fix

# check formatting without modifying (CI / pre-commit)
.\PSFoundation.ps1 format -Check

# run PSScriptAnalyzer lint checks
.\PSFoundation.ps1 lint
# also: .\PSFoundation.ps1 check | analyze

# build distribution archives (ZIP + tar.gz)
.\PSFoundation.ps1 build
# also: .\PSFoundation.ps1 bundle | package

# run all Pester tests
.\PSFoundation.ps1 test
# also: .\PSFoundation.ps1 pester

# publish module to PowerShell Gallery
.\PSFoundation.ps1 release -Version 1.0.0 -NuGetApiKey $env:NUGET_API_KEY
# also: .\PSFoundation.ps1 publish

# dry-run release (build + checksums without publishing)
.\PSFoundation.ps1 release -Version 1.0.0 -DryRun
```

### Registry policy files

Read or create `registry.pol` files without LGPO.exe. Conversion preserves record order and duplicate instructions; it does not apply
policy. Keys are relative to their hive, with Machine/User scope determined by the consumer. Reading protected system files may require
administrator permissions.

```pwsh
$entries = ConvertFrom-RegistryPolicy -Path '.\Machine\registry.pol'
$entries | Format-Table Key, ValueName, Type, Data

# Raw records preserve payload bytes, including unknown registry types.
ConvertFrom-RegistryPolicy -Path '.\source.pol' -Raw |
  ConvertTo-RegistryPolicy -Path '.\copy.pol'

# Existing destinations require -Force; -WhatIf validates without writing.
$entries | ConvertTo-RegistryPolicy -Path '.\copy.pol' -Force -WhatIf
```

Records contain `Key`, `ValueName`, numeric `Type`, and `Data`. Decoded data uses strings, string arrays, unsigned integers, or bytes;
expandable strings remain unexpanded. Zero-length decoded payloads are null. Raw records carry the `PSFoundation.RegistryPolicy.RawEntry`
type name, which the writer uses to preserve their byte arrays. Keep that type name when editing raw objects. Files are limited to 64 MiB
and payloads to 65535 bytes. Parent directories must already exist.

### Deployment error guidance

Translate caught errors or native exit codes. The tables are curated, not exhaustive. Unknown or ambiguous errors emit no result, so keep
the original error as a fallback. Codes are matched independently of message language. Explicit native codes require a domain.

```pwsh
Get-ErrorTranslation -Code 3010 -Domain Msi

try {
  Add-AppxPackage -Path '.\package.msix' -ErrorAction Stop
} catch {
  $translation = Get-ErrorTranslation -ErrorRecord $_ -Domain Appx
  $detail = if ($translation) { $translation.Detail } else { $_.Exception.Message }
  New-OperationResult -Target 'package.msix' -Action Install -Status Failed `
    -Detail $detail -ErrorMessage $_.Exception.Message
}
```

Results contain `Code`, `Domain`, `Benign`, `Detail`, and `Matched`. `Benign` is true only for unconditional success codes in the table; MSI
restart results still require following the restart guidance. A newer installed AppX version remains a conditional case with
`Benign = $false`: the caller must decide whether that version satisfies the requested state. WinGet wrapper codes and installer codes
belong to their respective domains. Multiple distinct recognized codes in message text produce no result; structured exception HRESULTs take
precedence. The translator does not retry or suppress errors.

### Native process results

`Invoke-SafeProcess` preserves argument boundaries on Windows PowerShell 5.1 and PowerShell 7, and reads stdout and stderr concurrently. Use
`-AsResult` for `ExitCode`, `StdOut`, `StdErr`, `Duration`, `TimedOut`, and `Cancelled`. Existing `-PassThru` combined text and
`-OutputPath` behavior remain available; `-AsResult` takes precedence over `-PassThru`.

```pwsh
$result = Invoke-SafeProcess -FilePath 'whoami.exe' -ArgumentList '/all' -TimeoutSeconds 30 -AsResult
$result | Select-Object ExitCode, Duration, TimedOut
```

`-CancellationToken` accepts a .NET cancellation token. Timeout and cancellation attempt to terminate the child and its descendants;
detached processes may survive. Output capture has a separate ten-second drain limit for inherited pipe handles. Structured mode throws on
start/capture failures and returns nonzero native exit codes as data. Output is buffered in memory, so use this for bounded command output.

### Registry snapshots and restoration

The existing `Export-RegistrySettingState` output stays compatible. Add `-Detailed` to capture version 1 snapshots containing explicit
`Exists`, `KeyExists`, `Type`, `Preferred`, and `View` fields. Expandable strings retain their raw text. An explicit registry view remains
attached to the snapshot across PowerShell architectures and JSON serialization.

```pwsh
$settings = @(@{ Path = 'HKCU:\Software\Example'; Name = 'Enabled' })
$before = @(Export-RegistrySettingState -Settings $settings -Detailed)
# Apply your selected registry changes, then capture the state they produced.
$after = @(Export-RegistrySettingState -Settings $settings -Detailed)
Compare-RegistrySettingState -Settings $before | Format-List
Restore-RegistrySettingState -Settings $before -ExpectedState $after -WhatIf
```

Remove `-WhatIf` to restore selected values. `-ExpectedState` detects intervening edits and reports `Conflict`; without it, restoration
overwrites current values. The checks are optimistic, not an atomic registry transaction. Keys and unrelated values are retained, including
empty keys created while restoring values. Comparison also accepts desired settings with `Path`, `Name`, `Type`, and `Preferred`; use
`Exists = $false` to request absence. Restoration requires detailed snapshots, avoiding ambiguity in legacy null values.

### Prerequisites and operation outcomes

```pwsh
$report = Get-HostPrerequisiteReport -MinBuild 22000 -RequireAdministrator `
  -RequiredModules @{ Pester = '5.0.0' } -RequiredCommands 'winget.exe' `
  -RequiredServices @{ wuauserv = 'Running' }
$report.Checks | Where-Object { -not $_.Satisfied } | Format-Table Check, Target, Actual, Expected, Guidance
```

The report includes `Applicable` and all requested checks. It discovers prerequisites without remediation; `Test-HostApplicability` still
provides the original Boolean gate. Edition, OS architecture (including Arm64), module versions, commands, services, and elevation are
supported.

`New-OperationResult`, `Add-OperationResult`, and `New-PackageLifecycleResult` accept optional `Changed`, `AlreadyCompliant`, `Before`,
`After`, `ExitCode`, `RebootRequired`, `Duration`, and `RunId`. Omitted fields remain absent; a missing `Changed` means unknown. Use the
same `RunId` across a batch, or supply it to `Write-OperationResultLog` for entries without their own identifier. Package failures retain
their original `ErrorRecord` and add `ErrorTranslation` when recognized.

`Install-Win32Program` now checks exit codes even without `-PassThru`. Failed exits report `Failed`; success codes default to 0, 1641 and
3010, with 1641/3010 also setting `RebootRequired`. Override `-SuccessExitCodes` and `-RebootExitCodes` for installers with other
conventions. The legacy `-PassThru` status remains `ExitCode:n`, with a separate `Succeeded` flag. `-NoWait` now reports `Started` and
`ProcessId`, because completion is not yet known. Callers that previously treated every result as `Installed` should check these outcomes.

### Self-elevation

An entry-point script using `Request-AdministratorPrivilege` must declare a reserved `[switch]$Elevated` parameter and forward it through
`-IsElevatedRelaunch`. Pass the caller's `$PSBoundParameters` and `$args`. The helper preserves strings, arrays, booleans, numbers, nulls,
and explicit false switches across Windows PowerShell 5.1 and PowerShell 7. Other argument types and oversized command lines are rejected
before launching. Arguments are encoded, not encrypted: do not pass secrets on the command line. The helper exits the original process after
the elevated child finishes; it is intended for entry-point scripts.

### Contributing

Contributions are welcome via GitHub's Pull Requests. Fork the repository and implement your changes within the forked repository, after
that you may submit a [Pull Request][gh_pr_fork_docs]. Refer to our [documentation for contributors][contributing] for contributing
guidelines, commit message formats and versioning tips.

### Maintainers

This project is owned and maintained by [Ad Noctem Collective](https://github.com/adnoctem) refer to the [`AUTHORS`][authors] or
[`CODEOWNERS`][owners] for more information. You may also use the linked contact details to reach out directly.

### Copyright

_Assets provided by:_ **[Microsoft Corporation][microsoft]**

<!-- File references -->

[license]: LICENSE
[contributing]: docs/CONTRIBUTING.md
[authors]: .github/AUTHORS
[owners]: .github/CODEOWNERS

<!-- General links -->

[org]: https://github.com/adnoctem
[microsoft]: https://www.microsoft.com/
[powershell]: https://github.com/PowerShell/PowerShell
[gh_pr_fork_docs]:
  https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/proposing-changes-to-your-work-with-pull-requests/creating-a-pull-request-from-a-fork
[github_releases]: https://github.com/adnoctem/PSFoundation/releases
[github_commits]: https://github.com/adnoctem/PSFoundation/commits/main/
[psgallery_package]: https://www.powershellgallery.com/packages/PSFoundation
[testing_workflow]: https://github.com/adnoctem/PSFoundation/actions/workflows/testing.yaml

<!-- Third-party -->

[semantic_release]: https://semantic-release.org/
[renovate]: https://renovatebot.com/
[precommit]: https://pre-commit.com/
[superlinter_action]: https://github.com/marketplace/actions/super-linter
