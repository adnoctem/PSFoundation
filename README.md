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
| `office.ps1`       | Office deployment planning, verified media, scoped ODT operations, and protected recovery records                              |
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

### Office deployment API

Office commands separate discovery, media preparation, installation, removal, migration, and maintenance authority. A plan is a reviewable
snapshot; execution rechecks the machine and media. These APIs are intended for thin orchestration wrappers such as winkit's Office scripts.

| Capability             | Commands                                                                                                                                                             |
| ---------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Tool provisioning      | `Resolve-OfficeDeploymentToolSource`, `Install-OfficeDeploymentTool`, `Test-OfficeDeploymentTool`, `Get-OfficeDeploymentToolHelp`                                    |
| Discovery and intent   | `Get-OfficeInventory`, `New-OfficeDeploymentConfiguration`, `Get-OfficeDeploymentPlan`                                                                               |
| Independent assessment | `Test-OfficeDeployment`, `Get-OfficeActivationStatus`, `Test-OfficeDeploymentMedia`                                                                                  |
| Media preparation      | `Save-OfficeDeploymentMedia`                                                                                                                                         |
| Product lifecycle      | `Install-Office`, `Uninstall-Office`, `Switch-OfficeDeployment`                                                                                                      |
| Maintenance            | `Update-Office`, `Set-OfficeUpdateConfiguration`, `Add-OfficeLanguage`, `Remove-OfficeLanguage`, `Set-OfficeApplicationSelection`, `Set-OfficeApplicationPreference` |
| Recovery               | `Get-OfficeDeploymentRecovery`, `Resume-OfficeInstallation`, `Resume-OfficeMigration`                                                                                |

```powershell
$configuration = New-OfficeDeploymentConfiguration `
  -TargetProductId Standard2024Volume `
  -Architecture 64 `
  -Language en-us,de-de

# Prepare is the only deployment operation that downloads Office payloads.
# The destination's parent must exist; a package is published only after verification.
$media = Save-OfficeDeploymentMedia `
  -Configuration $configuration `
  -SourcePath 'C:\Deployment\Office2024' `
  -OdtPath 'C:\Tools\ODT\setup.exe' `
  -Confirm

$plan = Get-OfficeDeploymentPlan `
  -Action Install `
  -Configuration $configuration `
  -SourcePath $media.Path

$plan | Format-List Action, State, Eligible, Blockers, Warnings, LanguageTransition
$plan | Install-Office -OdtPath 'C:\Tools\ODT\setup.exe' -WhatIf
```

Default language is exactly `en-us`, independent of the operating system or account. Explicit lists preserve order; the first language is
the primary shell language. `-AutoSourceLocales` opts into installed-Office discovery; add `-LocaleSource OperatingSystem` to use the
machine installation UI language from `HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language:InstallLanguage`. This is not the user's display
language, keyboard layout, or regional format. Ambiguous installed-Office language evidence is a blocker rather than an implicit fallback.

Executors accept only their matching plan action. `Uninstall-Office` requires exact `RemoveProductId` selections. `Switch-OfficeDeployment`
alone combines selected removal and installation, including explicitly authorized broad MSI removal. Updates preserve other deployment
dimensions; language operations preserve the primary language. An intentional primary-language replacement belongs to migration.
`Set-OfficeUpdateConfiguration` accepts explicit `Enabled`, `UpdatePath`, `TargetVersion`, and `Channel` settings. Deadlines are
deliberately excluded because they can forcibly close applications later. Application preferences accept validated Office
`REG_SZ`/`REG_DWORD` records under `Settings.Preferences`, with `Key`, `Name`, `Value`, `Type`, `App`, and `Id`; their scope includes
existing and future users.

All mutation commands support `-WhatIf`; deployment commands also support `-DryRun`. Previews create no files, journals, or logs and never
start an installer or stop applications. A compliant no-op returns `AlreadyCompliant`; an absent removal selection returns `AlreadyAbsent`.
`-Confirm:$false` acknowledges the displayed scope but does not disable validation or language warnings. The module never exits the host,
reboots it, invokes registry uninstall strings, or accepts arbitrary ODT XML or command-line arguments.

Media schema 2 records the pinned build, product, channel, architecture, available languages, tool version, and complete payload hashes.
Deployment language selection can be a subset of the available languages. Old `winkit-office-media.json` schema-1 packages require explicit
preparation into a new directory. Files and manifests require Administrators/SYSTEM ownership and write access; hashes do not authenticate
an attacker-replaced manifest. Local/UNC media must be accessible to the actual execution identity. Recovery records must remain local.

Operation results have `SchemaVersion = 1` and include `RunId`, `Action`, `Phase`, `Status`, `ReasonCode`, `Before`, `After`,
`Verification`, `Configuration`, `LanguageTransition`, `Activation`, `NativeResults`, `RebootRequired`, `RecoveryRequired`, `RecoveryPath`,
`LogPaths`, and cleanup details. `Changed = $null` with `ChangeKnown = $false` means the outcome is uncertain, including after an invoked
installer fails. Installation verification and activation are independent. Wrappers can use `WrapperExitCode` (`0`, `1`, or `3010`) while
retaining native exit codes. Use `ConvertTo-Json -Depth 30` for the complete nested result. Do not flatten unknowns into successful
compliance.

Recovery journals are written atomically under `%ProgramData%\PSFoundation-Office` by default. They contain no product keys. A key is
accepted only as `SecureString` and materialized in protected temporary XML for ODT; it is never placed on a process command line. Secure
erasure of storage and redaction of ODT's own logs cannot be guaranteed. Cleanup failures retain the original error and report protected
residue.

```powershell
$recovery = Get-OfficeDeploymentRecovery -RunId $result.RunId
$recovery | Resume-OfficeInstallation -OdtPath 'C:\Tools\ODT\setup.exe' -WhatIf
# Migration journals require Resume-OfficeMigration and fresh confirmation of remaining scope.
```

**Current validation limits:** ordinary native execution is restricted to x64 Windows 11 desktop. The explicitly scoped pilot below is the
only Windows 10 exception. Product IDs are deployment identifiers, not a claim of current vendor lifecycle support. Routine tests mock ODT
and do not certify any real Office installation. Standalone MSI removal is unsupported. Recovery supports verification of completed
installations, pre-launch continuation, and migration continuation after verified Click-to-Run removal. Replaying an uncertain partial
installer returns `UnsupportedRecoveryState`; Quick Repair, Online Repair, rollback, and journal-free mutation are not implemented.

Inventory uses Microsoft's documented `ClickToRun\Inventory\Office\16.0` product/build values and preserves incomplete registration and
resource evidence. It deliberately does **not** promote `VersionToReport`, `ClientCulture`, or per-user language preferences to proof of
complete installed languages or primary shell language. The native language/primary-language verification backend remains a validation gate:
these fields currently remain unknown, preventing full compliance and maintenance that depends on preserving them. Native inventory exposes
`VerificationLimitations`; installation and migration plans return `UnsupportedNativeVerification` before mutation when the backend cannot
verify these required postconditions. Installed-Office automatic locale preservation therefore remains blocked on native observations.
Update-policy and preference execution returns `AppliedUnverified` until effective settings can be independently verified. Validate these
scenarios on separately authorized disposable VMs before production adoption; a mocked passing suite is not that evidence.

Inventory distinguishes known Click-to-Run infrastructure and known Office add-ins in `RelatedComponents` from legacy Office entries in
`Msi`. Orphaned Click-to-Run infrastructure remains an unknown state. Known add-in registrations must survive deployment; this check does
not establish add-in compatibility with the destination architecture. Unrecognized Office components remain subject to conservative
classification and blocking. `RegisteredLanguages` contains candidates from the active product registration, and `LanguageEvidence`
preserves machine language observations. Neither field establishes complete installed languages or primary shell language.

These inventory fields are additive within schema 1. Existing plans must be recreated after inventory changes; execution revalidates current
observations. Existing recovery journals remain subject to their original authority and the current verification gates.

### Office 2007 migration pilot

`Get-OfficeDeploymentPlan -PilotMigration` and `Switch-OfficeDeployment -PilotMigration` require separate explicit consent for a narrow
pilot: x64 Windows 10 desktop build 19045, the observed Office Enterprise 2007 MSI suite/resources, and German Standard 2019 volume x64.
Select `-Language de-de` explicitly and authorize broad removal with `-RemoveMsi`. Automatic installed-language discovery remains strict.
The plan checks the current host, source registrations, German language evidence, and German/English/French/Italian proofing registrations.
Other destinations, source products, hosts, unknown inventory, absent consent, and unrelated verification limitations remain blocked.

The reviewed resource intent is stored in `Plan.PilotResources`, separately from full UI languages. German media supplies the intended
German UI and companion proofing; it does not request English/French/Italian UI packs or use runtime `MatchPreviousMSI`. Microsoft's
[companion-language table](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-deploying-languages-microsoft-365-apps#companion-proofing-languages)
lists those four proofing languages for German. Applying that companion set to the selected 2019 package is a pilot assumption to verify on
the target, not proof of installed resources. See also Microsoft's
[Office 2019 language deployment](https://learn.microsoft.com/en-us/office/2019/deploy#deploy-languages-for-office-2019).

```powershell
$target = New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 64 -Language de-de
# Prepare media explicitly with Save-OfficeDeploymentMedia before planning.
$plan = Get-OfficeDeploymentPlan -Action Migrate -Configuration $target -SourcePath C:\Media\Office2019 -RemoveMsi -PilotMigration
$plan | ConvertTo-Json -Depth 30
$plan | Switch-OfficeDeployment -OdtPath C:\ODT\setup.exe -PilotMigration -WhatIf
# Execute only after reviewing the plan and establishing the VM rollback point.
$result = $plan | Switch-OfficeDeployment -OdtPath C:\ODT\setup.exe -PilotMigration -Confirm
```

Normal signature, media identity/build/hash, staging, disk-space, lock, fresh-inventory, application and pending-reboot checks still apply.
Successful native execution with unresolved observations returns `AppliedUnverified`, `PilotVerificationRequired`, and wrapper exit **1**.
Proofing verification always remains manual in this pilot. Native 3010 and `RebootRequired` are preserved separately; unverified exit 1
takes precedence over 3010. Known mismatches, unexpected inventory, and native errors remain failures. Activation is reported separately.
Never interpret an unverified result as a retry instruction or fleet readiness.

Pilot plans/journals use schema **2**; ordinary plans/journals remain schema 1, and result/configuration schemas are unchanged.
`Get-OfficeDeploymentRecovery` can inspect pilot evidence, but resume commands reject it with `UnsupportedPilotRecovery`. Old plans cannot
gain pilot authority by adding fields. Save the result, protected journal/JSONL, collector report, and relevant ODT logs off the VM before
reverting its snapshot. ODT logs may contain sensitive data. Verify German UI, all four proofing languages, build, x64 apps, activation,
retained add-ins, representative documents and the user's Outlook profile manually. No actual Office migration is exercised by unit tests.

The current migration validation target is Office Enterprise 2007 to Standard 2019, 64-bit, with language preservation on Windows 10 x64.
The Windows 10 execution path, language/proofing preservation, and interrupted-installer recovery still require validation; the execution
host gate remains Windows 11 x64. The Office 2019 reference captures do not contain the installed-build inventory key. Their telemetry build
is not substituted as proof. This is a validation target, not a completed deployment capability or a vendor lifecycle support statement.

Microsoft references: [ODT operations](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-office-deployment-tool),
[configuration and language behavior](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options),
[installed-build inventory](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/microsoft-guidance-on-office-build-install), and
[MSI migration](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/upgrade-from-msi-version).

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
