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

| Capability             | Commands                                                                                                                                                                         |
| ---------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Tool provisioning      | `Resolve-OfficeDeploymentToolSource`, `Test-OfficeDeploymentToolSourceAvailability`, `Install-OfficeDeploymentTool`, `Test-OfficeDeploymentTool`, `Get-OfficeDeploymentToolHelp` |
| Discovery and intent   | `Get-OfficeInventory`, `New-OfficeDeploymentConfiguration`, `Get-OfficeDeploymentPlan`                                                                                           |
| Independent assessment | `Test-OfficeDeployment`, `Get-OfficeActivationStatus`, `Test-OfficeDeploymentMedia`                                                                                              |
| Media preparation      | `Save-OfficeDeploymentMedia`                                                                                                                                                     |
| Product lifecycle      | `Install-Office`, `Uninstall-Office`, `Switch-OfficeDeployment`                                                                                                                  |
| Maintenance            | `Update-Office`, `Set-OfficeUpdateConfiguration`, `Add-OfficeLanguage`, `Remove-OfficeLanguage`, `Set-OfficeApplicationSelection`, `Set-OfficeApplicationPreference`             |
| Recovery               | `Get-OfficeDeploymentRecovery`, `Resume-OfficeInstallation`, `Resume-OfficeMigration`                                                                                            |

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

ODT acquisition follows the same resolve/check/install separation as LGPO. `Resolve-OfficeDeploymentToolSource` returns a reviewed,
versioned Microsoft URL without network access; `Test-OfficeDeploymentToolSourceAvailability` probes it with HEAD. A reachable URL is not
signature validation. `Install-OfficeDeploymentTool` downloads the self-extracting package, checks its Microsoft Authenticode signature
before running extraction, validates the extracted `setup.exe`, and returns its validation result including `Path`, `Version` and `Valid`.
It provisions the tool only; it does not install Office. Run from an elevated session with an existing destination parent:

```powershell
Resolve-OfficeDeploymentToolSource
Test-OfficeDeploymentToolSourceAvailability
$tool = Install-OfficeDeploymentTool -Destination 'C:\Tools\ODT' -Confirm
Test-OfficeDeploymentTool -OdtPath $tool.Path
```

`-WhatIf` and `-DryRun` acquire nothing. An existing validated tool is reused; an untrusted existing file is rejected without replacement.
To acquire the reviewed build separately, select a new dedicated directory. Like LGPO source resolution, this is a reviewed source table,
not a live latest-version lookup. Deployment commands continue to require an explicit `OdtPath` and never silently acquire tools.

Both `Groove` and `OneDrive` remain valid ODT exclusion IDs. `Groove` targets the legacy OneDrive for Business sync client; use `OneDrive`
for the modern sync client. To exclude both during Office deployment, specify `-ExcludeApp Groove,OneDrive` consistently for configuration,
preparation and execution. Changing an existing `Groove`-only selection to both is an intentional application-selection change. See
[Microsoft's modern OneDrive example](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-office-deployment-tool#exclude-onedrive-when-installing-microsoft-365-apps-or-other-applications)
and [legacy client guidance](https://learn.microsoft.com/en-us/sharepoint/exclude-or-uninstall-previous-sync-client).

Default language is exactly `en-us`, independent of the operating system or account. Explicit lists preserve order; the first language is
the primary shell language. `-AutoSourceLocales` opts into installed-Office discovery; add `-LocaleSource OperatingSystem` to use the
machine installation UI language from `HKLM\SYSTEM\CurrentControlSet\Control\Nls\Language:InstallLanguage`. This is not the user's display
language, keyboard layout, or regional format. Ambiguous installed-Office language evidence is a blocker rather than an implicit fallback.
For multiple Click-to-Run products, automatic sourcing requires complete per-product language observations and one agreed primary language;
it preserves that primary first and includes the union of observed languages. Conflicting primaries, unknown inventory, and legacy MSI
sources require explicit language intent. Explicit `-Language` and `-AutoSourceLocales` are mutually exclusive. UI languages, language
interface packs, and proofing resources are distinct; this discovery does not establish complete MSI or proofing-resource coverage.

`Get-OfficeInventory` retains App Paths observations in `AppPathEvidence`: registry view/key, raw and resolved target, `State` and `Reason`.
Confirmed missing executable references do not by themselves count as Click-to-Run residue. Present or uncertain Click-to-Run references
still block when configuration is absent, as do genuine remaining Click-to-Run registry artifacts. Discovery never removes registry entries.
Path probing uses literal local fixed-drive paths, bounded environment expansion by registry view and ancestor checks; network paths,
reparse points, access failures and ambiguous targets cannot establish absence. Windows system paths and short-name aliases in 32-bit
PowerShell on x64 Windows remain uncertain to avoid redirection errors. Absolute Office installation paths are not remapped according to the
invoking process's bitness.

An ordinary migration journal at completed `Remove` can continue with `Resume-OfficeMigration` after fresh inventory establishes removal.
Recovery hashes its historical `Before` snapshot unchanged, revalidates current state and media, and intersects original removal IDs with
current products. When none remain, continuation installs without repeating removal and writes a new run/journal, preserving the original.
Use a fresh recovery descriptor after reloading an updated module; preview first and supply a fresh `SecureString` key for execution when
needed. This does not enable replay of uncertain partial installations or historical pilot journals.

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

Media metadata can be `Office/Data/v<architecture>.cab` or `Office/Data/v<architecture>_<pinned-version>.cab`, with the exact architecture
and manifest build. Neutral and every declared language stream remain required. Missing metadata or neutral content reports `MissingMedia`;
missing language content reports `MissingLanguageMedia`. The readable `Error` lists missing paths and, when metadata is absent, both CAB
alternatives. Failed preparation preserves that diagnostic while removing temporary downloads and leaving the destination unpublished. This
validates package structure and integrity; native local-source behavior still needs validation with the selected ODT and build. Execution
keeps an exact version and `AllowCdnFallback=FALSE`. Microsoft notes that absent generic CAB metadata can lead to CDN downloads in some
deployment flows; see
[package anatomy and installation options](https://learn.microsoft.com/en-us/microsoft-365-apps/best-practices/install-options).

Operation results have `SchemaVersion = 1` and include `RunId`, `Action`, `Phase`, `Status`, `ReasonCode`, `Before`, `After`,
`Verification`, `Configuration`, `LanguageTransition`, `Activation`, `NativeResults`, `RebootRequired`, `RecoveryRequired`, `RecoveryPath`,
`LogPaths`, and cleanup details. `Changed = $null` with `ChangeKnown = $false` means the outcome is uncertain, including after an invoked
installer fails. Installation verification and activation are independent. Wrappers can use `WrapperExitCode` (`0`, `1`, or `3010`) while
retaining native exit codes. Use `ConvertTo-Json -Depth 30` for the complete nested result. Do not flatten unknowns into successful
compliance.

Workflow results also include `Execution` (`ModuleVersion`, `ModulePath`, `PowerShellVersion`, `PowerShellEdition`, `ProcessBitness`) from
the executing module and process. Journal result/log metadata retains this context; an earlier preparation report does not establish the
version used for a later execution. This is additive within schema 1 and does not rewrite historical inventory or journal fingerprints.

`Diagnostic` adds safe structured detail where available. Media trust failures identify the object/path, owner or write-grant SID, rights,
and inheritance. Recovery validation failures identify the read/JSON/schema/identity/context/settings/fingerprint stage and safe error
category; direct reader errors expose this under `Exception.Data['OfficeDiagnostic']`, alongside the existing `OfficeReason`. Malformed
journal contents and raw parser exception text are not echoed. These diagnostics do not repair ACLs or relax validation. New plans retain
the assessment under optional `Media` metadata. Preparation, plan revalidation, staging and recovery preserve its safe diagnostic on
validation failure; a valid assessment with only a changed fingerprint does not become an ACL/JSON error. Older plans and journals without
`Media` remain readable. Recorded assessments are informational; execution and recovery always validate media afresh.

Recovery journals are written atomically under `%ProgramData%\PSFoundation-Office` by default. They contain no product keys. A key is
accepted only as `SecureString` and materialized in protected temporary XML for ODT; it is never placed on a process command line. Secure
erasure of storage and redaction of ODT's own logs cannot be guaranteed. Cleanup failures retain the original error and report protected
residue.

```powershell
$recovery = Get-OfficeDeploymentRecovery -RunId $result.RunId
$recovery | Resume-OfficeInstallation -OdtPath 'C:\Tools\ODT\setup.exe' -WhatIf
# Migration journals require Resume-OfficeMigration and fresh confirmation of remaining scope.
```

**Host requirements:** ordinary native execution accepts x64 Windows 10 22H2 (build 19045) and later Windows desktop hosts, including
Windows 11. Older Windows 10 builds, Windows Server and ARM hosts remain outside this backend's scope. Elevated 64-bit PowerShell, no
pending reboot, and no active deployment are still required. Product IDs are deployment identifiers, not a claim of current vendor lifecycle
support. Routine tests mock ODT and do not certify any real Office installation. Standalone MSI removal is unsupported. Recovery supports
verification of completed installations, pre-launch continuation, and migration continuation after verified Click-to-Run removal. Replaying
an uncertain partial installer returns `UnsupportedRecoveryState`; Quick Repair, Online Repair, rollback, and journal-free mutation are not
implemented.

ODT execution deliberately has no forced timeout or cancellation kill policy: `setup.exe` may delegate work to shared services. Tests of
generic process cancellation and the ODT download extractor do not establish safe cancellation of a live Office installation. Application
closure requires explicit consent. A held deployment lock or a changed inventory blocks execution before any installer launch.

| Recovery observation                                                      | Decision                                                                                  |
| ------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------- |
| Fresh state fully verifies the recorded destination, including activation | Complete verification without another ODT invocation                                      |
| Supported pre-launch checkpoint                                           | Revalidate authority, host, media and state; preview/continue the original operation      |
| Migration checkpoint after confirmed completed C2R removal                | Continue installation of the recorded destination without repeating removal               |
| Busy deployment or pending reboot                                         | Stop and wait/resolve the reported prerequisite                                           |
| Uncertain partial installation or unverified unsupported checkpoint       | `UnsupportedRecoveryState`; inspect native state/logs before a separately reviewed action |

These are decisions constrained by the recorded action and current gates, not permission to replay any failed journal.

Inventory resolves installed builds in evidence order:

1. Microsoft's documented `ClickToRun\Inventory\Office\16.0:OfficePackageVersion`, with matching `OfficeProductReleaseIds`.
2. When that key or version value is absent, agreeing `Version` values under
   `ClickToRun\ProductReleaseIDs\<ActiveConfiguration>\<ProductId>.16\<resource>`. The selected product must register `x-none` and at least
   one language; every listed resource must have exactly one valid four-part version, and all must agree in the same registry view.

`Products[].VersionSource` identifies `ClickToRunInventory` or `ActiveProductResources`; `Evidence` records the selected path or why the
version remains unknown. Resource versions are retained in `RegisteredResources`. Missing evidence can fall through; malformed documented
versions, conflicting product identities, and duplicate inventory records cannot. Inactive trees, another product's resources, shared
`culture` leaves alone, and `VersionToReport` cannot supply the fallback. Valid documented inventory takes precedence over resource data.

The fallback is based on Office Standard 2019 x86 and x64 observations whose resource and collected binary versions agree. It recognizes
that registry structure without a product-year allowlist; other releases still need native validation. It is a derived observation, not a
Microsoft-documented contract or proof of application health. No assumption is made that every pre-Microsoft-365 product lacks the inventory
key. Microsoft's
[installed-version guidance](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/microsoft-guidance-on-office-build-install)
explicitly excludes telemetry as installation-state evidence.

`Languages` is derived from the active per-product culture registrations. A single registered language supplies `PrimaryLanguage`; with
multiple registered languages, the current reader uses `ClientCulture` only when it belongs to that set. Missing or ambiguous observations
remain unknown. These inference rules still need broader native validation, especially bilingual installations. Per-user language
preferences and requested XML are not installation evidence. `VerificationLimitations` reflects unresolved language fields on observed
Click-to-Run products. A clean or MSI-only source can now reach an eligible ordinary plan; that is not a promise of verified
post-installation compliance. Installed build and application exclusions can still be unknown, and legacy MSI automatic locale sourcing is
not implemented. `VersionToReport` is not substituted for an absent installed-build observation. Update-policy and preference execution
returns `AppliedUnverified` until effective settings can be independently verified. Validate these scenarios on separately authorized
disposable VMs before production adoption; a mocked passing suite is not that evidence.

Inventory distinguishes known Click-to-Run infrastructure and known Office add-ins in `RelatedComponents` from legacy Office entries in
`Msi`. Orphaned Click-to-Run infrastructure remains an unknown state. Known add-in registrations must survive deployment; this check does
not establish add-in compatibility with the destination architecture. Unrecognized Office components remain subject to conservative
classification and blocking. Infrastructure classification requires a Microsoft publisher, an Office 16 product-code structure, and a
specific Click-to-Run Licensing, Extensibility or Localization component name. It does not rely on a list of component SKU fragments such as
`007E` or `008F`. Both have been observed for Licensing Components; neither is canonical. Actual product codes remain in inventory as
identities. Unrecognized names or other Office families are not automatically treated as infrastructure. `RegisteredLanguages` retains the
culture registrations used by language inference; `LanguageEvidence` preserves machine language observations separately from user
preferences.

These inventory fields are additive within schema 1. Existing plans must be recreated after inventory changes; execution revalidates current
observations. Existing recovery journals remain subject to their original authority and the current verification gates.

Per-product `ExcludedApps` registration is observed separately from requested configuration: an explicitly empty string yields an empty
`ExcludeApp` list; an absent, null or malformed value stays unknown. Valid strings are normalized for whitespace, case, order and
duplicates, with the selected registry value recorded in `Evidence`. Missing data does not prove every application is installed, and shared
executable presence alone cannot establish a SKU's application selection. No unsupported fallback is inferred from the requested XML.

### Migration and historical pilot records

Migrations use ordinary schema-1 plans and `Switch-OfficeDeployment`. The former `-PilotMigration` option and scenario-specific
host/resource assumptions have been removed. Select source Click-to-Run products explicitly with `-RemoveProductId`; use `-RemoveMsi` only
when authorizing supported legacy MSI removal. Missing observations still prevent verified compliance.

Historical schema-2 pilot journals remain inspectable with `Get-OfficeDeploymentRecovery`, but resume commands reject them with
`UnsupportedPilotRecovery`. They cannot be executed as ordinary plans or converted into new removal authority. A completed installation with
failed verification must not be replayed to change its reporting status.

An Office Enterprise 2007 x86 to Standard 2019 x64 migration on Windows 10 22H2 x64 completed with native exit 0 and licensed activation.
The operator verified German UI, the intended proofing resources, applications, add-ins, and Outlook profile/data/settings preservation.
Post-installation reporting exposed an infrastructure-classification defect and a remaining legacy File Validation Add-In registration;
native success and manual acceptance are therefore recorded separately from automated compliance. The Office 2019 captures lack the
installed-build inventory key. The pilot's missing exclusion evidence and the add-in's actual Windows Installer state remain separate
read-only follow-ups.

A subsequent Standard 2019 x86 to x64 migration is accepted through release 1.8.5: ordinary post-removal recovery completed the pinned
German 16.0.10417.20211 target, automated compliance and licensed activation passed, and the operator accepted applications, mail,
signatures and settings. An unchanged execution and preview both returned `Completed` / `AlreadyCompliant`, `Changed=false`,
`ChangeKnown=true`, exit 0 and no native results. This validates that migration and continuation; it does not establish CDN independence or
arbitrary interrupted-install replay. Identical migration requests with a now-absent, different source product ID also return the verified
no-op; that variation has imported-module regression coverage and still requires native validation for the next source SKU.

| Scenario                                                                                            | Evidence and remaining boundary                                                                    |
| --------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| Enterprise 2007 x86 to Standard 2019 x64, de-de, Windows 10 22H2 x64                                | Native migration and manual preservation accepted; pilot exclusions/add-in state remain unresolved |
| Standard 2019 x86 to x64, de-de, pinned .20211                                                      | Native continuation, compliance, activation, manual acceptance and unchanged repeat accepted       |
| Clean install, selected removal with retained C2R products, maintenance                             | Automated coverage; native shared-resource/preservation and effective-policy checks remain         |
| English/bilingual order, multiple C2R products, different source IDs, newer targets                 | Synthetic regression coverage where implemented; native product/language matrix remains            |
| SYSTEM execution, network shares, offline-only payload installation                                 | Requires dedicated native identity/access/network evidence                                         |
| Standalone MSI removal, uncertain partial-installer replay, Quick/Online Repair, automatic rollback | Unsupported                                                                                        |
| Outlook archive/optimization, legacy MAPI, targeted ScanPST                                         | Separate native workflows; successful Office migration or process exit is not repair proof         |

Tests import the module under strict mode for acceptance-level regressions. Sanitized registry observations and synthetic variants are
identified in [fixture provenance](tests/fixtures/office/README.md). Passing mocked tests does not expand the host or native acceptance
matrix.

Microsoft references: [ODT operations](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/overview-office-deployment-tool),
[configuration and language behavior](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options),
[installed-build inventory](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/microsoft-guidance-on-office-build-install), and
[MSI migration](https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/upgrade-from-msi-version).

### Existing Outlook PST sources

`Open-OutlookPstStore -Namespace $context.Namespace -LiteralPath <existing.pst>` reuses a matching attachment or temporarily attaches the
existing file. `Close-OutlookPstStore -Context $source` releases its root and removes only the attachment that Open created, identified by
StoreID and normalized file path. Existing attachments keep their names and remain attached. Neither helper releases the borrowed MAPI
namespace/application, changes the default store, or deletes, renames, repairs or compacts the file. Destination creation through
`Add-OutlookStoreRoot` remains unchanged and uses Unicode.

The live context has `Path`, `StoreId`, `DisplayName`, `Root`, `Namespace`, `AttachedByCall` and `Closed`, with type name
`PSFoundation.OutlookPstStoreContext`. Report only the scalar fields; do not serialize the context or modify its private `_State`. Release
all child folder/item references before closing. The creating caller must keep its context alive until overlapping users finish; these
contexts are not reference-counted attachment leases.

```powershell
$source = $null
try {
  $source = Open-OutlookPstStore -Namespace $context.Namespace -LiteralPath '.\Combined archive.pst'
  if ($null -ne $source) {
    # Use $source.Root with the existing folder-plan/transfer APIs.
    # Compare normalized source/destination paths and StoreIDs before transferring.
    $source | Select-Object Path, StoreId, DisplayName, AttachedByCall
  }
}
finally {
  if ($null -ne $source) { Close-OutlookPstStore -Context $source }
}
```

Open honors ShouldProcess: a suppressed/declined attachment produces no context. A pre-existing match still returns a context. Close is
required lifetime cleanup, so it performs no second confirmation and inherited WhatIf cannot suppress detachment/release. Close marks the
context closed and releases its root even when detachment fails; it throws an error naming the path, and the attachment may remain. A second
Close does nothing. Consumers should include cleanup failure in their reports without hiding the original processing error.

An explicitly documented wrapper preview may need a temporary attachment to inspect messages. Only that source-opening call may use
`-WhatIf:$false -Confirm:$false`, followed by Close in finally. Destination creation and message transfer retain preview suppression.
Outlook can update PST metadata during inspection, so this is not byte-preserving or offline parsing. User/profile/elevation checks and
source/destination separation remain consumer responsibilities.

Sources must be existing writable local PST files. Relative paths resolve against PowerShell's current location; matching ignores Windows
path case and expands short names. UNC/mapped network sources and reparse-point traversal are rejected; unreadable nonempty store paths
block inspection rather than imply absence. Hard-link aliases are not detected. Do not replace/delete sources or concurrently alter profile
attachments while a context is live: existence rechecks and observed StoreIDs cannot provide atomic file/profile guarantees.

The helpers use [AddStore](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.addstore) for existing files without selecting
a creation format, and [RemoveStore](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.removestore) only for profile
detachment. No format conversion is requested. Automated tests use synthetic files and mocked COM; native attach/reuse/detach, preview,
duplicate display names and existing ANSI/Unicode PSTs still need disposable-file validation. No splitting or transfer engine is included.

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
