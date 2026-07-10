<p align="center">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/Powershell_256.png">
      <img src="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/Powershell_256.png" width="225">
    </picture>
    <h1 align="center">PSFoundation</h1>
</p>

[![License](https://img.shields.io/github/license/adnoctem/PSFoundation?label=License)][license]
[![Language](https://img.shields.io/github/languages/top/adnoctem/PSFoundation?label=PowerShell)][powershell]
[![PSGallery Version](https://img.shields.io/powershellgallery/v/PSFoundation)][psgallery_package]
[![GitHub Release](https://img.shields.io/github/v/release/adnoctem/PSFoundation?label=Release)][github_releases]
[![GitHub Activity](https://img.shields.io/github/commit-activity/m/adnoctem/PSFoundation?label=Commits)][github_commits]
[![Semantic Release](https://img.shields.io/badge/Semantic_Release-enabled-brightgreen?logo=semanticrelease&logoColor=E5E4E7)][semantic_release]
[![Renovate](https://img.shields.io/badge/Renovate-enabled-brightgreen?logo=renovate&logoColor=1A1F6C)][renovate]
[![PreCommit](https://img.shields.io/badge/PreCommit-enabled-brightgreen?logo=precommit&logoColor=FAB040)][precommit]

`PSFoundation` is an open-source [MIT][license]-licensed [PowerShell][powershell] module library written and maintained by the [Ad Noctem Collective][org] for Windows system administration, configuration management, and automation. The module targets both desktop Windows installations and Windows Server environments and supports [PowerShell][powershell] 5.1 and above, including Windows PowerShell 5.1 as well as newer PowerShell 7+ releases. It is published to the [PowerShell Gallery][psgallery_package] for easy discovery and installation.

The [`src`](src) directory contains the module source code — a collection of PowerShell functions organized by domain (registry, networking, security, packages, system, etc.) — bundled together as a single importable module. The [`tools`](tools) directory contains the repository's development tooling for building, formatting, linting, testing, and publishing the module.

### Module Coverage

PSFoundation provides functions across these domains:

| Module File       | Domain                                                           |
| ----------------- | ---------------------------------------------------------------- |
| `common.ps1`      | Operation result helpers and registry setting state management   |
| `data.ps1`        | Data transformation utilities (quote conversion, object merging) |
| `devices.ps1`     | Print and scan device enumeration and management                 |
| `interop.ps1`     | COM interop and Outlook automation                               |
| `log.ps1`         | Console logging helpers                                          |
| `networking.ps1`  | IP validation, network adapter resolution, address calculation   |
| `packages.ps1`    | Win32 and AppX package lifecycle management                      |
| `permissions.ps1` | Elevation detection and privilege requests                       |
| `policies.ps1`    | LGPO (Local Group Policy Object) integration                     |
| `registry.ps1`    | Registry key and value CRUD with path resolution                 |
| `security.ps1`    | Defender, firewall, event log analysis, security auditing        |
| `settings.ps1`    | Default application associations                                 |
| `system.ps1`      | OS version, memory, disk, uptime, and hostname queries           |
| `updates.ps1`     | Windows Update and Microsoft Store update management             |
| `user.ps1`        | User and SID information retrieval                               |

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

### Contributing

Contributions are welcome via GitHub's Pull Requests. Fork the repository and implement your changes within the forked repository, after that you may submit a [Pull Request][gh_pr_fork_docs]. Refer to our [documentation for contributors][contributing] for contributing guidelines, commit message formats and versioning tips.

### Maintainers

This project is owned and maintained by [Ad Noctem Collective](https://github.com/adnoctem) refer to the [`AUTHORS`][authors] or [`CODEOWNERS`][owners] for more information. You may also use the linked contact details to reach out directly.

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
[powershell_docs]: https://learn.microsoft.com/de-de/powershell/
[gh_pr_fork_docs]: https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/proposing-changes-to-your-work-with-pull-requests/creating-a-pull-request-from-a-fork
[github_releases]: https://github.com/adnoctem/PSFoundation/releases
[github_commits]: https://github.com/adnoctem/PSFoundation/commits/main/
[psgallery_package]: https://www.powershellgallery.com/packages/PSFoundation

<!-- Third-party -->

[semantic_release]: https://semantic-release.org/
[renovate]: https://renovatebot.com/
[precommit]: https://pre-commit.com/
