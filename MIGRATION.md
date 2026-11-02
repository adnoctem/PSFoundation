# Migrating from PSFoundation to substrate

> [!WARNING]
>
> **PSFoundation is frozen and no longer maintained.** No further features, bug fixes, security updates, or releases will be provided.
> Continue with [adnoctem/substrate](https://github.com/adnoctem/substrate), the open-source successor.

## Moving existing PowerShell scripts

Current substrate versions provide a one-to-one match for PSFoundation's PowerShell API through `AdNoctem.Substrate.PowerShell`. The adapter
preserves the original command surface, including its 188 commands and four aliases, while the successor can add further APIs. Migration
centers on updating the module dependency and import; existing command usage can remain in place.

The module is available on [PowerShell Gallery](https://www.powershellgallery.com/packages/AdNoctem.Substrate.PowerShell), and the package
is also published on [NuGet](https://www.nuget.org/packages/AdNoctem.Substrate.PowerShell). For script use, install and import the module
through PowerShell Gallery. The following example pins substrate 0.2.1:

```powershell
Install-Module -Name AdNoctem.Substrate.PowerShell -RequiredVersion 0.2.1 -Repository PSGallery -Scope CurrentUser
Import-Module -Name AdNoctem.Substrate.PowerShell -RequiredVersion 0.2.1
```

Replace `Import-Module PSFoundation` and any PSFoundation module requirement in scripts, manifests, deployment configuration, or CI with the
successor's identity. Run existing script checks in a fresh PowerShell process to verify the migrated dependency. The successor supports
Windows PowerShell 5.1 and PowerShell 7 on x64 Windows; consult its documentation for operation-specific requirements.

Substrate's `0.x` releases are still evolving. Pin the version you adopt and review its release notes before updating. The compatibility
statement here describes the PowerShell API at PSFoundation's freeze, rather than a guarantee about every future substrate release.

## What this repository preserves

This archive retains PSFoundation's legacy PowerShell implementation. Release 1.8.7 is the compatibility baseline used by the successor. The
existing source, tests, and historical documentation remain useful for understanding legacy behavior.

The C# rewrite used a temporary PSFoundation v2 staging branch to validate its implementation against the legacy tests. It subsequently
became the independent substrate project. That staging work was not released as a PSFoundation v2 package. Its C# projects, generated
reference pages, and build tooling are not part of this legacy archive.

This document consolidates the useful migration context from that staging work. Incremental checkpoint logs and temporary implementation
plans are superseded by the successor's maintained documentation.

## Why the rewrite became a separate project

PSFoundation began as shared Windows administration functions used by [winkit](https://github.com/adnoctem/winkit). The rewrite made that
behavior available to applications through reusable C# domain libraries, with PowerShell as one adapter.

- Typed requests, values, and snapshots make ownership, missing data, native errors, and partial completion explicit.
- Shared registry, process, filesystem, trust, and diagnostic foundations reduce duplicated administration logic.
- Libraries keep application policy, credentials, prompts, logging destinations, and elevation with the caller.
- Native handles, processes, and COM resources have explicit lifetimes. Cancellation has documented limits and does not imply rollback.
- C# XML comments and reviewed PowerShell help feed generated reference documentation and packaged `Get-Help`.

The adapter retains small compatibility functions where PowerShell binding or presentation differs from compiled cmdlets. Domain behavior
belongs in the C# libraries. The native Windows capabilities still require Windows and the relevant permissions and components.

## Compatibility and persisted state

PowerShell command compatibility does not make the reusable C# assemblies binary-compatible with PSFoundation types. Direct .NET consumers
should use the successor's `AdNoctem.Substrate.*` packages and namespaces, following its current library documentation.

Some PSFoundation identifiers intentionally remain in substrate because they identify persisted state or compatibility objects. Examples
include registry-policy entry type names, Outlook context tags, the Office deployment mutex, and existing Office storage and journal
formats. These names do not indicate that the old module must remain installed. Changing persisted formats or identifiers requires a
deliberate state migration.

Keep existing deployment records when migrating. Follow the successor's documented recovery and validation rules before using them;
historical plans, journals, or pilot results do not grant new execution authority.

## Verification and further guidance

The rewrite was compared with the frozen PowerShell implementation using managed tests and fresh PowerShell hosts. Those comparisons
establish specific compatibility evidence; they do not by themselves establish successful live Office deployment, Windows Update servicing,
or every remote administrative workflow. Use the successor's current documentation for supported behavior and validation limits.

Continue with:

- [substrate repository and releases](https://github.com/adnoctem/substrate)
- [PowerShell installation and workflows](https://github.com/adnoctem/substrate/blob/main/docs/www/guides/powershell.md)
- [Reusable library guide](https://github.com/adnoctem/substrate/blob/main/docs/www/guides/libraries.md)
- [Migration history and design rationale](https://github.com/adnoctem/substrate/blob/main/docs/MIGRATION.md)
- [Contributor and development workflow](https://github.com/adnoctem/substrate/blob/main/docs/CONTRIBUTING.md)
