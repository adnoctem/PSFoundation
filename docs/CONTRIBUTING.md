# Ad Noctem Collective - `PSFoundation` Repository Contributing Guidelines

Contributions are welcome via GitHub's Pull Requests. This document outlines the process to help get your contribution accepted.

## Building

The project uses the `PSFoundation.ps1` launcher script in the repository root to drive all development workflows. No external build tools (Make, CMake, etc.) are required — only PowerShell and the launcher script.

Before running anything else, you must initialize the project. This downloads the PowerShell module dependencies declared in
[`src/PSFoundation.psd1`](../src/PSFoundation.psd1):

```pwsh
.\PSFoundation.ps1 init
```

Additional aliases recognised for this command are `initialize`, `setup`, and `bootstrap`.

The launcher maps short, familiar command names to the scripts located in the [`tools/`](../tools) directory:

| Command   | Aliases                            | Target                 | Purpose                                             |
| --------- | ---------------------------------- | ---------------------- | --------------------------------------------------- |
| `init`    | `initialize`, `setup`, `bootstrap` | `tools/initialize.ps1` | Install required PowerShell modules                 |
| `format`  | `fmt`, `fix`                       | `tools/format.ps1`     | Format all PowerShell sources with PSScriptAnalyzer |
| `lint`    | `check`, `analyze`                 | `tools/lint.ps1`       | Run PSScriptAnalyzer rule checks                    |
| `build`   | `bundle`, `package`                | `tools/build.ps1`      | Create distribution archives                        |
| `test`    | `tests`, `pester`                  | `tools/test.ps1`       | Run Pester tests for the module                     |
| `release` | `publish`                          | `tools/release.ps1`    | Publish module to PowerShell Gallery                |

Any arguments supplied after the command are forwarded directly to the underlying script. For example, `.\PSFoundation.ps1 format -Check` is equivalent to running `.\tools\format.ps1 -Check`.

### Source Formatting

Format all PowerShell sources in the repository in place:

```pwsh
.\PSFoundation.ps1 format
```

To check formatting without modifying files, suitable for CI jobs and pre-commit hooks:

```pwsh
.\PSFoundation.ps1 format -Check
```

The formatter delegates whitespace, brace, indentation, and casing rules entirely to PSScriptAnalyzer via
[`PSScriptAnalyzerSettings.psd1`](../PSScriptAnalyzerSettings.psd1) and performs no repository-specific
post-processing beyond encoding and line-ending normalization on write.

Output defaults to:

- **Encoding**: UTF-8 with BOM (required for reliable parsing under Windows PowerShell 5.1)
- **Line endings**: CRLF
- **Indentation**: 2 spaces
- **Excluded directories**: `.git`, `.idea`, `dist`, `build`, `secrets` (unless `-IncludeSecrets` is supplied)

To limit the scope, pass explicit paths:

```pwsh
.\PSFoundation.ps1 format -Path ./src,./tests
```

### Linting

Run PSScriptAnalyzer against all PowerShell sources under the repository root (the `src/` and `tests/`
directories, the `tools/` scripts, and `PSFoundation.ps1`). The `.git`, `.idea`, `dist`, `build`, and `secrets`
directories are excluded by default:

```pwsh
.\PSFoundation.ps1 lint
```

Target specific files or directories:

```pwsh
.\PSFoundation.ps1 lint -Path ./src,./tests
```

The linter exits with code `1` when any analyzer findings remain, making it suitable for CI and pre-commit usage.
Findings are printed as a table listing the rule name, severity, file, line, and message.

### Building Distribution Archives

Create clean deployment archives containing the `src/` directory with its relative layout preserved:

```pwsh
.\PSFoundation.ps1 build
```

By default, both a `.zip` and a `.tar.gz` archive are written to the `dist/` directory. Existing archives with the
same names are overwritten.

Build only a specific format:

```pwsh
.\PSFoundation.ps1 build -Format Zip
```

Override the output directory or archive base name:

```pwsh
.\PSFoundation.ps1 build -OutputDirectory C:\Temp -Name PSFoundation-v0.1.0
```

### Pre-Commit Hooks

The repository ships a pre-configured [`.pre-commit-config.yaml`](../.pre-commit-config.yaml) that runs formatting and
linting checks automatically before each commit. After installing [pre-commit](https://pre-commit.com/), activate the
hooks from the repository root:

```pwsh
pre-commit install
```

The hooks invoke `.\PSFoundation.ps1 format -Check` and `.\PSFoundation.ps1 lint` with zero additional configuration beyond having
run `.\PSFoundation.ps1 init` to install the module dependencies.

### Running Tests

Run all Pester tests:

```pwsh
.\PSFoundation.ps1 test
```

Run a specific test file:

```pwsh
.\PSFoundation.ps1 test -Path .\tests\registry.Tests.ps1
```

Tests require Pester 5.0 or higher, which is installed automatically with `.\PSFoundation.ps1 init`. The test runner exits
with the number of failed tests as its exit code, making it suitable for CI pipelines.

### Publishing a Release

Build artifacts, generate SHA256 checksums, and publish the module to the PowerShell Gallery:

```pwsh
.\PSFoundation.ps1 release -Version 1.0.0 -NuGetApiKey $env:NUGET_API_KEY
```

Preview what would happen without making changes:

```pwsh
.\PSFoundation.ps1 release -Version 1.0.0 -DryRun
```

Skip the build step when archives are already present:

```pwsh
.\PSFoundation.ps1 release -Version 1.0.0 -NuGetApiKey $env:NUGET_API_KEY -SkipBuild
```

The release tool generates `dist/CHECKSUMS_SHA256.txt` for all built archives, suitable for CI artifact validation
and the semantic-release asset pipeline defined in [`.releaserc`](../.releaserc).

## Commit Message Format

This specification is inspired by and supersedes the **AngularJS commit message format**.

We have very precise rules over how our Git commit messages must be formatted.
This format leads to **easier to read commit history**.

Each commit message consists of a **header**, a **body**, and a **footer**.

```text
<header>
<BLANK LINE>
<body>
<BLANK LINE>
<footer>
```

The `header` is mandatory and must conform to the [Commit Message Header](#commit-header) format.

The `body` is mandatory for all commits except for those of type "docs".
When the body is present it must be at least 20 characters long and must conform to
the [Commit Message Body](#commit-body) format.

The `footer` is optional. The [Commit Message Footer](#commit-footer) format describes what the footer is used for and
the structure it must have.

### <a name="commit-header"></a>Commit Message Header

```text
<type>(<scope>): <short summary>
  │       │             │
  │       │             └─⫸ Summary in present tense. Not capitalized. No period at the end.
  │       │
  │       └─⫸ Commit Scope: src|tools|tests|config|docs
  │
  └─⫸ Commit Type: build|ci|docs|feat|fix|perf|refactor|test|chore
```

The `<type>` and `<summary>` fields are mandatory, the `(<scope>)` field is optional.

#### Type

Must be one of the following:

- **feat**: New features
- **fix**: Bugfixes
- **docs**: Documentation changes
- **refactor**: Code changes which neither add features nor fix bugs
- **test**: Adding tests or improving upon existing tests
- **chore**: Miscellaneous maintenance tasks which can generally be ignored
- **build**: Changes or improvements to the build tool or to the project's dependencies (_supported Scopes_: `tools`)
- **ci**: Changes to CI configuration files and scripts (_supported Scopes_: `actions`)

#### Scopes

The following is the list of supported scopes:

- `src` — Changes to the module source code (`src/`)
- `tools` — Changes to development tooling (`tools/`)
- `tests` — Changes to the test suite (`tests/`)
- `config` — Changes to configuration files (`.editorconfig`, `.gitattributes`, `.pre-commit-config.yaml`, etc.)
- `docs` — Documentation changes (`README.md`, `docs/`, `CONTRIBUTING.md`)

#### Summary

Use the summary field to provide a succinct description of the change:

- use the imperative, present tense: "change" not "changed" nor "changes"
- don't capitalize the first letter
- no dot (.) at the end

#### <a name="commit-body"></a>Commit Message Body

Just as in the summary, use the imperative, present tense: "fix" not "fixed" nor "fixes".

Explain the motivation for the change in the commit message body. This commit message should explain _why_ you are
making the change.
You can include a comparison of the previous behavior with the new behavior in order to illustrate the impact of the
change.

#### <a name="commit-footer"></a>Commit Message Footer

The footer can contain information about breaking changes and deprecations and is also the place to reference GitHub
issues, Jira tickets, and other PRs that this commit closes or is related to.
For example:

```text
BREAKING CHANGE: <breaking change summary>
<BLANK LINE>
<breaking change description + migration instructions>
<BLANK LINE>
<BLANK LINE>
Fixes #<issue number>
```

or

```text
DEPRECATED: <what is deprecated>
<BLANK LINE>
<deprecation description + recommended update path>
<BLANK LINE>
<BLANK LINE>
Closes #<pr number>
```

Breaking Change section should start with the phrase "BREAKING CHANGE: " followed by a summary of the breaking change, a
blank line, and a detailed description of the breaking change that also includes migration instructions.

Similarly, a Deprecation section should start with "DEPRECATED: " followed by a short description of what is deprecated,
a blank line, and a detailed description of the deprecation that also mentions the recommended update path.

#### Revert commits

If the commit reverts a previous commit, it should begin with `revert:`, followed by the header of the reverted commit.

The content of the commit message body should contain:

- information about the SHA of the commit being reverted in the following format: `This reverts commit <SHA>`,
- a clear description of the reason for reverting the commit message.

## How to Contribute

1. Fork this repository, develop, and test your changes
2. Run `.\PSFoundation.ps1 format` and `.\PSFoundation.ps1 lint` to ensure your changes pass all checks
3. Add your GitHub username to the [`AUTHORS`](../.github/AUTHORS) and [`CODEOWNERS`](../.github/CODEOWNERS) files
4. Submit a pull request

_**NOTE**_: In order to make testing and merging of PRs easier, please submit changes to unrelated areas of the
repository in separate PRs.

### Technical Requirements

- Must target PowerShell 5.0 or higher (every script must include `#Requires -Version 5.0`)
- Must pass `.\PSFoundation.ps1 format -Check` with no formatting drift
- Must pass `.\PSFoundation.ps1 lint` with zero PSScriptAnalyzer findings
- Module functions under `src/` must include comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`)
- New or updated functions must have corresponding Pester tests in `tests/`

### Versioning

The `PSFoundation` PowerShell module declared in [`src/PSFoundation.psd1`](../src/PSFoundation.psd1) follows [SemVer](https://semver.org/).

Any change to the module source (`src/`), the module manifest, or a change that alters the public API surface of the
module requires a version bump in the manifest. Documentation-only changes do not require a bump.

Breaking (backwards incompatible) changes to the module must:

1. Bump the MAJOR version in the module manifest
2. Describe the breaking change and migration instructions in the commit message footer
3. Update the module release notes if they exist

New features and non-breaking enhancements bump the MINOR version. Bugfixes and documentation bumps increment the
PATCH version.
