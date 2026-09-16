# AGENTS.md - PSFoundation

`PSFoundation` is a PowerShell module library for Windows administration (registry, networking, security, packages, system, ...) supporting
Windows PowerShell 5.1 and PowerShell 7+. `src/` is the module source, `tools/` the dev tooling, `tests/` the Pester suite. See
`docs/CONTRIBUTING.md` for full details.

## All dev tasks go through the launcher

```pwsh
.\PSFoundation.ps1 <command>   # forwards to tools/<command>.ps1, passes remaining args through
```

| Command   | Aliases                            | Purpose                                                                                              |
| --------- | ---------------------------------- | ---------------------------------------------------------------------------------------------------- |
| `init`    | `initialize`, `setup`, `bootstrap` | Install module deps (`src/PSFoundation.psd1`) + dev deps (`tools/dev-dependencies.json`). Run first. |
| `format`  | `fmt`, `fix`                       | Format all PowerShell sources with PSScriptAnalyzer. `-Check` verifies without writing.              |
| `lint`    | `check`, `analyze`                 | PSScriptAnalyzer rule check. **Exits 1 on any finding** (CI + pre-commit rely on this).              |
| `build`   | `bundle`, `package`                | Create `dist/` archives (zip + tar.gz).                                                              |
| `test`    | `tests`, `pester`                  | Run Pester 5 tests; exits with the failed count. `-Path` limits to one file.                         |
| `deps`    | `dependencies`                     | Check/update PowerShell Gallery dependencies.                                                        |
| `release` | `publish`                          | Publish to PSGallery. Dry-run: `-DryRun`.                                                            |

Verification order: `.\PSFoundation.ps1 format -Check` then `lint` then `test`.

Before preparing a commit message or handing changes back, run `pre-commit run --all-files`. Resolve findings and rerun after any hook
modifies files so the final check passes on the proposed changes.

## Adding a public function

- Extend an existing domain file and its matching test file when they are a suitable home for new functions. Create a new source file only
  when no existing domain fits. Keep repository filenames free of spaces; use temporary test paths when testing whitespace handling.
- Every `*.ps1` in `src/` is dot-sourced automatically by `src/PSFoundation.psm1` (except `common.ps1`, sourced first) — no manual wiring.
- Register the function in **both** `$publicFunctions` in `src/PSFoundation.psm1` **and** `FunctionsToExport` in `src/PSFoundation.psd1`
  (aliases likewise: `$publicAliases` / `AliasesToExport`). Missing either means the function is not exported.
- Add comment-based help (`.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`) and a matching `tests/<domain>.Tests.ps1` Pester test.

## Conventions & gotchas

- Formatting is PSScriptAnalyzer-driven: 2-space indent, UTF-8 **with BOM**, CRLF; `format` normalizes encoding/line endings on write. UTF-8
  BOM is required for Windows PowerShell 5.1 parsing.
- `format`/`lint` exclude `.git`, `.idea`, `dist`, `build`, `secrets`. `secrets/` is gitignored — never commit its contents.
- Target PowerShell 5.1+; avoid PS7-only syntax. New `src/` scripts start with `#Requires -Version 5.0`; tests use `#Requires -Version 5.1`
  plus the Pester 5 module requirement.
- Commits: conventional commits `type(scope): summary`; types `feat|fix|docs|refactor|test|chore|build|ci`, scopes
  `src|tools|tests|config|docs`. Body mandatory (>= 20 chars) except for `docs` commits.
- Never bump `ModuleVersion` in `src/PSFoundation.psd1` manually — semantic-release (`.releaserc`, branches `main`/`next`) writes it via
  `tools/release.ps1 -Prepare`.
- pre-commit hooks (`.pre-commit-config.yaml`) run format/lint and prettier on Markdown (prettier runs via `bun`); install once with
  `pre-commit install`.

## Security review rules (non-negotiable)

- Never commit secrets, credentials, tokens, or private machine configuration. Keep test fixtures synthetic or sanitized. Do not log secrets
  or pass them in process arguments; encoding is not encryption.
- Verify downloaded executable content against an independently trusted hash or signature before execution. Do not disable certificate
  validation or bypass integrity failures to make a download succeed.
- Do not use `Invoke-Expression` on string-built commands. Keep executable code separate from data; when crossing process boundaries,
  constrain supported value types and quote all user-controlled values.

## Layout

- `src/<domain>.ps1` — module function library (common, registry, networking, security, ...) plus `PSFoundation.psd1`/`.psm1`
- `tests/<domain>.Tests.ps1` — Pester 5 tests mirroring `src/` per domain
- `tools/` — dev scripts behind the launcher, `dev-dependencies.json`
- `dist/`, `build/` — gitignored build output
- `docs/` — project docs. **This file lives at `docs/AGENTS.md`; the root `AGENTS.md` is a symlink to it — edit `docs/AGENTS.md`.**
