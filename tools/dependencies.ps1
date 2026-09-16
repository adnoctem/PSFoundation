#Requires -Version 5.0

<#
.SYNOPSIS
  Checks declared PowerShell module dependencies against the PowerShell Gallery
  and opens a pull request with the version updates.

.DESCRIPTION
  Reads the runtime dependencies declared in src/PSFoundation.psd1
  (RequiredModules) and the dev dependencies declared in
  tools/dev-dependencies.json, queries the PowerShell Gallery for the latest
  published stable version of each module, and reports which declared versions
  are out of date.

  Exact RequiredVersion pins take precedence over minimum versions. Updates
  are scoped to the dependency declaration, and both proposed files are parsed
  and checked before either file is written.

  In update mode (the default) the declared versions are bumped in place and a
  pull request is opened against the default branch with the version diffs. The
  pull request is picked up by the repository's regular CI on its own. No other
  workflows are dispatched.

  In -Check mode no files are modified and the script exits with code 1 when any
  dependency is outdated, making it suitable for local use and CI checks.

  The gallery is queried through the NuGet v2 OData endpoint. When the
  PSGALLERY_API_KEY environment variable is set it is sent as the X-NuGet-ApiKey
  header to avoid rate limiting. The gh CLI and the git push use the GitHub
  token taken from the conventional environment variables GH_TOKEN,
  GITHUB_TOKEN, GH_ENTERPRISE_TOKEN, or GITHUB_ENTERPRISE_TOKEN, in that order
  of precedence.

  The script is intended to run from the scheduled dependencies.yaml workflow
  on a clean checkout of the default branch.

.PARAMETER Check
  Report outdated dependencies without modifying files. Exits with code 1 when
  any dependency is outdated.

.EXAMPLE
  PS> ./dependencies.ps1 -Check
  Reports outdated dependencies without making any changes.

.EXAMPLE
  PS> ./dependencies.ps1
  Bumps declared versions and opens a pull request with the diffs.

.LINK
  https://github.com/adnoctem/PSFoundation

.NOTES
  Author: MVProwess <info@mvprowess.com>
  License: MIT
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'This script is intended for CI output and Write-Host is appropriate for user feedback.')]
[CmdletBinding()]
param(
  [switch]$Check
)

function Get-PSGalleryLatestVersion {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$Name
  )

  $headers = @{ Accept = 'application/atom+xml' }
  if ($env:PSGALLERY_API_KEY) {
    $headers['X-NuGet-ApiKey'] = $env:PSGALLERY_API_KEY
  }

  $uri = "https://www.powershellgallery.com/api/v2/FindPackagesById()?id='{0}'" -f [uri]::EscapeDataString($Name)
  $response = Invoke-WebRequest -Uri $uri -Headers $headers -UseBasicParsing

  $document = [xml]$response.Content
  $namespace = New-Object System.Xml.XmlNamespaceManager($document.NameTable)
  $namespace.AddNamespace('a', 'http://www.w3.org/2005/Atom')
  $namespace.AddNamespace('m', 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata')
  $namespace.AddNamespace('d', 'http://schemas.microsoft.com/ado/2007/08/dataservices')

  $versions = @()
  foreach ($entry in $document.SelectNodes('//a:entry', $namespace)) {
    $versionNode = $entry.SelectSingleNode('m:properties/d:Version', $namespace)
    $prereleaseNode = $entry.SelectSingleNode('m:properties/d:IsPrerelease', $namespace)
    if (-not $versionNode) {
      continue
    }
    if ($prereleaseNode -and $prereleaseNode.InnerText -eq 'true') {
      continue
    }

    try {
      $null = [version]$versionNode.InnerText
      $versions += $versionNode.InnerText
    }
    catch {
      Write-Host "  Skipping unparsable version '$($versionNode.InnerText)' for $Name" -ForegroundColor DarkGray
    }
  }

  if ($versions.Count -eq 0) {
    return $null
  }

  return ($versions | Sort-Object -Property { [version]$_ } -Descending | Select-Object -First 1)
}

function Get-GitHubToken {
  [CmdletBinding()]
  param()

  foreach ($token in @($env:GH_TOKEN, $env:GITHUB_TOKEN, $env:GH_ENTERPRISE_TOKEN, $env:GITHUB_ENTERPRISE_TOKEN)) {
    if ($token) {
      return $token
    }
  }

  return $null
}

function Get-DependencyDocument {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestContent,

    [Parameter(Mandatory = $true)]
    [string]$DevContent
  )

  $tokens = $null
  $parseErrors = $null
  $ast = [System.Management.Automation.Language.Parser]::ParseInput($ManifestContent, [ref]$tokens, [ref]$parseErrors)
  if ($parseErrors.Count -gt 0) {
    throw "Invalid dependency manifest: $($parseErrors[0].Message)"
  }
  $manifestAst = $ast.EndBlock.Statements[0].PipelineElements[0].Expression
  if ($ast.EndBlock.Statements.Count -ne 1 -or $manifestAst -isnot [System.Management.Automation.Language.HashtableAst]) {
    throw 'The dependency manifest must contain a single data hashtable.'
  }
  $manifest = $manifestAst.SafeGetValue()
  if (-not $DevContent.TrimStart().StartsWith('[')) {
    throw 'The dev dependency document must be a JSON array.'
  }
  # Windows PowerShell 5.1 emits a JSON array as one pipeline object; normalize
  # after assignment so both engines expose a flat array, including [] and [one].
  $devDependencies = ConvertFrom-Json -InputObject $DevContent -ErrorAction Stop
  if ($null -eq $devDependencies) {
    $devDependencies = @()
  }
  else {
    $devDependencies = @($devDependencies)
  }
  $declared = New-Object System.Collections.Generic.List[object]

  foreach ($source in @('runtime', 'dev')) {
    $modules = if ($source -eq 'runtime') { @($manifest.RequiredModules) } else { $devDependencies }
    $index = 0
    foreach ($module in $modules) {
      $name = if ($source -eq 'runtime') {
        if ($module -is [string]) { $module } else { $module.ModuleName }
      }
      else { $module.Name }
      if (-not $name) {
        throw "Missing module name in $source dependency $index."
      }
      $field = if ($module.RequiredVersion) { 'RequiredVersion' }
      elseif ($source -eq 'runtime' -and $module.ModuleVersion) { 'ModuleVersion' }
      elseif ($source -eq 'dev' -and $module.MinimumVersion) { 'MinimumVersion' }
      else { $null }
      $declared.Add([pscustomobject]@{
          Name = $name
          Source = $source
          Index = $index
          VersionField = $field
          DeclaredVersion = if ($field) { [string]$module.$field } else { $null }
        })
      $index++
    }
  }

  return [pscustomobject]@{
    ManifestAst = $manifestAst
    DevDependencies = $devDependencies
    Dependencies = $declared.ToArray()
  }
}

function Get-DependencyUpdate {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestContent,

    [Parameter(Mandatory = $true)]
    [string]$DevContent,

    [Parameter(Mandatory = $true)]
    [object[]]$Updates
  )

  $document = Get-DependencyDocument -ManifestContent $ManifestContent -DevContent $DevContent
  $edits = New-Object System.Collections.Generic.List[object]
  $expected = @{}
  $devChanged = $false

  foreach ($update in $Updates) {
    $declarations = @($document.Dependencies | Where-Object { $_.Source -eq $update.Source -and $_.Name -eq $update.Name })
    if ($declarations.Count -ne 1) {
      throw "Expected exactly one $($update.Source) declaration for '$($update.Name)'; found $($declarations.Count)."
    }
    $dependency = $declarations[0]
    $field = $dependency.VersionField
    if (-not $field -or $field -ne $update.VersionField -or $dependency.DeclaredVersion -cne $update.DeclaredVersion) {
      throw "The declared version field for '$($update.Name)' does not match the proposed update."
    }
    $latest = [string]$update.LatestVersion
    if ([version]$latest -le [version]$dependency.DeclaredVersion) {
      throw "The proposed version for '$($update.Name)' must be newer than the declared version."
    }
    $key = '{0}:{1}' -f $dependency.Source, $dependency.Index
    if ($expected.ContainsKey($key)) {
      throw "Duplicate update for '$($update.Name)'."
    }
    $expected[$key] = $latest

    if ($dependency.Source -eq 'dev') {
      $document.DevDependencies[$dependency.Index].$field = $latest
      $devChanged = $true
      continue
    }

    # Search only RequiredModules, never root metadata, comments, or adjacent declarations.
    $requiredModules = @($document.ManifestAst.KeyValuePairs | Where-Object { $_.Item1.SafeGetValue() -eq 'RequiredModules' })
    $tables = @($requiredModules[0].Item2.FindAll({
          param($node)
          $node -is [System.Management.Automation.Language.HashtableAst]
        }, $false) | Where-Object { $_.SafeGetValue().ModuleName -eq $dependency.Name })
    if ($tables.Count -ne 1) {
      throw "Cannot locate a unique RequiredModules declaration for '$($dependency.Name)'."
    }
    $pair = @($tables[0].KeyValuePairs | Where-Object { $_.Item1.SafeGetValue() -eq $field })
    $valueAst = $pair[0].Item2.PipelineElements[0].Expression
    if ($valueAst -isnot [System.Management.Automation.Language.StringConstantExpressionAst] -or $valueAst.Value -cne $dependency.DeclaredVersion) {
      throw "The $field value for '$($dependency.Name)' must be a literal version string."
    }
    $quote = $valueAst.Extent.Text.Substring(0, 1)
    if ($quote -notin @("'", '"')) {
      throw "The $field value for '$($dependency.Name)' must be quoted."
    }
    $edits.Add([pscustomobject]@{
        Start = $valueAst.Extent.StartOffset
        Length = $valueAst.Extent.EndOffset - $valueAst.Extent.StartOffset
        Text = $quote + $latest + $quote
      })
  }

  # Apply offsets from the end so earlier extents remain valid after longer bumps.
  foreach ($edit in ($edits | Sort-Object Start -Descending)) {
    $ManifestContent = $ManifestContent.Remove($edit.Start, $edit.Length).Insert($edit.Start, $edit.Text)
  }
  if ($devChanged) {
    $DevContent = (ConvertTo-Json -InputObject @($document.DevDependencies) -Depth 100) -replace '\r?\n', "`r`n"
    $DevContent += "`r`n"
  }

  # Reparse both proposed documents before the caller can write or stage either one.
  $proposed = Get-DependencyDocument -ManifestContent $ManifestContent -DevContent $DevContent
  if ($proposed.Dependencies.Count -ne $document.Dependencies.Count) {
    throw 'Dependency declarations changed unexpectedly while preparing updates.'
  }
  for ($i = 0; $i -lt $document.Dependencies.Count; $i++) {
    $before = $document.Dependencies[$i]
    $after = $proposed.Dependencies[$i]
    $key = '{0}:{1}' -f $before.Source, $before.Index
    $version = if ($expected.ContainsKey($key)) { $expected[$key] } else { $before.DeclaredVersion }
    if ($before.Name -cne $after.Name -or $before.Source -ne $after.Source -or $before.VersionField -ne $after.VersionField -or $version -cne $after.DeclaredVersion) {
      throw "Dependency update verification failed for '$($before.Name)'."
    }
  }

  return [pscustomobject]@{
    ManifestContent = $ManifestContent
    DevContent = $DevContent
  }
}

# Dot-sourcing exposes the helpers for tests without checking, writing, or publishing.
if ($MyInvocation.InvocationName -eq '.') {
  return
}

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$RepoRoot = Split-Path -Path $PSScriptRoot -Parent
$ModuleManifestPath = Join-Path -Path $RepoRoot -ChildPath 'src/PSFoundation.psd1'
$DevDependenciesPath = Join-Path -Path $PSScriptRoot -ChildPath 'dev-dependencies.json'

Write-Host 'Checking declared PowerShell module dependencies against the PowerShell Gallery...' -ForegroundColor Cyan

$manifestContent = [System.IO.File]::ReadAllText($ModuleManifestPath)
$devContent = [System.IO.File]::ReadAllText($DevDependenciesPath)
$document = Get-DependencyDocument -ManifestContent $manifestContent -DevContent $devContent

$outdated = New-Object System.Collections.Generic.List[object]

Write-Host ('  {0,-24} {1,-8} {2,-12} {3,-12} {4}' -f 'Name', 'Source', 'Declared', 'Latest', 'Status')

foreach ($dep in $document.Dependencies) {
  $latest = Get-PSGalleryLatestVersion -Name $dep.Name

  if (-not $latest) {
    Write-Host ('  {0,-24} {1,-8} {2,-12} {3,-12} {4}' -f $dep.Name, $dep.Source, '--', '--', 'UNKNOWN') -ForegroundColor Red
    continue
  }

  $isOutdated = $false
  if ($dep.DeclaredVersion) {
    $isOutdated = [version]$latest -gt [version]$dep.DeclaredVersion
  }

  if ($isOutdated) {
    $outdated.Add([pscustomobject]@{
        Name = $dep.Name
        Source = $dep.Source
        VersionField = $dep.VersionField
        DeclaredVersion = $dep.DeclaredVersion
        LatestVersion = $latest
      })
    Write-Host ('  {0,-24} {1,-8} {2,-12} {3,-12} {4}' -f $dep.Name, $dep.Source, $dep.DeclaredVersion, $latest, 'OUTDATED') -ForegroundColor Yellow
  }
  else {
    Write-Host ('  {0,-24} {1,-8} {2,-12} {3,-12} {4}' -f $dep.Name, $dep.Source, $dep.DeclaredVersion, $latest, 'OK') -ForegroundColor Green
  }
}

if ($outdated.Count -eq 0) {
  Write-Host 'All declared dependencies are up to date.' -ForegroundColor Green
  exit 0
}

if ($Check) {
  Write-Host ('{0} dependency update(s) required.' -f $outdated.Count) -ForegroundColor Yellow
  exit 1
}

Write-Host ('Bumping {0} dependency version(s) and opening a pull request...' -f $outdated.Count) -ForegroundColor Cyan

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  throw 'git is required to open a dependency update pull request.'
}
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
  throw 'gh (GitHub CLI) is required to open a dependency update pull request.'
}
$githubToken = Get-GitHubToken
if (-not $githubToken) {
  throw 'No GitHub token found; set GH_TOKEN or GITHUB_TOKEN to push the branch and open the pull request.'
}

$proposed = Get-DependencyUpdate -ManifestContent $manifestContent -DevContent $devContent -Updates $outdated.ToArray()
if ($proposed.ManifestContent -cne $manifestContent) {
  [System.IO.File]::WriteAllText($ModuleManifestPath, $proposed.ManifestContent, (New-Object System.Text.UTF8Encoding($true)))
}
if ($proposed.DevContent -cne $devContent) {
  [System.IO.File]::WriteAllText($DevDependenciesPath, $proposed.DevContent, (New-Object System.Text.UTF8Encoding($false)))
}

$branch = 'chore/psdeps/{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss')
$commitMessage = 'chore(deps): update PowerShell module dependencies'

& git -C $RepoRoot checkout -b $branch
if ($LASTEXITCODE -ne 0) {
  throw "Failed to create branch '$branch'."
}

& git -C $RepoRoot add -- $ModuleManifestPath $DevDependenciesPath
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to stage dependency manifest changes.'
}

& git -C $RepoRoot -c user.name='github-actions[bot]' -c user.email='41898282+github-actions[bot]@users.noreply.github.com' commit -m $commitMessage
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to commit dependency updates.'
}

$repoName = (& gh repo view --json nameWithOwner --jq '.nameWithOwner' 2>$null).Trim()
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to determine the repository name.'
}

& git -C $RepoRoot push "https://x-access-token:$($githubToken)@github.com/$repoName.git" $branch
if ($LASTEXITCODE -ne 0) {
  throw "Failed to push branch '$branch'."
}

$bodyLines = @(
  '| Dependency | Source | Declared | Latest |',
  '|---|---|---|---|'
)
foreach ($dep in $outdated) {
  $bodyLines += '| {0} | {1} | {2} | {3} |' -f $dep.Name, $dep.Source, $dep.DeclaredVersion, $dep.LatestVersion
}
$bodyLines += ''
$bodyLines += 'Generated by the scheduled PSGallery dependency check in `.github/workflows/dependencies.yaml`.'
$body = $bodyLines -join "`n"

$prOutput = & gh pr create --repo $repoName --base main --head $branch --title $commitMessage --body $body
if ($LASTEXITCODE -ne 0) {
  throw 'Failed to open the dependency update pull request.'
}

$prNumber = $null
if ($prOutput -match 'pull/(\d+)') {
  $prNumber = $Matches[1]
}

if ($prNumber) {
  Write-Host ("Opened pull request #{0} with {1} dependency update(s)." -f $prNumber, $outdated.Count) -ForegroundColor Green
}
else {
  Write-Host ("Pushed dependency updates to branch '{0}' and opened a pull request." -f $branch) -ForegroundColor Green
}
