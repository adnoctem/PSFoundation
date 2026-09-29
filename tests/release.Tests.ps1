#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  # Load only the helpers, never the tool's publishing entry point.
  $tokens = $null
  $parseErrors = $null
  $ast = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot '../tools/release.ps1'), [ref]$tokens, [ref]$parseErrors
  )
  foreach ($name in @('Write-DistChecksum', 'Get-ManifestAlignmentWidth', 'Write-ReleaseManifest')) {
    $helper = $ast.Find({
        param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
      }.GetNewClosure(), $false)
    $helper | Should -Not -BeNullOrEmpty -Because "release.ps1 must define $name"
    . ([scriptblock]::Create($helper.Extent.Text))
  }

  # Mirrors the real manifest's alignment: 22 at the top level, 25 inside PSData.
  function Write-TestManifest {
    param (
      [Parameter(Mandatory = $true)]
      [string]$Path,

      [switch]$LivePrerelease
    )

    $prerelease = "      # Prerelease = 'beta'"
    if ($LivePrerelease) {
      $prerelease = "      Prerelease               = '1.9.0-beta.2'"
    }
    $lines = @(
      '@{',
      '',
      '  # Script module or binary module file associated with this manifest.',
      "  RootModule            = 'PSFoundation.psm1'",
      '',
      '  # Version number of this module.',
      "  ModuleVersion         = '1.0.0'",
      '',
      "  CompatiblePSEditions  = @('Desktop', 'Core')",
      '',
      '  PrivateData           = @{',
      '    PSData = @{',
      "      ReleaseNotes             = 'https://example.invalid/releases'",
      '',
      '      # Prerelease string of this module',
      $prerelease,
      '',
      '      RequireLicenseAcceptance = $false',
      '    }',
      '  }',
      '}'
    )
    [IO.File]::WriteAllText($Path, ($lines -join "`r`n"), (New-Object Text.UTF8Encoding($true)))
    $Path
  }
}

Describe 'Release archive collection handling' {
  It 'handles <Number> archives under strict mode' -ForEach @(@{ Number = 0 }, @{ Number = 1 }, @{ Number = 2 }) {
    $script:distPath = Join-Path $TestDrive "dist-$Number"
    $script:checksumPath = Join-Path $distPath 'CHECKSUMS_SHA256.txt'
    New-Item -ItemType Directory -Path $distPath -Force | Out-Null
    for ($i = 0; $i -lt $Number; $i++) {
      Set-Content -LiteralPath (Join-Path $distPath "synthetic$i.zip") -Value 'Synthetic archive bytes'
    }
    Set-StrictMode -Version Latest
    Write-DistChecksum -WarningVariable warnings -WarningAction SilentlyContinue | Out-Null
    if ($Number -eq 0) {
      ($warnings -join ' ') | Should -Match 'No archive files found'
      Test-Path -LiteralPath $checksumPath | Should -BeFalse
    }
    else {
      $lines = @(Get-Content -LiteralPath $checksumPath)
      $lines.Count | Should -Be $Number
      $lines[0] | Should -Match '^[0-9A-F]{64}  synthetic0\.zip$'
    }
  }
}

Describe 'Release manifest version synchronization' {
  BeforeEach {
    $script:manifest = Join-Path $TestDrive 'PSFoundation.psd1'
  }

  It 'reports the operator column the formatter aligns <Indent> to' -ForEach @(
    @{ Indent = 'top-level keys'; Spaces = '  '; Expected = 22 }
    @{ Indent = 'PSData keys'; Spaces = '      '; Expected = 25 }
  ) {
    Set-StrictMode -Version Latest
    $content = Get-Content -LiteralPath (Write-TestManifest -Path $manifest) -Raw
    Get-ManifestAlignmentWidth -Content $content -Indent $Spaces | Should -Be $Expected
  }

  It 'preserves the aligned assignment operator when bumping the version' {
    Set-StrictMode -Version Latest
    $null = Write-ReleaseManifest -Path (Write-TestManifest -Path $manifest) -CoreVersion '1.8.0'

    # The whole point: semantic-release must not collapse the padding, because
    # PSAlignAssignmentStatement then fails the repository's own format gate.
    $lines = @(Get-Content -LiteralPath $manifest)
    $lines | Should -Contain "  ModuleVersion         = '1.8.0'"
    $lines | Should -Not -Contain "  ModuleVersion = '1.8.0'"
    @($lines | Where-Object { $_ -match "^  RootModule            = 'PSFoundation.psm1'$" }).Count | Should -Be 1
  }

  It 'writes the manifest as UTF-8 with a BOM' {
    Set-StrictMode -Version Latest
    $null = Write-ReleaseManifest -Path (Write-TestManifest -Path $manifest) -CoreVersion '1.8.0'

    # Windows PowerShell 5.1 needs the BOM to parse the manifest.
    $bytes = [IO.File]::ReadAllBytes($manifest)
    @($bytes[0], $bytes[1], $bytes[2]) | Should -Be @(0xEF, 0xBB, 0xBF)
  }

  It 'aligns an uncommented Prerelease key to its own block' {
    Set-StrictMode -Version Latest
    $null = Write-ReleaseManifest -Path (Write-TestManifest -Path $manifest) -CoreVersion '1.9.0' -Prerelease 'beta.2'

    $lines = @(Get-Content -LiteralPath $manifest)
    $lines | Should -Contain "      Prerelease               = 'beta.2'"
    $lines | Should -Contain "  ModuleVersion         = '1.9.0'"
    # The neighbouring explanatory comment must survive untouched.
    $lines | Should -Contain '      # Prerelease string of this module'
  }

  It 'comments a live Prerelease out again for a stable release' {
    Set-StrictMode -Version Latest
    $null = Write-ReleaseManifest -Path (Write-TestManifest -Path $manifest -LivePrerelease) -CoreVersion '2.0.0'

    $lines = @(Get-Content -LiteralPath $manifest)
    $lines | Should -Contain "      # Prerelease = ''"
    @($lines | Where-Object { $_ -match "^\s*Prerelease\s*=" }).Count | Should -Be 0
    $lines | Should -Contain "  ModuleVersion         = '2.0.0'"
  }

  It 'refuses a manifest without a top-level <Key>' -ForEach @(
    @{ Key = 'RootModule'; Pattern = 'indentation' }
    @{ Key = 'ModuleVersion'; Pattern = 'ModuleVersion' }
  ) {
    Set-StrictMode -Version Latest
    $path = Write-TestManifest -Path $manifest
    $stripped = @(Get-Content -LiteralPath $path) | Where-Object { $_ -notmatch "^  $Key\s" }
    [IO.File]::WriteAllText($path, ($stripped -join "`r`n"), (New-Object Text.UTF8Encoding($true)))

    { Write-ReleaseManifest -Path $path -CoreVersion '1.8.0' } | Should -Throw -ExpectedMessage "*$Pattern*"
  }
}
