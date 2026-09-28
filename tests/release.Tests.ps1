#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  # Load only the checksum helper, never the tool's publishing entry point.
  $tokens = $null
  $parseErrors = $null
  $ast = [Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot '../tools/release.ps1'), [ref]$tokens, [ref]$parseErrors
  )
  $helper = $ast.Find({
      param($node)
      $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Write-DistChecksum'
    }, $false)
  . ([scriptblock]::Create($helper.Extent.Text))
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
