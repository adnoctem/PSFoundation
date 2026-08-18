#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Guard against manifest/psm1 export drift. The module's public surface is
# defined in two places: the $publicFunctions list in src/PSFoundation.psm1
# (via Export-ModuleMember) and FunctionsToExport in src/PSFoundation.psd1.
# PowerShell applies the intersection of the two, so a function missing from
# either list is invisible to consumers - Get-PrintDevice was lost this way
# once. These tests fail on any mismatch.

BeforeAll {
  $script:ManifestPath = Join-Path $PSScriptRoot '../src/PSFoundation.psd1'
  $script:Manifest = Test-ModuleManifest -Path $ManifestPath -ErrorAction Stop
  $script:Module = Import-Module -Name $ManifestPath -Force -PassThru -ErrorAction Stop
}

AfterAll {
  Remove-Module -Name PSFoundation -Force -ErrorAction SilentlyContinue
}

Describe 'Module manifest exports' {
  It 'makes Get-PrintDevice available to consumers' {
    Get-Command -Name Get-PrintDevice -Module PSFoundation -ErrorAction Stop | Should -Not -BeNullOrEmpty
  }

  It 'exports the same function set from the manifest and the module' {
    $manifestFunctions = @($Manifest.ExportedFunctions.Keys) | Sort-Object
    $moduleFunctions = @($Module.ExportedFunctions.Keys) | Sort-Object
    $manifestFunctions | Should -Be $moduleFunctions
  }

  It 'exports every function that the psm1 declares public' {
    $tokens = $null
    $errors = $null
    $psm1Path = Join-Path $PSScriptRoot '../src/PSFoundation.psm1'
    [System.Management.Automation.Language.Parser]::ParseFile($psm1Path, [ref]$tokens, [ref]$errors) | Out-Null

    $publicFunctionsIndex = -1
    for ($i = 0; $i -lt $tokens.Count; $i++) {
      if ($tokens[$i].Kind -eq 'Variable' -and $tokens[$i].Text -eq '$publicFunctions') {
        $publicFunctionsIndex = $i
        break
      }
    }
    $publicFunctionsIndex | Should -BeGreaterThan 0

    $declared = @()
    for ($i = $publicFunctionsIndex; $i -lt $tokens.Count; $i++) {
      $token = $tokens[$i]
      if ($token.Kind -eq 'StringLiteral') {
        $declared += $token.Text.Trim("'")
      }
      elseif ($token.Kind -eq 'RParen' -and $i -gt $publicFunctionsIndex) {
        break
      }
    }
    $declared.Count | Should -BeGreaterThan 100

    $moduleFunctions = @($Module.ExportedFunctions.Keys)
    foreach ($name in $declared) {
      $name | Should -BeIn $moduleFunctions
    }
  }

  It 'contains no stale function entries' {
    $manifestFunctions = @($Manifest.ExportedFunctions.Keys)
    $moduleFunctions = @($Module.ExportedFunctions.Keys)
    foreach ($name in $manifestFunctions) {
      $name | Should -BeIn $moduleFunctions
    }
  }
}
