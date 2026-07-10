#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/permissions.ps1
}

Describe 'Test-Elevation' {
  It 'returns a boolean on Windows' {
    $result = Test-Elevation
    $result | Should -BeOfType [bool]
  }

  It 'does not throw on supported platform' {
    { Test-Elevation } | Should -Not -Throw
  }
}
