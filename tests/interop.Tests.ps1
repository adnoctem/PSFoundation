#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/interop.ps1
}

Describe 'Remove-ComObject' {
  It 'does not throw when passed null' {
    { Remove-ComObject $null } | Should -Not -Throw
  }

  It 'does not throw when passed multiple values including null' {
    { Remove-ComObject $null, 'non-com-object', $null } | Should -Not -Throw
  }

  It 'does not throw when no arguments are passed' {
    { Remove-ComObject } | Should -Not -Throw
  }
}

Describe 'Invoke-ComGarbageCollection' {
  It 'does not throw' {
    { Invoke-ComGarbageCollection } | Should -Not -Throw
  }
}
