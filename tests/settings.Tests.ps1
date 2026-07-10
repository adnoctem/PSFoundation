#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/settings.ps1
}

Describe 'Get-DefaultApp' {
  It 'returns the default application command for a known extension' {
    Mock Get-ItemProperty {
      [PSCustomObject]@{ ProgId = 'txtfile' }
    } -ParameterFilter { $Path -like 'HKCU:*UserChoice' }

    Mock Get-ItemProperty {
      [PSCustomObject]@{ '(default)' = 'C:\Windows\system32\NOTEPAD.EXE %1' }
    } -ParameterFilter { $Path -like 'HKCR:*shell\open\command' }

    $result = Get-DefaultApp -FileExtension '.txt'
    $result | Should -Be 'C:\Windows\system32\NOTEPAD.EXE %1'
  }

  It 'writes an error when the UserChoice key is not found' {
    Mock Get-ItemProperty { throw 'Registry key not found' } -ParameterFilter { $Path -like 'HKCU:*UserChoice' }

    { Get-DefaultApp -FileExtension '.fake' -ErrorAction Stop } | Should -Throw
  }
}
