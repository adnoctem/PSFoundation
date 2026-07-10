#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/registry.ps1
}

Describe 'ConvertTo-RegistryProviderPath' {
  Context 'HKLM paths' {
    It 'returns canonical HKLM: for short HKLM format' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKLM\Software\Microsoft'
      $result | Should -Be 'HKLM:\Software\Microsoft'
    }

    It 'returns unchanged for HKLM: PS drive format' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKLM:\Software\Microsoft'
      $result | Should -Be 'HKLM:\Software\Microsoft'
    }

    It 'returns canonical for Registry:: prefix with long hive name' {
      $result = ConvertTo-RegistryProviderPath -Path 'Registry::HKEY_LOCAL_MACHINE\Software\Microsoft'
      $result | Should -Be 'HKLM:\Software\Microsoft'
    }

    It 'returns canonical for long .NET hive name' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKEY_LOCAL_MACHINE\Software\Microsoft'
      $result | Should -Be 'HKLM:\Software\Microsoft'
    }

    It 'returns hive root when subkey is empty' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKLM'
      $result | Should -Be 'HKLM:'
    }
  }

  Context 'HKCU paths' {
    It 'returns canonical for short HKCU format' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKCU\Control Panel\Desktop'
      $result | Should -Be 'HKCU:\Control Panel\Desktop'
    }

    It 'returns canonical for HKEY_CURRENT_USER long name' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKEY_CURRENT_USER\Software'
      $result | Should -Be 'HKCU:\Software'
    }
  }

  Context 'HKCR paths' {
    It 'returns Registry:: prefix for HKCR' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKCR\*\shell\open\command'
      $result | Should -Match '^Registry::HKEY_CLASSES_ROOT\\'
    }
  }

  Context 'HKU paths' {
    It 'returns Registry:: prefix for HKU' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKU\S-1-5-18\Software'
      $result | Should -Match '^Registry::HKEY_USERS\\'
    }
  }

  Context 'HKCC paths' {
    It 'returns Registry:: prefix for HKCC' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKCC\Software'
      $result | Should -Match '^Registry::HKEY_CURRENT_CONFIG\\'
    }
  }

  Context 'Registry:: prefixed inputs' {
    It 'strips duplicate Registry:: prefix' {
      $result = ConvertTo-RegistryProviderPath -Path 'Registry::HKEY_CURRENT_USER\Software'
      $result | Should -Be 'HKCU:\Software'
    }
  }

  Context 'edge cases' {
    It 'collapses duplicate path separators' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKLM\\Software\\\\Microsoft'
      $result | Should -Be 'HKLM:\Software\Microsoft'
    }

    It 'trims trailing backslash' {
      $result = ConvertTo-RegistryProviderPath -Path 'HKLM:\Software\Microsoft\'
      $result | Should -Be 'HKLM:\Software\Microsoft'
    }

    It 'returns null for unrecognised hive' {
      $result = ConvertTo-RegistryProviderPath -Path 'NONSENSE\Path' -ErrorAction SilentlyContinue
      $result | Should -Be $null
    }
  }
}

Describe 'Test-RegistryPath' {
  It 'returns true for an existing registry key' {
    $result = Test-RegistryPath -Path 'HKLM:\Software\Microsoft'
    $result | Should -BeTrue
  }

  It 'returns false for a non-existent registry key' {
    $result = Test-RegistryPath -Path 'HKLM:\Software\NonExistentFOOBAR12345'
    $result | Should -BeFalse
  }

  It 'returns false for an invalid path without throwing' {
    $result = Test-RegistryPath -Path 'NONSENSE\Path' -ErrorAction SilentlyContinue
    $result | Should -BeFalse
  }
}

Describe 'Test-RegistryValue' {
  It 'returns false for a non-existent key' {
    $result = Test-RegistryValue -Path 'HKLM:\Software\NonExistentFOOBAR' -Name 'AnyValue'
    $result | Should -BeFalse
  }

  It 'returns false when the key exists but the value does not' {
    $result = Test-RegistryValue -Path 'HKLM:\Software\Microsoft' -Name 'NonExistentValue12345'
    $result | Should -BeFalse
  }
}
