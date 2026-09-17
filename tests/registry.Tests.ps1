#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/registry.ps1
}

Describe 'Registry setting snapshots and restoration' {
  BeforeAll { . $PSScriptRoot/../src/common.ps1 }
  BeforeEach {
    $subkey = 'Software\PSFoundation.Tests\' + [guid]::NewGuid().ToString('N')
    $script:registryPath = 'HKCU:\' + $subkey
    $script:testKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($subkey)
  }
  AfterEach {
    $testKey.Dispose()
    [Microsoft.Win32.Registry]::CurrentUser.DeleteSubKeyTree($subkey, $false)
  }

  It 'round trips raw types and ordered values through JSON and restores them' {
    $testKey.SetValue('Text', '%TEMP%\test', [Microsoft.Win32.RegistryValueKind]::ExpandString)
    $testKey.SetValue('Binary', [byte[]]@(0, 255, 7), [Microsoft.Win32.RegistryValueKind]::Binary)
    $testKey.SetValue('Multi', [string[]]@('A', 'b'), [Microsoft.Win32.RegistryValueKind]::MultiString)
    $testKey.SetValue('Large', [long]4294967296, [Microsoft.Win32.RegistryValueKind]::QWord)
    $testKey.SetValue('', '', [Microsoft.Win32.RegistryValueKind]::String)
    $settings = @('Text', 'Binary', 'Multi', 'Large', '', 'Missing') | ForEach-Object { @{ Path = $registryPath; Name = $_; Group = 'Preserved' } }
    $before = @(Export-RegistrySettingState -Settings $settings -Detailed)
    $before[0].Preferred | Should -Be '%TEMP%\test'
    $before[0].Group | Should -Be 'Preserved'
    $before[0].View | Should -BeIn @('Registry32', 'Registry64')
    $before[5].Exists | Should -BeFalse
    foreach ($name in @('Text', 'Binary', 'Multi', 'Large', '', 'Missing')) { $testKey.SetValue($name, 'changed', [Microsoft.Win32.RegistryValueKind]::String) }
    $expected = @(Export-RegistrySettingState -Settings $settings -Detailed)
    $restored = $before | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $result = @(Restore-RegistrySettingState -Settings $restored -ExpectedState $expected -Confirm:$false)
    @($result | Where-Object Status -EQ 'Restored').Count | Should -Be 6
    @(Compare-RegistrySettingState -Settings $before | Where-Object Changed).Count | Should -Be 0
    $testKey.GetValueNames() | Should -Not -Contain 'Missing'
  }

  It 'previews differences without mutation and protects intervening edits' {
    $testKey.SetValue('Value', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
    $setting = @{ Path = $registryPath; Name = 'Value' }
    $before = Export-RegistrySettingState -Settings $setting -Detailed
    $testKey.SetValue('Value', 2)
    $expected = Export-RegistrySettingState -Settings $setting -Detailed
    $preview = Compare-RegistrySettingState -Settings $before
    $preview.Before.Preferred | Should -Be 2
    $preview.After.Preferred | Should -Be 1
    (Restore-RegistrySettingState -Settings $before -ExpectedState $expected -WhatIf).Status | Should -Be 'Skipped'
    $testKey.GetValue('Value') | Should -Be 2
    $testKey.SetValue('Value', 3)
    (Restore-RegistrySettingState -Settings $before -ExpectedState $expected).Status | Should -Be 'Conflict'
    $testKey.GetValue('Value') | Should -Be 3
    (Restore-RegistrySettingState -Settings $before -ExpectedState @()).Status | Should -Be 'Conflict'
  }

  It 'detects case, array order and type changes in desired configuration' {
    $testKey.SetValue('Value', [string[]]@('A', 'B'), [Microsoft.Win32.RegistryValueKind]::MultiString)
    $desired = @{ Path = $registryPath; Name = 'Value'; Type = 'MultiString'; Preferred = @('B', 'A') }
    (Compare-RegistrySettingState -Settings $desired).Changed | Should -BeTrue
    $desired.Preferred = @('a', 'B')
    (Compare-RegistrySettingState -Settings $desired).Changed | Should -BeTrue
    $desired.Type = 'String'
    $desired.Preferred = 'A B'
    (Compare-RegistrySettingState -Settings $desired).Changed | Should -BeTrue
  }

  It 'does not remove keys or unrelated values when restoring absence' {
    $snapshot = Export-RegistrySettingState -Settings @{ Path = $registryPath; Name = 'Missing' } -Detailed
    $testKey.SetValue('Missing', 'created')
    $testKey.SetValue('Unrelated', 'keep')
    (Restore-RegistrySettingState -Settings $snapshot).Status | Should -Be 'Restored'
    $testKey.GetValue('Unrelated') | Should -Be 'keep'
    (Restore-RegistrySettingState -Settings $snapshot).AlreadyCompliant | Should -BeTrue
  }

  It 'rejects old snapshots and ambiguous desired values before writing' {
    { Restore-RegistrySettingState -Settings @{ Path = $registryPath; Name = 'Value'; Preferred = 1; Type = 'DWord' } } | Should -Throw '*version 1*'
    { Compare-RegistrySettingState -Settings @{ Path = $registryPath; Name = 'Value'; Type = 'DWord'; Preferred = $null } } | Should -Throw '*non-null*'
    { Compare-RegistrySettingState -Settings @{ Path = $registryPath; Name = 'Value'; Exists = 'false' } } | Should -Throw '*Boolean*'
  }

  It 'does not interpret access errors as missing values' {
    Mock Resolve-RegistryPath { throw 'Access denied' }
    { Export-RegistrySettingState -Settings @{ Path = $registryPath; Name = 'Value' } -Detailed } | Should -Throw '*Access denied*'
  }

  It 'recreates a missing key and retains an explicitly selected registry view' {
    $nested = $testKey.CreateSubKey('Nested')
    $nested.SetValue('Value', -1, [Microsoft.Win32.RegistryValueKind]::DWord)
    $nested.Dispose()
    $snapshot = Export-RegistrySettingState -Settings @{ Path = "$registryPath\Nested"; Name = 'Value' } -Detailed -View Registry32
    $snapshot.View | Should -Be 'Registry32'
    $testKey.DeleteSubKey('Nested')
    (Restore-RegistrySettingState -Settings $snapshot -Confirm:$false).Status | Should -Be 'Restored'
    (Compare-RegistrySettingState -Settings $snapshot).Changed | Should -BeFalse
  }
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
