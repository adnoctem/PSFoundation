#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/registry.ps1
}

Describe 'New-OperationResult' {
  It 'produces a PSCustomObject with Target, Action, and Status' {
    $result = New-OperationResult -Target 'TestTarget' -Action 'Install' -Status 'Completed'
    $result | Should -BeOfType [PSCustomObject]
    $result.Target | Should -Be 'TestTarget'
    $result.Action | Should -Be 'Install'
    $result.Status | Should -Be 'Completed'
  }

  It 'omits Source when not supplied' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Z'
    $result.PSObject.Properties.Name -contains 'Source' | Should -BeFalse
  }

  It 'includes Source when supplied' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Z' -Source 'Registry'
    $result.Source | Should -Be 'Registry'
  }

  It 'omits Scope when not supplied' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Z'
    $result.PSObject.Properties.Name -contains 'Scope' | Should -BeFalse
  }

  It 'includes Scope when supplied' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Z' -Scope 'CurrentUser'
    $result.Scope | Should -Be 'CurrentUser'
  }

  It 'includes Detail when supplied' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Skipped' -Detail 'No match'
    $result.Detail | Should -Be 'No match'
  }

  It 'omits Detail when not supplied' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Z'
    $result.PSObject.Properties.Name -contains 'Detail' | Should -BeFalse
  }

  It 'includes SkippedReason when supplied' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Skipped' -SkippedReason 'AlreadyExists'
    $result.SkippedReason | Should -Be 'AlreadyExists'
  }

  It 'renames ErrorMessage parameter to Error property' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Failed' -ErrorMessage 'That failed'
    $result.Error | Should -Be 'That failed'
  }

  It 'appends extra properties from -Property without overwriting core fields' {
    $result = New-OperationResult -Target 'X' -Action 'Y' -Status 'Z' -Property @{ Extra = 'value'; Target = 'ignored' }
    $result.Extra | Should -Be 'value'
    $result.Target | Should -Be 'X'
  }

  It 'handles property hashtable without specifying any optional fields' {
    $result = New-OperationResult -Target 'Basic' -Action 'Log' -Status 'Done' -Property @{ Custom = 'data' }
    $result.Target | Should -Be 'Basic'
    $result.Custom | Should -Be 'data'
  }
}

Describe 'Add-OperationResult' {
  It 'appends a result to an ArrayList' {
    $results = New-Object System.Collections.ArrayList
    Add-OperationResult -Results $results -Target 'A' -Action 'Set' -Status 'Done'
    $results.Count | Should -Be 1
    $results[0].Target | Should -Be 'A'
  }

  It 'passes through the result when -PassThru is set' {
    $results = New-Object System.Collections.ArrayList
    $out = Add-OperationResult -Results $results -Target 'B' -Action 'Remove' -Status 'Skipped' -PassThru
    $out.Target | Should -Be 'B'
    $results.Count | Should -Be 1
  }

  It 'does not write pipeline output without -PassThru' {
    $results = New-Object System.Collections.ArrayList
    $out = Add-OperationResult -Results $results -Target 'C' -Action 'Test' -Status 'Ok'
    $out | Should -BeNullOrEmpty
  }

  It 'forwards optional parameters to New-OperationResult' {
    $results = New-Object System.Collections.ArrayList
    Add-OperationResult -Results $results -Target 'D' -Action 'Install' -Status 'Failed' -Source 'WinGet' -Detail 'Not found'
    $results[0].Source | Should -Be 'WinGet'
    $results[0].Detail | Should -Be 'Not found'
  }
}

Describe 'Export-RegistrySettingState' {
  It 'returns a snapshot with Preferred populated from current registry value' {
    Mock Get-RegistryValue { 'mock-current-value' }
    Mock Test-RegistryValue { $true }
    Mock Get-RegistryValueKind { 'String' }

    $setting = @{
      Path = 'HKLM:\Software\Test'
      Name = 'TestValue'
      Type = 'String'
      Default = 'default'
    }

    $result = Export-RegistrySettingState -Settings @($setting)
    $result.Path | Should -Be 'HKLM:\Software\Test'
    $result.Name | Should -Be 'TestValue'
    $result.Preferred | Should -Be 'mock-current-value'
    $result.Type | Should -Be 'String'
  }

  It 'sets Preferred to null for missing registry values' {
    Mock Test-RegistryValue { $false }

    $setting = @{
      Path = 'HKLM:\Software\Test'
      Name = 'MissingValue'
    }

    $result = Export-RegistrySettingState -Settings @($setting)
    $result.Preferred | Should -BeNull
  }

  It 'errors when Path or Name is missing' {
    Mock Test-RegistryValue { $false }

    $incomplete = @{ Name = 'NoPath' }
    { Export-RegistrySettingState -Settings @($incomplete) -ErrorAction Stop } | Should -Throw
  }
}

Describe 'ConvertTo-RegistrySettingResult' {
  It 'builds Skipped/DryRun results in DryRun mode' {
    $setting = [PSCustomObject]@{
      Path = 'HKLM:\Software\Test'
      Name = 'Setting1'
      Preferred = 'enabled'
      Default = 'disabled'
      Description = 'A test setting'
    }

    $results = ConvertTo-RegistrySettingResult -Settings @($setting) -DryRun
    $results.Count | Should -Be 1
    $results[0].Status | Should -Be 'Skipped'
    $results[0].Detail | Should -Be 'DryRun'
    $results[0].Target | Should -Match 'Setting1'
  }

  It 'builds undo results with RemoveValue action when Default is null' {
    $setting = [PSCustomObject]@{
      Path = 'HKLM:\Software\Test'
      Name = 'Setting2'
      Preferred = 'enabled'
      Default = $null
    }

    $results = ConvertTo-RegistrySettingResult -Settings @($setting) -Undo -DryRun
    $results[0].Action | Should -Be 'RemoveValue'
  }

  It 'builds SetValue action for normal apply with Preferred' {
    $setting = [PSCustomObject]@{
      Path = 'HKLM:\Software\Test'
      Name = 'Setting3'
      Preferred = '42'
      Default = 0
    }

    $results = ConvertTo-RegistrySettingResult -Settings @($setting) -DryRun
    $results[0].Action | Should -Be 'SetValue'
  }

  It 'skips null settings' {
    $null | ConvertTo-RegistrySettingResult -ErrorAction SilentlyContinue
    $true | Should -BeTrue
  }

  It 'marks result as Removed when undo successfully removes a value' {
    Mock Test-RegistryValue { $false }

    $setting = [PSCustomObject]@{
      Path = 'HKLM:\Software\Test'
      Name = 'Setting4'
      Preferred = 'value'
      Default = $null
    }

    $results = ConvertTo-RegistrySettingResult -Settings @($setting) -Undo
    $results[0].Status | Should -Be 'Removed'
    $results[0].Action | Should -Be 'RemoveValue'
  }
}
