#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/registry.ps1
}

Describe 'New-OperationResult' {
  It 'keeps optional outcome fields absent unless explicitly supplied' {
    $plain = New-OperationResult -Target 'A' -Action 'Set' -Status 'Completed'
    $plain.PSObject.Properties.Name | Should -Not -Contain 'Changed'
    $result = New-OperationResult -Target 'A' -Action 'Set' -Status 'Completed' -Changed $false -AlreadyCompliant $true -Before $null -After 0 -ExitCode 0 -RebootRequired $false -Duration ([timespan]::FromSeconds(2)) -RunId 'run-1'
    $result.Changed | Should -BeFalse
    $result.AlreadyCompliant | Should -BeTrue
    $result.PSObject.Properties.Name | Should -Contain 'Before'
    $result.Before | Should -BeNull
    $result.After | Should -Be 0
    $result.ExitCode | Should -Be 0
    $result.RebootRequired | Should -BeFalse
    $result.Duration.TotalSeconds | Should -Be 2
    $result.RunId | Should -Be 'run-1'
  }

  It 'forwards outcome metadata through the collection and JSON Lines helpers' {
    $results = [Collections.ArrayList]::new()
    Add-OperationResult -Results $results -Target 'A' -Action 'Set' -Status 'Completed' -Changed $false -Before $null -RunId 'operation-run'
    Add-OperationResult -Results $results -Target 'B' -Action 'Set' -Status 'Completed' -ExitCode 3010 -RebootRequired $true
    $path = Write-OperationResultLog -Results $results -Path (Join-Path $TestDrive 'results.jsonl') -RunId 'batch-run'
    $lines = @(Get-Content -LiteralPath $path | ForEach-Object { $_ | ConvertFrom-Json })
    $lines.Count | Should -Be 2
    $lines[0].RunId | Should -Be 'operation-run'
    $lines[0].Changed | Should -BeFalse
    $lines[1].RunId | Should -Be 'batch-run'
    $lines[1].RebootRequired | Should -BeTrue
  }

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
      Path    = 'HKLM:\Software\Test'
      Name    = 'TestValue'
      Type    = 'String'
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
      Path        = 'HKLM:\Software\Test'
      Name        = 'Setting1'
      Preferred   = 'enabled'
      Default     = 'disabled'
      Description = 'A test setting'
    }

    $results = @(ConvertTo-RegistrySettingResult -Settings @($setting) -DryRun)
    $results.Count | Should -Be 1
    $results[0].Status | Should -Be 'Skipped'
    $results[0].Detail | Should -Be 'DryRun'
    $results[0].Target | Should -Match 'Setting1'
  }

  It 'builds undo results with RemoveValue action when Default is null' {
    $setting = [PSCustomObject]@{
      Path      = 'HKLM:\Software\Test'
      Name      = 'Setting2'
      Preferred = 'enabled'
      Default   = $null
    }

    $results = ConvertTo-RegistrySettingResult -Settings @($setting) -Undo -DryRun
    $results[0].Action | Should -Be 'RemoveValue'
  }

  It 'builds SetValue action for normal apply with Preferred' {
    $setting = [PSCustomObject]@{
      Path      = 'HKLM:\Software\Test'
      Name      = 'Setting3'
      Preferred = '42'
      Default   = 0
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
      Path      = 'HKLM:\Software\Test'
      Name      = 'Setting4'
      Preferred = 'value'
      Default   = $null
    }

    $results = ConvertTo-RegistrySettingResult -Settings @($setting) -Undo
    $results[0].Status | Should -Be 'Removed'
    $results[0].Action | Should -Be 'RemoveValue'
  }
}


Describe 'Write-OperationResultLog default log directory' {
  BeforeEach {
    $script:originalTemp = $env:TEMP
    $script:fakeTemp = Join-Path $TestDrive "temp-$([guid]::NewGuid().ToString('N'))"
    $null = [IO.Directory]::CreateDirectory($script:fakeTemp)
    $env:TEMP = $script:fakeTemp
  }

  AfterEach {
    $env:TEMP = $script:originalTemp
  }

  It 'writes under this module''s own directory, never another product''s' {
    $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
    $path = Write-OperationResultLog -Results @($result) -ScriptName 'Synthetic-Script'

    # A library must not log into a consuming product's directory by default.
    $path | Should -BeLike (Join-Path $fakeTemp 'PSFoundation\logs\*')
    $path | Should -Not -BeLike '*\winkit\*'
    Test-Path -LiteralPath $path | Should -BeTrue
  }

  It 'lets a consuming product choose its own directory with <Name>' -ForEach @(
    @{ Name = 'winkit' }
    @{ Name = 'some.other-product' }
  ) {
    $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
    $path = Write-OperationResultLog -Results @($result) -ScriptName 'Synthetic-Script' -Name $Name

    $path | Should -BeLike (Join-Path $fakeTemp "$Name\logs\*")
    Test-Path -LiteralPath $path | Should -BeTrue
  }

  It 'still uses its own directory when ScriptName is omitted' {
    # The ScriptName fallback derives from Name too, but a caller with a script
    # of its own supplies that name, so only the directory is asserted here.
    $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
    $path = Write-OperationResultLog -Results @($result)

    $path | Should -BeLike (Join-Path $fakeTemp 'PSFoundation\logs\*')
    $path | Should -BeLike '*.jsonl'
  }

  It 'rejects <Description> in Name so it cannot escape the log root' -ForEach @(
    @{ Value = '..\..\Windows'; Description = 'traversal' }
    @{ Value = 'winkit\logs'; Description = 'a separator' }
    @{ Value = 'C:\Temp'; Description = 'a qualified path' }
    @{ Value = '*'; Description = 'a wildcard' }
  ) {
    $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
    { Write-OperationResultLog -Results @($result) -Name $Value } |
      Should -Throw -ExpectedMessage '*does not match*'
  }
}

Describe 'Write-OperationResultLog path handling' {
  It 'resolves relative destinations against the PowerShell location' {
    $nativeRoot = Join-Path $TestDrive 'native'
    $shellRoot = Join-Path $TestDrive 'shell'
    $null = [IO.Directory]::CreateDirectory($nativeRoot)
    $null = [IO.Directory]::CreateDirectory($shellRoot)
    $previous = [Environment]::CurrentDirectory
    Push-Location $shellRoot
    try {
      [Environment]::CurrentDirectory = $nativeRoot
      $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
      $path = Write-OperationResultLog -Results @($result) -Path 'review.jsonl'
      $path | Should -Be (Join-Path $shellRoot 'review.jsonl')
      Test-Path -LiteralPath $path | Should -BeTrue
      Test-Path -LiteralPath (Join-Path $nativeRoot 'review.jsonl') | Should -BeFalse
    }
    finally {
      [Environment]::CurrentDirectory = $previous
      Pop-Location
    }
  }

  It 'preserves an existing log when path normalization fails' {
    $path = Join-Path $TestDrive 'existing.jsonl'
    [IO.File]::WriteAllText($path, 'original content')
    Mock Resolve-LongPath { throw 'Synthetic path expansion failure' }
    $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
    { Write-OperationResultLog -Results @($result) -Path $path } | Should -Throw '*Synthetic path expansion failure*'
    [IO.File]::ReadAllText($path) | Should -Be 'original content'
  }

  It 'creates nested literal directories through a filesystem PSDrive' {
    $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
    $path = Write-OperationResultLog -Results @($result) -Path 'TestDrive:\new [literal]\nested\results.jsonl'
    $path | Should -Be (Join-Path $TestDrive 'new [literal]\nested\results.jsonl')
    (Get-Content -LiteralPath $path | ConvertFrom-Json).Target | Should -Be 'Synthetic'
  }

  It 'rejects non-filesystem destinations before changing their values' {
    $env:PSFOUNDATION_TEST_LOG = 'original content'
    try {
      $result = [PSCustomObject]@{ Target = 'Synthetic'; Status = 'Completed' }
      { Write-OperationResultLog -Results @($result) -Path Env:PSFOUNDATION_TEST_LOG } | Should -Throw '*filesystem*'
      $env:PSFOUNDATION_TEST_LOG | Should -Be 'original content'
    }
    finally {
      Remove-Item Env:PSFOUNDATION_TEST_LOG
    }
  }
}

Describe 'Resolve-LongPath' {
  It 'resolves an existing literal path containing spaces and brackets' {
    $directory = Join-Path $TestDrive 'Long directory [literal]'
    $null = [IO.Directory]::CreateDirectory($directory)
    $path = Join-Path $directory 'long filename.txt'
    [IO.File]::WriteAllText($path, 'fixture')
    Resolve-LongPath -LiteralPath $path | Should -Be $path
  }

  It 'rejects a missing path without creating it' {
    $path = Join-Path $TestDrive 'missing.txt'
    { Resolve-LongPath -LiteralPath $path } | Should -Throw
    Test-Path -LiteralPath $path | Should -BeFalse
  }

  It 'rejects non-filesystem provider paths' {
    { Resolve-LongPath -LiteralPath Env:TEMP } | Should -Throw '*filesystem*'
  }
}
