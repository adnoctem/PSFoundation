#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/errors.ps1
  . $PSScriptRoot/../src/security.ps1
  . $PSScriptRoot/../src/packages.ps1
}

Describe 'Install-Win32Program outcomes' {
  BeforeEach {
    $installer = Join-Path $TestDrive 'Synthetic.msi'
    Set-Content -LiteralPath $installer -Value 'Mocked installer; never executed.'
    Mock Start-Process { [PSCustomObject]@{ ExitCode = 0; Id = 123 } }
  }

  It 'checks the exit code even without the legacy PassThru switch' {
    Mock Start-Process { [PSCustomObject]@{ ExitCode = 1603; Id = 123 } }
    $result = Install-Win32Program -Path $installer
    $result.Status | Should -Be 'Failed'
    $result.ExitCode | Should -Be 1603
    $result.Succeeded | Should -BeFalse
    $result.Error | Should -Match '1603'
    $result.ErrorTranslation.Domain | Should -Be 'Msi'
    Should -Invoke Start-Process -Times 1 -ParameterFilter { $Wait -and $PassThru }
  }

  It 'reports reboot success separately and preserves PassThru status' {
    Mock Start-Process { [PSCustomObject]@{ ExitCode = 3010; Id = 123 } }
    $result = Install-Win32Program -Path $installer -PassThru -RunId 'batch-1'
    $result.Status | Should -Be 'ExitCode:3010'
    $result.RebootRequired | Should -BeTrue
    $result.Succeeded | Should -BeTrue
    $result.RunId | Should -Be 'batch-1'
  }

  It 'quotes each installer argument and preserves empty strings' {
    $null = Install-Win32Program -Path $installer -ArgumentList @('/quiet', 'NAME=two words', '', 'C:\path with spaces\')
    Should -Invoke Start-Process -Times 1 -ParameterFilter { $ArgumentList -eq '"/quiet" "NAME=two words" "" "C:\path with spaces\\"' }
  }

  It 'does not claim installation completed when NoWait is selected' {
    $result = Install-Win32Program -Path $installer -NoWait
    $result.Status | Should -Be 'Started'
    $result.ProcessId | Should -Be 123
    $result.PSObject.Properties.Name | Should -Not -Contain 'ExitCode'
    Should -Invoke Start-Process -Times 1 -ParameterFilter { -not $Wait }
  }

  It 'preserves caught errors and performs no process start during WhatIf' {
    $preview = Install-Win32Program -Path $installer -WhatIf -RunId 'preview-run'
    $preview.Changed | Should -BeFalse
    $preview.RunId | Should -Be 'preview-run'
    Should -Invoke Start-Process -Times 0
    Mock Start-Process { throw 'Synthetic start failure' }
    $result = Install-Win32Program -Path $installer -RunId 'failed-run'
    $result.Status | Should -Be 'Failed'
    $result.ErrorRecord | Should -BeOfType [System.Management.Automation.ErrorRecord]
    $result.Error | Should -Be 'Synthetic start failure'
    $result.RunId | Should -Be 'failed-run'
  }
}

Describe 'Package discovery and error guidance' {
  It 'retains the original Appx failure and adds guidance' {
    try { throw 'Synthetic 0x80073D02 failure' }
    catch { $result = New-PackageLifecycleResult -Target 'Synthetic' -Source 'UPFAppxPackage' -Action 'Install' -Status 'Failed' -ErrorRecord $_ }
    $result.Error | Should -Be 'Synthetic 0x80073D02 failure'
    $result.ErrorTranslation.Code | Should -Be '0x80073D02'
    $result.ErrorTranslation.Benign | Should -BeFalse
  }

  It 'reports an unavailable WinGet provider without importing it' {
    Mock Get-Module { $null } -ParameterFilter { $Name -eq 'Microsoft.WinGet.Client' }
    Mock Import-Module { throw 'Must not import' }
    (Install-Win32ProgramFromWinGet -Id 'Synthetic.Package').SkippedReason | Should -Be 'Microsoft.WinGet.ClientUnavailable'
    Should -Invoke Import-Module -Times 0
  }

  It 'reports an unavailable update command without attempting an upgrade' {
    Mock Get-Module { [PSCustomObject]@{ Name = 'Microsoft.WinGet.Client' } }
    Mock Import-Module { }
    Mock Get-Command { $null } -ParameterFilter { $Name -eq 'Update-WinGetPackage' }
    (Update-Win32ProgramFromWinGet -Id 'Synthetic.Package').SkippedReason | Should -Be 'UpdateWinGetPackageUnavailable'
  }
}
