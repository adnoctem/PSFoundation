#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll { . $PSScriptRoot/../src/maintenance.ps1 }

Describe 'Module maintenance safety and failures' {
  BeforeEach {
    Mock Write-Log { }
    Mock Ensure-PSResourceGet { }
    Mock Test-PSResourceGetAvailable { $false }
    Mock Install-PackageProvider { }
    Mock Install-Module { }
    Mock Uninstall-Module { }
  }

  It 'does not bootstrap providers or install modules during WhatIf' {
    Add-PSModule -Name 'Synthetic' -WhatIf
    Should -Invoke Ensure-PSResourceGet -Times 0
    Should -Invoke Install-PackageProvider -Times 0
    Should -Invoke Install-Module -Times 0
  }

  It 'rejects conflicting version requirements before bootstrapping' {
    { Add-PSModule -Name 'Synthetic' -Version '1.0' -MinimumVersion '2.0' } | Should -Throw '*mutually exclusive*'
    Should -Invoke Install-PackageProvider -Times 0
  }

  It 'forwards exact and minimum versions to PowerShellGet' {
    Add-PSModule -Name 'Synthetic' -Version '1.2.3'
    Should -Invoke Install-Module -Times 1 -ParameterFilter { $RequiredVersion -eq '1.2.3' -and -not $MinimumVersion }
    Add-PSModule -Name 'Synthetic' -MinimumVersion '2.0.0'
    Should -Invoke Install-Module -Times 1 -ParameterFilter { $MinimumVersion -eq '2.0.0' -and -not $RequiredVersion }
  }

  It 'reports a failed restore entry and continues with later entries' {
    $path = Join-Path $TestDrive 'modules.json'
    @(@{ Name = 'Broken'; Version = '1.0' }, @{ Name = 'Working'; Version = '2.0' }) | ConvertTo-Json | Set-Content -LiteralPath $path
    Mock Install-Module { Write-Error 'Synthetic provider failure' } -ParameterFilter { $Name -eq 'Broken' }
    Add-PSModule -FromFile $path -WarningVariable warnings -WarningAction SilentlyContinue
    ($warnings -join ' ') | Should -Match 'Failed to install Broken'
    Should -Invoke Install-Module -Times 1 -ParameterFilter { $Name -eq 'Working' }
  }

  It 'does not bootstrap or uninstall while previewing removal' {
    Mock Get-Module {
      [PSCustomObject]@{ Name = 'Synthetic'; Version = [version]'1.0'; ModuleType = 'Script'; ModuleBase = 'C:\SyntheticModules\Synthetic\1.0'; Path = 'C:\SyntheticModules\Synthetic\1.0\Synthetic.psd1' }
      [PSCustomObject]@{ Name = 'Synthetic'; Version = [version]'2.0'; ModuleType = 'Script'; ModuleBase = 'C:\SyntheticModules\Synthetic\2.0'; Path = 'C:\SyntheticModules\Synthetic\2.0\Synthetic.psd1' }
    }
    Mock Test-Path { $true }
    Remove-PSModule -Path 'C:\SyntheticModules' -WhatIf
    Should -Invoke Ensure-PSResourceGet -Times 0
    Should -Invoke Uninstall-Module -Times 0
  }
}
