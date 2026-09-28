#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll { . $PSScriptRoot/../src/maintenance.ps1 }

Describe 'Imported module collection contracts' {
  BeforeAll {
    Remove-Module -Name PSFoundation -Force -ErrorAction SilentlyContinue
    Import-Module "$PSScriptRoot/../src/PSFoundation.psd1" -Force
  }

  BeforeEach {
    Mock Write-Log { } -ModuleName PSFoundation
    Mock Test-PSResourceGetAvailable { $false } -ModuleName PSFoundation
    Mock Ensure-PSResourceGet { } -ModuleName PSFoundation
    Mock Resolve-ModuleDirectory { 'C:\SyntheticModules' } -ModuleName PSFoundation
    Mock Test-Path { $true } -ModuleName PSFoundation
    Mock Get-Module { $script:exportModules } -ModuleName PSFoundation -ParameterFilter { $ListAvailable }
    Mock Get-Item { [PSCustomObject]@{ CreationTime = [datetime]'2026-01-01' } } -ModuleName PSFoundation
    Mock Install-PackageProvider { } -ModuleName PSFoundation
    Mock Install-Module { } -ModuleName PSFoundation
  }

  It 'round-trips <Number> exported modules as a JSON array under module strict mode' -ForEach @(
    @{ Number = 0 }, @{ Number = 1 }, @{ Number = 3 }
  ) {
    $script:exportModules = @(for ($i = 0; $i -lt $Number; $i++) {
        [PSCustomObject]@{
          Name                     = "Synthetic$i"
          Version                  = [version]'1.0'
          ModuleType               = 'Script'
          ModuleBase               = "C:\SyntheticModules\Synthetic$i"
          Path                     = "C:\SyntheticModules\Synthetic$i\Synthetic$i.psd1"
          RepositorySourceLocation = 'Synthetic'
        }
      })
    $path = Join-Path $TestDrive 'modules.json'
    PSFoundation\Get-PSModule -Path $path
    $json = Get-Content -LiteralPath $path -Raw
    $json.Trim() | Should -Match '^\['
    $decoded = ConvertFrom-Json -InputObject $json
    if ($Number -eq 0) { $json.Trim() | Should -Match '^\[\s*\]$' }
    else { @($decoded).Count | Should -Be $Number }
    PSFoundation\Add-PSModule -FromFile $path -Confirm:$false
    Should -Invoke Install-Module -Times $Number -Exactly -ModuleName PSFoundation
    foreach ($entry in $script:exportModules) {
      $expectedName = $entry.Name
      Should -Invoke Install-Module -Times 1 -Exactly -ModuleName PSFoundation -ParameterFilter {
        $Name -eq $expectedName -and $RequiredVersion -eq '1.0' -and $Scope -eq 'CurrentUser'
      }
    }
    Should -Invoke Write-Log -Times 1 -ModuleName PSFoundation -ParameterFilter { $Message -like "Exported $Number module(s)*" }
    Should -Invoke Write-Log -Times 1 -ModuleName PSFoundation -ParameterFilter { $Message -like "Restored $Number module(s)*" }
  }

  It 'previews a legacy single-object export without installing or bootstrapping' {
    $path = Join-Path $TestDrive 'legacy.json'
    '{"Name":"Synthetic","Version":"1.0","Scope":"CurrentUser"}' | Set-Content -LiteralPath $path
    PSFoundation\Add-PSModule -FromFile $path -WhatIf
    Should -Invoke Install-Module -Times 0 -Exactly -ModuleName PSFoundation
    Should -Invoke Install-PackageProvider -Times 0 -Exactly -ModuleName PSFoundation
    Should -Invoke Ensure-PSResourceGet -Times 0 -Exactly -ModuleName PSFoundation
    Should -Invoke Write-Log -Times 1 -ModuleName PSFoundation -ParameterFilter { $Message -like 'Restored 1 module(s)*' }
  }
}

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

Describe 'Remove-PSModule version retention under strict mode' {
  BeforeEach {
    Mock Write-Log { }
    Mock Test-PSResourceGetAvailable { $false }
    Mock Resolve-ModuleDirectory { 'C:\SyntheticModules' }
    Mock Test-Path { $true }
    Mock Get-Module {
      foreach ($version in $script:RetentionVersions) {
        [PSCustomObject]@{
          Name       = 'Synthetic'
          Version    = [version]$version
          ModuleType = 'Script'
          ModuleBase = "C:\SyntheticModules\Synthetic\$version"
          Path       = "C:\SyntheticModules\Synthetic\$version\Synthetic.psd1"
        }
      }
    }
    Mock Uninstall-Module { }
  }

  It 'keeps <Keep> of <Versions> without scalar or null Count errors' -ForEach @(
    @{ Versions = @('1.6.1'); Keep = 1; Removed = @() }
    @{ Versions = @('1.6.1'); Keep = 3; Removed = @() }
    @{ Versions = @('1.6.0', '1.6.1'); Keep = 1; Removed = @('1.6.0') }
    @{ Versions = @('1.5.0', '1.6.0', '1.6.1'); Keep = 1; Removed = @('1.6.0', '1.5.0') }
    @{ Versions = @('1.5.0', '1.6.0', '1.6.1'); Keep = 2; Removed = @('1.5.0') }
  ) {
    $script:RetentionVersions = $Versions
    Set-StrictMode -Version Latest

    { Remove-PSModule -Name '^Synthetic$' -LatestToKeep $Keep -Scope AllUsers -Confirm:$false } | Should -Not -Throw

    Should -Invoke Uninstall-Module -Times $Removed.Count -Exactly
    foreach ($version in $Removed) {
      Should -Invoke Uninstall-Module -Times 1 -Exactly -ParameterFilter {
        $Name -eq 'Synthetic' -and $RequiredVersion -eq $version
      }
    }
    Should -Invoke Uninstall-Module -Times 0 -Exactly -ParameterFilter { $RequiredVersion -eq '1.6.1' }
  }

  It 'does not uninstall a single older version during WhatIf under strict mode' {
    $script:RetentionVersions = @('1.6.0', '1.6.1')
    Set-StrictMode -Version Latest

    { Remove-PSModule -Name '^Synthetic$' -LatestToKeep 1 -Scope AllUsers -WhatIf } | Should -Not -Throw

    Should -Invoke Uninstall-Module -Times 0 -Exactly
  }
}
