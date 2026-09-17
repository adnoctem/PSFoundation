#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  # Load the real provider's command metadata so mocks enforce its parameters.
  Import-Module PSWindowsUpdate -Prefix PSFProvider -ErrorAction Stop
  . $PSScriptRoot/../src/updates.ps1
  . $PSScriptRoot/../src/log.ps1
}

Describe 'Windows Update provider boundaries' {
  BeforeEach {
    Mock Test-PSWindowsUpdateAvailable { $true }
    Mock Import-Module { }
    Mock PSWindowsUpdate\Get-PSFProviderWindowsUpdate { }
  }

  It 'does not call the installation provider during WhatIf' {
    $update = [PSCustomObject]@{ Identity = @{ UpdateID = '11111111-1111-1111-1111-111111111111' }; Title = 'Synthetic' }
    $update | & (Get-Command Install-WindowsUpdate -CommandType Function) -WhatIf
    Should -Invoke PSWindowsUpdate\Get-PSFProviderWindowsUpdate -Times 0
  }

  It 'preserves partial failure and reboot information from the provider' {
    Mock PSWindowsUpdate\Get-PSFProviderWindowsUpdate {
      [PSCustomObject]@{ KB = 'KB1'; Title = 'First'; HResult = 0; Result = 'Installed'; RebootRequired = $true }
      [PSCustomObject]@{ KB = 'KB2'; Title = 'Second'; HResult = -1; Result = 'Failed'; RebootRequired = $false }
    }
    $update = [PSCustomObject]@{ Identity = @{ UpdateID = '11111111-1111-1111-1111-111111111111' } }
    $results = @($update | & (Get-Command Install-WindowsUpdate -CommandType Function) -IgnoreReboot)
    $results.Count | Should -Be 2
    $results[0].RebootRequired | Should -BeTrue
    $results[1].Result | Should -Be 'Failed'
    Should -Invoke PSWindowsUpdate\Get-PSFProviderWindowsUpdate -Times 1 -ParameterFilter { $Install -and $AcceptAll -and $IgnoreReboot -and $UpdateID.Count -eq 1 -and $UpdateID[0] -eq '11111111-1111-1111-1111-111111111111' }
  }

  It 'rejects missing identities rather than installing unfiltered updates' {
    { [PSCustomObject]@{ Title = 'Unknown' } | & (Get-Command Install-WindowsUpdate -CommandType Function) -ErrorAction Stop } | Should -Throw '*UpdateID*'
    Should -Invoke PSWindowsUpdate\Get-PSFProviderWindowsUpdate -Times 0
  }

  It 'accepts a direct UpdateID and filters queries by the requested categories' {
    [PSCustomObject]@{ UpdateID = '22222222-2222-2222-2222-222222222222' } | & (Get-Command Install-WindowsUpdate -CommandType Function)
    Should -Invoke PSWindowsUpdate\Get-PSFProviderWindowsUpdate -Times 1 -ParameterFilter { $Install -and $UpdateID[0] -eq '22222222-2222-2222-2222-222222222222' }
    Mock PSWindowsUpdate\Get-PSFProviderWindowsUpdate {
      [PSCustomObject]@{ Title = 'Security'; Categories = @([PSCustomObject]@{ Name = 'Security' }) }
      [PSCustomObject]@{ Title = 'Drivers'; Categories = @([PSCustomObject]@{ Name = 'Drivers' }) }
      [PSCustomObject]@{ Title = 'Unknown'; Categories = @() }
    }
    $result = & (Get-Command Get-WindowsUpdate -CommandType Function) -Category 'Security'
    @($result).Count | Should -Be 1
    $result.Title | Should -Be 'Security'
  }

  It 'hides only the selected update and honors WhatIf' {
    $update = [PSCustomObject]@{ KB = 'KB1'; Title = 'Synthetic'; UpdateID = '11111111-1111-1111-1111-111111111111' }
    $update | & (Get-Command Hide-WindowsUpdate -CommandType Function) -WhatIf
    Should -Invoke PSWindowsUpdate\Get-PSFProviderWindowsUpdate -Times 0
    $update | & (Get-Command Hide-WindowsUpdate -CommandType Function) -Confirm:$false
    Should -Invoke PSWindowsUpdate\Get-PSFProviderWindowsUpdate -Times 1 -ParameterFilter { $Hide -and $UpdateID[0] -eq '11111111-1111-1111-1111-111111111111' }
  }

  It 'filters history locally and uses the real uninstall provider contract' {
    Mock PSWindowsUpdate\Get-PSFProviderWUHistory {
      [PSCustomObject]@{ Title = 'Synthetic (KB123)' }
      [PSCustomObject]@{ Title = 'Synthetic (KB1234)' }
    }
    Mock PSWindowsUpdate\Remove-PSFProviderWindowsUpdate { }
    @(Get-WindowsUpdateHistory -KBArticleID 'KB123').Count | Should -Be 1
    & (Get-Command Uninstall-WindowsUpdate -CommandType Function) -KBArticleID 'KB123' -WhatIf
    Should -Invoke PSWindowsUpdate\Remove-PSFProviderWindowsUpdate -Times 0
    & (Get-Command Uninstall-WindowsUpdate -CommandType Function) -KBArticleID 'KB123' -Confirm:$false
    Should -Invoke PSWindowsUpdate\Remove-PSFProviderWindowsUpdate -Times 1 -ParameterFilter { $KBArticleID -eq 'KB123' }
  }

  It 'stops cleanly when the module is unavailable' {
    Mock Test-PSWindowsUpdateAvailable { $false }
    { & (Get-Command Install-WindowsUpdate -CommandType Function) -AcceptAll -ErrorAction Stop } | Should -Throw '*not installed*'
    Should -Invoke PSWindowsUpdate\Get-PSFProviderWindowsUpdate -Times 0
  }
}

Describe 'Store update previews' {
  It 'does not initialize WinRT or request an update during WhatIf' {
    Mock _getAppInstallManager { throw 'Must not initialize WinRT' }
    [PSCustomObject]@{ PackageFamilyName = 'Synthetic.Package' } | Install-MSStoreUpdate -WhatIf
    Should -Invoke _getAppInstallManager -Times 0
  }
}
