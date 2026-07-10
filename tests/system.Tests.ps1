#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/system.ps1
}

Describe 'Get-Hostname' {
  It 'returns an object with a non-empty Hostname string' {
    $result = Get-Hostname
    $result | Should -BeOfType [PSCustomObject]
    $result.Hostname | Should -BeOfType [string]
    $result.Hostname | Should -Not -BeNullOrEmpty
  }
}

Describe 'Get-OSBuildNumber' {
  It 'returns a positive integer' {
    $result = Get-OSBuildNumber
    $result | Should -BeOfType [int]
    $result | Should -BeGreaterThan 0
  }
}

Describe 'Get-OSDisplayVersion' {
  It 'returns a non-empty string' {
    $result = Get-OSDisplayVersion
    $result | Should -BeOfType [string]
    $result | Should -Not -BeNullOrEmpty
  }
}

Describe 'Get-OSEdition' {
  It 'returns a non-empty string' {
    $result = Get-OSEdition
    $result | Should -BeOfType [string]
    $result | Should -Not -BeNullOrEmpty
  }
}

Describe 'Get-OSProductName' {
  It 'returns a non-empty string containing Windows or Server' {
    $result = Get-OSProductName
    $result | Should -BeOfType [string]
    $result | Should -Not -BeNullOrEmpty
    $result | Should -Match 'Windows|Server'
  }
}

Describe 'Get-OSVersionInfo' {
  It 'returns a PSCustomObject with expected keys' {
    $result = Get-OSVersionInfo
    $result | Should -BeOfType [PSCustomObject]
    $result.ProductName | Should -Not -BeNullOrEmpty
    $result.EditionID | Should -Not -BeNullOrEmpty
    $result.CurrentBuild | Should -BeGreaterThan 0
  }
}

Describe 'Get-SystemInfo' {
  It 'returns a PSCustomObject with expected system keys' {
    $result = Get-SystemInfo
    $result | Should -BeOfType [PSCustomObject]
    $result.Hostname | Should -Not -BeNullOrEmpty
    $result.OSProductName | Should -Not -BeNullOrEmpty
    $result.OSBuild | Should -BeGreaterThan 0
    $result.TotalMemoryGiB | Should -BeGreaterThan 0
  }
}
