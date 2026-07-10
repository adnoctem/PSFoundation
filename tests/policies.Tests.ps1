#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/policies.ps1
}

Describe 'Resolve-LGPOSource' {
  It 'returns a PSCustomObject with the expected metadata fields' {
    $result = Resolve-LGPOSource -Source 'SCT-LGPO-Standalone'
    $result | Should -BeOfType [PSCustomObject]
    $result.Name | Should -Not -BeNullOrEmpty
    $result.Url | Should -Match '^https://'
    $result.Sha256 | Should -Not -BeNullOrEmpty
    $result.ExpectedBinaryPath | Should -Not -BeNullOrEmpty
    $result.LastVerified | Should -Not -BeNullOrEmpty
  }

  It 'throws for an unknown source name' {
    { Resolve-LGPOSource -Source 'NonExistent' } | Should -Throw
  }
}

Describe 'Test-LGPOSourceAvailability' {
  It 'returns a PSCustomObject with Available, Url, and CheckedAt fields' {
    Mock Invoke-WebRequest {
      [PSCustomObject]@{ StatusCode = 200; Headers = @{ 'Content-Length' = '12345' } }
    }

    $result = Test-LGPOSourceAvailability
    $result | Should -BeOfType [PSCustomObject]
    $result.Available | Should -BeTrue
    $result.StatusCode | Should -Be 200
    $result.Url | Should -Match '^https://'
    $result.CheckedAt | Should -BeOfType [datetime]
  }

  It 'returns Available=false when the URL is not reachable' {
    Mock Invoke-WebRequest { throw 'Network error' }

    $result = Test-LGPOSourceAvailability
    $result.Available | Should -BeFalse
    $result.Error | Should -Not -BeNullOrEmpty
    $result.CheckedAt | Should -BeOfType [datetime]
  }
}
