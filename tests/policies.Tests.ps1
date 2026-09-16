#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/policies.ps1
}

Describe 'Registry policy conversion' {
  BeforeAll {
    $fixtureHex = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'fixtures/policies/ordered-dword.hex') -Raw
    $fixtureBytes = [byte[]]@($fixtureHex.Trim() -split '\s+' | ForEach-Object { [Convert]::ToByte($_, 16) })
    $fixturePath = Join-Path $TestDrive 'fixture.pol'
    [IO.File]::WriteAllBytes($fixturePath, $fixtureBytes)
  }

  It 'reads an independently authored fixture with duplicate records in order' {
    $entries = @(ConvertFrom-RegistryPolicy -Path $fixturePath)
    $entries.Count | Should -Be 2
    $entries[0].Key | Should -Be 'Software\Test'
    $entries[0].ValueName | Should -Be 'Flag'
    $entries[0].Type | Should -Be 4
    $entries[0].Data | Should -Be ([uint32]::MaxValue)
    $entries[1].Data | Should -Be 1
  }

  It 'reproduces independent fixture bytes from raw and decoded records' {
    foreach ($raw in @($true, $false)) {
      $destination = Join-Path $TestDrive "roundtrip-$raw.pol"
      ConvertFrom-RegistryPolicy -Path $fixturePath -Raw:$raw | ConvertTo-RegistryPolicy -Path $destination
      [IO.File]::ReadAllBytes($destination) | Should -Be $fixtureBytes
    }
  }

  It 'round-trips <Label> payloads without expanding variables or losing values' -ForEach @(
    @{ Label = 'Unicode'; Type = 1; Data = 'Grüße 世界' },
    @{ Label = 'empty string'; Type = 1; Data = '' },
    @{ Label = 'expandable string'; Type = 2; Data = '%SystemRoot%\test' },
    @{ Label = 'binary delimiters'; Type = 3; Data = [byte[]]@(91, 0, 59, 0, 93, 0, 255) },
    @{ Label = 'DWORD'; Type = 4; Data = [uint32]::MaxValue },
    @{ Label = 'big-endian DWORD'; Type = 5; Data = [uint32]305419896 },
    @{ Label = 'multiple strings'; Type = 7; Data = [string[]]@('one', '世界') },
    @{ Label = 'empty multiple strings'; Type = 7; Data = [string[]]@() },
    @{ Label = 'QWORD'; Type = 11; Data = [uint64]::MaxValue },
    @{ Label = 'opaque type'; Type = 42; Data = [byte[]]@(1, 2, 255) },
    @{ Label = 'key-only'; Type = 0; Data = $null }
  ) {
    $destination = Join-Path $TestDrive "$Label.pol"
    $entry = [PSCustomObject]@{ Key = 'Software\Policies\Nested\PackageFamily'; ValueName = ''; Type = $Type; Data = $Data }
    ConvertTo-RegistryPolicy -InputObject $entry -Path $destination
    $result = ConvertFrom-RegistryPolicy -Path $destination
    $result.Key | Should -Be $entry.Key
    $result.ValueName | Should -Be ''
    $result.Type | Should -Be $Type
    $result.Data | Should -Be $Data
  }

  It 'preserves directive names and order without interpreting them' {
    $entries = @(
      [PSCustomObject]@{ Key = 'Software\Test'; ValueName = '**Del.Flag'; Type = 1; Data = '' },
      [PSCustomObject]@{ Key = 'Software\Test'; ValueName = 'Flag'; Type = 4; Data = 1 },
      [PSCustomObject]@{ Key = 'Software\Test'; ValueName = '**SecureKey'; Type = 4; Data = 1 }
    )
    $destination = Join-Path $TestDrive 'directives.pol'
    $entries | ConvertTo-RegistryPolicy -Path $destination
    @(ConvertFrom-RegistryPolicy -Path $destination).ValueName | Should -Be @('**Del.Flag', 'Flag', '**SecureKey')
  }

  It 'writes and reads an empty policy, including an empty pipeline' {
    $destination = Join-Path $TestDrive 'empty.pol'
    @() | ConvertTo-RegistryPolicy -Path $destination
    [IO.File]::ReadAllBytes($destination) | Should -Be ([byte[]]@(80, 82, 101, 103, 1, 0, 0, 0))
    @(ConvertFrom-RegistryPolicy -Path $destination).Count | Should -Be 0
  }

  It 'keeps raw zero-length and unknown payloads as byte arrays' {
    $destination = Join-Path $TestDrive 'raw-empty.pol'
    ConvertTo-RegistryPolicy -InputObject ([PSCustomObject]@{ Key = 'K'; ValueName = ''; Type = 42; Data = [byte[]]@() }) -Path $destination
    $entry = ConvertFrom-RegistryPolicy -Path $destination -Raw
    $entry.Data.GetType() | Should -Be ([byte[]])
    $entry.PSObject.TypeNames | Should -Contain 'PSFoundation.RegistryPolicy.RawEntry'
    $copy = Join-Path $TestDrive 'raw-copy.pol'
    $entry | ConvertTo-RegistryPolicy -Path $copy
    [IO.File]::ReadAllBytes($copy) | Should -Be ([IO.File]::ReadAllBytes($destination))
  }

  It 'rejects bad signature, version, delimiters, and truncated input with offsets' {
    foreach ($index in @(0, 4, 8)) {
      $broken = [byte[]]$fixtureBytes.Clone()
      $broken[$index] = 255
      $path = Join-Path $TestDrive "bad-$index.pol"
      [IO.File]::WriteAllBytes($path, $broken)
      { ConvertFrom-RegistryPolicy -Path $path } | Should -Throw '*at byte*'
    }
    foreach ($length in @(0, 3, 7, 9, 20, ($fixtureBytes.Length - 1))) {
      $path = Join-Path $TestDrive "truncated-$length.pol"
      $broken = [byte[]]::new($length)
      [array]::Copy($fixtureBytes, $broken, $length)
      [IO.File]::WriteAllBytes($path, $broken)
      { ConvertFrom-RegistryPolicy -Path $path } | Should -Throw '*at byte*'
    }
  }

  It 'validates decoded data while raw mode preserves uninterpreted payload bytes' {
    $path = Join-Path $TestDrive 'malformed-string.pol'
    $entry = [PSCustomObject]@{ Key = 'K'; ValueName = 'V'; Type = 1; Data = [byte[]]@(65) }
    $entry.PSObject.TypeNames.Insert(0, 'PSFoundation.RegistryPolicy.RawEntry')
    $entry | ConvertTo-RegistryPolicy -Path $path
    { ConvertFrom-RegistryPolicy -Path $path } | Should -Throw '*odd byte count*'
    (ConvertFrom-RegistryPolicy -Path $path -Raw).Data | Should -Be ([byte[]]@(65))
  }

  It 'rejects invalid records before replacing an existing destination' {
    $path = Join-Path $TestDrive 'protected.pol'
    [IO.File]::WriteAllBytes($path, $fixtureBytes)
    $invalid = @(
      [PSCustomObject]@{ Key = 'HKLM\Software'; ValueName = 'V'; Type = 4; Data = 1 },
      [PSCustomObject]@{ Key = 'K'; ValueName = 'V'; Type = 4; Data = 1.5 },
      [PSCustomObject]@{ Key = 'K'; ValueName = 'V'; Type = 4; Data = -1 },
      [PSCustomObject]@{ Key = 'K'; ValueName = 'V'; Type = 4; Data = [uint64]::MaxValue },
      [PSCustomObject]@{ Key = 'K'; ValueName = 'V'; Type = 3; Data = [byte[]]::new(65536) },
      [PSCustomObject]@{ Key = 'K'; ValueName = 'V'; Type = 1; Data = "bad`0value" },
      [PSCustomObject]@{ Key = 'K'; ValueName = 'V'; Type = 1 }
    )
    foreach ($entry in $invalid) {
      { $entry | ConvertTo-RegistryPolicy -Path $path -Force } | Should -Throw
      [IO.File]::ReadAllBytes($path) | Should -Be $fixtureBytes
    }
  }

  It 'requires Force to overwrite and leaves no temporary file after replacement' {
    $folder = New-Item -Path (Join-Path $TestDrive 'replace') -ItemType Directory
    $path = Join-Path $folder.FullName 'policy.pol'
    [IO.File]::WriteAllBytes($path, $fixtureBytes)
    { ConvertTo-RegistryPolicy -Path $path } | Should -Throw '*-Force*'
    ConvertTo-RegistryPolicy -Path $path -Force
    @(ConvertFrom-RegistryPolicy -Path $path).Count | Should -Be 0
    @(Get-ChildItem -LiteralPath $folder.FullName).Count | Should -Be 1
  }

  It 'honors WhatIf for new and existing destinations' {
    $path = Join-Path $TestDrive 'whatif.pol'
    ConvertTo-RegistryPolicy -Path $path -WhatIf
    Test-Path -LiteralPath $path | Should -BeFalse
    [IO.File]::WriteAllBytes($path, $fixtureBytes)
    ConvertTo-RegistryPolicy -Path $path -Force -WhatIf
    [IO.File]::ReadAllBytes($path) | Should -Be $fixtureBytes
  }

  It 'preserves the destination and removes staging files if replacement fails' {
    $folder = New-Item -Path (Join-Path $TestDrive 'locked') -ItemType Directory
    $path = Join-Path $folder.FullName 'policy.pol'
    [IO.File]::WriteAllBytes($path, $fixtureBytes)
    $lock = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
      { ConvertTo-RegistryPolicy -Path $path -Force } | Should -Throw
      [IO.File]::ReadAllBytes($path) | Should -Be $fixtureBytes
      @(Get-ChildItem -LiteralPath $folder.FullName).Count | Should -Be 1
    }
    finally { $lock.Dispose() }
  }
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
