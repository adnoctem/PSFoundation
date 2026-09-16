#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'Test fixtures construct in-memory credentials with placeholder passwords; no real secrets involved.')]
param()

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/user.ps1
  . $PSScriptRoot/../src/permissions.ps1
}

Describe 'Test-Elevation' {
  It 'returns a boolean on Windows' {
    $result = Test-Elevation
    $result | Should -BeOfType [bool]
  }

  It 'does not throw on supported platform' {
    { Test-Elevation } | Should -Not -Throw
  }
}

Describe 'Request-AdministratorPrivilege' {
  BeforeAll {
    $hostExecutable = (Get-Process -Id $PID).Path
    $fixture = Join-Path $PSScriptRoot 'fixtures/elevation/Invoke-ElevationFixture.ps1'
    $source = Join-Path $PSScriptRoot '../src/permissions.ps1'
    # Keep source filenames space-free; exercise spaced paths only in TestDrive.
    $targetDirectory = New-Item -Path (Join-Path $TestDrive 'elevation path with spaces') -ItemType Directory
    $targetScript = Join-Path $targetDirectory.FullName 'ElevationTarget.ps1'
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'fixtures/elevation/ElevationTarget.ps1') -Destination $targetScript
    function Invoke-ElevationTest {
      param ([string]$Kind)
      $report = Join-Path $TestDrive "$Kind-result.xml"
      $launchReport = Join-Path $TestDrive "$Kind-launch.xml"
      # PS 5.1 turns native stderr into ErrorRecords. Capture expected failures
      # without Pester's Stop preference aborting the parent test harness.
      $ErrorActionPreference = 'Continue'
      $output = & $hostExecutable -NoProfile -ExecutionPolicy Bypass -File $fixture -Source $source -Report $report -LaunchReport $launchReport -Kind $Kind -Target $targetScript 2>&1
      [PSCustomObject]@{
        ExitCode = $LASTEXITCODE
        Output = $output | Out-String
        Launch = if (Test-Path -LiteralPath $launchReport) { Import-Clixml -LiteralPath $launchReport } else { $null }
        Result = if (Test-Path -LiteralPath $report) { Import-Clixml -LiteralPath $report } else { $null }
      }
    }
  }

  It 'accepts <Kind> dictionaries and always injects one true marker' -ForEach @(
    @{ Kind = 'Bound' }, @{ Kind = 'Ordered' }, @{ Kind = 'Hashtable' }, @{ Kind = 'Empty' }, @{ Kind = 'Null' }
  ) {
    $result = Invoke-ElevationTest -Kind $Kind
    $result.ExitCode | Should -Be 37
    $result.Launch.Verb | Should -Be 'RunAs'
    $result.Launch.HostPath | Should -Be $hostExecutable
    $result.Launch.Directory | Should -Be (Get-Location).Path
    $result.Launch.PassThru | Should -BeTrue
    $result.Launch.Wait | Should -BeTrue
    $command = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($result.Launch.Arguments[-1]))
    ([regex]::Matches($command, "'Elevated'=\`$true")).Count | Should -Be 1
  }

  It 'round-trips values through a real child PowerShell process' {
    $result = Invoke-ElevationTest -Kind Execute
    $result.ExitCode | Should -Be 37
    $result.Result.Bound.Payload | Should -Be "spaces ' quotes `" and `$([throw]) ; #"
    $smartQuotes = 'smart ' + [char]0x2019 + "; throw 'untrusted' # " + [char]0x2018 + [char]0x201a + [char]0x201b
    $result.Result.Bound.Items | Should -Be @('one', 'two words', '', $smartQuotes)
    [bool]$result.Result.Bound.Enabled | Should -BeFalse
    $result.Result.Bound.Flag | Should -BeFalse
    $result.Result.Bound.ContainsKey('Optional') | Should -BeTrue
    $result.Result.Bound.Optional | Should -BeNullOrEmpty
    [bool]$result.Result.Bound.Elevated | Should -BeTrue
    $result.Result.Extra | Should -Be @('loose value', "quote'`"`$()", '')
    $result.Result.Directory | Should -Be (Get-Location).Path
    $result.Result.ScriptPath | Should -Be $targetScript
  }

  It 'returns without launching when already elevated' {
    $result = Invoke-ElevationTest -Kind Already
    $result.ExitCode | Should -Be 0
    $result.Launch | Should -BeNullOrEmpty
  }

  It 'stops a relaunch loop without launching' {
    $result = Invoke-ElevationTest -Kind Loop
    $result.ExitCode | Should -Be 1
    $result.Launch | Should -BeNullOrEmpty
    $result.Output | Should -Match 're-launch loop'
  }

  It 'propagates UAC cancellation even with ErrorActionPreference Stop' {
    $result = Invoke-ElevationTest -Kind Cancelled
    $result.ExitCode | Should -Be 1223
    $result.Output | Should -Match 'cancelled'
  }

  It 'does not disguise other process launch errors as cancellation' {
    $result = Invoke-ElevationTest -Kind LaunchFailure
    $result.ExitCode | Should -Be 1
    $result.Output | Should -Not -Match 'Elevation was cancelled'
  }

  It 'rejects unsupported argument types before launch' {
    $result = Invoke-ElevationTest -Kind Unsupported
    $result.ExitCode | Should -Be 1
    $result.Launch | Should -BeNullOrEmpty
    $result.Output | Should -Match 'Cannot forward elevation argument'
  }
}

Describe 'Set-RegistryOwner' {
  It 'throws when the process is not elevated' {
    Mock Test-Elevation { return $false }
    { Set-RegistryOwner -Hive 'HKLM' -Key 'SOFTWARE\PSFTest\DoesNotExist' -WhatIf } | Should -Throw
  }

  It 'reports DryRun without touching the registry when elevated with -WhatIf' {
    Mock Test-Elevation { return $true }
    Mock Get-UserSID { return 'S-1-5-32-544' }
    $result = Set-RegistryOwner -Hive 'HKLM' -Key 'SOFTWARE\PSFTest\DoesNotExist' -WhatIf
    $result | Should -BeOfType [PSCustomObject]
    $result.Status | Should -Be 'DryRun'
    $result.Target | Should -Be 'HKLM\SOFTWARE\PSFTest\DoesNotExist'
  }

  It 'rejects an unknown hive' {
    Mock Test-Elevation { return $true }
    Mock Get-UserSID { return 'S-1-5-32-544' }
    { Set-RegistryOwner -Hive 'NOPE' -Key 'x' -WhatIf } | Should -Throw
  }
}

Describe 'Set-ItemOwner' {
  It 'throws when the process is not elevated' {
    Mock Test-Elevation { return $false }
    { Set-ItemOwner -Path "$env:TEMP\psf-test-missing.txt" -WhatIf } | Should -Throw
  }

  It 'reports DryRun when elevated with -WhatIf' {
    Mock Test-Elevation { return $true }
    $testFile = Join-Path $env:TEMP 'psf-itemowner-dryrun.txt'
    Set-Content -LiteralPath $testFile -Value 'test' -Force
    try {
      $result = Set-ItemOwner -Path $testFile -WhatIf
      $result | Should -BeOfType [PSCustomObject]
      $result.Status | Should -Be 'DryRun'
    }
    finally {
      Remove-Item -LiteralPath $testFile -Force -ErrorAction SilentlyContinue
    }
  }

  It 'reports Failed for a path that does not exist' {
    Mock Test-Elevation { return $true }
    $result = Set-ItemOwner -Path "$env:TEMP\psf-test-missing.txt"
    $result | Should -BeOfType [PSCustomObject]
    $result.Status | Should -Be 'Failed'
  }
}

Describe 'New-EncryptedCredentialFile / Get-EncryptedCredentialFile' {
  It 'round-trips a credential through encrypted files' {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'Unavoidable during testing.')]

    $blobFile = Join-Path $env:TEMP 'psf-cred.bin'
    $keyFile = Join-Path $env:TEMP 'psf-cred.key'
    Remove-Item -LiteralPath $blobFile, $keyFile -Force -ErrorAction SilentlyContinue
    try {
      $credential = [System.Management.Automation.PSCredential]::new('DOMAIN\svc-test', (ConvertTo-SecureString 'S3cr3t!' -AsPlainText -Force))
      $writeResult = New-EncryptedCredentialFile -Path $blobFile -KeyPath $keyFile -Credential $credential
      $writeResult.Status | Should -Be 'Completed'
      Test-Path -LiteralPath $blobFile | Should -BeTrue
      Test-Path -LiteralPath $keyFile | Should -BeTrue

      $readBack = Get-EncryptedCredentialFile -Path $blobFile -KeyPath $keyFile
      $readBack | Should -BeOfType [System.Management.Automation.PSCredential]
      $readBack.UserName | Should -Be 'DOMAIN\svc-test'
      [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($readBack.Password)) | Should -Be 'S3cr3t!'
    }
    finally {
      Remove-Item -LiteralPath $blobFile, $keyFile -Force -ErrorAction SilentlyContinue
    }
  }

  It 'reports DryRun with -WhatIf and writes nothing' {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification = 'Unavoidable during testing.')]

    $blobFile = Join-Path $env:TEMP 'psf-cred-dry.bin'
    $keyFile = Join-Path $env:TEMP 'psf-cred-dry.key'
    Remove-Item -LiteralPath $blobFile, $keyFile -Force -ErrorAction SilentlyContinue
    try {
      $credential = [System.Management.Automation.PSCredential]::new('DOMAIN\svc-test', (ConvertTo-SecureString 'x' -AsPlainText -Force))
      $result = New-EncryptedCredentialFile -Path $blobFile -KeyPath $keyFile -Credential $credential -WhatIf
      $result.Status | Should -Be 'DryRun'
      Test-Path -LiteralPath $blobFile | Should -BeFalse
    }
    finally {
      Remove-Item -LiteralPath $blobFile, $keyFile -Force -ErrorAction SilentlyContinue
    }
  }

  It 'throws when neither credential nor username is supplied' {
    { New-EncryptedCredentialFile -Path 'x' -KeyPath 'y' } | Should -Throw
  }

  It 'throws when the key file is missing' {
    { Get-EncryptedCredentialFile -Path 'x' -KeyPath 'y' } | Should -Throw
  }
}
