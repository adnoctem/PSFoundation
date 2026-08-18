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
