#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/errors.ps1
  function New-TestErrorRecord {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Constructs an in-memory error fixture only.')]
    param ([string]$Message, [Exception]$Exception)
    if ($null -eq $Exception) { $Exception = [Exception]::new($Message) }
    [Management.Automation.ErrorRecord]::new($Exception, 'TestError', [Management.Automation.ErrorCategory]::NotSpecified, $null)
  }
}

Describe 'Get-ErrorTranslation' {
  It 'never treats the AppX package-open failure as already installed' {
    $result = Get-ErrorTranslation -Code '0x80073CF0' -Domain Appx
    $result.Code | Should -Be '0x80073CF0'
    $result.Domain | Should -Be 'Appx'
    $result.Matched | Should -BeTrue
    $result.Benign | Should -BeFalse
    $result.Detail | Should -Match 'could not be opened'
  }

  It 'normalizes signed, unsigned, and hexadecimal HRESULTs' {
    foreach ($code in @('0x80073d06', -2147009274L, 2147958022L)) {
      $result = Get-ErrorTranslation -Code $code -Domain Appx
      $result.Code | Should -Be '0x80073D06'
      $result.Benign | Should -BeFalse
    }
  }

  It 'recognizes <Domain> code <Code> with the expected benign classification' -ForEach @(
    @{ Domain = 'Winget'; Code = '0x8A150011'; Benign = $false },
    @{ Domain = 'Winget'; Code = '0x8A15002B'; Benign = $false },
    @{ Domain = 'Dism'; Code = '0x800F081F'; Benign = $false },
    @{ Domain = 'Dism'; Code = '0x800F0906'; Benign = $false },
    @{ Domain = 'Msi'; Code = 1603; Benign = $false },
    @{ Domain = 'Msi'; Code = 1618; Benign = $false },
    @{ Domain = 'Msi'; Code = 1638; Benign = $false },
    @{ Domain = 'Msi'; Code = 0; Benign = $true },
    @{ Domain = 'Msi'; Code = 1641; Benign = $true },
    @{ Domain = 'Msi'; Code = 3010; Benign = $true }
  ) {
    $result = Get-ErrorTranslation -Code $Code -Domain $Domain
    $result.Domain | Should -Be $Domain
    $result.Benign | Should -Be $Benign
  }

  It 'keeps reboot guidance even for successful installer results' {
    (Get-ErrorTranslation -Code 3010 -Domain Msi).Detail | Should -Match 'restart is required'
    (Get-ErrorTranslation -Code 1641 -Domain Msi).Detail | Should -Match 'initiated a restart'
  }

  It 'returns no output for unknown codes and wrong domains' {
    Get-ErrorTranslation -Code '0xDEADBEEF' -Domain Appx | Should -BeNullOrEmpty
    Get-ErrorTranslation -Code '0x80073CF0' -Domain Msi | Should -BeNullOrEmpty
    Get-ErrorTranslation -Code 3010 -Domain Winget | Should -BeNullOrEmpty
  }

  It 'rejects malformed, fractional, and out-of-range codes' {
    foreach ($code in @('oops', 1.5, 4294967296L, -2147483649L, '0x100000000')) {
      { Get-ErrorTranslation -Code $code -Domain Msi } | Should -Throw
    }
  }

  It 'recognizes codes in localized messages and deduplicates repeats' {
    $record = New-TestErrorRecord -Message 'Fehler 0x80073CF0; détail: 0x80073cf0'
    ($record | Get-ErrorTranslation).Code | Should -Be '0x80073CF0'
    $record = New-TestErrorRecord -Message 'HRESULT: -2147009274'
    (Get-ErrorTranslation -ErrorRecord $record).Code | Should -Be '0x80073D06'
  }

  It 'uses structured HRESULTs before message codes, including nested exceptions' {
    $inner = [Runtime.InteropServices.COMException]::new('0x80073CF0', -2147009274)
    $outer = [Exception]::new('Wrapper with 0x80073CF9', $inner)
    $record = New-TestErrorRecord -Exception $outer
    (Get-ErrorTranslation -ErrorRecord $record).Code | Should -Be '0x80073D06'
  }

  It 'reads ErrorDetails and refuses ambiguous message matches' {
    $record = New-TestErrorRecord -Message 'Wrapper'
    $record.ErrorDetails = [Management.Automation.ErrorDetails]::new('0x800F081F')
    (Get-ErrorTranslation -ErrorRecord $record).Domain | Should -Be Dism
    $record = New-TestErrorRecord -Message '0x80073CF0 and 0x800F081F'
    Get-ErrorTranslation -ErrorRecord $record | Should -BeNullOrEmpty
    (Get-ErrorTranslation -ErrorRecord $record -Domain Appx).Code | Should -Be '0x80073CF0'
  }

  It 'does not match partial tokens or arbitrary decimal counts' {
    foreach ($message in @('file0x80073CF0', '0x80073CF01', '1603 objects', '3010 ms', 'ordinary failure')) {
      Get-ErrorTranslation -ErrorRecord (New-TestErrorRecord -Message $message) | Should -BeNullOrEmpty
    }
  }

  It 'processes records after unknown errors in a pipeline' {
    $records = @((New-TestErrorRecord -Message 'Unknown'), (New-TestErrorRecord -Message '0x80073D02'))
    $results = @($records | Get-ErrorTranslation)
    $results.Count | Should -Be 1
    $results[0].Code | Should -Be '0x80073D02'
  }
}
