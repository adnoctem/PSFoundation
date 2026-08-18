#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/provisioning.ps1
}

Describe 'New-DjoinFile' {
  It 'writes a blob in the djoin-compatible format' {
    $destFile = Join-Path $env:TEMP 'psf-djoin-test.djoin'
    Remove-Item -LiteralPath $destFile -Force -ErrorAction SilentlyContinue
    try {
      $result = New-DjoinFile -Blob 'PSF-TEST-BLOB' -DestinationFile $destFile
      $result.Status | Should -Be 'Completed'
      Test-Path -LiteralPath $destFile | Should -BeTrue

      $bytes = [System.IO.File]::ReadAllBytes($destFile)
      $bytes[0] | Should -Be 255
      $bytes[1] | Should -Be 254
      $content = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 4)
      $content | Should -Be 'PSF-TEST-BLOB'
      $bytes[$bytes.Length - 1] | Should -Be 0
      $bytes[$bytes.Length - 2] | Should -Be 0
    }
    finally {
      Remove-Item -LiteralPath $destFile -Force -ErrorAction SilentlyContinue
    }
  }

  It 'reports DryRun with -WhatIf and writes nothing' {
    $destFile = Join-Path $env:TEMP 'psf-djoin-dry.djoin'
    Remove-Item -LiteralPath $destFile -Force -ErrorAction SilentlyContinue
    try {
      $result = New-DjoinFile -Blob 'BLOB' -DestinationFile $destFile -WhatIf
      $result.Status | Should -Be 'DryRun'
      Test-Path -LiteralPath $destFile | Should -BeFalse
    }
    finally {
      Remove-Item -LiteralPath $destFile -Force -ErrorAction SilentlyContinue
    }
  }
}

Describe 'New-OfflineDomainJoinBlob' {
  It 'reports DryRun with -WhatIf without provisioning anything' {
    $testComputerName = 'PSFTEST01'
    $result = New-OfflineDomainJoinBlob -ComputerName $testComputerName -Domain 'contoso.com' -WhatIf
    $result | Should -BeOfType [PSCustomObject]
    $result.Status | Should -Be 'DryRun'
    $result.Target | Should -Be $testComputerName
  }

  It 'declares required parameters as mandatory' {
    $command = Get-Command New-OfflineDomainJoinBlob
    foreach ($requiredParameter in 'ComputerName', 'Domain') {
      $parameter = $command.Parameters[$requiredParameter]
      $mandatory = ($parameter.Attributes |
          Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] }).Mandatory
      $mandatory | Should -BeTrue
    }
  }
}
