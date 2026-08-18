#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/user.ps1
  . $PSScriptRoot/../src/permissions.ps1
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
    $result.InstallationType | Should -Not -BeNullOrEmpty
    $result.CurrentBuild | Should -BeGreaterThan 0
  }
}

Describe 'Test-HostApplicability' {
  It 'returns a boolean' {
    $result = Test-HostApplicability
    $result | Should -BeOfType [bool]
  }

  It 'passes when no constraints are supplied' {
    Test-HostApplicability | Should -BeTrue
  }

  It 'honours a MinBuild constraint below the current build' {
    Test-HostApplicability -MinBuild 1 | Should -BeTrue
  }
}

Describe 'Get-DotNetVersion' {
  It 'returns a structured object' {
    $result = Get-DotNetVersion
    $result | Should -BeOfType [PSCustomObject]
    $result.FrameworkVersion | Should -Not -BeNullOrEmpty
    $result.Releases.Count | Should -BeGreaterThan 0
    foreach ($release in $result.Releases) {
      $release.Product | Should -BeOfType [string]
    }
  }

  It 'maps release numbers to friendly product names' {
    Resolve-NetFrameworkProductName -Release 528449 | Should -Be '4.8 Windows 11 or Windows Server 2022'
    Resolve-NetFrameworkProductName -Release 461808 | Should -Be '4.7.2'
    Resolve-NetFrameworkProductName -Release 394806 | Should -Be '4.6.2'
    Resolve-NetFrameworkProductName -Release 378389 | Should -Be '4.5 Original Release'
    Resolve-NetFrameworkProductName -ChildName 'v3.5' -ServicePack 1 | Should -Be '3.5 ServicePack 1'
  }

  It 'returns null for an unknown release number' {
    Resolve-NetFrameworkProductName -Release 1 | Should -BeNullOrEmpty
  }
}

Describe 'New-DriveMapping' {
  It 'rejects a drive letter that is a physical volume' {
    $physical = (Get-Volume | Where-Object { $_.DriveLetter } | Select-Object -First 1).DriveLetter
    if (-not $physical) {
      { New-DriveMapping -DriveLetter 'C' -Path "$env:TEMP" } | Should -Throw
    }
    else {
      { New-DriveMapping -DriveLetter $physical -Path "$env:TEMP" } | Should -Throw
    }
  }

  It 'rejects a root-level folder path' {
    Mock Test-Elevation { return $true }
    $freeLetter = $null
    foreach ($letter in 'Z', 'Y', 'X', 'W', 'V') {
      if (-not (Get-Volume -DriveLetter $letter -ErrorAction SilentlyContinue)) {
        $freeLetter = $letter
        break
      }
    }
    if (-not $freeLetter) {
      { New-DriveMapping -DriveLetter 'Z' -Path 'C:\' } | Should -Throw -ExpectedMessage '*root level folder*'
    }
    else {
      { New-DriveMapping -DriveLetter $freeLetter -Path 'C:\' } | Should -Throw -ExpectedMessage '*root level folder*'
    }
  }

  It 'throws when the process is not elevated' {
    Mock Test-Elevation { return $false }
    $freeLetter = $null
    foreach ($letter in 'Z', 'Y', 'X', 'W', 'V') {
      if (-not (Get-Volume -DriveLetter $letter -ErrorAction SilentlyContinue)) {
        $freeLetter = $letter
        break
      }
    }
    $testPath = Join-Path $env:TEMP 'psf-drivemap-test'
    New-Item -ItemType Directory -Path $testPath -Force | Out-Null
    try {
      { New-DriveMapping -DriveLetter $freeLetter -Path $testPath } | Should -Throw
    }
    finally {
      Remove-Item -LiteralPath $testPath -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

Describe 'Remove-DriveMapping' {
  It 'rejects a drive letter that is not mapped' {
    { Remove-DriveMapping -DriveLetter 'Q' } | Should -Throw -ExpectedMessage '*does not represent a mapped drive*'
  }

  It 'throws when the process is not elevated for a mapped drive' {
    Mock Test-Elevation { return $false }
    { Remove-DriveMapping -DriveLetter 'Q' -WhatIf } | Should -Throw
  }
}

Describe 'Test-PendingReboot' {
  It 'returns a structured result' {
    $result = Test-PendingReboot
    $result | Should -BeOfType [PSCustomObject]
    $result.PendingReboot | Should -BeOfType [bool]
    ($result.Indicators -is [string[]]) | Should -BeTrue
  }

  It 'does not throw when the SCCM client is absent' {
    { Test-PendingReboot } | Should -Not -Throw
  }

  It 'detects registry value presence without throwing under StrictMode' {
    Test-RegistryValuePresent -KeyPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ValueName 'CurrentBuild' | Should -BeTrue
    Test-RegistryValuePresent -KeyPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ValueName 'NoSuchValuePSFTest' | Should -BeFalse
    Test-RegistryValuePresent -KeyPath 'HKLM:\NoSuchKeyPSFTest' -ValueName 'x' | Should -BeFalse
  }
}

Describe 'Resolve-SystemFileIntegrityVerdict' {
  It 'maps the found-corrupt-files marker to Failed' {
    $verdict = Resolve-SystemFileIntegrityVerdict -SfcOutput @('Windows Resource Protection found corrupt files and was unable to fix some of them.') -ExitCode 1
    $verdict.Status | Should -Be 'Failed'
    $verdict.Detail | Should -Be 'integrity violations found'
  }

  It 'maps the clean marker to Completed even with a non-zero exit code' {
    $verdict = Resolve-SystemFileIntegrityVerdict -SfcOutput @('Windows Resource Protection did not find any integrity violations.') -ExitCode 1
    $verdict.Status | Should -Be 'Completed'
  }

  It 'maps the could-not-run marker to Failed' {
    $verdict = Resolve-SystemFileIntegrityVerdict -SfcOutput @('Windows Resource Protection could not perform the requested operation.') -ExitCode 0
    $verdict.Status | Should -Be 'Failed'
    $verdict.Detail | Should -Be 'could not run'
  }

  It 'falls back to the exit code when no marker matches' {
    Resolve-SystemFileIntegrityVerdict -ExitCode 0 | Select-Object -ExpandProperty Status | Should -Be 'Completed'
    Resolve-SystemFileIntegrityVerdict -ExitCode 1 | Select-Object -ExpandProperty Status | Should -Be 'Failed'
    Resolve-SystemFileIntegrityVerdict -ExitCode 2 | Select-Object -ExpandProperty Detail | Should -Be 'could not run'
    Resolve-SystemFileIntegrityVerdict -ExitCode 3 | Select-Object -ExpandProperty Detail | Should -Be 'sfc exited with code 3'
  }

  It 'matches markers from the CBS log excerpt too' {
    $verdict = Resolve-SystemFileIntegrityVerdict -LogExcerpt @('2016-01-01 blah', 'Windows Resource Protection found corrupt files') -ExitCode 0
    $verdict.Status | Should -Be 'Failed'
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

Describe 'Convert-RobocopyExitCode' {
  It 'decodes success variants' {
    Convert-RobocopyExitCode -ExitCode 0 | Should -Be 'No Change'
    Convert-RobocopyExitCode -ExitCode 1 | Should -Be 'OKCOPY'
    Convert-RobocopyExitCode -ExitCode 7 | Should -Be 'OKCOPY + MISMATCHES + XTRA'
  }

  It 'decodes failure variants' {
    Convert-RobocopyExitCode -ExitCode 8 | Should -Be 'FAIL'
    Convert-RobocopyExitCode -ExitCode 16 | Should -Be '***FATAL ERROR***'
  }

  It 'maps unknown codes to Unknown' {
    Convert-RobocopyExitCode -ExitCode 99 | Should -Be 'Unknown'
  }
}

Describe 'Find-ServiceAccountUsage' {
  It 'returns a structured result per searched name' {
    $result = Find-ServiceAccountUsage -Name 'PSFNoSuchAccount' -ComputerName 'localhost'
    $result | Should -BeOfType [PSCustomObject]
    $result.Name | Should -Be 'PSFNoSuchAccount'
    ($null -ne $result.Services) | Should -BeTrue
    ($null -ne $result.SchTasks) | Should -BeTrue
  }
}

Describe 'ConvertTo-FontFallbackName' {
  It 'derives a spaced name from a tokenized file name' {
    ConvertTo-FontFallbackName -FileName 'FiraCode-NF-Bold.ttf' | Should -Be 'FiraCode NF Bold'
  }

  It 'maps the wght token to Variable' {
    ConvertTo-FontFallbackName -FileName 'Roboto-VF.ttf' | Should -Be 'Roboto VF'
  }

  It 'collapses duplicate whitespace' {
    ConvertTo-FontFallbackName -FileName 'My  Font--Regular.ttf' | Should -Be 'My Font Regular'
  }

  It 'returns $null for an unusable name' {
    ConvertTo-FontFallbackName -FileName '.ttf' | Should -BeNullOrEmpty
  }
}

Describe 'Get-FontRegistryName' {
  It 'returns a non-empty display name for a plausible font file' {
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("winkit-font-test-$(New-Guid)")
    $null = New-Item -ItemType Directory -Path $tempDir
    try {
      $fakeFont = Join-Path $tempDir 'FiraCode-NF-Bold.ttf'
      $null = [System.IO.File]::WriteAllText($fakeFont, 'not a real font')
      $fontFile = [System.IO.FileInfo]::new($fakeFont)
      $result = Get-FontRegistryName -FontFile $fontFile
      $result | Should -Not -BeNullOrEmpty
      $result | Should -Match 'Fira Code NF Bold|FiraCode'
    }
    finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

Describe 'Install-Font' {
  It 'skips when no font files are present' {
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("winkit-font-empty-$(New-Guid)")
    $null = New-Item -ItemType Directory -Path $tempDir
    try {
      $result = @(Install-Font -FontSourceFolder $tempDir -WhatIf)
      $result[0].Status | Should -Be 'Skipped'
      $result[0].Detail | Should -Match 'No font files'
    }
    finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  It 'skips per-font under -WhatIf when font files exist' {
    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("winkit-font-whatif-$(New-Guid)")
    $null = New-Item -ItemType Directory -Path $tempDir
    try {
      $fakeFont = Join-Path $tempDir 'TestFont-Regular.ttf'
      $null = [System.IO.File]::WriteAllText($fakeFont, 'not a real font')
      $result = @(Install-Font -FontSourceFolder $tempDir -WhatIf)
      $result[0].Status | Should -Be 'Skipped'
      $result[0].Detail | Should -Be 'WhatIf'
    }
    finally {
      Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

Describe 'Set-ServiceStartupState' {
  It 'skips services that do not exist' {
    $result = @(Set-ServiceStartupState -Name 'NoSuchServiceXYZ' -StartupType Automatic -WhatIf)
    $result[0].Status | Should -Be 'Skipped'
    $result[0].Detail | Should -Be 'Service not found.'
  }

  It 'excludes services listed in -Filter' {
    Mock Get-Service { [pscustomobject]@{ Name = 'Spooler' } }
    $result = @(Set-ServiceStartupState -Name 'Spooler' -StartupType Automatic -Filter 'Spooler' -WhatIf)
    $result[0].Status | Should -Be 'Skipped'
    $result[0].Detail | Should -Be 'Excluded via -Filter.'
  }

  It 'refuses to change a protected service to Automatic' {
    $result = @(Set-ServiceStartupState -Name 'RemoteRegistry' -StartupType Automatic -WhatIf)
    $result[0].Status | Should -Be 'Refused'
  }

  It 'allows Disabled for a protected service' {
    Mock Get-Service { [pscustomobject]@{ Name = 'RemoteRegistry' } }
    Mock Set-Service { }
    $result = @(Set-ServiceStartupState -Name 'RemoteRegistry' -StartupType Disabled -Confirm:$false)
    $result[0].Status | Should -Be 'Completed'
  }

  It 'skips under -WhatIf for a normal service' {
    Mock Get-Service { [pscustomobject]@{ Name = 'Spooler' } }
    $result = @(Set-ServiceStartupState -Name 'Spooler' -StartupType Automatic -WhatIf)
    $result[0].Status | Should -Be 'Skipped'
    $result[0].Detail | Should -Be 'WhatIf'
  }
}

Describe 'Set-ScheduledTaskState' {
  It 'skips tasks that do not exist' {
    $result = @(Set-ScheduledTaskState -TaskName 'NoSuchTaskXYZ' -State Disabled -WhatIf)
    $result[0].Status | Should -Be 'Skipped'
    $result[0].Detail | Should -Be 'Task not found.'
  }

  It 'excludes tasks listed in -Filter' {
    Mock Get-ScheduledTask { [pscustomobject]@{ TaskName = 'SampleTask' } }
    $result = @(Set-ScheduledTaskState -TaskName 'SampleTask' -State Disabled -Filter 'SampleTask' -WhatIf)
    $result[0].Status | Should -Be 'Skipped'
    $result[0].Detail | Should -Be 'Excluded via -Filter.'
  }

  It 'skips under -WhatIf for an existing task' {
    Mock Get-ScheduledTask { [pscustomobject]@{ TaskName = 'SampleTask' } }
    $result = @(Set-ScheduledTaskState -TaskName 'SampleTask' -State Disabled -WhatIf)
    $result[0].Status | Should -Be 'Skipped'
    $result[0].Detail | Should -Be 'WhatIf'
  }
}
