#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/security.ps1
  . $PSScriptRoot/../src/system.ps1
  . $PSScriptRoot/../src/office.ps1

  function New-TestOfficeInventory {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates synthetic in-memory test data.')]
    [CmdletBinding()]
    param ($Configuration)

    $products = @()
    if ($Configuration) {
      $products = @([PSCustomObject]@{
          ProductId       = $Configuration.TargetProductId
          Architecture    = $Configuration.Architecture
          Channel         = $Configuration.Channel
          Version         = $Configuration.Version
          Languages       = @($Configuration.Language)
          PrimaryLanguage = $Configuration.PrimaryLanguage
          ExcludeApp      = @($Configuration.ExcludeApp)
          Evidence        = @('Synthetic complete observation')
        })
    }
    [PSCustomObject][ordered]@{
      SchemaVersion = 1
      MachineId     = 'synthetic-machine'
      Products      = $products
      Msi           = @()
      Unknowns      = @()
    }
  }
}

Describe 'Office configuration and authority contracts' {
  BeforeEach {
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Version 16.0.17932.20162
    $script:inventory = New-TestOfficeInventory
    Mock Get-OfficeInventory { $script:inventory }
    Mock Get-PSFOfficeOsLocale {
      [PSCustomObject]@{ Languages = @('de-de'); PrimaryLanguage = 'de-de'; Evidence = 'Synthetic machine UI language' }
    }
  }

  It 'defaults to en-us without consulting machine or caller culture' {
    $target.Language | Should -Be @('en-us')
    $target.LocaleSource | Should -Be Default
    Should -Invoke Get-PSFOfficeOsLocale -Times 0
    Should -Invoke Get-OfficeInventory -Times 0
  }

  It 'preserves bilingual order and removes duplicates case-insensitively' -ForEach @(
    @{ Languages = @('EN-US', 'de-de', 'en-us'); Primary = 'en-us' }
    @{ Languages = @('de-de', 'en-us', 'DE-DE'); Primary = 'de-de' }
  ) {
    $result = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language $Languages
    $result.Language.Count | Should -Be 2
    $result.PrimaryLanguage | Should -Be $Primary
    $result.Language[0] | Should -Be $Primary
  }

  It 'requires explicit locale discovery and records its evidence' {
    $result = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -AutoSourceLocales -LocaleSource OperatingSystem
    $result.Language | Should -Be @('de-de')
    $result.LocaleEvidence | Should -Contain 'Synthetic machine UI language'
  }

  It 'rejects conflicting or unsupported locale inputs' {
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language de-de -AutoSourceLocales } | Should -Throw
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -LocaleSource OperatingSystem } | Should -Throw
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language @() } | Should -Throw
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language xx-xx } | Should -Throw
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -AutoSourceLocales } | Should -Throw
  }

  It 'discovers installed languages only when their primary language is known' {
    $script:inventory = New-TestOfficeInventory (New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language de-de, en-us)
    $result = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -AutoSourceLocales
    $result.Language | Should -Be @('de-de', 'en-us')
    $script:inventory.Products[0].PrimaryLanguage = $null
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -AutoSourceLocales } | Should -Throw
  }

  It 'rejects a channel that belongs to another product family' {
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Channel Current } | Should -Throw
    { New-OfficeDeploymentConfiguration -TargetProductId O365ProPlusRetail -Channel PerpetualVL2024 } | Should -Throw
  }

  It 'rejects unknown serialized configuration fields' {
    $target | Add-Member -NotePropertyName RemoveAll -NotePropertyValue $true
    { Test-OfficeDeployment -Configuration $target } | Should -Throw '*Unexpected*'
  }

  It 'rejects removal authority on <Action>' -ForEach @(
    @{ Action = 'Install' }, @{ Action = 'Update' }, @{ Action = 'AddLanguage' },
    @{ Action = 'SetApplicationSelection' }, @{ Action = 'SetApplicationPreference' }
  ) {
    { Get-OfficeDeploymentPlan -Action $Action -Configuration $target -RemoveProductId O365ProPlusRetail } | Should -Throw '*cannot remove*'
    { Get-OfficeDeploymentPlan -Action $Action -Configuration $target -RemoveMsi } | Should -Throw '*cannot remove*'
  }

  It 'rejects arbitrary settings or preferences outside their schema' {
    { Get-OfficeDeploymentPlan -Action Install -Configuration $target -Settings @{ Xml = '<Remove All="TRUE" />' } } | Should -Throw
    { Get-OfficeDeploymentPlan -Action SetUpdateConfiguration -Configuration $target -Settings @{ Enabled = 'false' } } | Should -Throw
    { Get-OfficeDeploymentPlan -Action SetUpdateConfiguration -Configuration $target -Settings @{ UpdatePath = 'http://example.invalid' } } | Should -Throw
    { Get-OfficeDeploymentPlan -Action SetApplicationPreference -Configuration $target -Settings @{ Preferences = @(@{ Key = 'software\other'; Name = 'Test'; Value = 1; Type = 'REG_DWORD'; App = 'excel16'; Id = 'Test' }) } } | Should -Throw
  }
}

Describe 'Office inventory and compliance evidence' {
  BeforeEach {
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Version 16.0.17932.20162
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' }
  }

  It 'does not treat an empty C2R registration as a clean machine' {
    Mock Get-PSFOfficeRegistrySnapshot {
      [PSCustomObject]@{ View = 'Registry64'; Path = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; Values = [PSCustomObject]@{} }
    }
    $observed = Get-OfficeInventory
    $observed.Unknowns | Should -Contain IncompleteClickToRunRegistration
    (Get-OfficeDeploymentPlan -Action Install -Configuration $target -Inventory $observed).Eligible | Should -BeFalse
  }

  It 'recognizes Office 2007 setup wrappers but does not execute uninstall strings' {
    Mock Get-PSFOfficeRegistrySnapshot {
      [PSCustomObject]@{
        View   = 'Registry32'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ENTERPRISE'
        Values = [PSCustomObject]@{
          Publisher        = 'Microsoft Corporation'
          DisplayName      = 'Microsoft Office Enterprise 2007'
          DisplayVersion   = '12.0.0.0'
          WindowsInstaller = 0
          UninstallString  = '"C:\Program Files\Common Files\Microsoft Shared\OFFICE12\Office Setup Controller\setup.exe" /uninstall ENTERPRISE'
        }
      }
    }
    $observed = Get-OfficeInventory
    $observed.Msi[0].ProductCode | Should -Be ENTERPRISE
    ($observed | ConvertTo-Json -Depth 8) | Should -Not -Match UninstallString
  }

  It 'does not promote discovery errors to clean inventory' {
    Mock Get-PSFOfficeRegistrySnapshot { throw 'Synthetic inaccessible registry' }
    (Get-OfficeInventory).Unknowns | Should -Contain RegistryDiscoveryFailed
  }

  It 'requires every compliance dimension, independent of activation' {
    $observed = New-TestOfficeInventory $target
    (Test-OfficeDeployment -Configuration $target -Inventory $observed).Compliant | Should -BeTrue
    $observed.Products[0].PrimaryLanguage = $null
    $result = Test-OfficeDeployment -Configuration $target -Inventory $observed
    $result.Compliant | Should -BeFalse
    $result.Unknowns | Should -Contain PrimaryLanguage
    $observed.Products[0].Channel = 'Current'
    (Test-OfficeDeployment -Configuration $target -Inventory $observed).Discrepancies | Should -Contain Channel
  }

  It 'does not infer full language or primary-language evidence from ClientCulture' {
    Mock Get-PSFOfficeRegistrySnapshot {
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        Values = [PSCustomObject]@{ ProductReleaseIds = 'Standard2024Volume'; Platform = 'x64'; ClientCulture = 'de-de'; VersionToReport = '16.0.17932.20162' }
      }
    }
    $observed = Get-OfficeInventory
    $observed.Products[0].Languages | Should -BeNullOrEmpty
    $observed.Products[0].PrimaryLanguage | Should -BeNullOrEmpty
  }

  It 'keeps Microsoft 365 activation separate from volume licensing' {
    Mock Get-CimInstance { throw 'Must not query licenses' }
    (Get-OfficeActivationStatus O365ProPlusRetail).Status | Should -Be UserActivationRequired
    Should -Invoke Get-CimInstance -Times 0
  }
}

Describe 'Office plans and narrow XML generation' {
  BeforeEach {
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language en-us, de-de -Version 16.0.17932.20162
    $script:inventory = New-TestOfficeInventory
    Mock Get-OfficeInventory { $script:inventory }
    Mock Test-OfficeDeploymentMedia {
      [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.17932.20162'; Files = @() } }
    }
  }

  It 'pins prepared media and plans a clean installation' {
    $target.Version = $null
    $plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath C:\Media
    $plan.Eligible | Should -BeTrue
    $plan.Configuration.Version | Should -Be '16.0.17932.20162'
    $plan.State | Should -Be Clean
  }

  It 'reports a genuinely compliant installation without requiring media' {
    $script:inventory = New-TestOfficeInventory $target
    $plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target
    $plan.State | Should -Be Compliant
    $plan.Eligible | Should -BeTrue
  }

  It 'blocks ordinary install when configuration evidence is incomplete' {
    $script:inventory = New-TestOfficeInventory $target
    $script:inventory.Products[0].Languages = $null
    $plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath C:\Media
    $plan.Blockers | Should -Contain RecoveryRequired
  }

  It 'requires explicit selection even when replacing the same product ID' {
    $script:inventory = New-TestOfficeInventory $target
    $script:inventory.Products[0].PrimaryLanguage = 'de-de'
    (Get-OfficeDeploymentPlan -Action Migrate -Configuration $target -SourcePath C:\Media).Blockers | Should -Contain UnapprovedProducts
    $plan = Get-OfficeDeploymentPlan -Action Migrate -Configuration $target -SourcePath C:\Media -RemoveProductId Standard2024Volume
    $plan.Eligible | Should -BeTrue
    $plan.Warnings.Count | Should -BeGreaterThan 0
  }

  It 'blocks standalone MSI removal and rejects empty removal selections' {
    (Get-OfficeDeploymentPlan -Action Remove -RemoveMsi).Blockers | Should -Contain UnsupportedStandaloneMsi
    { Get-OfficeDeploymentPlan -Action Remove } | Should -Throw
    { Get-OfficeDeploymentPlan -Action Remove -RemoveProductId 'bad*id' } | Should -Throw
  }

  It 'does not let updates alter other configuration dimensions' {
    $script:inventory = New-TestOfficeInventory $target
    $script:inventory.Products[0].Architecture = '32'
    (Get-OfficeDeploymentPlan -Action Update -Configuration $target -SourcePath C:\Media).Blockers | Should -Contain UnverifiedPreservedConfiguration
  }

  It 'requires migration for removal of the primary language' {
    $script:inventory = New-TestOfficeInventory $target
    $plan = Get-OfficeDeploymentPlan -Action RemoveLanguage -Configuration $target -Language en-us
    $plan.Blockers | Should -Contain PrimaryLanguageRequiresMigration
  }

  It 'preserves XML language order and escapes paths without permitting removal' {
    $document = New-PSFOfficeXml -Action Install -Configuration $target -MediaPath 'C:\Media & Tools'
    @($document.Configuration.Add.Product.Language.ID) | Should -Be @('en-us', 'de-de')
    $document.OuterXml | Should -Match '&amp;'
    $document.SelectNodes('//Remove|//RemoveMSI|//*[@MigrateArch]').Count | Should -Be 0
    $document.Configuration.Add.AllowCdnFallback | Should -Be FALSE
    { New-PSFOfficeXml -Action Install -Configuration $target -MediaPath C:\Media -RemoveMsi $true } | Should -Throw
  }

  It 'limits removal XML to selected products and languages' {
    $document = New-PSFOfficeXml -Action Remove -RemoveProductId O365ProPlusRetail
    $document.Configuration.Remove.All | Should -Be FALSE
    $document.Configuration.Remove.Product.ID | Should -Be O365ProPlusRetail
    $document.SelectNodes('//Add').Count | Should -Be 0
    $document = New-PSFOfficeXml -Action RemoveLanguage -Configuration $target -Language de-de
    $document.Configuration.Remove.Product.Language.ID | Should -Be de-de
  }

  It 'emits update-only and preference-only configurations' {
    $document = New-PSFOfficeXml -Action SetUpdateConfiguration -Settings @{ Enabled = $false }
    $document.SelectNodes('//Add|//Remove|//AppSettings').Count | Should -Be 0
    $document.Configuration.Updates.Enabled | Should -Be False
    $settings = @{ Preferences = @(@{ Key = 'software\microsoft\office\16.0\excel\options'; Name = 'test'; Value = 'a & b'; Type = 'REG_SZ'; App = 'excel16'; Id = 'Test' }) }
    $document = New-PSFOfficeXml -Action SetApplicationPreference -Settings $settings
    $document.SelectNodes('//Add|//Remove|//Updates').Count | Should -Be 0
    $document.Configuration.AppSettings.User.Value | Should -Be 'a & b'
  }

  It 'rechecks inventory and rejects foreign actions and extra plan fields' {
    $plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath C:\Media
    { Confirm-PSFOfficePlan -Plan $plan -Action Remove } | Should -Throw
    $script:inventory = New-TestOfficeInventory $target
    { Confirm-PSFOfficePlan -Plan $plan -Action Install } | Should -Throw '*changed*'
    $plan | Add-Member -NotePropertyName SkipValidation -NotePropertyValue $true
    { Confirm-PSFOfficePlan -Plan $plan -Action Install } | Should -Throw '*Unexpected*'
  }
}

Describe 'Office media package integrity' {
  BeforeEach {
    Mock Assert-PSFOfficeProtectedPath { }
    $script:mediaRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    $null = New-Item -Path (Join-Path $mediaRoot 'Office\Data\16.0.17932.20162') -ItemType Directory -Force
    foreach ($relative in @('Office\Data\v64.cab', 'Office\Data\16.0.17932.20162\stream.x64.x-none.dat', 'Office\Data\16.0.17932.20162\stream.x64.en-us.dat', 'Office\Data\16.0.17932.20162\stream.x64.de-de.dat')) {
      [IO.File]::WriteAllText((Join-Path $mediaRoot $relative), 'synthetic payload')
    }
    $script:manifest = [ordered]@{
      SchemaVersion      = 2
      Product            = 'Standard2024Volume'
      Architecture       = '64'
      Channel            = 'PerpetualVL2024'
      AvailableLanguages = @('en-us', 'de-de')
      Version            = '16.0.17932.20162'
      ToolVersion        = '16.0.20326.20112'
      Files              = @(Get-PSFOfficeMediaFile $mediaRoot | Sort-Object Path)
    }
    $script:manifestPath = Join-Path $mediaRoot 'psfoundation-office-media.json'
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume
  }

  It 'allows a bilingual package to serve a single-language request' {
    $result = Test-OfficeDeploymentMedia -SourcePath $mediaRoot -Configuration $target
    $result.Valid | Should -BeTrue
    $target.Language | Should -Be @('en-us')
  }

  It 'detects altered or extra source payloads' {
    [IO.File]::WriteAllText((Join-Path $mediaRoot 'Office\Data\v64.cab'), 'tampered')
    (Test-OfficeDeploymentMedia $mediaRoot).ReasonCode | Should -Be MediaIntegrityFailed
  }

  It 'rejects traversal and case-colliding paths before reading files' {
    $manifest.Files[0].Path = 'Office/Data/../../outside.exe'
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    (Test-OfficeDeploymentMedia $mediaRoot).ReasonCode | Should -Be UnsafeManifest
  }

  It 'rejects manifests that bless missing language payloads' {
    $missing = 'Office/Data/16.0.17932.20162/stream.x64.de-de.dat'
    Remove-Item -LiteralPath (Join-Path $mediaRoot $missing)
    $manifest.Files = @($manifest.Files | Where-Object { $_.Path -ne $missing })
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    (Test-OfficeDeploymentMedia $mediaRoot).ReasonCode | Should -Be MissingLanguageMedia
  }

  It 'rejects interrupted preparation and schema-1 packages' {
    Remove-Item -LiteralPath $manifestPath
    (Test-OfficeDeploymentMedia $mediaRoot).ReasonCode | Should -Be ReprepareMedia
    @{ Schema = 1 } | ConvertTo-Json | Set-Content -LiteralPath $manifestPath
    (Test-OfficeDeploymentMedia $mediaRoot).Valid | Should -BeFalse
  }
}

Describe 'Office executor lifecycle' {
  BeforeEach {
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Version 16.0.17932.20162
    $script:inventory = New-TestOfficeInventory
    Mock Get-OfficeInventory { $script:inventory }
    Mock Test-OfficeDeploymentMedia {
      [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.17932.20162'; Files = @([PSCustomObject]@{ Length = 1 }) } }
    }
    Mock Test-OfficeDeploymentTool { [PSCustomObject]@{ Valid = $true; Version = '16.0.20326.20112' } }
    Mock Assert-PSFOfficeHost { }
    Mock Get-PSFOfficeActivity { [PSCustomObject]@{ Busy = $false; Apps = @() } }
    Mock New-PSFOfficeProtectedDirectory { }
    Mock Copy-Item { }
    $script:checkpoints = @()
    Mock Write-PSFOfficeJson { $script:checkpoints += $Value.Phase }
    Mock Write-OperationResultLog { 'synthetic-log' }
    Mock Remove-PSFOfficeWorkDirectory { }
    Mock Enter-PSFOfficeLock {
      $lock = [PSCustomObject]@{}
      $lock | Add-Member ScriptMethod ReleaseMutex { }
      $lock | Add-Member ScriptMethod Dispose { }
      $lock
    }
    Mock Invoke-PSFOfficeConfiguration {
      $script:inventory = New-TestOfficeInventory $script:target
      [PSCustomObject]@{ ExitCode = 0 }
    }
    Mock Get-OfficeActivationStatus { [PSCustomObject]@{ Status = 'Licensed' } }
    $script:plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath (Join-Path $TestDrive 'Media')
    $script:logRoot = Join-Path $TestDrive 'Journal'
  }

  It 'keeps WhatIf and DryRun completely free of writes and process actions' {
    $preview = Install-Office -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -WhatIf
    $preview.Status | Should -Be Preview
    $preview = Install-Office -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -DryRun
    $preview.Status | Should -Be Preview
    Should -Invoke New-PSFOfficeProtectedDirectory -Times 0
    Should -Invoke Write-PSFOfficeJson -Times 0
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
    Should -Invoke Enter-PSFOfficeLock -Times 0
  }

  It 'returns one final verified result and a recovery record' {
    $result = @(Install-Office -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -Confirm:$false)
    $result.Count | Should -Be 1
    $result[0].Status | Should -Be Completed
    $result[0].Changed | Should -BeTrue
    $result[0].RecoveryRequired | Should -BeFalse
    @($script:checkpoints | Where-Object { $_ -eq 'StageMedia' }).Count | Should -Be 2
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1
  }

  It 'keeps unknown change state after a native failure' {
    Mock Invoke-PSFOfficeConfiguration { [PSCustomObject]@{ ExitCode = 1603 } }
    $result = Install-Office -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -Confirm:$false
    $result.Status | Should -Be Failed
    $result.Changed | Should -BeNullOrEmpty
    $result.ChangeKnown | Should -BeFalse
    $result.RecoveryRequired | Should -BeTrue
    $result.ExitCode | Should -Be 1603
  }

  It 'preserves native reboot success without restarting the host' {
    Mock Invoke-PSFOfficeConfiguration {
      $script:inventory = New-TestOfficeInventory $script:target
      [PSCustomObject]@{ ExitCode = 3010 }
    }
    $result = Install-Office -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -Confirm:$false
    $result.RebootRequired | Should -BeTrue
    $result.WrapperExitCode | Should -Be 3010
  }

  It 'preserves the original error when staging cleanup also fails' {
    Mock Invoke-PSFOfficeConfiguration { [PSCustomObject]@{ ExitCode = 1603 } }
    Mock Remove-PSFOfficeWorkDirectory { throw 'Synthetic cleanup failure' }
    $result = Install-Office -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -Confirm:$false
    $result.Error | Should -Match 1603
    $result.CleanupErrors | Should -Contain 'Synthetic cleanup failure'
    $result.Residue.Count | Should -Be 1
    $result.WrapperExitCode | Should -Be 1
  }

  It 'blocks bad signatures before staging or application closure' {
    Mock Test-OfficeDeploymentTool { [PSCustomObject]@{ Valid = $false } }
    $result = Install-Office -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -Confirm:$false
    $result.ReasonCode | Should -Be UntrustedTool
    Should -Invoke New-PSFOfficeProtectedDirectory -Times 0
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
  }

  It 'prevents one public command from executing another action' {
    { Uninstall-Office -Plan $plan -OdtPath C:\ODT\setup.exe -WhatIf } | Should -Throw '*does not match*'
    { Update-Office -Plan $plan -OdtPath C:\ODT\setup.exe -WhatIf } | Should -Throw '*does not match*'
  }

  It 'returns AlreadyAbsent and preserves unrelated inventory' {
    $script:inventory = New-TestOfficeInventory $target
    $remove = Get-OfficeDeploymentPlan -Action Remove -RemoveProductId O365ProPlusRetail
    $result = Uninstall-Office -Plan $remove -OdtPath C:\ODT\setup.exe -Confirm:$false
    $result.ReasonCode | Should -Be AlreadyAbsent
    $result.Changed | Should -BeFalse
    $result.After.Products[0].ProductId | Should -Be Standard2024Volume
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
  }

  It 'halts migration after removal failure or a reboot request' -ForEach @(
    @{ ExitCode = 1603 }, @{ ExitCode = 3010 }
  ) {
    $script:inventory = New-TestOfficeInventory $target
    $script:inventory.Products[0].ProductId = 'O365ProPlusRetail'
    $migration = Get-OfficeDeploymentPlan -Action Migrate -Configuration $target -SourcePath (Join-Path $TestDrive 'Media') -RemoveProductId O365ProPlusRetail
    $script:nativeExit = $ExitCode
    Mock Invoke-PSFOfficeConfiguration { [PSCustomObject]@{ ExitCode = $script:nativeExit } }
    $result = Switch-OfficeDeployment -Plan $migration -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -Confirm:$false
    $result.Status | Should -Be Failed
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0 -ParameterFilter { $Document.Configuration.Add }
  }
}

Describe 'Office recovery boundaries' {
  BeforeEach {
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Version 16.0.17932.20162
    $script:inventory = New-TestOfficeInventory
    Mock Get-OfficeInventory { $script:inventory }
    Mock Get-PSFOfficeActivity { [PSCustomObject]@{ Busy = $false; Apps = @() } }
    Mock Test-PendingReboot { [PSCustomObject]@{ PendingReboot = $false } }
    Mock Get-OfficeActivationStatus { [PSCustomObject]@{ Status = 'Licensed' } }
    Mock Test-OfficeDeploymentMedia {
      [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.17932.20162' } }
    }
    $script:plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath C:\Media
    $script:recovery = [PSCustomObject]@{
      RunId   = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
      LogRoot = (Join-Path $TestDrive 'Recovery')
      Path    = 'synthetic-journal'
      Record  = [PSCustomObject]@{
        Action           = 'Install'
        Plan             = $plan
        Phase            = 'Install'
        PhaseCompleted   = $false
        MediaFingerprint = 'synthetic-media'
      }
    }
    Mock Get-OfficeDeploymentRecovery { $script:recovery }
    Mock Invoke-PSFOfficeWorkflow { [PSCustomObject]@{ Continued = $true; Plan = $Plan; ExpectedAction = $ExpectedAction } }
    Mock Invoke-PSFOfficeConfiguration { throw 'Recovery must not invoke ODT directly' }
  }

  It 'verifies completion after native success before the final checkpoint' {
    $script:inventory = New-TestOfficeInventory $target
    $result = Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe
    $result.Status | Should -Be Completed
    $result.Changed | Should -BeFalse
    $result.AlreadyCompliant | Should -BeTrue
    Should -Invoke Invoke-PSFOfficeWorkflow -Times 0
  }

  It 'does not reinstall when only activation remains unverified' {
    $script:inventory = New-TestOfficeInventory $target
    Mock Get-OfficeActivationStatus { [PSCustomObject]@{ Status = 'NotVerified' } }
    $result = Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe
    $result.ReasonCode | Should -Be ActivationNotVerified
    $result.WrapperExitCode | Should -Be 1
    Should -Invoke Invoke-PSFOfficeWorkflow -Times 0
  }

  It 'continues pre-launch installation with the recorded target and ordered languages' {
    $script:recovery.Record.Phase = 'StageMedia'
    $result = Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe -WhatIf
    $result.Continued | Should -BeTrue
    $result.ExpectedAction | Should -Be Install
    $result.Plan.Configuration.Language | Should -Be @('en-us')
    Should -Invoke Invoke-PSFOfficeWorkflow -Times 1 -ParameterFilter { $Plan.RemoveProductId.Count -eq 0 -and -not $Plan.RemoveMsi }
  }

  It 'does not replay an uncertain interrupted installer' {
    $script:inventory = New-TestOfficeInventory $target
    $script:inventory.Products[0].Languages = $null
    $result = Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe
    $result.ReasonCode | Should -Be UnsupportedRecoveryState
    $result.RecoveryRequired | Should -BeTrue
    Should -Invoke Invoke-PSFOfficeWorkflow -Times 0
  }

  It 'never imports migration removal authority into installation recovery' {
    $script:recovery.Record.Action = 'Migrate'
    { Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe } | Should -Throw '*cannot resume*'
  }

  It 'blocks recovery during external deployment or pending reboot' {
    Mock Get-PSFOfficeActivity { [PSCustomObject]@{ Busy = $true; Apps = @() } }
    (Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe).ReasonCode | Should -Be DeploymentBusy
    Mock Get-PSFOfficeActivity { [PSCustomObject]@{ Busy = $false; Apps = @() } }
    Mock Test-PendingReboot { [PSCustomObject]@{ PendingReboot = $true } }
    (Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe).WrapperExitCode | Should -Be 3010
    Should -Invoke Invoke-PSFOfficeWorkflow -Times 0
  }

  It 'rejects changed media and newly discovered products' {
    Mock Test-OfficeDeploymentMedia { [PSCustomObject]@{ Valid = $true; Fingerprint = 'different' } }
    { Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe } | Should -Throw '*media*'
    Mock Test-OfficeDeploymentMedia { [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media' } }
    $script:inventory = New-TestOfficeInventory $target
    $script:inventory.Products[0].ProductId = 'VisioProRetail'
    (Resume-OfficeInstallation -Recovery $recovery -OdtPath C:\ODT\setup.exe).ReasonCode | Should -Be Conflict
  }
}

Describe 'Office protected files and secret handling' {
  BeforeEach {
    Mock Assert-PSFOfficeProtectedPath { }
  }

  It 'round-trips an atomic journal and rejects foreign machine identity' {
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' }
    $root = Join-Path $TestDrive 'journal'
    $null = New-Item -Path $root -ItemType Directory -Force
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Version 16.0.17932.20162
    $plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -Inventory (New-TestOfficeInventory)
    $record = [ordered]@{
      SchemaVersion            = 1
      RunId                    = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
      MachineId                = 'synthetic-machine'
      Action                   = 'Install'
      Plan                     = $plan
      ConfigurationFingerprint = Get-PSFOfficeFingerprint $target
      MediaFingerprint         = $null
      Phase                    = 'StageMedia'
      PhaseCompleted           = $false
      NativeResults            = @()
      RebootRequired           = $false
      CreatedAt                = '2026-01-01T00:00:00Z'
      UpdatedAt                = '2026-01-01T00:00:00Z'
      Result                   = $null
    }
    $path = Join-Path $root ($record.RunId + '.json')
    Write-PSFOfficeJson -Path $path -Value $record
    (Get-OfficeDeploymentRecovery -RunId $record.RunId -LogRoot $root).Record.Action | Should -Be Install
    $record.PhaseCompleted = $true
    Write-PSFOfficeJson -Path $path -Value $record
    (Get-OfficeDeploymentRecovery -RunId $record.RunId -LogRoot $root).Record.PhaseCompleted | Should -BeTrue
    @(Get-ChildItem -LiteralPath $root -Filter '*.tmp').Count | Should -Be 0
    Mock Get-PSFOfficeMachineId { 'another-machine' }
    { Get-OfficeDeploymentRecovery -RunId $record.RunId -LogRoot $root } | Should -Throw
  }

  It 'cleans temporary XML after native failure without passing a key in arguments' {
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Version 16.0.17932.20162
    $document = New-PSFOfficeXml -Action Install -Configuration $target -MediaPath C:\Media
    $key = New-Object Security.SecureString
    foreach ($character in 'AAAAA-BBBBB-CCCCC-DDDDD-EEEEE'.ToCharArray()) {
      $key.AppendChar($character)
    }
    Mock Invoke-PSFOfficeTool {
      $ConfigurationPath | Should -Not -Match AAAAA
      $xml = [xml](Get-Content -LiteralPath $ConfigurationPath -Raw)
      $xml.Configuration.Add.Product.PIDKEY | Should -Be 'AAAAA-BBBBB-CCCCC-DDDDD-EEEEE'
      throw 'Synthetic installer failure'
    }
    { Invoke-PSFOfficeConfiguration -OdtPath C:\ODT\setup.exe -Document $document -Directory $TestDrive -ProductKey $key } | Should -Throw '*Synthetic installer failure*'
    $document.SelectNodes('//*[@PIDKEY]').Count | Should -Be 0
    @(Get-ChildItem -LiteralPath $TestDrive -Filter 'configuration-*.xml').Count | Should -Be 0
    $key.Dispose()
  }

  It 'rejects traversal, alternate streams, and cleanup outside the operation root' {
    { Assert-PSFOfficePath 'C:\Media\..\outside' } | Should -Throw
    { Assert-PSFOfficePath 'C:\Media\setup.exe:stream' } | Should -Throw
    { Remove-PSFOfficeWorkDirectory -Path $TestDrive -Parent (Join-Path $TestDrive 'child') } | Should -Throw '*escaped*'
  }
}

Describe 'Office preparation and acquisition boundaries' {
  BeforeEach {
    Mock Assert-PSFOfficeProtectedPath { }
    Mock Test-OfficeDeploymentTool { [PSCustomObject]@{ Valid = $true; Version = '16.0.20326.20112' } }
    Mock New-PSFOfficeProtectedDirectory {
      $null = [IO.Directory]::CreateDirectory($Path)
    }
    $script:setupPath = Join-Path $TestDrive 'synthetic-setup.exe'
    [IO.File]::WriteAllText($setupPath, 'Synthetic executable fixture; never executed.')
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language en-us, de-de
    $script:destination = Join-Path $TestDrive ('Prepared media ' + [guid]::NewGuid().ToString('N'))
    Mock Invoke-PSFOfficeConfiguration {
      $data = Join-Path $Directory 'Office\Data\16.0.17932.20162'
      $null = [IO.Directory]::CreateDirectory($data)
      [IO.File]::WriteAllText((Join-Path (Split-Path $data -Parent) 'v64.cab'), 'Synthetic base catalog')
      foreach ($language in @('x-none', 'en-us', 'de-de')) {
        [IO.File]::WriteAllText((Join-Path $data "stream.x64.$language.dat"), "Synthetic $language payload")
      }
      [PSCustomObject]@{ ExitCode = 0 }
    }
  }

  It 'publishes a pinned package only after successful preparation and reuses it' {
    $media = Save-OfficeDeploymentMedia -Configuration $target -SourcePath $destination -OdtPath $setupPath -Confirm:$false
    $media.Valid | Should -BeTrue
    $media.Path | Should -Be $destination
    $media.Manifest.Version | Should -Be '16.0.17932.20162'
    $media.Manifest.AvailableLanguages | Should -Be @('en-us', 'de-de')
    Test-Path -LiteralPath (Join-Path $destination 'setup.exe') | Should -BeFalse
    $reuse = Save-OfficeDeploymentMedia -Configuration $target -SourcePath $destination -OdtPath $setupPath -Confirm:$false
    $reuse.Valid | Should -BeTrue
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1
  }

  It 'does not publish an interrupted download as ready' {
    Mock Invoke-PSFOfficeConfiguration { [PSCustomObject]@{ ExitCode = 1603 } }
    { Save-OfficeDeploymentMedia -Configuration $target -SourcePath $destination -OdtPath $setupPath -Confirm:$false } | Should -Throw '*1603*'
    Test-Path -LiteralPath $destination | Should -BeFalse
    @(Get-ChildItem -LiteralPath $TestDrive -Directory -Filter 'PSFOfficePrepare-*').Count | Should -Be 0
  }

  It 'never downloads or creates staging during preparation preview' {
    $preview = Save-OfficeDeploymentMedia -Configuration $target -SourcePath $destination -OdtPath $setupPath -WhatIf
    $preview.Status | Should -Be Preview
    Should -Invoke New-PSFOfficeProtectedDirectory -Times 0
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
  }

  It 'does not execute an untrusted downloaded extractor' {
    Mock Invoke-WebRequest { [IO.File]::WriteAllText($OutFile, 'Synthetic untrusted extractor') }
    Mock Get-AuthenticodeSignature { [PSCustomObject]@{ Status = 'NotSigned'; SignerCertificate = $null } }
    Mock Invoke-SafeProcess { throw 'Must not execute untrusted download' }
    { Install-OfficeDeploymentTool -Destination $destination -Confirm:$false } | Should -Throw '*signature*'
    Should -Invoke Invoke-SafeProcess -Times 0
  }

  It 'does not acquire or execute the tool during WhatIf' {
    Mock Invoke-WebRequest { throw 'Must not download' }
    Mock Invoke-SafeProcess { throw 'Must not execute' }
    (Install-OfficeDeploymentTool -Destination $destination -WhatIf).Status | Should -Be Preview
    Should -Invoke Invoke-WebRequest -Times 0
    Should -Invoke Invoke-SafeProcess -Times 0
    Should -Invoke New-PSFOfficeProtectedDirectory -Times 0
  }
}
