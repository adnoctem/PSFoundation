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

Describe 'Imported Office nullable collections' {
  BeforeAll {
    Remove-Module -Name PSFoundation -Force -ErrorAction SilentlyContinue
    Import-Module "$PSScriptRoot/../src/PSFoundation.psd1" -Force
  }

  BeforeEach {
    $script:countTarget = PSFoundation\New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Language de-de -Version 16.0.10417.20208
    $script:countInventory = New-TestOfficeInventory $script:countTarget
    $script:countInventory | Add-Member -NotePropertyName VerificationLimitations -NotePropertyValue @()
    Mock Get-OfficeInventory { $script:countInventory } -ModuleName PSFoundation
  }

  It 'reports unknown installed languages through the intended locale error' {
    $script:countInventory.Products[0].Languages = $null
    { PSFoundation\New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -AutoSourceLocales } |
      Should -Throw '*Supply Language explicitly*'
  }

  It 'treats null optional plan selections as empty rather than granting removal authority' {
    $plan = PSFoundation\Get-OfficeDeploymentPlan -Action Install -Configuration $script:countTarget -RemoveProductId $null -Language $null
    $plan.RemoveProductId.Count | Should -Be 0
    $plan.Language.Count | Should -Be 0
    $plan.RemoveMsi | Should -BeFalse
    $plan.Eligible | Should -BeTrue
    $plan.State | Should -Be Compliant
  }
}

Describe 'Imported Office sparse registry values' {
  BeforeAll {
    Remove-Module -Name PSFoundation -Force -ErrorAction SilentlyContinue
    Import-Module "$PSScriptRoot/../src/PSFoundation.psd1" -Force
  }

  BeforeEach {
    $script:sparseRegistry = @()
    Mock Get-PSFOfficeRegistrySnapshot { $script:sparseRegistry } -ModuleName PSFoundation
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' } -ModuleName PSFoundation
  }

  It 'recognizes the Office 2007 controller without a WindowsInstaller value' {
    $script:sparseRegistry = @([PSCustomObject]@{
        View   = 'Registry32'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ENTERPRISE'
        Values = [PSCustomObject]@{
          Publisher       = 'Microsoft Corporation'
          DisplayName     = 'Microsoft Office Enterprise 2007'
          UninstallString = '"C:\Program Files\Common Files\Microsoft Shared\OFFICE12\Office Setup Controller\setup.exe" /uninstall ENTERPRISE'
        }
      })

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Msi.Count | Should -Be 1
    $observed.Msi[0].ProductCode | Should -Be ENTERPRISE
    $observed.Msi[0].Version | Should -BeNullOrEmpty
    $observed.Unknowns.Count | Should -Be 0
    ($observed | ConvertTo-Json -Depth 8) | Should -Not -Match UninstallString
  }

  It 'ignores empty and unrelated uninstall keys without assuming MSI evidence' {
    $script:sparseRegistry = @(
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\empty'
        Values = [PSCustomObject]@{}
      }
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\unrelated'
        Values = [PSCustomObject]@{ DisplayName = 'Unrelated application' }
      }
    )

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Msi.Count | Should -Be 0
    $observed.Products.Count | Should -Be 0
    $observed.Unknowns.Count | Should -Be 0
  }

  It 'keeps unclassified Office registrations blocked when installer evidence is absent' {
    $script:sparseRegistry = @([PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\unknown-office'
        Values = [PSCustomObject]@{
          Publisher   = 'Microsoft Corporation'
          DisplayName = 'Microsoft Office Unknown Component'
        }
      })

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Msi.Count | Should -Be 0
    $observed.Unknowns | Should -Contain 'UnclassifiedOfficeRegistration:unknown-office'
    $target = PSFoundation\New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume
    (PSFoundation\Get-OfficeDeploymentPlan -Action Install -Configuration $target -Inventory $observed).Eligible | Should -BeFalse
  }

  It 'classifies Office 2007 patch registrations without a WindowsInstaller value' {
    # Sanitized from an Office Enterprise 2007 SP3 workstation: the patch keys
    # carry no WindowsInstaller value and a localizable service-pack name.
    $script:sparseRegistry = @(
      [PSCustomObject]@{
        View   = 'Registry32'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-001B-0407-0000-0000000FF1CE}'
        Values = [PSCustomObject]@{
          Publisher        = 'Microsoft'
          DisplayName      = 'Microsoft Office Word MUI (German) 2007'
          WindowsInstaller = 1
        }
      }
      [PSCustomObject]@{
        View   = 'Registry32'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-001B-0407-0000-0000000FF1CE}_ENTERPRISER_{DB2ACBD1-65B1-4FC5-881E-4E75C668E7E2}'
        Values = [PSCustomObject]@{
          Publisher         = 'Microsoft'
          DisplayName       = 'Microsoft Office 2007 Service Pack 3 (SP3)'
          ParentKeyName     = '{90120000-001B-0407-0000-0000000FF1CE}'
          ParentDisplayName = 'Microsoft Office Word MUI (German) 2007'
          SystemComponent   = 1
          UninstallString   = 'msiexec /package {90120000-001B-0407-0000-0000000FF1CE} /uninstall {DB2ACBD1-65B1-4FC5-881E-4E75C668E7E2}'
        }
      }
    )

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Msi.Count | Should -Be 1
    $observed.Msi[0].ResourceKind | Should -Be LanguageResource
    $patch = @($observed.RelatedComponents | Where-Object Role -EQ PatchRegistration)
    $patch.Count | Should -Be 1
    $patch[0].ParentProductCode | Should -Be '{90120000-001B-0407-0000-0000000FF1CE}'
    $patch[0].SystemComponent | Should -BeTrue
    $observed.Unknowns.Count | Should -Be 0
  }

  It 'keeps a patch registration unknown when its product is absent' {
    $script:sparseRegistry = @([PSCustomObject]@{
        View   = 'Registry32'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90120000-001B-0407-0000-0000000FF1CE}_ENTERPRISER_{DB2ACBD1-65B1-4FC5-881E-4E75C668E7E2}'
        Values = [PSCustomObject]@{
          Publisher       = 'Microsoft'
          DisplayName     = 'Microsoft Office 2007 Service Pack 3 (SP3)'
          SystemComponent = 1
        }
      })

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Msi.Count | Should -Be 0
    $observed.Unknowns | Should -Contain 'OrphanedOfficePatchRegistration:{90120000-001B-0407-0000-0000000FF1CE}_ENTERPRISER_{DB2ACBD1-65B1-4FC5-881E-4E75C668E7E2}'
  }

  It 'recognizes Click-to-Run infrastructure without a component SKU allowlist (<Code>/<Name>)' -ForEach @(
    @{ Code = '007E'; Name = 'Office 16 Click-to-Run Licensing Component' }
    @{ Code = '008F'; Name = 'Office 16 Click-to-Run Licensing Component' }
    @{ Code = '0BAD'; Name = 'Office 16 Click-to-Run Licensing Component' }
    @{ Code = '0BAD'; Name = 'Office 16 Click-to-Run Extensibility Component 64-bit Registration' }
    @{ Code = '0BAD'; Name = 'Office 16 Click-to-Run Localization Component' }
  ) {
    # 007E and 008F are native observations from two 2019 installations.
    # 0BAD is deliberately synthetic: classification must not depend on a
    # growing list of observed SKU fragments.
    $script:sparseRegistry = @([PSCustomObject]@{
        View   = 'Registry64'
        Path   = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-$Code-0000-1000-0000000FF1CE}"
        Values = [PSCustomObject]@{
          Publisher   = 'Microsoft Corporation'
          DisplayName = $Name
        }
      })

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Msi.Count | Should -Be 0
    @($observed.RelatedComponents | Where-Object Role -EQ ClickToRunInfrastructure).Count | Should -Be 1
    $observed.Unknowns | Should -Contain ClickToRunInfrastructureWithoutConfiguration
  }

  It 'does not hide an unrelated or inconsistent component as Click-to-Run infrastructure (<Name>/<Code>/<Publisher>)' -ForEach @(
    @{ Code = '{90160000-007E-0000-1000-0000000FF1CE}'; Name = 'Microsoft Office Professional Plus 2016'; Publisher = 'Microsoft Corporation' }
    @{ Code = '{90160000-0BAD-0000-1000-0000000FF1CE}'; Name = 'Office 16 Click-to-Run Unknown Component'; Publisher = 'Microsoft Corporation' }
    @{ Code = '{90160000-007E-0000-1000-0000000FF1CE}'; Name = 'Office 16 Click-to-Run Licensing Component'; Publisher = 'Other publisher' }
    @{ Code = '{90140000-2005-0000-0000-0000000FF1CE}'; Name = 'Office 16 Click-to-Run Licensing Component'; Publisher = 'Microsoft Corporation' }
    @{ Code = '{90140000-2005-0000-0000-0000000FF1CE}'; Name = 'Microsoft Office File Validation Add-In'; Publisher = 'Microsoft Corporation' }
  ) {
    $script:sparseRegistry = @([PSCustomObject]@{
        View   = 'Registry64'
        Path   = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$Code"
        Values = [PSCustomObject]@{
          Publisher        = $Publisher
          DisplayName      = $Name
          WindowsInstaller = 1
        }
      })
    $observed = PSFoundation\Get-OfficeInventory
    @($observed.RelatedComponents | Where-Object Role -EQ ClickToRunInfrastructure).Count | Should -Be 0
    if ($Publisher -eq 'Microsoft Corporation') {
      $observed.Msi.Count | Should -Be 1
    }
  }

  It 'does not accept a ClientCulture the resource registration contradicts' {
    # Two registered languages make the shell language ambiguous, so it needs
    # ClientCulture to agree. A ClientCulture outside the registered set is a
    # conflict and must stay unknown rather than become the answer.
    $script:sparseRegistry = @(
      [PSCustomObject]@{
        View    = 'Registry64'
        Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        Values  = [PSCustomObject]@{ ProductReleaseIds = 'Standard2019Volume'; Platform = 'x64'; ClientCulture = 'fr-fr' }
        SubKeys = @()
      }
      [PSCustomObject]@{
        View    = 'Registry64'
        Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
        Values  = [PSCustomObject]@{ ActiveConfiguration = '69ddc2fc-20c6-4c3e-ad5b-68f318d54698' }
        SubKeys = @('69ddc2fc-20c6-4c3e-ad5b-68f318d54698')
      }
      [PSCustomObject]@{
        View    = 'Registry64'
        Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\69ddc2fc-20c6-4c3e-ad5b-68f318d54698\Standard2019Volume.16'
        Values  = [PSCustomObject]@{}
        SubKeys = @('de-de', 'en-us', 'x-none')
      }
    )

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Products[0].Languages | Should -Be @('de-de', 'en-us')
    $observed.Products[0].PrimaryLanguage | Should -BeNullOrEmpty
    $observed.VerificationLimitations | Should -Contain PrimaryLanguage
    $observed.VerificationLimitations | Should -Not -Contain Languages
    ($observed.Products[0].Evidence -join ' ') | Should -BeLike '*ClientCulture=fr-fr is not in the registered languages*'
  }

  It 'settles an ambiguous shell language when ClientCulture agrees' {
    $script:sparseRegistry = @(
      [PSCustomObject]@{
        View    = 'Registry64'
        Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        Values  = [PSCustomObject]@{ ProductReleaseIds = 'Standard2019Volume'; Platform = 'x64'; ClientCulture = 'en-us' }
        SubKeys = @()
      }
      [PSCustomObject]@{
        View    = 'Registry64'
        Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
        Values  = [PSCustomObject]@{ ActiveConfiguration = '69ddc2fc-20c6-4c3e-ad5b-68f318d54698' }
        SubKeys = @('69ddc2fc-20c6-4c3e-ad5b-68f318d54698')
      }
      [PSCustomObject]@{
        View    = 'Registry64'
        Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\69ddc2fc-20c6-4c3e-ad5b-68f318d54698\Standard2019Volume.16'
        Values  = [PSCustomObject]@{}
        SubKeys = @('de-de', 'en-us', 'x-none')
      }
    )

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Products[0].Languages | Should -Be @('de-de', 'en-us')
    $observed.Products[0].PrimaryLanguage | Should -Be 'en-us'
    $observed.VerificationLimitations.Count | Should -Be 0
  }

  It 'does not read Office 2007 App Paths as Click-to-Run residue' {
    $script:sparseRegistry = @(
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
        Values = [PSCustomObject]@{ '(default)' = 'C:\PROGRA~2\MICROS~2\Office12\WINWORD.EXE' }
      }
      [PSCustomObject]@{
        View   = 'Registry32'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
        Values = [PSCustomObject]@{ '(default)' = 'C:\PROGRA~2\MICROS~2\Office12\OUTLOOK.EXE' }
      }
    )

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Unknowns | Should -Not -Contain OfficeResidueWithoutConfiguration
  }

  It 'still reports Click-to-Run App Paths without a configuration as residue' {
    Mock Get-PSFOfficePathAttribute {
      if ($LiteralPath -like '*.exe') { [IO.FileAttributes]::Normal }
      else { [IO.FileAttributes]::Directory }
    } -ModuleName PSFoundation
    $script:sparseRegistry = @([PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
        Values = [PSCustomObject]@{ '(default)' = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' }
      })

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Unknowns | Should -Contain OfficeResidueWithoutConfiguration
  }

  It 'preserves incomplete Click-to-Run registration as unknown evidence' {
    $script:sparseRegistry = @([PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        Values = [PSCustomObject]@{}
      })

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Unknowns | Should -Contain IncompleteClickToRunRegistration
  }

  It 'preserves missing configuration, installed-version and language fields as unknown' {
    $script:sparseRegistry = @(
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
        Values = [PSCustomObject]@{ ProductReleaseIds = 'Standard2019Volume' }
      }
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0'
        Values = [PSCustomObject]@{ OfficeProductReleaseIds = 'Standard2019Volume' }
      }
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
        Values = [PSCustomObject]@{}
      }
      [PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources'
        Values = [PSCustomObject]@{ SKULanguage = 1031 }
      }
    )

    $observed = PSFoundation\Get-OfficeInventory
    $observed.Products.Count | Should -Be 1
    $observed.Products[0].Architecture | Should -BeNullOrEmpty
    $observed.Products[0].Version | Should -BeNullOrEmpty
    $observed.Products[0].Channel | Should -BeNullOrEmpty
    $observed.Products[0].Languages | Should -BeNullOrEmpty
    $observed.Products[0].RegisteredLanguages.Count | Should -Be 0
    ($null -eq $observed.Products[0].ExcludeApp) | Should -BeTrue
    $observed.RegisteredResources[0].ActiveConfiguration | Should -BeNullOrEmpty
    $observed.LanguageEvidence[0].InstallLanguage | Should -BeNullOrEmpty
    $observed.LanguageEvidence[0].SKULanguage | Should -Be 1031

    $script:sparseRegistry[0].Values | Add-Member -NotePropertyName 'Standard2019Volume.ExcludedApps' -NotePropertyValue ''
    $observed = PSFoundation\Get-OfficeInventory
    ($null -eq $observed.Products[0].ExcludeApp) | Should -BeFalse
    $observed.Products[0].ExcludeApp.Count | Should -Be 0
  }

  It 'plans ordinary MSI-only Office 2007 migration with empty Click-to-Run collections' {
    $script:sparseRegistry = Get-Content "$PSScriptRoot/fixtures/office/office2007-registry.json" -Raw | ConvertFrom-Json
    Mock Test-OfficeDeploymentMedia {
      [PSCustomObject]@{
        Valid       = $true
        Fingerprint = 'synthetic-media'
        Manifest    = [PSCustomObject]@{ Version = '16.0.10417.20211' }
      }
    } -ModuleName PSFoundation

    $target = PSFoundation\New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 64 -Language de-de
    $planArgs = @{
      Action        = 'Migrate'
      Configuration = $target
      SourcePath    = 'C:\Media\Office2019'
      RemoveMsi     = $true
    }

    $plan = PSFoundation\Get-OfficeDeploymentPlan @planArgs
    $plan.Eligible | Should -BeTrue
    $plan.LanguageTransition.Known | Should -BeFalse
    $plan.LanguageTransition.PrimaryBefore.Count | Should -Be 0
    $plan.Configuration.Version | Should -Be '16.0.10417.20211'

    $plan = PSFoundation\Get-OfficeDeploymentPlan @planArgs -RemoveProductId Standard2019Volume
    $plan.Eligible | Should -BeFalse
    $plan.Blockers | Should -Contain StaleRemovalSelection
  }

  It 'reads the Office <Release> fixture through the strict imported module' -ForEach @(
    @{ Release = '2007'; MsiExpected = $true; ProductCount = 0; RelatedCount = 2 }
    @{ Release = '2019'; MsiExpected = $false; ProductCount = 1; RelatedCount = 5 }
  ) {
    $script:sparseRegistry = Get-Content "$PSScriptRoot/fixtures/office/office$Release-registry.json" -Raw | ConvertFrom-Json

    $observed = PSFoundation\Get-OfficeInventory
    ($observed.Msi.Count -gt 0) | Should -Be $MsiExpected
    $observed.Products.Count | Should -Be $ProductCount
    $observed.RelatedComponents.Count | Should -Be $RelatedCount
    $observed.Unknowns.Count | Should -Be 0
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

Describe 'Office native registry regression fixtures' {
  BeforeEach {
    $script:registryFixture = @()
    Mock Get-PSFOfficeRegistrySnapshot { $script:registryFixture }
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' }
  }

  It 'keeps Office 2007 proofing languages separate from UI resources and add-ins' {
    $script:registryFixture = Get-Content "$PSScriptRoot/fixtures/office/office2007-registry.json" -Raw | ConvertFrom-Json
    $observed = Get-OfficeInventory
    $observed.Unknowns.Count | Should -Be 0
    @($observed.Msi | Where-Object Name -Match 'Teams Meeting|Office Live').Count | Should -Be 0
    @($observed.RelatedComponents | Where-Object Role -EQ AddIn).Count | Should -Be 2
    @($observed.Msi | Where-Object ResourceKind -EQ Proofing | Select-Object -ExpandProperty LanguageId -Unique | Sort-Object) | Should -Be @('de-de', 'en-us', 'fr-fr', 'it-it')
    @($observed.Msi | Where-Object ResourceKind -EQ LanguageResource | Select-Object -ExpandProperty LanguageId -Unique) | Should -Be @('de-de')
    @($observed.Msi | Where-Object Name -EQ 'Microsoft Office Enterprise 2007').Count | Should -Be 1
    @($observed.LanguageEvidence.SKULanguage | Where-Object { $_ }) | Should -Be @(1031)
    $script:nativeObserved = $observed
    Mock Get-OfficeInventory { $script:nativeObserved }
    { New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -AutoSourceLocales } | Should -Throw '*preservation*'
  }

  It 'does not classify Click-to-Run infrastructure as legacy MSI Office' {
    $script:registryFixture = Get-Content "$PSScriptRoot/fixtures/office/office2019-registry.json" -Raw | ConvertFrom-Json
    $observed = Get-OfficeInventory
    $observed.Msi.Count | Should -Be 0
    $observed.Unknowns.Count | Should -Be 0
    @($observed.RelatedComponents | Where-Object Role -EQ ClickToRunInfrastructure).Count | Should -Be 4
    @($observed.RelatedComponents | Where-Object Role -EQ AddIn).Count | Should -Be 1
    $observed.Products[0].ProductId | Should -Be Standard2019Volume
    $observed.Products[0].Architecture | Should -Be '32'
    $observed.Products[0].RegisteredLanguages | Should -Be @('de-de')
    $observed.Products[0].Version | Should -Be '16.0.10417.20208'
    $observed.Products[0].VersionSource | Should -Be ActiveProductResources
    # Derived from the active per-product resource registration; the single
    # registered language also settles the shell language without ClientCulture.
    $observed.Products[0].Languages | Should -Be @('de-de')
    $observed.Products[0].PrimaryLanguage | Should -Be 'de-de'
    $observed.VerificationLimitations.Count | Should -Be 0
    $observed.Products[0].ExcludeApp | Should -Be @('groove')
  }

  It 'uses only the active product language registration' {
    $script:registryFixture = Get-Content "$PSScriptRoot/fixtures/office/office2019-registry.json" -Raw | ConvertFrom-Json
    $script:registryFixture += [PSCustomObject]@{
      View    = 'Registry64'
      Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs\00000000-0000-0000-0000-000000000002\Standard2019Volume.16'
      Values  = [PSCustomObject]@{}
      SubKeys = @('en-us', 'fr-fr', 'x-none')
    }
    (Get-OfficeInventory).Products[0].RegisteredLanguages | Should -Be @('de-de')
    ($script:registryFixture | Where-Object Path -EQ 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs').Values.ActiveConfiguration = 'invalid'
    (Get-OfficeInventory).Products[0].RegisteredLanguages.Count | Should -Be 0
  }

  It 'does not treat orphaned Click-to-Run infrastructure as a clean machine' {
    $script:registryFixture = Get-Content "$PSScriptRoot/fixtures/office/office2019-registry.json" -Raw | ConvertFrom-Json
    $script:registryFixture = @($script:registryFixture | Where-Object { $_.Path -like '*Uninstall*' })
    $observed = Get-OfficeInventory
    $observed.Msi.Count | Should -Be 0
    $observed.Unknowns | Should -Contain ClickToRunInfrastructureWithoutConfiguration
  }

  It 'keeps unfamiliar add-ins blocked rather than ignoring arbitrary Office registrations' {
    $script:registryFixture = @([PSCustomObject]@{
        View   = 'Registry64'
        Path   = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\synthetic-office-component'
        Values = [PSCustomObject]@{ Publisher = 'Microsoft Corporation'; DisplayName = 'Microsoft Office Unknown Add-in'; DisplayVersion = '1.0'; WindowsInstaller = 1 }
      })
    $observed = Get-OfficeInventory
    $observed.Msi.Count | Should -Be 1
    $observed.RelatedComponents.Count | Should -Be 0
  }

  It 'blocks native installation before mutation while required verification is unavailable' {
    # An installed product whose active configuration cannot be resolved has no
    # derivable languages. That is a real limitation and must still block before
    # mutation, even though a machine with nothing unverifiable no longer does.
    $script:registryFixture = Get-Content "$PSScriptRoot/fixtures/office/office2019-registry.json" -Raw | ConvertFrom-Json
    ($script:registryFixture | Where-Object Path -EQ 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs').Values.ActiveConfiguration = 'invalid'
    $observed = Get-OfficeInventory
    $observed.Products[0].Languages | Should -BeNullOrEmpty
    $observed.Products[0].PrimaryLanguage | Should -BeNullOrEmpty
    $observed.VerificationLimitations | Should -Contain Languages
    $observed.VerificationLimitations | Should -Contain PrimaryLanguage

    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 64 -Language de-de -Version 16.0.10417.20208
    Mock Test-OfficeDeploymentMedia { [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.10417.20208' } } }
    $plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath C:\Media -Inventory $observed
    $plan.Eligible | Should -BeFalse
    $plan.Blockers | Should -Contain UnsupportedNativeVerification
  }

  It 'does not block a clean machine that has nothing unverifiable to observe' {
    $observed = Get-OfficeInventory
    $observed.Products.Count | Should -Be 0
    $observed.VerificationLimitations.Count | Should -Be 0

    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 64 -Language de-de -Version 16.0.10417.20208
    Mock Test-OfficeDeploymentMedia { [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.10417.20208' } } }
    $plan = Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath C:\Media -Inventory $observed
    $plan.Blockers | Should -Not -Contain UnsupportedNativeVerification
    $plan.Eligible | Should -BeTrue
  }

  It 'reports migration verification limits without false unsupported-MSI blockers' {
    $script:registryFixture = Get-Content "$PSScriptRoot/fixtures/office/office2007-registry.json" -Raw | ConvertFrom-Json
    $observed = Get-OfficeInventory
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 64 -Language de-de -Version 16.0.10417.20208
    Mock Test-OfficeDeploymentMedia { [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.10417.20208' } } }
    $plan = Get-OfficeDeploymentPlan -Action Migrate -RemoveMsi -Configuration $target -SourcePath C:\Media -Inventory $observed
    $plan.Blockers | Should -Not -Contain UnsupportedMsiComponent
    # An MSI-only source has no Click-to-Run product to sample, so nothing
    # observed is unverifiable and ordinary migration is no longer blocked.
    # Any post-deployment gap still surfaces as a verification Unknown.
    $plan.Blockers | Should -Not -Contain UnsupportedNativeVerification
    $plan.Eligible | Should -BeTrue
    # Source-side MSI language derivation is not implemented yet, so the
    # transition remains explicitly unknown rather than assumed.
    $plan.LanguageTransition.Known | Should -BeFalse
  }

  It 'detects disappearance of an add-in outside removal authority' {
    $component = [PSCustomObject]@{ ProductCode = 'synthetic-addin'; Name = 'Synthetic Add-in'; Version = '1.0'; RegistryView = 'Registry64'; Role = 'AddIn' }
    $before = New-TestOfficeInventory
    $before | Add-Member -NotePropertyName RelatedComponents -NotePropertyValue @($component)
    $after = New-TestOfficeInventory
    $after | Add-Member -NotePropertyName RelatedComponents -NotePropertyValue @($component)
    $plan = [PSCustomObject]@{ Action = 'Remove'; Before = $before; RemoveProductId = @('Standard2019Volume') }
    (Test-PSFOfficePostcondition $plan $after).Compliant | Should -BeTrue
    $after.RelatedComponents = @()
    $result = Test-PSFOfficePostcondition $plan $after
    $result.Compliant | Should -BeFalse
    $result.Discrepancies | Should -Contain RelatedAddInChanged
  }
}

Describe 'Imported Office installed-version evidence precedence' {
  BeforeAll {
    Remove-Module -Name PSFoundation -Force -ErrorAction SilentlyContinue
    Import-Module "$PSScriptRoot/../src/PSFoundation.psd1" -Force
  }

  BeforeEach {
    $script:versionRegistry = Get-Content "$PSScriptRoot/fixtures/office/office2019-registry.json" -Raw | ConvertFrom-Json
    $script:versionRoot = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
    $script:versionProduct = $versionRoot + '\00000000-0000-0000-0000-000000000001\Standard2019Volume.16'
    $script:documentedVersion = [PSCustomObject]@{
      View    = 'Registry64'
      Path    = 'SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0'
      Values  = [PSCustomObject]@{ OfficeProductReleaseIds = 'Standard2019Volume'; OfficePackageVersion = '16.0.10417.20299' }
      SubKeys = @()
    }
    Mock Get-PSFOfficeRegistrySnapshot { $script:versionRegistry } -ModuleName PSFoundation
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' } -ModuleName PSFoundation
  }

  It 'derives the observed <Platform>/<Build> resource shape independently of telemetry' -ForEach @(
    @{ Platform = 'x86'; Build = '16.0.10417.20208'; Architecture = '32' }
    @{ Platform = 'x64'; Build = '16.0.10417.20211'; Architecture = '64' }
  ) {
    $configuration = ($versionRegistry | Where-Object { $_.Path -like '*\Configuration' }).Values
    $configuration.Platform = $Platform
    $configuration.VersionToReport = '16.0.99999.99999'
    foreach ($record in $versionRegistry | Where-Object { $_.Values.PSObject.Properties['Version'] }) {
      $record.Values.Version = $Build
    }
    $observed = PSFoundation\Get-OfficeInventory
    $observed.Products[0].Version | Should -Be $Build
    $observed.Products[0].VersionSource | Should -Be ActiveProductResources
    $observed.Products[0].Architecture | Should -Be $Architecture
    ($observed.Products[0].Evidence -join ' ') | Should -Match 'derived from agreeing Version values'
    @($observed.RegisteredResources | Where-Object { $_.Version -eq $Build }).Count | Should -Be 4
  }

  It 'prefers valid documented inventory over lower-priority resource versions' {
    $script:versionRegistry += $documentedVersion
    $product = (PSFoundation\Get-OfficeInventory).Products[0]
    $product.Version | Should -Be '16.0.10417.20299'
    $product.VersionSource | Should -Be ClickToRunInventory
  }

  It 'falls through when the documented version value is absent and product identity agrees' {
    $documentedVersion.Values.PSObject.Properties.Remove('OfficePackageVersion')
    $script:versionRegistry += $documentedVersion
    $product = (PSFoundation\Get-OfficeInventory).Products[0]
    $product.Version | Should -Be '16.0.10417.20208'
    $product.VersionSource | Should -Be ActiveProductResources
  }

  It 'does not fall through a malformed documented version (<Label>)' -ForEach @(
    @{ Label = 'blank'; Value = '' }
    @{ Label = 'null'; Value = $null }
    @{ Label = 'partial'; Value = '16.0' }
    @{ Label = 'overflow'; Value = '16.0.99999999999.1' }
    @{ Label = 'array'; Value = @('16.0.10417.20208') }
  ) {
    $documentedVersion.Values.OfficePackageVersion = $Value
    $script:versionRegistry += $documentedVersion
    $product = (PSFoundation\Get-OfficeInventory).Products[0]
    $product.Version | Should -BeNullOrEmpty
    $product.VersionSource | Should -BeNullOrEmpty
    ($product.Evidence -join ' ') | Should -Match 'malformed.*OfficePackageVersion'
  }

  It 'does not fall through conflicting documented product identity' {
    $documentedVersion.Values.OfficeProductReleaseIds = 'ProPlus2019Volume'
    $script:versionRegistry += $documentedVersion
    $observed = PSFoundation\Get-OfficeInventory
    $observed.Products[0].Version | Should -BeNullOrEmpty
    $observed.Unknowns | Should -Contain ConflictingInstalledProductIdentity
  }

  It 'does not choose arbitrarily between duplicate documented records' {
    $script:versionRegistry += @($documentedVersion, $documentedVersion)
    (PSFoundation\Get-OfficeInventory).Products[0].Version | Should -BeNullOrEmpty
  }

  It 'keeps missing or conflicting active product evidence unknown (<Change>)' -ForEach @(
    @{ Change = 'MissingVersion' }
    @{ Change = 'MalformedVersion' }
    @{ Change = 'MixedVersions' }
    @{ Change = 'MissingLeaf' }
    @{ Change = 'DuplicateLeaf' }
    @{ Change = 'MissingNeutral' }
    @{ Change = 'NeutralOnly' }
    @{ Change = 'AdditionalLanguageWithoutVersion' }
    @{ Change = 'InvalidActive' }
    @{ Change = 'WrongView' }
    @{ Change = 'NoProductLeaves' }
  ) {
    $leaf = $versionRegistry | Where-Object { $_.Path -eq ($versionProduct + '\de-de') }
    $parent = $versionRegistry | Where-Object { $_.Path -eq $versionProduct }
    switch ($Change) {
      MissingVersion { $leaf.Values.PSObject.Properties.Remove('Version') }
      MalformedVersion { $leaf.Values.Version = '16.0' }
      MixedVersions { $leaf.Values.Version = '16.0.10417.20211' }
      MissingLeaf { $script:versionRegistry = @($versionRegistry | Where-Object { $_.Path -ne $leaf.Path }) }
      DuplicateLeaf { $script:versionRegistry += $leaf }
      MissingNeutral { $parent.SubKeys = @('de-de') }
      NeutralOnly { $parent.SubKeys = @('x-none') }
      AdditionalLanguageWithoutVersion { $parent.SubKeys += 'en-us' }
      InvalidActive { ($versionRegistry | Where-Object Path -EQ $versionRoot).Values.ActiveConfiguration = 'invalid' }
      WrongView { $leaf.View = 'Registry32' }
      NoProductLeaves { $script:versionRegistry = @($versionRegistry | Where-Object { -not $_.Path.StartsWith($versionProduct + '\') }) }
    }
    $observed = PSFoundation\Get-OfficeInventory
    $observed.Products[0].Version | Should -BeNullOrEmpty
    $target = PSFoundation\New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 32 -Language de-de -ExcludeApp Groove -Version 16.0.10417.20208
    $verification = PSFoundation\Test-OfficeDeployment -Configuration $target -Inventory $observed
    $verification.Compliant | Should -BeFalse
    $verification.Unknowns | Should -Contain Version
  }

  It 'ignores inactive resource trees even when they match the desired build' {
    $inactive = $versionRegistry | Where-Object { $_.Path.StartsWith($versionProduct) } | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    foreach ($record in $inactive) {
      $record.Path = $record.Path.Replace('00000000-0000-0000-0000-000000000001', '00000000-0000-0000-0000-000000000002')
    }
    $script:versionRegistry += $inactive
    ($versionRegistry | Where-Object { $_.Path -eq ($versionProduct + '\de-de') }).Values.PSObject.Properties.Remove('Version')
    (PSFoundation\Get-OfficeInventory).Products[0].Version | Should -BeNullOrEmpty
  }

  It 'does not borrow another products version or require a product-year allowlist' {
    $other = $versionRegistry | Where-Object { $_.Path.StartsWith($versionProduct) } | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    foreach ($record in $other) {
      $record.Path = $record.Path.Replace('Standard2019Volume', 'Standard2024Volume')
    }
    ($versionRegistry | Where-Object { $_.Path -like '*\Configuration' }).Values.ProductReleaseIds = 'Standard2019Volume,Standard2024Volume'
    $script:versionRegistry += $other
    ($versionRegistry | Where-Object { $_.Path -eq ($versionProduct + '\de-de') }).Values.PSObject.Properties.Remove('Version')
    $observed = PSFoundation\Get-OfficeInventory
    ($observed.Products | Where-Object ProductId -EQ Standard2019Volume).Version | Should -BeNullOrEmpty
    ($observed.Products | Where-Object ProductId -EQ Standard2024Volume).Version | Should -Be '16.0.10417.20208'
  }

  It 'reports a compliant repeat run without launching an installer' {
    $target = PSFoundation\New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 32 -Language de-de -ExcludeApp Groove -Version 16.0.10417.20208
    $plan = PSFoundation\Get-OfficeDeploymentPlan -Action Install -Configuration $target
    $plan.State | Should -Be Compliant
    Mock Get-OfficeActivationStatus { [PSCustomObject]@{ Status = 'Licensed' } } -ModuleName PSFoundation
    Mock Invoke-PSFOfficeConfiguration { throw 'Must not launch ODT' } -ModuleName PSFoundation
    Mock Assert-PSFOfficeHost { throw 'Must not enter mutation checks' } -ModuleName PSFoundation
    $result = PSFoundation\Install-Office -Plan $plan -OdtPath C:\Unused\setup.exe -Confirm:$false
    $result.AlreadyCompliant | Should -BeTrue
    $result.Changed | Should -BeFalse
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0 -ModuleName PSFoundation
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

Describe 'Office App Paths target evidence' {
  BeforeEach {
    $script:appKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    Mock Get-PSFOfficePathEnvironment { @{ ProgramFiles = 'C:\NativeApps'; 'ProgramFiles(x86)' = 'C:\Apps32' } }
    Mock Get-PSFOfficePathAttribute {
      if ($LiteralPath -like '*.exe') { [IO.FileAttributes]::Normal }
      else { [IO.FileAttributes]::Directory }
    }
  }

  It 'resolves quoted and expanded paths literally for <View>' -ForEach @(
    @{ View = 'Registry32'; Expected = 'C:\Apps32\Office16\WINWORD.EXE' }
    @{ View = 'Registry64'; Expected = 'C:\NativeApps\Office16\WINWORD.EXE' }
  ) {
    Mock Get-PSFOfficePathEnvironment {
      if ($View -eq 'Registry32') { @{ ProgramFiles = 'C:\Apps32' } }
      else { @{ ProgramFiles = 'C:\NativeApps' } }
    }
    $raw = '"%ProgramFiles%\Office16\WINWORD.EXE"'
    $result = Get-PSFOfficeAppPathEvidence -RawTarget $raw -View $View -RegistryPath $appKey
    $result.State | Should -Be Present
    $result.RawTarget | Should -Be $raw
    $result.ResolvedTarget | Should -Be $Expected
    $result.RegistryView | Should -Be $View
    Should -Invoke Get-PSFOfficePathAttribute -Times 1 -ParameterFilter { $LiteralPath -eq $Expected }
  }

  It 'reports <Failure> as <State> rather than silently treating every failure as absence' -ForEach @(
    @{ Failure = 'FileNotFound'; State = 'Missing'; Reason = 'TargetNotFound' }
    @{ Failure = 'DirectoryNotFound'; State = 'Missing'; Reason = 'TargetNotFound' }
    @{ Failure = 'Denied'; State = 'Uncertain'; Reason = 'AccessDenied' }
    @{ Failure = 'IO'; State = 'Uncertain'; Reason = 'ProbeFailed' }
    @{ Failure = 'Directory'; State = 'Uncertain'; Reason = 'UnexpectedPathType' }
  ) {
    $script:probeFailure = $Failure
    Mock Get-PSFOfficePathAttribute {
      switch ($script:probeFailure) {
        FileNotFound { throw [IO.FileNotFoundException]::new('Synthetic missing file') }
        DirectoryNotFound { throw [IO.DirectoryNotFoundException]::new('Synthetic missing directory') }
        Denied { throw [UnauthorizedAccessException]::new('Synthetic access failure') }
        IO { throw [IO.IOException]::new('Synthetic I/O failure') }
        Directory { [IO.FileAttributes]::Directory }
      }
    } -ParameterFilter { $LiteralPath -like '*.exe' }
    $result = Get-PSFOfficeAppPathEvidence -RawTarget 'C:\Apps\Office16\WINWORD.EXE' -View Registry64 -RegistryPath $appKey
    $result.State | Should -Be $State
    $result.Reason | Should -Be $Reason
  }

  It 'does not traverse a reparse ancestor or mistake an intermediate file for absence' -ForEach @(
    @{ Attributes = [IO.FileAttributes]::ReparsePoint; Reason = 'ReparsePoint' }
    @{ Attributes = [IO.FileAttributes]::Normal; Reason = 'UnexpectedPathType' }
  ) {
    $script:ancestorAttributes = $Attributes
    Mock Get-PSFOfficePathAttribute { $script:ancestorAttributes } -ParameterFilter { $LiteralPath -eq 'C:\Apps' }
    $result = Get-PSFOfficeAppPathEvidence -RawTarget 'C:\Apps\Office16\WINWORD.EXE' -View Registry64 -RegistryPath $appKey
    $result.State | Should -Be Uncertain
    $result.Reason | Should -Be $Reason
    Should -Invoke Get-PSFOfficePathAttribute -Times 0 -ParameterFilter { $LiteralPath -like 'C:\Apps\*' }
  }

  It 'rejects unsupported or ambiguous target <Raw> without filesystem access' -ForEach @(
    @{ Raw = '\\server\share\Office16\WINWORD.EXE' }
    @{ Raw = '\\?\C:\Office16\WINWORD.EXE' }
    @{ Raw = 'Office16\WINWORD.EXE' }
    @{ Raw = 'C:\Office16\..\WINWORD.EXE' }
    @{ Raw = 'C:\Office16\WINWORD.EXE:stream' }
    @{ Raw = 'C:\Office16\WINWORD.EXE /argument' }
    @{ Raw = 'C:\Office*\WINWORD.EXE' }
    @{ Raw = 'C:\Office16\WINWORD.EXE.' }
    @{ Raw = 'C:\Office16\EXCEL.EXE' }
    @{ Raw = 'C:\missing\CON\Office16\WINWORD.EXE' }
    @{ Raw = ('C:\' + ('x' * 256) + '\Office16\WINWORD.EXE') }
    @{ Raw = '%UNKNOWN_OFFICE_ROOT%\Office16\WINWORD.EXE' }
    @{ Raw = '' }
    @{ Raw = @('C:\Office16\WINWORD.EXE', 'other') }
  ) {
    (Get-PSFOfficeAppPathEvidence -RawTarget $Raw -View Registry64 -RegistryPath $appKey).State | Should -Be Uncertain
    Should -Invoke Get-PSFOfficePathAttribute -Times 0
  }

  It 'does not expand unapproved nested environment tokens' {
    Mock Get-PSFOfficePathEnvironment { @{ ProgramFiles = '%UNRESOLVED%' } }
    (Get-PSFOfficeAppPathEvidence -RawTarget '%ProgramFiles%\Office16\WINWORD.EXE' -View Registry32 -RegistryPath $appKey).State | Should -Be Uncertain
    Should -Invoke Get-PSFOfficePathAttribute -Times 0
  }

  It 'does not classify an inaccessible root as a missing executable' {
    Mock Get-PSFOfficePathAttribute { throw [IO.DirectoryNotFoundException]::new('Synthetic unavailable root') } -ParameterFilter { $LiteralPath -eq 'C:\' }
    $result = Get-PSFOfficeAppPathEvidence -RawTarget 'C:\Office16\WINWORD.EXE' -View Registry64 -RegistryPath $appKey
    $result.State | Should -Be Uncertain
    $result.Reason | Should -Be ProbeFailed
  }

  It 'does not infer absence from WOW64 system-path redirection or short aliases' -ForEach @(
    @{ Target = 'system' }, @{ Target = 'alias' }
  ) {
    $path = Join-Path $env:SystemRoot 'System32\Office16\WINWORD.EXE'
    if ($Target -eq 'alias') { $path = 'C:\WINDOW~1\System32\Office16\WINWORD.EXE' }
    $result = Get-PSFOfficeAppPathEvidence -RawTarget $path -View Registry64 -RegistryPath $appKey
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
      $result.State | Should -Be Uncertain
      $result.Reason | Should -Be FilesystemRedirection
      Should -Invoke Get-PSFOfficePathAttribute -Times 0
    }
    else { $result.State | Should -Be Present }
  }

  It 'selects Program Files consistently from the registry view' {
    # Exercise the real environment mapping in its own module scope.
    $native = & (Get-Module PSFoundation) { Get-PSFOfficePathEnvironment Registry64 }
    $x86 = & (Get-Module PSFoundation) { Get-PSFOfficePathEnvironment Registry32 }
    if ([Environment]::Is64BitOperatingSystem) {
      $native.ProgramFiles | Should -Be $env:ProgramW6432
      $x86.ProgramFiles | Should -Be ${env:ProgramFiles(x86)}
    }
    else {
      $native.ProgramFiles | Should -Be $env:ProgramFiles
      $x86.ProgramFiles | Should -Be $env:ProgramFiles
    }
  }

  It 'distinguishes an existing and deleted synthetic executable using native attributes' {
    $file = Join-Path $TestDrive 'WINWORD.EXE'
    [IO.File]::WriteAllText($file, 'Synthetic data; never executed.')
    $present = & (Get-Module PSFoundation) {
      param ($TargetPath)
      Get-PSFOfficeAppPathEvidence -RawTarget $TargetPath -View Registry64 -RegistryPath 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    } $file
    $present.State | Should -Be Present
    Remove-Item -LiteralPath $file
    $missing = & (Get-Module PSFoundation) {
      param ($TargetPath)
      Get-PSFOfficeAppPathEvidence -RawTarget $TargetPath -View Registry32 -RegistryPath 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    } $file
    $missing.State | Should -Be Missing
    $missing.Reason | Should -Be TargetNotFound
  }
}

Describe 'Office stale App Paths inventory' {
  BeforeEach {
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' }
    $script:appRecords = @()
    Mock Get-PSFOfficeRegistrySnapshot { $script:appRecords }
    Mock Get-PSFOfficePathAttribute {
      if ($LiteralPath -like '*.exe') { throw [IO.FileNotFoundException]::new('Synthetic missing file') }
      [IO.FileAttributes]::Directory
    }
  }

  It 'retains <Count> missing references without inventing installed products under strict mode' -ForEach @(
    @{ Count = 0 }, @{ Count = 1 }, @{ Count = 6 }
  ) {
    $script:appRecords = @(foreach ($view in @('Registry32', 'Registry64')) {
        foreach ($name in @('WINWORD.EXE', 'EXCEL.EXE', 'OUTLOOK.EXE')) {
          [PSCustomObject]@{ View = $view; Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$name"; Values = [PSCustomObject]@{ '(default)' = "C:\Program Files (x86)\Microsoft Office\Root\Office16\$name" }; SubKeys = @() }
        }
      })
    $script:appRecords = @($script:appRecords | Select-Object -First $Count)
    Set-StrictMode -Version Latest
    $observed = Get-OfficeInventory
    $observed.Unknowns.Count | Should -Be 0
    $observed.Products.Count | Should -Be 0
    $observed.AppPathEvidence.Count | Should -Be $Count
    @($observed.AppPathEvidence | Where-Object State -NE Missing).Count | Should -Be 0
    if ($Count) {
      $observed.AppPathEvidence[0].RegistryPath | Should -Be $appRecords[0].Path
      $observed.AppPathEvidence[0].RawTarget | Should -Be $appRecords[0].Values.'(default)'
    }
  }

  It 'keeps mixed missing and <Case> references blocking' -ForEach @(
    @{ Case = 'present' }, @{ Case = 'denied' }, @{ Case = 'invalid' }
  ) {
    $script:appRecords = @(foreach ($view in @('Registry32', 'Registry64')) {
        [PSCustomObject]@{ View = $view; Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'; Values = [PSCustomObject]@{ '(default)' = "C:\Apps\Office16\$view\WINWORD.EXE" } }
      })
    $script:appCase = $Case
    if ($Case -eq 'invalid') { $script:appRecords[1].Values.'(default)' = 'invalid Office16 target' }
    Mock Get-PSFOfficePathAttribute {
      if ($script:appCase -eq 'denied') { throw [UnauthorizedAccessException]::new('Synthetic denial') }
      [IO.FileAttributes]::Normal
    } -ParameterFilter { $LiteralPath -eq 'C:\Apps\Office16\Registry64\WINWORD.EXE' }
    $observed = Get-OfficeInventory
    $observed.Unknowns | Should -Contain OfficeResidueWithoutConfiguration
    $observed.AppPathEvidence[0].State | Should -Be Missing
  }

  It 'keeps genuine <Key> registry residue blocking' -ForEach @(
    @{ Key = 'ProductReleaseIDs' }, @{ Key = 'Inventory\Office\16.0' }, @{ Key = 'Configuration' }
  ) {
    $script:appRecords = @([PSCustomObject]@{ View = 'Registry64'; Path = "SOFTWARE\Microsoft\Office\ClickToRun\$Key"; Values = [PSCustomObject]@{}; SubKeys = @() })
    (Get-OfficeInventory).Unknowns.Count | Should -BeGreaterThan 0
  }

  It 'distinguishes uncertain custom paths from known legacy MSI paths (<Folder>)' -ForEach @(
    @{ Folder = 'Custom'; Blocking = $true }, @{ Folder = 'Office12'; Blocking = $false }
  ) {
    $script:appRecords = @([PSCustomObject]@{ View = 'Registry64'; Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'; Values = [PSCustomObject]@{ '(default)' = "C:\$Folder\WINWORD.EXE" } })
    Mock Get-PSFOfficePathAttribute { throw [UnauthorizedAccessException]::new('Synthetic denial') }
    $observed = Get-OfficeInventory
    ($observed.Unknowns -contains 'OfficeResidueWithoutConfiguration') | Should -Be $Blocking
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
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language de-de
    $result = Test-OfficeDeploymentMedia -SourcePath $mediaRoot -Configuration $target
    $result.Valid | Should -BeTrue
    $target.Language | Should -Be @('de-de')
  }

  It 'accepts <Layout> metadata for architecture <Architecture>' -ForEach @(
    foreach ($architecture in @('32', '64')) {
      foreach ($layout in @('generic', 'versioned', 'both')) {
        @{ Architecture = $architecture; Layout = $layout }
      }
    }
  ) {
    $manifest.Architecture = $Architecture
    if ($Architecture -eq '32') {
      foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $mediaRoot 'Office\Data\16.0.17932.20162') -File)) {
        Rename-Item -LiteralPath $file.FullName -NewName $file.Name.Replace('x64', 'x86')
      }
      Rename-Item -LiteralPath (Join-Path $mediaRoot 'Office\Data\v64.cab') -NewName 'v32.cab'
    }
    if ($Layout -ne 'generic') {
      [IO.File]::WriteAllText((Join-Path $mediaRoot "Office\Data\v${Architecture}_16.0.17932.20162.cab"), 'Synthetic pinned catalog')
    }
    if ($Layout -eq 'versioned') {
      Remove-Item -LiteralPath (Join-Path $mediaRoot "Office\Data\v$Architecture.cab")
    }
    $manifest.Files = @(Get-PSFOfficeMediaFile $mediaRoot | Sort-Object Path)
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Architecture $Architecture -Language de-de -Version $manifest.Version

    $result = Test-OfficeDeploymentMedia -SourcePath $mediaRoot -Configuration $target
    $result.Valid | Should -BeTrue
    $result.Fingerprint | Should -Not -BeNullOrEmpty
    $result.ReasonCode | Should -BeNullOrEmpty
    $result.Error | Should -BeNullOrEmpty
    $xml = New-PSFOfficeXml -Action Install -Configuration $target -MediaPath $result.Path
    $xml.Configuration.Add.Version | Should -Be $manifest.Version
    $xml.Configuration.Add.SourcePath | Should -Be $mediaRoot
    $xml.Configuration.Add.AllowCdnFallback | Should -Be 'FALSE'
  }

  It 'rejects <Catalog> as a substitute for exact metadata on architecture <Architecture>' -ForEach @(
    foreach ($architecture in @('32', '64')) {
      foreach ($catalog in @('absent', 'wrong-version', 'wrong-architecture', 'arbitrary')) {
        @{ Architecture = $architecture; Catalog = $catalog }
      }
    }
  ) {
    $manifest.Architecture = $Architecture
    Remove-Item -LiteralPath (Join-Path $mediaRoot 'Office\Data\v64.cab')
    if ($Architecture -eq '32') {
      foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $mediaRoot 'Office\Data\16.0.17932.20162') -File)) {
        Rename-Item -LiteralPath $file.FullName -NewName $file.Name.Replace('x64', 'x86')
      }
    }
    $otherArchitecture = '32'
    if ($Architecture -eq '32') { $otherArchitecture = '64' }
    $wrongCab = switch ($Catalog) {
      'wrong-version' { "v${Architecture}_16.0.17932.20161.cab" }
      'wrong-architecture' { "v${otherArchitecture}_16.0.17932.20162.cab" }
      'arbitrary' { 'unrelated.cab' }
    }
    if ($wrongCab) {
      [IO.File]::WriteAllText((Join-Path $mediaRoot "Office\Data\$wrongCab"), 'Synthetic unrelated metadata')
    }
    $manifest.Files = @(Get-PSFOfficeMediaFile $mediaRoot | Sort-Object Path)
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $result = Test-OfficeDeploymentMedia $mediaRoot
    $result.Valid | Should -BeFalse
    $result.ReasonCode | Should -Be MissingMedia
    $result.Error | Should -BeLike "*Office/Data/v$Architecture.cab*"
    $result.Error | Should -BeLike "*Office/Data/v${Architecture}_16.0.17932.20162.cab*"
  }

  It 'reports every missing payload and both CAB alternatives together' {
    foreach ($relative in @('Office/Data/v64.cab', 'Office/Data/16.0.17932.20162/stream.x64.x-none.dat', 'Office/Data/16.0.17932.20162/stream.x64.de-de.dat')) {
      Remove-Item -LiteralPath (Join-Path $mediaRoot $relative)
    }
    $manifest.Files = @(Get-PSFOfficeMediaFile $mediaRoot | Sort-Object Path)
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $result = Test-OfficeDeploymentMedia $mediaRoot
    $result.ReasonCode | Should -Be MissingMedia
    foreach ($relative in @('Office/Data/v64.cab', 'Office/Data/v64_16.0.17932.20162.cab', 'Office/Data/16.0.17932.20162/stream.x64.x-none.dat', 'Office/Data/16.0.17932.20162/stream.x64.de-de.dat')) {
      $result.Error | Should -BeLike "*$relative*"
    }
  }

  It 'distinguishes missing <Locale> streams on architecture <Architecture>' -ForEach @(
    foreach ($architecture in @('32', '64')) {
      @{ Architecture = $architecture; Locale = 'x-none'; Reason = 'MissingMedia' }
      @{ Architecture = $architecture; Locale = 'de-de'; Reason = 'MissingLanguageMedia' }
    }
  ) {
    $manifest.Architecture = $Architecture
    $platform = 'x64'
    if ($Architecture -eq '32') { $platform = 'x86' }
    foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $mediaRoot 'Office\Data\16.0.17932.20162') -File)) {
      if ($Architecture -eq '32') { Rename-Item -LiteralPath $file.FullName -NewName $file.Name.Replace('x64', 'x86') }
    }
    Rename-Item -LiteralPath (Join-Path $mediaRoot 'Office\Data\v64.cab') -NewName "v${Architecture}_16.0.17932.20162.cab"
    $missing = "Office/Data/16.0.17932.20162/stream.$platform.$Locale.dat"
    Remove-Item -LiteralPath (Join-Path $mediaRoot $missing)
    $manifest.Files = @(Get-PSFOfficeMediaFile $mediaRoot | Sort-Object Path)
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $result = Test-OfficeDeploymentMedia $mediaRoot
    $result.ReasonCode | Should -Be $Reason
    $result.Error | Should -BeLike "*$missing*"
  }

  It 'reports undeclared requested languages even if their streams exist' {
    $manifest.AvailableLanguages = @('en-us')
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language de-de
    $result = Test-OfficeDeploymentMedia $mediaRoot -Configuration $target
    $result.ReasonCode | Should -Be MissingLanguageMedia
    $result.Error | Should -BeLike '*not declared in AvailableLanguages: de-de*'
  }

  It 'identifies unavailable requested languages and their missing paths' {
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language fr-fr, it-it
    $result = Test-OfficeDeploymentMedia $mediaRoot -Configuration $target
    $result.ReasonCode | Should -Be MissingLanguageMedia
    $result.Error | Should -BeLike '*not declared in AvailableLanguages: fr-fr, it-it*'
    foreach ($locale in @('fr-fr', 'it-it')) {
      $result.Error | Should -BeLike "*Office/Data/16.0.17932.20162/stream.x64.$locale.dat*"
    }
  }

  It 'retains integrity checks for versioned metadata with <Change>' -ForEach @(
    @{ Change = 'hash'; Reason = 'MediaIntegrityFailed' }
    @{ Change = 'length'; Reason = 'MediaIntegrityFailed' }
    @{ Change = 'extra'; Reason = 'MediaIntegrityFailed' }
    @{ Change = 'empty'; Reason = 'InvalidMedia' }
    @{ Change = 'collision'; Reason = 'UnsafeManifest' }
  ) {
    $cab = Join-Path $mediaRoot 'Office\Data\v64_16.0.17932.20162.cab'
    Rename-Item -LiteralPath (Join-Path $mediaRoot 'Office\Data\v64.cab') -NewName 'v64_16.0.17932.20162.cab'
    $manifest.Files = @(Get-PSFOfficeMediaFile $mediaRoot | Sort-Object Path)
    switch ($Change) {
      'hash' { $manifest.Files[0].Hash = 'a' * 64 }
      'length' { $manifest.Files[0].Length++ }
      'extra' { [IO.File]::WriteAllText((Join-Path $mediaRoot 'Office\Data\extra.cab'), 'unexpected') }
      'empty' { [IO.File]::WriteAllText($cab, '') }
      'collision' {
        $manifest.Files += [PSCustomObject]@{ Path = $manifest.Files[0].Path.ToUpperInvariant(); Length = 1; Hash = 'a' * 64 }
      }
    }
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    (Test-OfficeDeploymentMedia $mediaRoot).ReasonCode | Should -Be $Reason
  }

  It 'retains target matching for <Dimension>' -ForEach @(
    @{ Dimension = 'Product'; Arguments = @{ TargetProductId = 'ProPlus2024Volume' } }
    @{ Dimension = 'Architecture'; Arguments = @{ TargetProductId = 'Standard2024Volume'; Architecture = '32' } }
    @{ Dimension = 'Channel'; Arguments = @{ TargetProductId = 'O365ProPlusRetail'; Channel = 'MonthlyEnterprise' } }
    @{ Dimension = 'Version'; Arguments = @{ TargetProductId = 'Standard2024Volume'; Version = '16.0.17932.20161' } }
  ) {
    if ($Dimension -eq 'Channel') {
      $manifest.Product = 'O365ProPlusRetail'
      $manifest.Channel = 'Current'
      $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    }
    $different = New-OfficeDeploymentConfiguration @Arguments
    (Test-OfficeDeploymentMedia $mediaRoot -Configuration $different).ReasonCode | Should -Be MediaMismatch
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

Describe 'Office media path protection' {
  It 'rejects <Case> before reading a manifest' -ForEach @(
    @{ Case = 'untrusted owner'; Sddl = 'O:S-1-1-0G:S-1-5-32-544D:P(A;;FA;;;S-1-5-32-544)'; Reason = 'UntrustedMedia' }
    @{ Case = 'untrusted write permission'; Sddl = 'O:S-1-5-32-544G:S-1-5-32-544D:P(A;;FA;;;S-1-1-0)'; Reason = 'UntrustedMedia' }
    @{ Case = 'reparse traversal'; Sddl = ''; Reason = 'UnsafePath' }
  ) {
    $root = Join-Path $TestDrive 'protected-media'
    $null = New-Item -ItemType Directory -Path $root -Force
    $script:mediaAcl = New-Object Security.AccessControl.DirectorySecurity
    if ($Sddl) { $script:mediaAcl.SetSecurityDescriptorSddlForm($Sddl) }
    Mock Get-Acl { $script:mediaAcl }
    Mock Get-Content { throw 'Must reject before reading a manifest' }
    if ($Case -eq 'reparse traversal') {
      Mock Get-Item { [PSCustomObject]@{ Attributes = [IO.FileAttributes]::ReparsePoint } } -ParameterFilter { $LiteralPath -eq $root }
    }
    $result = Test-OfficeDeploymentMedia -SourcePath $root
    $result.ReasonCode | Should -Be $Reason
    Should -Invoke Get-Content -Times 0
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

  It 'does not turn an unknown installed build into verified migration success (<Code>)' -ForEach @(
    @{ Code = 0 }, @{ Code = 3010 }
  ) {
    $script:nativeExit = $Code
    Mock Invoke-PSFOfficeConfiguration {
      $script:inventory = New-TestOfficeInventory $script:target
      $script:inventory.Products[0].Version = $null
      [PSCustomObject]@{ ExitCode = $script:nativeExit }
    }
    $migration = Get-OfficeDeploymentPlan -Action Migrate -Configuration $target -SourcePath (Join-Path $TestDrive 'Media')
    $result = Switch-OfficeDeployment -Plan $migration -OdtPath C:\ODT\setup.exe -LogRoot $logRoot -Confirm:$false
    $result.Status | Should -Be Failed
    $result.ReasonCode | Should -Be VerificationFailed
    $result.WrapperExitCode | Should -Be 1
    $result.ExitCode | Should -Be $Code
    $result.RebootRequired | Should -Be ($Code -eq 3010)
    $result.Verification.Unknowns | Should -Contain Version
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1
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

Describe 'Office pilot retirement' {
  BeforeEach {
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' }
    $script:registryFixture = Get-Content "$PSScriptRoot/fixtures/office/office2007-registry.json" -Raw | ConvertFrom-Json
    Mock Get-PSFOfficeRegistrySnapshot { $script:registryFixture }
    $script:inventory = Get-OfficeInventory
    Mock Get-OfficeInventory { $script:inventory }
    $script:target = New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 64 -Language de-de
    Mock Test-OfficeDeploymentMedia {
      [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.10417.20208' } }
    }
    $script:planArgs = @{
      Action        = 'Migrate'
      Configuration = $target
      SourcePath    = 'C:\Media\Office2019'
      RemoveMsi     = $true
    }
    Mock Invoke-PSFOfficeConfiguration { throw 'Must not execute an installer' }
  }

  It 'uses ordinary migration for the former MSI source without scenario-specific resource claims' {
    $plan = Get-OfficeDeploymentPlan @planArgs
    $plan.Eligible | Should -BeTrue
    $plan.SchemaVersion | Should -Be 1
    $plan.LanguageTransition.Known | Should -BeFalse
    $plan.PSObject.Properties.Name | Should -Not -Contain PilotResources
    $plan.PSObject.Properties.Name | Should -Not -Contain PilotMigration
    $plan.Configuration.Version | Should -Be '16.0.10417.20208'
    $xml = New-PSFOfficeXml -Action Migrate -Configuration $plan.Configuration -RemoveMsi $true -MediaPath $planArgs.SourcePath
    $xml.Configuration.Add.Product.Language.ID | Should -Be 'de-de'
    $xml.OuterXml | Should -Not -Match 'MatchPreviousMSI'
  }

  It 'retains ordinary blocker <Expected>' -ForEach @(
    @{ Change = 'Consent'; Expected = 'MsiConsentRequired' }
    @{ Change = 'Unknown'; Expected = 'UnknownInventory' }
    @{ Change = 'Backend'; Expected = 'UnsupportedNativeVerification' }
    @{ Change = 'Media'; Expected = 'MissingMedia' }
  ) {
    switch ($Change) {
      Consent { $planArgs.RemoveMsi = $false }
      Unknown { $inventory.Unknowns = @('SyntheticUnknown') }
      Backend { $inventory.VerificationLimitations += 'Languages' }
      Media { $planArgs.Remove('SourcePath') }
    }
    $plan = Get-OfficeDeploymentPlan @planArgs
    $plan.Eligible | Should -BeFalse
    $plan.Blockers | Should -Contain $Expected
  }

  It 'rejects the retired parameter and schema-2 execution plans before any installer launch' {
    { Get-OfficeDeploymentPlan @planArgs -PilotMigration } | Should -Throw '*parameter*PilotMigration*'
    $plan = Get-OfficeDeploymentPlan @planArgs
    { Switch-OfficeDeployment -Plan $plan -OdtPath C:\ODT\setup.exe -PilotMigration } | Should -Throw '*parameter*PilotMigration*'
    $plan.SchemaVersion = 2
    { Switch-OfficeDeployment -Plan $plan -OdtPath C:\ODT\setup.exe -Confirm:$false } | Should -Throw '*schema*'
    $plan.SchemaVersion = 1
    $plan | Add-Member -NotePropertyName PilotMigration -NotePropertyValue $true
    { Switch-OfficeDeployment -Plan $plan -OdtPath C:\ODT\setup.exe -Confirm:$false } | Should -Throw
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
  }

  It 'reads historical schema-2 evidence but refuses recovery before any replay' {
    $plan = Get-OfficeDeploymentPlan @planArgs
    $plan.SchemaVersion = 2
    $plan | Add-Member -NotePropertyName PilotMigration -NotePropertyValue $true
    $plan | Add-Member -NotePropertyName PilotResources -NotePropertyValue ([PSCustomObject]@{
        UiLanguages       = @('de-de')
        PrimaryLanguage   = 'de-de'
        ProofingLanguages = @('de-de', 'en-us', 'fr-fr', 'it-it')
        Provisioning      = 'GermanCompanionProofing'
        Verification      = 'ManualRequired'
      })
    $record = [ordered]@{
      SchemaVersion            = 2
      RunId                    = 'aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
      MachineId                = 'synthetic-machine'
      Action                   = 'Migrate'
      Plan                     = $plan
      ConfigurationFingerprint = Get-PSFOfficeFingerprint $plan.Configuration
      MediaFingerprint         = $plan.MediaFingerprint
      Phase                    = 'Verify'
      PhaseCompleted           = $false
      NativeResults            = @([PSCustomObject]@{ ExitCode = 0 })
      RebootRequired           = $false
      CreatedAt                = '2026-01-01T00:00:00Z'
      UpdatedAt                = '2026-01-01T00:00:00Z'
      Result                   = $null
    }
    $root = Join-Path $TestDrive 'Historical'
    $null = New-Item -Path $root -ItemType Directory
    Mock Assert-PSFOfficeProtectedPath { }
    Write-PSFOfficeJson -Path (Join-Path $root ($record.RunId + '.json')) -Value $record
    $recovery = Get-OfficeDeploymentRecovery -RunId $record.RunId -LogRoot $root
    $recovery.Record.SchemaVersion | Should -Be 2
    { Resume-OfficeMigration -Recovery $recovery -OdtPath C:\ODT\setup.exe } | Should -Throw '*evidence only*'
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
  }
}

Describe 'Office ordinary host assessment' {
  It 'accepts x64 desktop Windows 10 22H2 or later (<Build>/<Type>/<Cpu>)' -ForEach @(
    @{ Build = '19045'; Type = 1; Cpu = 9; Expected = $true }
    @{ Build = '22000'; Type = 1; Cpu = 9; Expected = $true }
    @{ Build = '26100'; Type = 1; Cpu = 9; Expected = $true }
    @{ Build = '19044'; Type = 1; Cpu = 9; Expected = $false }
    @{ Build = '17763'; Type = 1; Cpu = 9; Expected = $false }
    @{ Build = '19045'; Type = 1; Cpu = 0; Expected = $false }
    @{ Build = '26100'; Type = 1; Cpu = 12; Expected = $false }
    @{ Build = '26100'; Type = 3; Cpu = 9; Expected = $false }
    @{ Build = '26100'; Type = 2; Cpu = 9; Expected = $false }
  ) {
    $script:hostBuild = $Build
    $script:hostType = $Type
    $script:hostCpu = $Cpu
    Mock Get-CimInstance { [PSCustomObject]@{ ProductType = $script:hostType; BuildNumber = $script:hostBuild } } -ParameterFilter { $ClassName -eq 'Win32_OperatingSystem' }
    Mock Get-CimInstance { [PSCustomObject]@{ Architecture = $script:hostCpu } } -ParameterFilter { $ClassName -eq 'Win32_Processor' }
    Test-PSFOfficeHost | Should -Be $Expected
  }

  It 'propagates host discovery failure instead of assuming eligibility' {
    Mock Get-CimInstance { throw 'Synthetic host lookup failure' }
    { Test-PSFOfficeHost } | Should -Throw '*Synthetic host lookup failure*'
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

Describe 'Office pre-fix migration journal continuation' {
  BeforeEach {
    Mock Get-PSFOfficeMachineId { 'synthetic-machine' }
    Mock Assert-PSFOfficeProtectedPath { }
    Mock Assert-PSFOfficeHost { }
    Mock Get-PSFOfficeActivity { [PSCustomObject]@{ Busy = $false; Apps = @() } }
    Mock Test-PendingReboot { [PSCustomObject]@{ PendingReboot = $false } }
    Mock Test-OfficeDeploymentTool { [PSCustomObject]@{ Valid = $true; Version = '16.0.20326.20112' } }
    Mock Test-OfficeDeploymentMedia {
      [PSCustomObject]@{ Valid = $true; Fingerprint = 'synthetic-media'; Manifest = [PSCustomObject]@{ Version = '16.0.10417.20211'; Files = @([PSCustomObject]@{ Length = 1 }) } }
    }
    Mock Get-OfficeActivationStatus { [PSCustomObject]@{ Status = 'Licensed' } }
    Mock Get-PSFOfficePathAttribute {
      if ($LiteralPath -like '*.exe') { throw [IO.FileNotFoundException]::new('Synthetic removed executable') }
      [IO.FileAttributes]::Directory
    }
    $script:staleRecords = @(foreach ($view in @('Registry32', 'Registry64')) {
        foreach ($name in @('WINWORD.EXE', 'EXCEL.EXE', 'OUTLOOK.EXE')) {
          [PSCustomObject]@{ View = $view; Path = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$name"; Values = [PSCustomObject]@{ '(default)' = "C:\Program Files (x86)\Microsoft Office\Root\Office16\$name" }; SubKeys = @() }
        }
      })
    $script:recoveryRecords = $script:staleRecords
    Mock Get-PSFOfficeRegistrySnapshot { $script:recoveryRecords }
    $script:installedRecords = Get-Content "$PSScriptRoot/fixtures/office/office2019-registry.json" -Raw | ConvertFrom-Json
    foreach ($record in $script:installedRecords) {
      if ($record.Path -like '*ClickToRun\Configuration') {
        $record.Values.Platform = 'x64'
        $record.Values.'Standard2019Volume.ExcludedApps' = 'Groove,OneDrive,OneNote,Publisher'
      }
      if ($record.Values.PSObject.Properties['Version']) { $record.Values.Version = '16.0.10417.20211' }
    }
    $script:recoveryTarget = New-OfficeDeploymentConfiguration -TargetProductId Standard2019Volume -Architecture 64 -Language de-de -Version 16.0.10417.20211 -ExcludeApp Groove, OneDrive, OneNote, Publisher
    # Historical inventory layout: deliberately no AppPathEvidence field. Do not
    # obtain this snapshot from the new reader or inject new defaults before hashing.
    $script:historicalBefore = [PSCustomObject][ordered]@{
      SchemaVersion = 1
      MachineId     = 'synthetic-machine'
      Products      = @([PSCustomObject][ordered]@{
          ProductId = 'Standard2019Volume'; Architecture = '32'; Version = '16.0.10417.20208'
          VersionSource = 'ActiveProductResources'; Channel = 'PerpetualVL2019'
          Languages = @('de-de'); PrimaryLanguage = 'de-de'; RegisteredLanguages = @('de-de')
          ExcludeApp = @('Groove'); Evidence = @('Synthetic pre-fix observation')
        })
      Msi = @(); RelatedComponents = @(); Unknowns = @(); RegisteredResources = @()
      LanguageEvidence = @(); VerificationLimitations = @()
    }
    $originalPlan = Get-OfficeDeploymentPlan -Action Migrate -Configuration $recoveryTarget -Inventory $historicalBefore -SourcePath (Join-Path $TestDrive 'media') -RemoveProductId Standard2019Volume
    $script:originalRecord = [PSCustomObject][ordered]@{
      SchemaVersion = 1; RunId = 'bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'; MachineId = 'synthetic-machine'; Action = 'Migrate'
      Plan = $originalPlan; ConfigurationFingerprint = Get-PSFOfficeFingerprint $recoveryTarget; MediaFingerprint = 'synthetic-media'
      Phase = 'Remove'; PhaseCompleted = $true; NativeResults = @([PSCustomObject]@{ Phase = 'Remove'; ExitCode = 0 })
      RebootRequired = $false; CreatedAt = '2026-01-01T00:00:00Z'; UpdatedAt = '2026-01-01T00:01:00Z'
      Result = [PSCustomObject]@{ Status = 'Failed'; ReasonCode = 'VerificationFailed'; After = [PSCustomObject]@{ Unknowns = @('OfficeResidueWithoutConfiguration') } }
    }
    $script:recoveryRoot = Join-Path $TestDrive 'recovery'
    $null = New-Item -ItemType Directory -Path $recoveryRoot -Force
    $script:originalPath = Join-Path $recoveryRoot ($originalRecord.RunId + '.json')
    Write-PSFOfficeJson -Path $originalPath -Value $originalRecord
    $script:originalHash = (Get-FileHash -LiteralPath $originalPath).Hash
    $script:descriptor = Get-OfficeDeploymentRecovery -RunId $originalRecord.RunId -LogRoot $recoveryRoot
    Mock New-PSFOfficeProtectedDirectory { $null = [IO.Directory]::CreateDirectory($Path) }
    Mock Copy-Item { }
    Mock Write-OperationResultLog { 'synthetic-log' }
    Mock Enter-PSFOfficeLock {
      $lock = [PSCustomObject]@{}
      $lock | Add-Member ScriptMethod ReleaseMutex { }
      $lock | Add-Member ScriptMethod Dispose { }
      $lock
    }
    Mock Invoke-PSFOfficeConfiguration {
      if ($Document.Configuration.Remove) { $script:recoveryRecords = $script:staleRecords }
      else { $script:recoveryRecords = $script:installedRecords }
      [PSCustomObject]@{ ExitCode = 0 }
    }
  }

  It 'loads the historical fingerprint unchanged and previews without launching ODT' {
    $descriptor.Record.Plan.Before.PSObject.Properties.Name | Should -Not -Contain AppPathEvidence
    (Get-PSFOfficeFingerprint $descriptor.Record.Plan.Before) | Should -Be $originalRecord.Plan.InventoryFingerprint
    (Get-OfficeInventory).AppPathEvidence.Count | Should -Be 6
    $preview = Resume-OfficeMigration -Recovery $descriptor -OdtPath C:\ODT\setup.exe -DryRun
    $preview.Status | Should -Be Preview
    $preview.Configuration.Version | Should -Be $recoveryTarget.Version
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
    Should -Invoke New-PSFOfficeProtectedDirectory -Times 0
    (Get-FileHash -LiteralPath $originalPath).Hash | Should -Be $originalHash
  }

  It 'installs once without repeating removal and retains both journals' {
    $key = New-Object Security.SecureString
    foreach ($character in 'AAAAA-BBBBB-CCCCC-DDDDD-EEEEE'.ToCharArray()) { $key.AppendChar($character) }
    try {
      $result = Resume-OfficeMigration -Recovery $descriptor -OdtPath C:\ODT\setup.exe -ProductKey $key -Confirm:$false
      $result.Status | Should -Be Completed
      $result.Activation.Status | Should -Be Licensed
      $result.RunId | Should -Not -Be $originalRecord.RunId
      $newRecord = (Get-OfficeDeploymentRecovery -RunId $result.RunId -LogRoot $recoveryRoot).Record
      $newRecord.Plan.RemoveProductId.Count | Should -Be 0
      $newRecord.Plan.RemoveMsi | Should -BeFalse
      $newRecord.Plan.Configuration.ExcludeApp | Should -Be $recoveryTarget.ExcludeApp
      $newRecord.Plan.Configuration.Language | Should -Be @('de-de')
      (Get-Content -LiteralPath $result.RecoveryPath -Raw) | Should -Not -Match 'AAAAA'
      Should -Invoke Invoke-PSFOfficeConfiguration -Times 1 -Exactly
      Should -Invoke Invoke-PSFOfficeConfiguration -Times 0 -ParameterFilter { $Document.Configuration.Remove }
      Should -Invoke Invoke-PSFOfficeConfiguration -Times 1 -ParameterFilter { $null -ne $ProductKey -and $Document.Configuration.Add.Version -eq '16.0.10417.20211' }
      (Get-FileHash -LiteralPath $originalPath).Hash | Should -Be $originalHash
    }
    finally { $key.Dispose() }
  }

  It 'also completes a fresh migration when removal leaves only stale App Paths' {
    # Start with real classification of synthetic x86 registry evidence.
    $script:recoveryRecords = Get-Content "$PSScriptRoot/fixtures/office/office2019-registry.json" -Raw | ConvertFrom-Json
    $plan = Get-OfficeDeploymentPlan -Action Migrate -Configuration $recoveryTarget -SourcePath (Join-Path $TestDrive 'media') -RemoveProductId Standard2019Volume
    $result = Switch-OfficeDeployment -Plan $plan -OdtPath C:\ODT\setup.exe -LogRoot $recoveryRoot -Confirm:$false
    $result.Status | Should -Be Completed
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 2 -Exactly
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1 -ParameterFilter { $Document.Configuration.Remove }
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1 -ParameterFilter { $Document.Configuration.Add }
  }

  It 'keeps <Guard> blocking recovery without a native launch' -ForEach @(
    @{ Guard = 'busy'; Reason = 'DeploymentBusy' }
    @{ Guard = 'reboot'; Reason = 'RebootRequired' }
    @{ Guard = 'uncertain target'; Reason = 'Conflict' }
    @{ Guard = 'unsupported checkpoint'; Reason = 'UnsupportedRecoveryState' }
    @{ Guard = 'new product'; Reason = 'Conflict' }
  ) {
    switch ($Guard) {
      busy { Mock Get-PSFOfficeActivity { [PSCustomObject]@{ Busy = $true; Apps = @() } } }
      reboot { Mock Test-PendingReboot { [PSCustomObject]@{ PendingReboot = $true } } }
      'uncertain target' { Mock Get-PSFOfficePathAttribute { throw [UnauthorizedAccessException]::new('Synthetic denial') } }
      'unsupported checkpoint' {
        $originalRecord.Phase = 'Migrate'
        $originalRecord.PhaseCompleted = $false
        Write-PSFOfficeJson -Path $originalPath -Value $originalRecord
      }
      'new product' {
        $unexpected = New-TestOfficeInventory (New-OfficeDeploymentConfiguration -TargetProductId O365ProPlusRetail)
        Mock Get-OfficeInventory { $unexpected }
      }
    }
    (Resume-OfficeMigration -Recovery $descriptor -OdtPath C:\ODT\setup.exe -Confirm:$false).ReasonCode | Should -Be $Reason
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
  }

  It 'rejects <Guard> before continuing a historical journal' -ForEach @(
    @{ Guard = 'wrong machine' }, @{ Guard = 'changed media' }, @{ Guard = 'altered historical inventory' }
  ) {
    switch ($Guard) {
      'wrong machine' { Mock Get-PSFOfficeMachineId { 'different-machine' } }
      'changed media' { Mock Test-OfficeDeploymentMedia { [PSCustomObject]@{ Valid = $true; Fingerprint = 'different' } } }
      'altered historical inventory' {
        $originalRecord.Plan.Before | Add-Member NoteProperty AppPathEvidence @()
        Write-PSFOfficeJson -Path $originalPath -Value $originalRecord
      }
    }
    { Resume-OfficeMigration -Recovery $descriptor -OdtPath C:\ODT\setup.exe -Confirm:$false } | Should -Throw
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 0
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

  It 'names the expected product key shape without echoing the key <Case>' -ForEach @(
    @{ Case = 'empty from a dropped paste'; Value = '' }
    @{ Case = 'en-dash separators from a formatted document'; Value = "AAAAA`u{2013}BBBBB`u{2013}CCCCC`u{2013}DDDDD`u{2013}EEEEE" }
    @{ Case = 'no separators at all'; Value = 'AAAAABBBBBCCCCCDDDDDEEEEE' }
  ) {
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Version 16.0.17932.20162
    $document = New-PSFOfficeXml -Action Install -Configuration $target -MediaPath C:\Media
    $key = New-Object Security.SecureString
    foreach ($character in $Value.ToCharArray()) {
      $key.AppendChar($character)
    }

    try {
      $message = $null
      try {
        Invoke-PSFOfficeConfiguration -OdtPath C:\ODT\setup.exe -Document $document -Directory $TestDrive -ProductKey $key
      }
      catch {
        $message = $_.Exception.Message
      }

      $message | Should -BeLike '*five groups of five letters or digits separated by ASCII hyphens*'
      # The rejection must describe the shape, never the supplied value.
      $message | Should -Not -BeLike '*AAAAA*'
      $document.SelectNodes('//*[@PIDKEY]').Count | Should -Be 0
    }
    finally {
      $key.Dispose()
    }
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
    $script:prepareCatalog = 'v64.cab'
    $script:omitPrepare = @()
    Mock Invoke-PSFOfficeConfiguration {
      $data = Join-Path $Directory 'Office\Data\16.0.17932.20162'
      $null = [IO.Directory]::CreateDirectory($data)
      if ($script:prepareCatalog) {
        [IO.File]::WriteAllText((Join-Path (Split-Path $data -Parent) $script:prepareCatalog), 'Synthetic base catalog')
      }
      $platform = 'x64'
      if ($Document.Configuration.Add.OfficeClientEdition -eq '32') { $platform = 'x86' }
      foreach ($language in @('x-none', 'en-us', 'de-de')) {
        if ($language -notin $script:omitPrepare) {
          [IO.File]::WriteAllText((Join-Path $data "stream.$platform.$language.dat"), "Synthetic $language payload")
        }
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

  It 'publishes and reuses exact-version-only metadata for architecture <Architecture>' -ForEach @(
    @{ Architecture = '32' }
    @{ Architecture = '64' }
  ) {
    $script:prepareCatalog = "v${Architecture}_16.0.17932.20162.cab"
    $target = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Architecture $Architecture -Language en-us, de-de -Version 16.0.17932.20162
    $media = Save-OfficeDeploymentMedia -Configuration $target -SourcePath $destination -OdtPath $setupPath -Confirm:$false
    $media.Valid | Should -BeTrue
    $media.Manifest.Version | Should -Be $target.Version
    $media.Manifest.Files.Path | Should -Contain "Office/Data/$script:prepareCatalog"
    Test-Path -LiteralPath (Join-Path $destination "Office\Data\v$Architecture.cab") | Should -BeFalse
    Test-Path -LiteralPath (Join-Path $destination 'psfoundation-office-media.json') | Should -BeTrue
    Test-Path -LiteralPath (Join-Path $destination 'setup.exe') | Should -BeFalse
    $german = New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Architecture $Architecture -Language de-de -Version $target.Version
    $reuse = Save-OfficeDeploymentMedia -Configuration $german -SourcePath $destination -OdtPath $setupPath -Confirm:$false
    $reuse.Valid | Should -BeTrue
    $reuse.Fingerprint | Should -Be $media.Fingerprint
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1 -Exactly
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1 -Exactly -ParameterFilter {
      $Mode -eq '/download' -and $Document.Configuration.Add.Version -eq '16.0.17932.20162' -and
      $Document.Configuration.Add.AllowCdnFallback -eq 'FALSE' -and $Document.Configuration.Add.SourcePath -eq $Directory
    }
    @(Get-ChildItem -LiteralPath $TestDrive -Directory -Filter 'PSFOfficePrepare-*').Count | Should -Be 0
  }

  It 'keeps complete <Case> diagnostics after failed preparation cleanup' -ForEach @(
    @{ Case = 'base and language'; Catalog = ''; Omit = @('x-none', 'de-de'); Reason = 'MissingMedia' }
    @{ Case = 'language only'; Catalog = 'v64_16.0.17932.20162.cab'; Omit = @('de-de'); Reason = 'MissingLanguageMedia' }
  ) {
    $script:prepareCatalog = $Catalog
    $script:omitPrepare = $Omit
    $failure = $null
    try {
      $null = Save-OfficeDeploymentMedia -Configuration $target -SourcePath $destination -OdtPath $setupPath -Confirm:$false
    }
    catch { $failure = $_ }
    $failure | Should -Not -BeNullOrEmpty
    $failure.Exception.Data['OfficeReason'] | Should -Be $Reason
    foreach ($locale in $Omit) {
      $failure.Exception.Message | Should -BeLike "*Office/Data/16.0.17932.20162/stream.x64.$locale.dat*"
    }
    if (-not $Catalog) {
      $failure.Exception.Message | Should -BeLike '*Office/Data/v64.cab*'
      $failure.Exception.Message | Should -BeLike '*Office/Data/v64_16.0.17932.20162.cab*'
    }
    Test-Path -LiteralPath $destination | Should -BeFalse
    @(Get-ChildItem -LiteralPath $TestDrive -Directory -Filter 'PSFOfficePrepare-*').Count | Should -Be 0
    Should -Invoke Invoke-PSFOfficeConfiguration -Times 1 -Exactly
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

Describe 'Office Deployment Tool acquisition' {
  BeforeEach {
    $script:destination = Join-Path $TestDrive ('ODT tools ' + [guid]::NewGuid().ToString('N'))
    Mock Assert-PSFOfficePath { }
    Mock Assert-PSFOfficeProtectedPath { }
    Mock New-PSFOfficeProtectedDirectory { $null = New-Item -Path $Path -ItemType Directory -Force }
    Mock Invoke-WebRequest { [IO.File]::WriteAllText($OutFile, 'Synthetic extractor') }
    Mock Get-AuthenticodeSignature {
      [PSCustomObject]@{ Status = 'Valid'; SignerCertificate = [PSCustomObject]@{ Subject = 'CN=Microsoft Corporation, O=Microsoft Corporation, C=US' } }
    }
    Mock Invoke-SafeProcess {
      $extract = $ArgumentList[1].Substring('/extract:'.Length)
      [IO.File]::WriteAllText((Join-Path $extract 'setup.exe'), 'Synthetic ODT')
      [PSCustomObject]@{ ExitCode = 0; TimedOut = $false; Cancelled = $false }
    }
    Mock Test-OfficeDeploymentTool {
      [PSCustomObject]@{ Valid = $true; Path = $OdtPath; Version = '16.0.20326.20112' }
    }
  }

  It 'publishes verified setup and reuses it without downloading again' {
    $first = Install-OfficeDeploymentTool -Destination $destination -Confirm:$false
    $first.Valid | Should -BeTrue
    $first.Path | Should -Be (Join-Path $destination 'setup.exe')
    Test-Path -LiteralPath $first.Path | Should -BeTrue
    $second = Install-OfficeDeploymentTool -Destination $destination -Confirm:$false
    $second.Path | Should -Be $first.Path
    Should -Invoke Invoke-WebRequest -Times 1
    Should -Invoke Get-AuthenticodeSignature -Times 1
    Should -Invoke Invoke-SafeProcess -Times 1
    @(Get-ChildItem -LiteralPath $TestDrive -Directory -Filter 'PSFOfficeTool-*').Count | Should -Be 0
  }

  It 'rejects failed extraction even when a partial setup exists (<Failure>)' -ForEach @(
    @{ Failure = 'Exit'; Code = 1603; Timeout = $false; Cancel = $false }
    @{ Failure = 'Timeout'; Code = 0; Timeout = $true; Cancel = $false }
    @{ Failure = 'Cancel'; Code = 0; Timeout = $false; Cancel = $true }
    @{ Failure = 'NoExit'; Code = $null; Timeout = $false; Cancel = $false }
  ) {
    $script:extractResult = [PSCustomObject]@{ ExitCode = $Code; TimedOut = $Timeout; Cancelled = $Cancel }
    Mock Invoke-SafeProcess {
      $extract = $ArgumentList[1].Substring('/extract:'.Length)
      [IO.File]::WriteAllText((Join-Path $extract 'setup.exe'), 'Partial output')
      $script:extractResult
    }
    { Install-OfficeDeploymentTool -Destination $destination -Confirm:$false } | Should -Throw '*extraction*'
    Test-Path -LiteralPath $destination | Should -BeFalse
    Should -Invoke Test-OfficeDeploymentTool -Times 0
    @(Get-ChildItem -LiteralPath $TestDrive -Directory -Filter 'PSFOfficeTool-*').Count | Should -Be 0
  }

  It 'does not publish an untrusted extracted tool' {
    Mock Test-OfficeDeploymentTool { [PSCustomObject]@{ Valid = $false } }
    { Install-OfficeDeploymentTool -Destination $destination -Confirm:$false } | Should -Throw '*Extracted ODT*'
    Test-Path -LiteralPath $destination | Should -BeFalse
  }

  It 'leaves an untrusted existing tool untouched' {
    $null = New-Item -Path $destination -ItemType Directory
    $setup = Join-Path $destination 'setup.exe'
    [IO.File]::WriteAllText($setup, 'Existing file')
    Mock Test-OfficeDeploymentTool { [PSCustomObject]@{ Valid = $false } }
    { Install-OfficeDeploymentTool -Destination $destination -Confirm:$false } | Should -Throw '*Existing ODT*'
    [IO.File]::ReadAllText($setup) | Should -Be 'Existing file'
    Should -Invoke Invoke-WebRequest -Times 0
  }

  It 'keeps DryRun free of downloads and executable launches' {
    (Install-OfficeDeploymentTool -Destination $destination -DryRun).Status | Should -Be Preview
    Should -Invoke Invoke-WebRequest -Times 0
    Should -Invoke Invoke-SafeProcess -Times 0
    Should -Invoke New-PSFOfficeProtectedDirectory -Times 0
  }

  It 'probes source availability using HEAD without executing or downloading content' {
    Mock Invoke-WebRequest { [PSCustomObject]@{ StatusCode = 200; Headers = @{ 'Content-Length' = '1234' } } }
    $result = Test-OfficeDeploymentToolSourceAvailability
    $result.Available | Should -BeTrue
    $result.Uri | Should -Be (Resolve-OfficeDeploymentToolSource).Uri
    $result.ContentLength | Should -Be '1234'
    Should -Invoke Invoke-WebRequest -Times 1 -ParameterFilter { $Method -eq 'Head' -and -not $OutFile }
    Should -Invoke Invoke-SafeProcess -Times 0
  }

  It 'reports source failure without claiming availability' {
    Mock Invoke-WebRequest { throw 'Synthetic source unavailable' }
    $result = Test-OfficeDeploymentToolSourceAvailability
    $result.Available | Should -BeFalse
    $result.Error | Should -Match 'Synthetic source unavailable'
    Should -Invoke Invoke-SafeProcess -Times 0
  }
}

Describe 'Office Deployment Tool executable identity' {
  BeforeEach {
    $script:odtMetadata = [PSCustomObject]@{
      OriginalFilename = 'Bootstrapper.exe'
      FileDescription  = 'Microsoft 365 and Office'
      ProductName      = 'Microsoft Office'
      CompanyName      = 'Microsoft Corporation'
      FileVersion      = 'Localized version text is not used'
      FileMajorPart    = 16
      FileMinorPart    = 0
      FileBuildPart    = 20326
      FilePrivatePart  = 20112
    }
    $script:odtSignature = [PSCustomObject]@{
      Status            = 'Valid'
      SignerCertificate = [PSCustomObject]@{ Subject = 'CN=Microsoft Corporation, O=Microsoft Corporation, C=US' }
    }
    Mock Assert-PSFOfficePath { }
    Mock Get-Item {
      [PSCustomObject]@{ PSIsContainer = $false; Extension = '.exe'; VersionInfo = $script:odtMetadata }
    }
    Mock Get-AuthenticodeSignature { $script:odtSignature }
    Mock Invoke-SafeProcess { throw 'Validation must not execute the file' }
    Mock Invoke-WebRequest { throw 'Validation must not download' }
  }

  It 'accepts the current Microsoft ODT metadata without requiring its on-disk filename internally' {
    $result = Test-OfficeDeploymentTool -OdtPath 'C:\ODT\setup.exe'
    $result.Valid | Should -BeTrue
    $result.Version | Should -Be '16.0.20326.20112'
    $result.OriginalFilename | Should -Be 'Bootstrapper.exe'
    $result.SignatureStatus | Should -Be Valid
    $result.ReasonCode | Should -BeNullOrEmpty
    $result.Detail | Should -BeNullOrEmpty
    Should -Invoke Invoke-SafeProcess -Times 0
    Should -Invoke Invoke-WebRequest -Times 0
  }

  It 'retains the established identity and minimum supported build' {
    $script:odtMetadata.OriginalFilename = 'setup.exe'
    $script:odtMetadata.FileDescription = 'Microsoft Office Deployment Tool'
    $script:odtMetadata.FileBuildPart = 12827
    $script:odtMetadata.FilePrivatePart = 20258
    (Test-OfficeDeploymentTool -OdtPath 'C:\ODT\setup.exe').Valid | Should -BeTrue
  }

  It 'rejects an invalid signature even when all metadata matches' -ForEach @(
    @{ SignatureStatus = 'NotSigned' }
    @{ SignatureStatus = 'HashMismatch' }
    @{ SignatureStatus = 'NotTrusted' }
    @{ SignatureStatus = 'UnknownError' }
  ) {
    $script:odtSignature.Status = $SignatureStatus
    $result = Test-OfficeDeploymentTool -OdtPath 'C:\ODT\setup.exe'
    $result.Valid | Should -BeFalse
    $result.ReasonCode | Should -Be UntrustedTool
    $result.Detail | Should -Match 'signature is not valid'
  }

  It 'rejects another publisher or missing signer' -ForEach @(
    @{ Subject = 'CN=Microsoft Corporation, O=Other Corporation, C=US' }
    @{ Subject = 'CN=Other, O=Microsoft Corporation Impostor, C=US' }
    @{ Subject = $null }
  ) {
    $script:odtSignature.SignerCertificate = if ($Subject) { [PSCustomObject]@{ Subject = $Subject } } else { $null }
    $result = Test-OfficeDeploymentTool -OdtPath 'C:\ODT\setup.exe'
    $result.Valid | Should -BeFalse
    $result.Detail | Should -Match 'signer'
  }

  It 'rejects a mismatched bootstrapper identity field <Field>' -ForEach @(
    @{ Field = 'OriginalFilename'; Value = 'officedeploymenttool.exe' }
    @{ Field = 'FileDescription'; Value = 'Other Office bootstrapper' }
    @{ Field = 'ProductName'; Value = 'Other product' }
    @{ Field = 'CompanyName'; Value = 'Other Corporation' }
  ) {
    $script:odtMetadata.$Field = $Value
    $result = Test-OfficeDeploymentTool -OdtPath 'C:\ODT\setup.exe'
    $result.Valid | Should -BeFalse
    $result.Detail | Should -Match 'Unrecognized ODT identity'
  }

  It 'rejects a build immediately below the supported minimum' {
    $script:odtMetadata.FileBuildPart = 12827
    $script:odtMetadata.FilePrivatePart = 20257
    $result = Test-OfficeDeploymentTool -OdtPath 'C:\ODT\setup.exe'
    $result.Valid | Should -BeFalse
    $result.Detail | Should -Match 'below the supported minimum'
  }

  It 'reports unreadable or missing files without claiming trust' {
    Mock Get-Item { throw 'Synthetic missing executable' }
    $result = Test-OfficeDeploymentTool -OdtPath 'C:\ODT\missing.exe'
    $result.Valid | Should -BeFalse
    $result.Version | Should -BeNullOrEmpty
    $result.ReasonCode | Should -Be UntrustedTool
    $result.Detail | Should -Match 'Synthetic missing executable'
    Should -Invoke Get-AuthenticodeSignature -Times 0
  }
}
