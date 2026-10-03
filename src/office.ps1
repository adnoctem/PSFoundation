#Requires -Version 5.0

# Both product-key validation sites share this text. It states the expected
# shape so a rejection is self-explanatory, and never echoes the key itself:
# -AsSecureString masks input, so a truncated paste is otherwise invisible.
$script:ProductKeyFormatMessage = 'ProductKey has an invalid format. Supply 25 characters as five groups of five letters or digits separated by ASCII hyphens, for example ABCDE-FGHIJ-KLMNO-PQRST-UVWXY.'

function Install-Office {
  <#
    .SYNOPSIS
      Installs Office only on a clean machine or returns a verified compliant no-op.
    .DESCRIPTION
      Accepts only an Install plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .PARAMETER ProductKey
      Optional SecureString volume key; never serialized or passed on a command line.
    .EXAMPLE
      $plan | Install-Office -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun,

    [Security.SecureString]
    $ProductKey
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'Install'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
      ProductKey     = $ProductKey
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Uninstall-Office {
  <#
    .SYNOPSIS
      Removes only explicitly selected Click-to-Run products.
    .DESCRIPTION
      Accepts only a Remove plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .EXAMPLE
      $plan | Uninstall-Office -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'Remove'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Switch-OfficeDeployment {
  <#
    .SYNOPSIS
      Executes an approved replacement after staging and verifying all destination media.
    .DESCRIPTION
      Accepts only a Migrate plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .PARAMETER ProductKey
      Optional SecureString volume key; never serialized or passed on a command line.
    .EXAMPLE
      $plan | Switch-OfficeDeployment -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun,

    [Security.SecureString]
    $ProductKey
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'Migrate'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
      ProductKey     = $ProductKey
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Update-Office {
  <#
    .SYNOPSIS
      Updates a pinned Office build while preserving other deployment dimensions.
    .DESCRIPTION
      Accepts only an Update plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .EXAMPLE
      $plan | Update-Office -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'Update'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Set-OfficeUpdateConfiguration {
  <#
    .SYNOPSIS
      Applies explicitly selected update settings without installing Office.
    .DESCRIPTION
      Accepts only a SetUpdateConfiguration plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .EXAMPLE
      $plan | Set-OfficeUpdateConfiguration -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'SetUpdateConfiguration'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Add-OfficeLanguage {
  <#
    .SYNOPSIS
      Adds selected full-UI languages while preserving primary language and existing resources.
    .DESCRIPTION
      Accepts only an AddLanguage plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .EXAMPLE
      $plan | Add-OfficeLanguage -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'AddLanguage'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Remove-OfficeLanguage {
  <#
    .SYNOPSIS
      Removes selected non-primary languages without removing the suite.
    .DESCRIPTION
      Accepts only a RemoveLanguage plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .EXAMPLE
      $plan | Remove-OfficeLanguage -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'RemoveLanguage'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Set-OfficeApplicationSelection {
  <#
    .SYNOPSIS
      Changes application exclusions while preserving product, build, architecture, and languages.
    .DESCRIPTION
      Accepts only a SetApplicationSelection plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .EXAMPLE
      $plan | Set-OfficeApplicationSelection -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'SetApplicationSelection'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}

function Set-OfficeApplicationPreference {
  <#
    .SYNOPSIS
      Applies validated Office preferences to existing and future users through ODT customize.
    .DESCRIPTION
      Accepts only a SetApplicationPreference plan. Revalidates current inventory and media before
      confirmation and again under the shared deployment lock. Returns one final
      result after cleanup. Native failures may have changed machine state.
    .PARAMETER Plan
      Matching plan from Get-OfficeDeploymentPlan; pipeline input is supported.
    .PARAMETER OdtPath
      Existing Microsoft-signed ODT setup.exe; never downloaded automatically.
    .PARAMETER LogRoot
      Local Administrators/SYSTEM-only journal and log directory.
    .PARAMETER ForceCloseApps
      Explicitly authorize closing Office applications across sessions.
    .PARAMETER DryRun
      Return a read-only preview without journals, staging, or installer invocation.
    .EXAMPLE
      $plan | Set-OfficeApplicationPreference -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'The shared lifecycle calls ShouldProcess on the supplied PSCmdlet before any mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Plan,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office'),

    [switch]
    $ForceCloseApps,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Plan           = $Plan
      ExpectedAction = 'SetApplicationPreference'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      LogRoot        = $LogRoot
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
    }
    Invoke-PSFOfficeWorkflow @parameters
  }
}


# Office deployment contracts. No discovery or mutation occurs during module import.
function Stop-PSFOfficeOperation {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Throws a structured exception without changing machine state.')]
  [CmdletBinding()]
  param (
    [string]
    $Reason,

    [string]
    $Message,

    [object]
    $Diagnostic
  )

  $exception = New-Object InvalidOperationException($Message)
  $exception.Data['OfficeReason'] = $Reason
  if ($Diagnostic) { $exception.Data['OfficeDiagnostic'] = $Diagnostic }
  throw $exception
}

function Assert-PSFOfficeField {
  [CmdletBinding()]
  param (
    [object]
    $InputObject,

    [string[]]
    $Allowed,

    [string[]]
    $Required = @()
  )

  if ($null -eq $InputObject -or $InputObject -is [string]) {
    Stop-PSFOfficeOperation InvalidContract 'An Office contract must be a data object.'
  }
  if ($InputObject -is [Collections.IDictionary]) {
    $names = @($InputObject.Keys)
  }
  else {
    $names = @($InputObject.PSObject.Properties | ForEach-Object { $_.Name })
  }
  foreach ($name in $names) {
    if ($name -notin $Allowed) {
      Stop-PSFOfficeOperation InvalidContract "Unexpected Office contract field: $name."
    }
  }
  foreach ($name in $Required) {
    if ($name -notin $names) {
      Stop-PSFOfficeOperation InvalidContract "Missing Office contract field: $name."
    }
  }
}

function ConvertTo-PSFOfficeList {
  [CmdletBinding()]
  param (
    [AllowEmptyCollection()]
    [string[]]
    $Value,

    [switch]
    $Language
  )

  $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($item in $Value) {
    if ([string]::IsNullOrWhiteSpace($item)) {
      Stop-PSFOfficeOperation InvalidConfiguration 'Empty identifiers are not allowed.'
    }
    $normalized = $item.Trim()
    if ($Language) {
      $normalized = $normalized.ToLowerInvariant()
      # Deliberately bounded full-UI language support; no inferred proofing/LIP conversion.
      if ($normalized -notin @(
          'en-us',
          'de-de',
          'fr-fr',
          'es-es',
          'it-it',
          'nl-nl',
          'pt-br',
          'pt-pt',
          'ja-jp',
          'ko-kr',
          'zh-cn',
          'zh-tw',
          'pl-pl',
          'cs-cz',
          'da-dk',
          'fi-fi',
          'sv-se',
          'nb-no',
          'hu-hu',
          'tr-tr',
          'el-gr',
          'ro-ro',
          'sk-sk',
          'sl-si',
          'hr-hr',
          'bg-bg',
          'et-ee',
          'lv-lv',
          'lt-lt'
        )) {
        Stop-PSFOfficeOperation Unsupported "Language '$normalized' is outside the supported full-UI language catalog."
      }
    }
    elseif ($normalized -notmatch '^[A-Za-z0-9]+$') {
      Stop-PSFOfficeOperation InvalidConfiguration 'Product/application identifiers must be alphanumeric.'
    }
    if ($seen.Add($normalized)) {
      $normalized
    }
  }
}

function Get-PSFOfficeFingerprint {
  [CmdletBinding()]
  param (
    [object]
    $InputObject
  )

  $json = ConvertTo-Json -InputObject $InputObject -Depth 30 -Compress
  $sha = [Security.Cryptography.SHA256]::Create()
  try {
    ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($json)))).Replace('-', '').ToLowerInvariant()
  }
  finally {
    $sha.Dispose()
  }
}

function Open-PSFOfficeRegistryBase {
  [CmdletBinding()]
  param ([Microsoft.Win32.RegistryView]$View)

  [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $View)
}

function Get-PSFOfficeMachineId {
  [CmdletBinding()]
  param ()

  $view = [Microsoft.Win32.RegistryView]::Registry32
  if ([Environment]::Is64BitOperatingSystem) { $view = [Microsoft.Win32.RegistryView]::Registry64 }
  $base = $null
  $key = $null
  try {
    $base = Open-PSFOfficeRegistryBase -View $view
    $key = $base.OpenSubKey('SOFTWARE\Microsoft\Cryptography', $false)
    if ($null -eq $key) {
      Stop-PSFOfficeOperation MachineIdentityUnavailable "Machine identity key is missing in $view."
    }
    $identity = $key.GetValue('MachineGuid', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    $parsed = [guid]::Empty
    if ($identity -isnot [string] -or -not [guid]::TryParse($identity, [ref]$parsed) -or $parsed -eq [guid]::Empty) {
      Stop-PSFOfficeOperation InvalidMachineIdentity "Machine identity (MachineGuid) is missing or invalid in $view."
    }
    # Return the original string: normalization would change existing journal bindings.
    return $identity
  }
  catch {
    if ($_.Exception.Data['OfficeReason']) { throw }
    Stop-PSFOfficeOperation MachineIdentityUnavailable "Cannot read MachineGuid from HKLM\SOFTWARE\Microsoft\Cryptography ($view)."
  }
  finally {
    if ($null -ne $key) { $key.Dispose() }
    if ($null -ne $base) { $base.Dispose() }
  }
}

function Get-PSFOfficeOsLocale {
  [CmdletBinding()]
  param ()

  # Machine installation UI language, not an administrator's regional format/culture.
  $key = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language' -ErrorAction Stop
  $lcid = [Convert]::ToInt32($key.InstallLanguage, 16)
  $language = [Globalization.CultureInfo]::GetCultureInfo($lcid).Name.ToLowerInvariant()
  [PSCustomObject]@{
    Languages       = @($language)
    PrimaryLanguage = $language
    Evidence        = 'HKLM SYSTEM CurrentControlSet Control Nls Language:InstallLanguage (machine installation UI language)'
  }
}

function Get-PSFOfficeSourceLocale {
  [CmdletBinding()]
  param ([object]$Inventory)

  $products = @($Inventory.Products)
  $languages = @($products | ForEach-Object { $_.Languages } | Where-Object { $_ } | Sort-Object -Unique)
  $primaries = @($products | ForEach-Object { $_.PrimaryLanguage } | Where-Object { $_ } | Sort-Object -Unique)
  $incomplete = @($products | Where-Object {
      $null -eq $_.Languages -or -not @($_.Languages).Count -or -not $_.PrimaryLanguage -or $_.PrimaryLanguage -notin $_.Languages
    })
  $known = $products.Count -gt 0 -and -not $Inventory.Unknowns.Count -and -not $Inventory.Msi.Count -and
  -not $incomplete.Count -and $primaries.Count -eq 1
  [PSCustomObject]@{
    Known           = $known
    Languages       = $languages
    PrimaryLanguage = if ($known) { $primaries[0] } else { $null }
    Evidence        = @($products | ForEach-Object { if ($_.PSObject.Properties['Evidence']) { $_.Evidence } })
  }
}

function New-OfficeDeploymentConfiguration {
  <#
    .SYNOPSIS
      Creates an ordered, secret-free Office target configuration in memory.
    .DESCRIPTION
      Defaults to en-us regardless of the execution account. Automatic locale
      discovery is explicit. This object grants no authority to change a machine.
    .PARAMETER TargetProductId
      Supported suite product ID. No product is selected implicitly.
    .PARAMETER Architecture
      Office architecture, 32 or 64. Defaults to 64.
    .PARAMETER Channel
      Volume channel is derived from the product; subscriptions default to Current.
    .PARAMETER Language
      Ordered full-UI languages. The first language is the primary shell language.
    .PARAMETER Version
      Exact build; when omitted, preparation resolves and pins a build.
    .PARAMETER ExcludeApp
      Applications excluded from the target suite.
    .PARAMETER AutoSourceLocales
      Explicitly discover languages instead of using the en-us default.
    .PARAMETER LocaleSource
      InstalledOffice or OperatingSystem. Requires AutoSourceLocales.
    .EXAMPLE
      New-OfficeDeploymentConfiguration -TargetProductId Standard2024Volume -Language en-us,de-de
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Builds a data object without mutation.')]
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet(
      'Standard2019Volume',
      'ProPlus2019Volume',
      'Standard2021Volume',
      'ProPlus2021Volume',
      'Standard2024Volume',
      'ProPlus2024Volume',
      'O365ProPlusRetail',
      'O365BusinessRetail'
    )]
    [string]
    $TargetProductId,

    [ValidateSet('32', '64')]
    [string]
    $Architecture = '64',

    [ValidateSet('PerpetualVL2019', 'PerpetualVL2021', 'PerpetualVL2024', 'Current', 'MonthlyEnterprise', 'SemiAnnual')]
    [string]
    $Channel,

    [string[]]
    $Language,

    [ValidatePattern('^16\.0\.\d+\.\d+$')]
    [string]
    $Version,

    [ValidateSet('Access', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Teams', 'Word')]
    [string[]]
    $ExcludeApp = @(),

    [switch]
    $AutoSourceLocales,

    [ValidateSet('InstalledOffice', 'OperatingSystem')]
    [string]
    $LocaleSource = 'InstalledOffice'
  )

  if (($PSBoundParameters.ContainsKey('Language') -and $AutoSourceLocales) -or
    ($PSBoundParameters.ContainsKey('LocaleSource') -and -not $AutoSourceLocales)) {
    Stop-PSFOfficeOperation InvalidConfiguration 'Language and automatic discovery are mutually exclusive; LocaleSource requires AutoSourceLocales.'
  }
  $requested = @()
  if ($PSBoundParameters.ContainsKey('Language')) {
    $requested = @($Language)
  }
  $evidence = @()
  $source = 'Default'
  if ($AutoSourceLocales) {
    $source = $LocaleSource
    if ($LocaleSource -eq 'OperatingSystem') {
      $found = Get-PSFOfficeOsLocale
    }
    else {
      $inventory = Get-OfficeInventory
      $found = Get-PSFOfficeSourceLocale $inventory
      if (-not $found.Known) {
        Stop-PSFOfficeOperation LocaleDiscoveryFailed 'Installed Office language preservation requires fully observed Click-to-Run products with one agreed primary language and no MSI or unknown inventory. Supply Language explicitly.'
      }
    }
    if (-not $found.PrimaryLanguage -or $null -eq $found.Languages -or -not $found.Languages.Count -or $found.PrimaryLanguage -notin $found.Languages) {
      Stop-PSFOfficeOperation LocaleDiscoveryFailed 'No unambiguous primary Office language was observed. Supply Language explicitly.'
    }
    $Language = @($found.PrimaryLanguage) + @($found.Languages | Where-Object { $_ -ne $found.PrimaryLanguage })
    $evidence = @($found.Evidence)
  }
  elseif ($PSBoundParameters.ContainsKey('Language')) {
    $source = 'Explicit'
  }
  else {
    $Language = @('en-us')
  }
  $languages = @(ConvertTo-PSFOfficeList -Value $Language -Language)
  if (-not $languages.Count) {
    Stop-PSFOfficeOperation InvalidConfiguration 'At least one language is required.'
  }
  $expectedChannel = 'Current'
  if ($TargetProductId -match '(2019|2021|2024)Volume$') {
    $expectedChannel = 'PerpetualVL' + $Matches[1]
  }
  if (-not $Channel) {
    $Channel = $expectedChannel
  }
  if (($TargetProductId -like '*Volume' -and $Channel -ne $expectedChannel) -or
    ($TargetProductId -like 'O365*' -and $Channel -like 'Perpetual*')) {
    Stop-PSFOfficeOperation InvalidConfiguration 'Product and channel do not match.'
  }
  [PSCustomObject][ordered]@{
    SchemaVersion      = 1
    TargetProductId    = $TargetProductId
    Architecture       = $Architecture
    Channel            = $Channel
    Language           = $languages
    PrimaryLanguage    = $languages[0]
    Version            = $Version
    ExcludeApp         = @($ExcludeApp | Sort-Object -Unique)
    RequestedLanguages = $requested
    LocaleSource       = $source
    LocaleEvidence     = $evidence
  }
}

function ConvertTo-PSFOfficeConfiguration {
  [CmdletBinding()]
  param (
    [object]
    $Configuration
  )

  $fields = @(
    'SchemaVersion',
    'TargetProductId',
    'Architecture',
    'Channel',
    'Language',
    'PrimaryLanguage',
    'Version',
    'ExcludeApp',
    'RequestedLanguages',
    'LocaleSource',
    'LocaleEvidence'
  )
  Assert-PSFOfficeField $Configuration $fields $fields
  if ($Configuration.SchemaVersion -ne 1) {
    Stop-PSFOfficeOperation InvalidContract 'Unsupported configuration schema.'
  }
  $parameters = @{
    TargetProductId = $Configuration.TargetProductId
    Architecture    = $Configuration.Architecture
    Channel         = $Configuration.Channel
    Language        = @($Configuration.Language)
    ExcludeApp      = @($Configuration.ExcludeApp)
  }
  if ($Configuration.Version) {
    $parameters.Version = $Configuration.Version
  }
  $normalized = New-OfficeDeploymentConfiguration @parameters
  if ($Configuration.PrimaryLanguage -ne $normalized.PrimaryLanguage) {
    Stop-PSFOfficeOperation InvalidContract 'PrimaryLanguage must equal the first ordered language.'
  }
  if ($Configuration.LocaleSource -notin @('Default', 'Explicit', 'InstalledOffice', 'OperatingSystem', 'Recovery')) {
    Stop-PSFOfficeOperation InvalidContract 'Invalid locale source.'
  }
  $normalized.RequestedLanguages = @($Configuration.RequestedLanguages | ForEach-Object { [string]$_ })
  $normalized.LocaleSource = [string]$Configuration.LocaleSource
  $normalized.LocaleEvidence = @($Configuration.LocaleEvidence | ForEach-Object { [string]$_ })
  $normalized
}

function Get-PSFOfficeRegistrySnapshot {
  [CmdletBinding()]
  param ()

  foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    if ($view -eq [Microsoft.Win32.RegistryView]::Registry64 -and -not [Environment]::Is64BitOperatingSystem) {
      continue
    }
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
    try {
      $paths = @(
        'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
        'SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0',
        'SOFTWARE\Microsoft\Office\12.0\Common\LanguageResources',
        'SOFTWARE\Microsoft\Office\14.0\Common\LanguageResources',
        'SOFTWARE\Microsoft\Office\15.0\Common\LanguageResources',
        'SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources',
        'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE',
        'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
      )
      # Preserve registered resource evidence without treating proofing packs as UI
      # languages. The bounded walk also exposes interrupted product registrations.
      $releaseRoot = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
      $releasePaths = @($releaseRoot)
      for ($depth = 0; $depth -lt 4; $depth++) {
        $next = @()
        foreach ($releasePath in $releasePaths) {
          $releaseKey = $base.OpenSubKey($releasePath)
          if ($releaseKey) {
            try {
              $paths += $releasePath
              $next += @($releaseKey.GetSubKeyNames() | ForEach-Object { $releasePath + '\' + $_ })
              if ($next.Count -gt 4096) {
                Stop-PSFOfficeOperation UnknownInventory 'Office resource registration exceeds supported discovery bounds.'
              }
            }
            finally {
              $releaseKey.Dispose()
            }
          }
        }
        $releasePaths = $next
      }
      $uninstall = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
      if ($uninstall) {
        try {
          $paths += @($uninstall.GetSubKeyNames() | ForEach-Object { 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' + $_ })
        }
        finally {
          $uninstall.Dispose()
        }
      }
      foreach ($path in $paths) {
        $key = $base.OpenSubKey($path)
        if (-not $key) {
          continue
        }
        try {
          $values = [ordered]@{}
          foreach ($name in $key.GetValueNames()) {
            # Inventory is an allowlist; never collect registration keys/PIDKEYs.
            if ($name -in @(
                'OfficeProductReleaseIds',
                'OfficePackageVersion',
                'ProductReleaseIds',
                'Platform',
                'VersionToReport',
                'CDNBaseUrl',
                'UpdateChannel',
                'ClientCulture',
                'InstallLanguage',
                'SKULanguage',
                'Language',
                'ActiveConfiguration',
                'DisplayName',
                'DisplayVersion',
                'Publisher',
                'WindowsInstaller',
                'UninstallString',
                'SystemComponent',
                'ParentKeyName',
                'ParentDisplayName'
              ) -or $name -like '*.ExcludedApps' -or
              ($name -eq 'Version' -and $path.StartsWith($releaseRoot + '\', [StringComparison]::OrdinalIgnoreCase))) {
              $values[$name] = $key.GetValue($name)
            }
            elseif ($name -eq '' -and $path -like '*\App Paths\*') {
              # An App Paths default value is the registered executable path. It
              # separates Click-to-Run residue from ordinary MSI registration.
              $values['(default)'] = $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            }
          }
          [PSCustomObject]@{
            View    = [string]$view
            Path    = $path
            Values  = [PSCustomObject]$values
            SubKeys = @($key.GetSubKeyNames())
          }
        }
        finally {
          $key.Dispose()
        }
      }
    }
    finally {
      $base.Dispose()
    }
  }
}

function Resolve-PSFOfficeInstalledVersion {
  [CmdletBinding()]
  param (
    [object[]]
    $Records,

    [string]
    $View,

    [string]
    $ProductId,

    [string[]]
    $ConfiguredProductIds
  )

  $result = [PSCustomObject]@{
    Version  = $null
    Source   = $null
    Evidence = 'Installed version unknown: no complete active product resource registration'
    Issue    = $null
  }
  # Ordered evidence: absence may fall through, contradictions must not. The
  # resource fallback is an observed registry layout, not Microsoft's documented
  # inventory contract. Never derive a build from telemetry or requested XML.
  foreach ($source in @('ClickToRunInventory', 'ActiveProductResources')) {
    switch ($source) {
      'ClickToRunInventory' {
        $path = 'SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0'
        $installed = @($Records | Where-Object { $_.View -eq $View -and $_.Path -eq $path })
        if (-not $installed.Count) { continue }
        if ($installed.Count -ne 1) {
          $result.Evidence = "Installed version unknown: duplicate $View inventory records"
          return $result
        }
        $ids = @(([string]$installed[0].Values['OfficeProductReleaseIds'] -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ((@($ids | Sort-Object) -join ',') -ne (@($ConfiguredProductIds | Sort-Object) -join ',')) {
          $result.Issue = 'ConflictingInstalledProductIdentity'
          $result.Evidence = "Installed version unknown: $View inventory product identity differs from configuration"
          return $result
        }
        if (-not $installed[0].Values.ContainsKey('OfficePackageVersion')) { continue }
        $value = $installed[0].Values['OfficePackageVersion']
        $parsed = $null
        if ($value -isnot [string] -or $value -notmatch '^\d+\.\d+\.\d+\.\d+$' -or
          -not [version]::TryParse($value, [ref]$parsed)) {
          $result.Evidence = "Installed version unknown: malformed $View inventory OfficePackageVersion"
          return $result
        }
        $result.Version = $parsed.ToString()
        $result.Source = $source
        $result.Evidence = "Installed version from $View`: $path`:OfficePackageVersion"
        return $result
      }
      'ActiveProductResources' {
        $root = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
        $active = @($Records | Where-Object { $_.View -eq $View -and $_.Path -eq $root })
        if ($active.Count -ne 1 -or $ProductId -notmatch '^[a-z0-9]+$' -or
          [string]$active[0].Values['ActiveConfiguration'] -notmatch '^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$') {
          return $result
        }
        $path = $root + '\' + $active[0].Values['ActiveConfiguration'] + '\' + $ProductId + '.16'
        $product = @($Records | Where-Object { $_.View -eq $View -and $_.Path -eq $path })
        if ($product.Count -ne 1) { return $result }
        $cultures = @($product[0].SubKeys)
        if ('x-none' -notin $cultures -or $cultures.Count -lt 2 -or
          @($cultures | Where-Object { $_ -ne 'x-none' -and $_ -notmatch '^[a-z]{2,3}-[a-z]{2,4}$' }).Count) {
          return $result
        }
        $versions = @()
        foreach ($culture in $cultures) {
          $leafPath = $path + '\' + $culture
          $leaf = @($Records | Where-Object { $_.View -eq $View -and $_.Path -eq $leafPath })
          $parsed = $null
          if ($leaf.Count -ne 1 -or $leaf[0].Values['Version'] -isnot [string] -or
            $leaf[0].Values['Version'] -notmatch '^\d+\.\d+\.\d+\.\d+$' -or
            -not [version]::TryParse($leaf[0].Values['Version'], [ref]$parsed)) {
            $result.Evidence = "Installed version unknown: missing, duplicate or malformed Version at $View`: $leafPath"
            return $result
          }
          $versions += $parsed.ToString()
        }
        $versions = @($versions | Sort-Object -Unique)
        if ($versions.Count -ne 1) {
          $result.Evidence = "Installed version unknown: conflicting active resource versions at $View`: $path"
          return $result
        }
        $result.Version = $versions[0]
        $result.Source = $source
        $result.Evidence = "Installed version derived from agreeing Version values at $View`: $path resources [$($cultures -join ',')]"
        return $result
      }
    }
  }
  $result
}

function Get-PSFOfficePathEnvironment {
  [CmdletBinding()]
  param ([string]$View)

  # Expand against the registration's view, not the PowerShell process's bitness.
  $program = $env:ProgramFiles
  $common = $env:CommonProgramFiles
  if ([Environment]::Is64BitOperatingSystem) {
    $program = $env:ProgramW6432
    $common = $env:CommonProgramW6432
    if ($View -eq 'Registry32') {
      $program = ${env:ProgramFiles(x86)}
      $common = ${env:CommonProgramFiles(x86)}
    }
  }
  @{
    ProgramFiles              = $program
    CommonProgramFiles        = $common
    'ProgramFiles(x86)'       = ${env:ProgramFiles(x86)}
    'CommonProgramFiles(x86)' = ${env:CommonProgramFiles(x86)}
    ProgramW6432              = $env:ProgramW6432
    CommonProgramW6432        = $env:CommonProgramW6432
    SystemDrive               = $env:SystemDrive
    SystemRoot                = $env:SystemRoot
    windir                    = $env:SystemRoot
  }
}

function Get-PSFOfficePathAttribute {
  [CmdletBinding()]
  param ([string]$LiteralPath)

  # Unlike File.Exists/Test-Path, failures are not collapsed into absence.
  [IO.File]::GetAttributes($LiteralPath)
}

function Get-PSFOfficeAppPathEvidence {
  [CmdletBinding()]
  param (
    [object]$RawTarget,
    [string]$View,
    [string]$RegistryPath
  )

  $evidence = [PSCustomObject][ordered]@{
    RegistryView   = $View
    RegistryPath   = $RegistryPath
    RawTarget      = $RawTarget
    ResolvedTarget = $null
    State          = 'Uncertain'
    Reason         = 'InvalidTarget'
  }
  try {
    if ($RawTarget -isnot [string] -or [string]::IsNullOrWhiteSpace($RawTarget) -or $View -notin @('Registry32', 'Registry64')) {
      return $evidence
    }
    $path = $RawTarget.Trim()
    if ($path.StartsWith('"') -and $path.EndsWith('"') -and $path.Length -gt 1) {
      $path = $path.Substring(1, $path.Length - 2)
    }
    $environment = Get-PSFOfficePathEnvironment $View
    foreach ($token in @([regex]::Matches($path, '%([^%]+)%'))) {
      $name = $token.Groups[1].Value
      if (-not $environment.ContainsKey($name) -or [string]::IsNullOrWhiteSpace($environment[$name])) {
        $evidence.Reason = 'UnresolvedEnvironment'
        return $evidence
      }
      $path = $path.Replace($token.Value, [string]$environment[$name])
    }
    # Reject UNC/device/provider paths, wildcards, arguments, ADS and ambiguous
    # segments before filesystem access. No incidental network traversal is allowed.
    if ($path -notmatch '^[A-Za-z]:\\' -or $path -match '[%"*?<>|/\x00-\x1f]' -or
      $path.Substring(2).Contains(':') -or $path -match '\\\\|\\\.{1,2}(\\|$)|[. ](\\|$)' -or
      $path -match '\\(?:CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])(?:\\|\.|$)' -or
      @($path.Substring(3) -split '\\' | Where-Object { $_.Length -gt 255 }).Count) {
      return $evidence
    }
    $expectedName = ($RegistryPath -split '\\')[-1]
    if ($expectedName -notin @('WINWORD.EXE', 'EXCEL.EXE', 'OUTLOOK.EXE') -or [IO.Path]::GetFileName($path) -ne $expectedName) {
      return $evidence
    }
    $path = [IO.Path]::GetFullPath($path)
    $evidence.ResolvedTarget = $path
    $root = [IO.Path]::GetPathRoot($path)
    $drive = New-Object IO.DriveInfo($root)
    if ($drive.DriveType -ne [IO.DriveType]::Fixed -or -not $drive.IsReady) {
      $evidence.Reason = 'UnsupportedDrive'
      return $evidence
    }
    # Do not infer absence from a redirected Windows system path in WOW64.
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess -and
      (($env:SystemRoot -and $path.StartsWith($env:SystemRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) -or $path -match '~\d')) {
      $evidence.Reason = 'FilesystemRedirection'
      return $evidence
    }
    $cursor = $root
    $parts = @($path.Substring($root.Length) -split '\\')
    # Inspect ancestors before descending: junctions cannot induce network access
    # or turn an unavailable target into evidence of a removed executable.
    $attributes = Get-PSFOfficePathAttribute $cursor
    if ($attributes -band [IO.FileAttributes]::ReparsePoint) {
      $evidence.Reason = 'ReparsePoint'
      return $evidence
    }
    for ($index = 0; $index -lt $parts.Count; $index++) {
      $cursor = [IO.Path]::Combine($cursor, $parts[$index])
      try {
        $attributes = Get-PSFOfficePathAttribute $cursor
      }
      catch {
        $errorType = $_.Exception.GetBaseException()
        if ($errorType -is [IO.FileNotFoundException] -or $errorType -is [IO.DirectoryNotFoundException]) {
          $evidence.State = 'Missing'
          $evidence.Reason = 'TargetNotFound'
          return $evidence
        }
        throw
      }
      if ($attributes -band [IO.FileAttributes]::ReparsePoint) {
        $evidence.Reason = 'ReparsePoint'
        return $evidence
      }
      $isDirectory = [bool]($attributes -band [IO.FileAttributes]::Directory)
      if ($isDirectory -ne ($index -lt ($parts.Count - 1))) {
        $evidence.Reason = 'UnexpectedPathType'
        return $evidence
      }
    }
    $evidence.State = 'Present'
    $evidence.Reason = 'ExecutablePresent'
  }
  catch {
    $evidence.Reason = 'ProbeFailed'
    if ($_.Exception.GetBaseException() -is [UnauthorizedAccessException] -or $_.Exception.GetBaseException() -is [Security.SecurityException]) {
      $evidence.Reason = 'AccessDenied'
    }
  }
  $evidence
}

function Get-OfficeInventory {
  <#
    .SYNOPSIS
      Reads Office registrations without launching Office or Windows Installer.
    .DESCRIPTION
      Inspects both machine registry views and preserves incomplete or conflicting
      evidence. Language completeness and shell language are not inferred from
      ClientCulture alone. Unknown properties cannot establish compliance.
      Installed build prefers documented Click-to-Run inventory, falling back to
      agreeing Version values on every active per-product resource when the
      documented key or value is absent. VersionSource and Evidence identify the
      observation used; telemetry never establishes the installed build.
      AppPathEvidence retains registered targets and present, missing or uncertain
      filesystem observations. Confirmed missing references alone do not block a
      clean inventory. Discovery never deletes stale registrations.
    .EXAMPLE
      Get-OfficeInventory
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param ()

  $products = New-Object Collections.ArrayList
  $msi = New-Object Collections.ArrayList
  $related = New-Object Collections.ArrayList
  $unknowns = New-Object Collections.ArrayList
  $records = @()
  try {
    $records = @(Get-PSFOfficeRegistrySnapshot | ForEach-Object {
        # Registry values are optional. Dictionary indexers preserve missing
        # evidence as null even under StrictMode; do not invent MSI or locale data.
        $values = @{}
        foreach ($property in $_.Values.PSObject.Properties) {
          $values[$property.Name] = $property.Value
        }

        $subKeys = @()
        if ($_.PSObject.Properties['SubKeys']) {
          $subKeys = @($_.SubKeys)
        }

        [PSCustomObject]@{
          View    = $_.View
          Path    = $_.Path
          Values  = $values
          SubKeys = $subKeys
        }
      })
  }
  catch {
    [void]$unknowns.Add('RegistryDiscoveryFailed')
  }
  $configuredRecords = @($records | Where-Object { $_.Path -like '*ClickToRun\Configuration' })
  $appPaths = @($records | Where-Object { $_.Path -like '*\App Paths\*' } | ForEach-Object {
      Get-PSFOfficeAppPathEvidence -RawTarget $_.Values['(default)'] -View $_.View -RegistryPath $_.Path
    })
  # Genuine registry residue remains blocking. App Paths also occur on MSI
  # installs; only confirmed missing targets dismiss a Click-to-Run reference.
  $residueRecords = @($records | Where-Object {
      $_.Path -like '*ClickToRun\ProductReleaseIDs*' -or
      $_.Path -like '*ClickToRun\Inventory\Office\16.0'
    })
  $appResidue = @($appPaths | Where-Object {
      $_.State -ne 'Missing' -and
      ([string]$_.RawTarget -match 'Office16|ClickToRun' -or $_.ResolvedTarget -match 'Office16|ClickToRun' -or
      ($_.State -eq 'Uncertain' -and $_.ResolvedTarget -notmatch '\\Office(?:11|12|14|15)\\(?:WINWORD|EXCEL|OUTLOOK)\.EXE$'))
    })

  if (-not $configuredRecords.Count -and ($residueRecords.Count -or $appResidue.Count)) {
    [void]$unknowns.Add('OfficeResidueWithoutConfiguration')
  }
  $registeredResources = @($records | Where-Object { $_.Path -like '*ClickToRun\ProductReleaseIDs*' } | ForEach-Object {
      [PSCustomObject]@{
        View                = $_.View
        Path                = $_.Path
        SubKeys             = @($_.SubKeys)
        ActiveConfiguration = $_.Values['ActiveConfiguration']
        Version             = $_.Values['Version']
      }
    })
  foreach ($record in $records) {
    $values = $record.Values
    if ($record.Path -like '*ClickToRun\Configuration') {
      $ids = @(([string]$values['ProductReleaseIds'] -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
      if (-not $ids.Count) {
        [void]$unknowns.Add('IncompleteClickToRunRegistration')
      }
      foreach ($id in $ids) {
        $installedVersion = Resolve-PSFOfficeInstalledVersion -Records $records -View $record.View -ProductId $id -ConfiguredProductIds $ids
        if ($installedVersion.Issue) {
          [void]$unknowns.Add($installedVersion.Issue)
        }
        $architecture = $null
        if ($values['Platform'] -eq 'x64') {
          $architecture = '64'
        }
        elseif ($values['Platform'] -eq 'x86') {
          $architecture = '32'
        }
        $channels = @{
          '492350f6-3a01-4f97-b9c0-c7c6ddf67d60' = 'Current'
          '55336b82-a18d-4dd6-b5f6-9e5095c314a6' = 'MonthlyEnterprise'
          '7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' = 'SemiAnnual'
          'f2e724c1-748f-4b47-8fb8-8e0d210e9208' = 'PerpetualVL2019'
          '5030841d-c919-4594-8d2d-84ae4f96e58e' = 'PerpetualVL2021'
          '7983bac0-e531-40cf-be00-fd24fe66619c' = 'PerpetualVL2024'
        }
        $channel = $null
        foreach ($entry in $channels.GetEnumerator()) {
          if ([string]$values['CDNBaseUrl'] -like "*$($entry.Key)*") {
            $channel = $entry.Value
          }
        }
        $excluded = $null
        $exclusionEvidence = 'ExcludeApp unknown: registration is missing or not a string'
        if ($values.ContainsKey("$id.ExcludedApps") -and $values["$id.ExcludedApps"] -is [string]) {
          $excluded = @(($values["$id.ExcludedApps"] -split ',') | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
          $exclusionEvidence = "ExcludeApp from $($record.View):$($record.Path) value $id.ExcludedApps"
        }
        # These are registered candidates, not proof of complete installed UI
        # resources or the initial shell language. Ignore inactive configurations.
        $registeredLanguages = @()
        $releaseRoot = 'SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs'
        $active = @($records | Where-Object { $_.View -eq $record.View -and $_.Path -eq $releaseRoot })
        if ($active.Count -eq 1 -and [string]$active[0].Values['ActiveConfiguration'] -match '^[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}$') {
          $productPath = $releaseRoot + '\' + $active[0].Values['ActiveConfiguration'] + '\' + $id + '.16'
          $resource = @($records | Where-Object { $_.View -eq $record.View -and $_.Path -eq $productPath })
          if ($resource.Count -eq 1) {
            $registeredLanguages = @($resource[0].SubKeys | Where-Object { $_ -match '^[a-z]{2,3}-[a-z]{2,4}$' -and $_ -ne 'x-none' } | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object -Unique)
          }
        }
        # Installed UI languages come from the active configuration's per-product
        # resource registration, which lists the payloads actually present. The
        # shell language needs a second, agreeing source: a single registered
        # language is unambiguous on its own, and otherwise ClientCulture counts
        # only when the registration corroborates it. A ClientCulture outside the
        # registered set is a conflict and stays unknown - it is never the answer
        # by itself, and requested XML is never evidence.
        $languages = $null
        $primaryLanguage = $null
        $languageEvidence = 'Languages unknown: no active per-product resource registration'
        $primaryEvidence = 'PrimaryLanguage unknown: no corroborated shell language'
        $clientCulture = ([string]$values['ClientCulture']).ToLowerInvariant()

        if ($registeredLanguages.Count) {
          $languages = @($registeredLanguages)
          $languageEvidence = "Languages from $($record.View):$productPath subkeys [$($languages -join ',')]"

          if ($languages.Count -eq 1) {
            $primaryLanguage = $languages[0]
            $primaryEvidence = "PrimaryLanguage from a single registered language [$primaryLanguage]"
          }
          elseif ($clientCulture -and $clientCulture -in $languages) {
            $primaryLanguage = $clientCulture
            $primaryEvidence = "PrimaryLanguage from ClientCulture=$clientCulture corroborated by the resource registration"
          }
          elseif ($clientCulture) {
            $primaryEvidence = "PrimaryLanguage unknown: ClientCulture=$clientCulture is not in the registered languages [$($languages -join ',')]"
          }
        }

        [void]$products.Add([PSCustomObject][ordered]@{
            ProductId           = $id
            Architecture        = $architecture
            Version             = $installedVersion.Version
            VersionSource       = $installedVersion.Source
            Channel             = $channel
            Languages           = $languages
            PrimaryLanguage     = $primaryLanguage
            RegisteredLanguages = $registeredLanguages
            ExcludeApp          = $excluded
            Evidence            = @(
              "$($record.View):$($record.Path)",
              $installedVersion.Evidence,
              "Telemetry VersionToReport=$($values['VersionToReport']); not installation evidence",
              "ClientCulture=$($values['ClientCulture']); not proof of complete languages or shell UI",
              $languageEvidence,
              $primaryEvidence,
              $exclusionEvidence
            )
          })
      }
    }
    elseif ($record.Path -like '*\Uninstall\*') {
      $keyName = Split-Path $record.Path -Leaf
      $isMicrosoft = $values['Publisher'] -match '^Microsoft(?: Corporation)?$'
      $officeCode = $keyName -match '^\{9[01](12|14|15|16)0000-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{7}FF1CE\}$'
      $officeName = $values['DisplayName'] -match 'Office|Visio|Project|Access|SharePoint Designer|InfoPath|Lync'
      $controller = $values['UninstallString'] -match '\\OFFICE(12|14|15|16)\\Office Setup Controller\\setup\.exe"?\s+/uninstall\s'
      # Both 007E and 008F occur in native Licensing Component registrations.
      # Product codes identify releases/packages, not stable component roles.
      # Keep the Office 16 family, Microsoft publisher and exact known role
      # checks together; do not maintain a list of component SKU fragments.
      $infrastructure = $keyName -match '^\{9[01]160000-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{7}FF1CE\}$' -and
      $values['DisplayName'] -match '^Office 16 Click-to-Run (Licensing|Extensibility|Localization) Component(?: 64-bit Registration)?$'
      $addIn = $values['WindowsInstaller'] -eq 1 -and $values['DisplayName'] -in @(
        'Microsoft Teams Meeting Add-in for Microsoft Office',
        'Microsoft Office Live Add-in 1.5'
      )
      if ($isMicrosoft -and ($infrastructure -or $addIn)) {
        [void]$related.Add([PSCustomObject]@{
            ProductCode  = $keyName
            Name         = $values['DisplayName']
            Version      = $values['DisplayVersion']
            RegistryView = $record.View
            Role         = if ($infrastructure) { 'ClickToRunInfrastructure' } else { 'AddIn' }
          })
        if ($infrastructure -and -not $configuredRecords.Count) {
          [void]$unknowns.Add('ClickToRunInfrastructureWithoutConfiguration')
        }
        continue
      }

      # Windows Installer patch registrations (service packs, hotfixes) are
      # registered per product and carry the parent product code. They are not
      # separate Office products. Classify them by structure, never by display
      # name: the name is localized and cannot be matched reliably.
      $patchParent = $null

      if ($keyName -match '^(\{[0-9A-F-]{36}\})_.+_\{[0-9A-F-]{36}\}$') {
        $patchParent = $Matches[1]
      }
      elseif ([string]$values['UninstallString'] -match '/package\s+(\{[0-9A-F-]{36}\})\s+/uninstall\s+\{[0-9A-F-]{36}\}') {
        $patchParent = $Matches[1]
      }

      if ($patchParent -notmatch '^\{9[01](12|14|15|16)0000-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{7}FF1CE\}$') {
        $patchParent = $null
      }

      if ($isMicrosoft -and $patchParent) {
        [void]$related.Add([PSCustomObject]@{
            ProductCode       = $keyName
            Name              = $values['DisplayName']
            Version           = $values['DisplayVersion']
            RegistryView      = $record.View
            Role              = 'PatchRegistration'
            ParentProductCode = $patchParent
            SystemComponent   = ($values['SystemComponent'] -eq 1)
          })
        continue
      }

      if ($isMicrosoft -and ($officeCode -or ($officeName -and ($values['WindowsInstaller'] -eq 1 -or $controller)))) {
        $resourceKind = 'ProductOrComponent'
        if (($officeCode -and $keyName -match '^\{[^-]+-(001F|002C)-') -or $values['DisplayName'] -match '\bProof(?:ing)?\b') {
          $resourceKind = 'Proofing'
        }
        elseif ($values['DisplayName'] -match 'Language Interface Pack') {
          $resourceKind = 'LanguageInterfacePack'
        }
        elseif ($values['DisplayName'] -match 'Language Pack|MUI') {
          $resourceKind = 'LanguageResource'
        }
        $languageId = $null
        if ($officeCode -and $keyName -match '^\{[^-]+-[^-]+-([0-9A-F]{4})-') {
          try {
            $languageId = [Globalization.CultureInfo]::GetCultureInfo([Convert]::ToInt32($Matches[1], 16)).Name.ToLowerInvariant()
          }
          catch {
            $languageId = $null
          }
        }
        [void]$msi.Add([PSCustomObject]@{
            ProductCode  = $keyName
            Name         = $values['DisplayName']
            Version      = $values['DisplayVersion']
            RegistryView = $record.View
            LanguageId   = $languageId
            ResourceKind = $resourceKind
          })
      }
      elseif ($isMicrosoft -and $officeName -and $values['DisplayName'] -notmatch 'Update|Hotfix|Security|Language|Proofing' -and
        $values['UninstallString'] -notmatch 'OfficeClickToRun\.exe') {
        [void]$unknowns.Add("UnclassifiedOfficeRegistration:$keyName")
      }
    }
  }

  # Registry order is not guaranteed, so resolve patch parents only once every
  # product has been read. A patch whose product is absent is real missing
  # evidence and stays unknown.
  $msiProductCodes = @($msi | ForEach-Object { $_.ProductCode })

  foreach ($patch in @($related | Where-Object { $_.Role -eq 'PatchRegistration' })) {
    if ($patch.ParentProductCode -notin $msiProductCodes) {
      [void]$unknowns.Add("OrphanedOfficePatchRegistration:$($patch.ProductCode)")
    }
  }

  $unique = @()
  foreach ($group in @($products | Group-Object ProductId)) {
    $variants = @($group.Group | ForEach-Object { '{0}|{1}|{2}' -f $_.Architecture, $_.Version, $_.Channel } | Sort-Object -Unique)
    if ($variants.Count -gt 1) {
      [void]$unknowns.Add('ConflictingRegistryViews')
    }
    $unique += $group.Group[0]
  }

  # Report what this machine's evidence actually cannot establish, rather than a
  # blanket capability claim. An observed product whose languages or shell
  # language could not be derived is still a hard limitation; a machine with no
  # Click-to-Run product to sample reports none, and any post-deployment gap
  # surfaces as a verification Unknown instead of silently passing.
  $limitations = New-Object Collections.ArrayList

  foreach ($product in $unique) {
    if ($null -eq $product.Languages) {
      [void]$limitations.Add('Languages')
    }

    if (-not $product.PrimaryLanguage) {
      [void]$limitations.Add('PrimaryLanguage')
    }
  }

  [PSCustomObject][ordered]@{
    SchemaVersion           = 1
    MachineId               = Get-PSFOfficeMachineId
    Products                = @($unique | Sort-Object ProductId)
    Msi                     = @($msi | Sort-Object ProductCode -Unique)
    RelatedComponents       = @($related | Sort-Object ProductCode, RegistryView)
    Unknowns                = @($unknowns | Sort-Object -Unique)
    RegisteredResources     = $registeredResources
    AppPathEvidence         = $appPaths
    LanguageEvidence        = @($records | Where-Object { $_.Path -like '*\Common\LanguageResources' } | ForEach-Object {
        [PSCustomObject]@{ View = $_.View; Path = $_.Path; SKULanguage = $_.Values['SKULanguage']; InstallLanguage = $_.Values['InstallLanguage'] }
      })
    VerificationLimitations = @($limitations | Sort-Object -Unique)
  }
}

function Test-OfficeDeployment {
  <#
    .SYNOPSIS
      Compares observed Office state with an explicit target.
    .DESCRIPTION
      Returns detailed discrepancies and unknowns. A success marker or activation
      status never substitutes for current configuration evidence.
    .PARAMETER Configuration
      Configuration from New-OfficeDeploymentConfiguration.
    .PARAMETER Inventory
      Optional observation for offline assessment. Execution always rediscovers state.
    .EXAMPLE
      Test-OfficeDeployment -Configuration $configuration
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Configuration,

    [object]
    $Inventory
  )

  $target = ConvertTo-PSFOfficeConfiguration $Configuration
  if (-not $Inventory) {
    $Inventory = Get-OfficeInventory
  }
  $differences = New-Object Collections.ArrayList
  $unknowns = New-Object Collections.ArrayList
  foreach ($unknown in $Inventory.Unknowns) {
    [void]$unknowns.Add($unknown)
  }
  $product = @($Inventory.Products | Where-Object { $_.ProductId -eq $target.TargetProductId })
  if ($product.Count -ne 1) {
    [void]$differences.Add('ProductId')
  }
  else {
    $actual = $product[0]
    foreach ($field in @('Architecture', 'Channel', 'Version', 'PrimaryLanguage')) {
      if (-not $target.$field -or -not $actual.$field) {
        [void]$unknowns.Add($field)
      }
      elseif ($actual.$field -ne $target.$field) {
        [void]$differences.Add($field)
      }
    }
    foreach ($pair in @(@('Languages', 'Language'), @('ExcludeApp', 'ExcludeApp'))) {
      if ($null -eq $actual.($pair[0])) {
        [void]$unknowns.Add($pair[0])
      }
      elseif ((@($actual.($pair[0]) | Sort-Object) -join ',') -ne (@($target.($pair[1]) | Sort-Object) -join ',')) {
        [void]$differences.Add($pair[0])
      }
    }
  }
  if (@($Inventory.Products | Where-Object { $_.ProductId -ne $target.TargetProductId }).Count -or $Inventory.Msi.Count) {
    [void]$differences.Add('OtherProducts')
  }
  [PSCustomObject]@{
    SchemaVersion = 1
    Compliant     = ($differences.Count -eq 0 -and $unknowns.Count -eq 0)
    Discrepancies = @($differences)
    Unknowns      = @($unknowns)
    Inventory     = $Inventory
  }
}

function Get-OfficeActivationStatus {
  <#
    .SYNOPSIS
      Assesses target activation separately from installation compliance.
    .DESCRIPTION
      Subscription activation requires the licensed user's session. Volume license
      status is matched to the selected SKU without returning partial product keys.
    .PARAMETER TargetProductId
      Exact Office suite product ID.
    .EXAMPLE
      Get-OfficeActivationStatus -TargetProductId Standard2024Volume
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $TargetProductId
  )

  $status = 'Unknown'
  if ($TargetProductId -in @('O365ProPlusRetail', 'O365BusinessRetail')) {
    $status = 'UserActivationRequired'
  }
  elseif ($TargetProductId -match '^(Standard|ProPlus)(2019|2021|2024)Volume$') {
    $pattern = '\bOffice\s*\d+,\s*Office\d+' + $Matches[1] + $Matches[2] + 'VL_'
    $licenses = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationID='0ff1ce15-a989-479d-af46-f275c6370663'" -ErrorAction Stop |
        Where-Object { $_.Name -match $pattern -and $_.PartialProductKey })
    $status = 'NotVerified'
    if (@($licenses | Where-Object { $_.LicenseStatus -eq 1 }).Count) {
      $status = 'Licensed'
    }
  }
  [PSCustomObject]@{
    TargetProductId = $TargetProductId
    Status          = $status
  }
}

function Assert-PSFOfficeSetting {
  [CmdletBinding()]
  param (
    [string]
    $Action,

    [object]
    $Settings
  )

  $allowed = @()
  switch ($Action) {
    'SetUpdateConfiguration' { $allowed = @('Enabled', 'UpdatePath', 'TargetVersion', 'Channel') }
    'SetApplicationPreference' { $allowed = @('Preferences') }
  }
  Assert-PSFOfficeField $Settings $allowed
  if ($Settings -is [Collections.IDictionary]) {
    $names = @($Settings.Keys)
  }
  else {
    $names = @($Settings.PSObject.Properties | ForEach-Object { $_.Name })
  }
  if ($Action -eq 'SetUpdateConfiguration') {
    if (-not $names.Count) {
      Stop-PSFOfficeOperation InvalidConfiguration 'Specify at least one update setting.'
    }
    if ('Enabled' -in $names -and $Settings.Enabled -isnot [bool]) {
      Stop-PSFOfficeOperation InvalidConfiguration 'Enabled must be a Boolean.'
    }
    if ('UpdatePath' -in $names -and [string]$Settings.UpdatePath -notmatch '^(https://|[A-Za-z]:\\|\\\\)') {
      Stop-PSFOfficeOperation InvalidConfiguration 'UpdatePath must be HTTPS or an absolute local/UNC path.'
    }
    if ('TargetVersion' -in $names -and [string]$Settings.TargetVersion -notmatch '^16\.0\.\d+\.\d+$') {
      Stop-PSFOfficeOperation InvalidConfiguration 'TargetVersion must be an exact Office build.'
    }
    if ('Channel' -in $names -and $Settings.Channel -notin @(
        'Current',
        'MonthlyEnterprise',
        'SemiAnnual',
        'PerpetualVL2019',
        'PerpetualVL2021',
        'PerpetualVL2024'
      )) {
      Stop-PSFOfficeOperation InvalidConfiguration 'Unsupported update channel.'
    }
  }
  elseif ($Action -eq 'SetApplicationPreference') {
    if ('Preferences' -notin $names -or -not @($Settings.Preferences).Count) {
      Stop-PSFOfficeOperation InvalidConfiguration 'At least one application preference is required.'
    }
    foreach ($preference in $Settings.Preferences) {
      Assert-PSFOfficeField $preference @('Key', 'Name', 'Value', 'Type', 'App', 'Id') @('Key', 'Name', 'Value', 'Type', 'App', 'Id')
      if ($preference.Key -notmatch '^software\\microsoft\\office\\16\.0\\(word|excel|powerpoint|outlook|access|onenote)(\\[a-z0-9 _-]+)+$' -or
        $preference.Type -notin @('REG_DWORD', 'REG_SZ') -or
        $preference.App -notin @('word16', 'excel16', 'ppt16', 'outlook16', 'access16', 'onenote16') -or
        $preference.Name -notmatch '^[a-zA-Z0-9 _-]+$' -or $preference.Id -notmatch '^[a-zA-Z0-9_-]+$') {
        Stop-PSFOfficeOperation InvalidConfiguration 'Application preference is outside the supported Office preference schema.'
      }
      if ($preference.Value -isnot [string] -and $preference.Value -isnot [int] -and $preference.Value -isnot [long]) {
        Stop-PSFOfficeOperation InvalidConfiguration 'Preference values must be strings or integers.'
      }
      if ($preference.Type -eq 'REG_DWORD' -and [string]$preference.Value -notmatch '^\d{1,10}$') {
        Stop-PSFOfficeOperation InvalidConfiguration 'REG_DWORD requires an unsigned decimal value.'
      }
      if ($preference.Type -eq 'REG_DWORD' -and [long]$preference.Value -gt [uint32]::MaxValue) {
        Stop-PSFOfficeOperation InvalidConfiguration 'REG_DWORD exceeds the unsigned 32-bit range.'
      }
    }
  }
}

function Get-OfficeDeploymentPlan {
  <#
    .SYNOPSIS
      Builds a read-only Office operation plan with explicit authority.
    .DESCRIPTION
      Returns blockers rather than silently reconciling conflicting state. A plan
      captures observations, not execution authority; every executor rediscovers
      state and validates the action again. Recovery uses dedicated resume commands.
    .PARAMETER Action
      Install, Remove, Migrate, Update, or a narrowly scoped maintenance operation.
    .PARAMETER Configuration
      Ordered target configuration. Not required for product removal.
    .PARAMETER SourcePath
      Prepared media package; never downloaded by a deployment command.
    .PARAMETER RemoveProductId
      Exact Click-to-Run product selections. Only Remove and Migrate accept these.
    .PARAMETER RemoveMsi
      Authorize broad supported MSI removal during migration only.
    .PARAMETER Language
      Exact resources selected by AddLanguage or RemoveLanguage.
    .PARAMETER Settings
      Validated update settings or application preferences, specific to the action.
    .PARAMETER Inventory
      Optional observation for offline planning. Never trusted during execution.
    .EXAMPLE
      Get-OfficeDeploymentPlan -Action Install -Configuration $target -SourcePath C:\Media\Office
    .EXAMPLE
      Get-OfficeDeploymentPlan -Action Remove -RemoveProductId O365ProPlusRetail
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet(
      'Install',
      'Remove',
      'Migrate',
      'Update',
      'SetUpdateConfiguration',
      'AddLanguage',
      'RemoveLanguage',
      'SetApplicationSelection',
      'SetApplicationPreference'
    )]
    [string]
    $Action,

    [object]
    $Configuration,

    [string]
    $SourcePath,

    [string[]]
    $RemoveProductId = @(),

    [switch]
    $RemoveMsi,

    [string[]]
    $Language = @(),

    [object]
    $Settings = @{},

    [object]
    $Inventory
  )

  if ($null -eq $RemoveProductId) { $RemoveProductId = @() }
  if ($null -eq $Language) { $Language = @() }
  if ($Action -notin @('Remove', 'Migrate') -and ($RemoveProductId.Count -or $RemoveMsi)) {
    Stop-PSFOfficeOperation InvalidAuthority 'This action cannot remove products.'
  }
  if ($Action -notin @('AddLanguage', 'RemoveLanguage') -and $Language.Count) {
    Stop-PSFOfficeOperation InvalidAuthority 'Language selection belongs only to language operations.'
  }
  Assert-PSFOfficeSetting $Action $Settings
  $selection = @(ConvertTo-PSFOfficeList $RemoveProductId)
  foreach ($id in $selection) {
    if ($id -notmatch '^(O365(ProPlus|Business)|ProPlus\d{0,4}|Standard\d{0,4}|Professional\d{0,4}|HomeBusiness\d{0,4}|HomeStudent\d{0,4}|Personal\d{0,4}|Visio(Pro|Std)\d{0,4}|Project(Pro|Std)\d{0,4}|Access\d{0,4}|Excel\d{0,4}|Outlook\d{0,4}|PowerPoint\d{0,4}|Word\d{0,4}|Publisher\d{0,4})(Retail|Volume)$') {
      Stop-PSFOfficeOperation Unsupported 'Removal selection is outside the supported Office product-ID families.'
    }
  }
  $languages = @(ConvertTo-PSFOfficeList $Language -Language)
  $target = $null
  if ($Configuration) {
    $target = ConvertTo-PSFOfficeConfiguration $Configuration
  }
  if ($Action -ne 'Remove' -and -not $target) {
    Stop-PSFOfficeOperation InvalidConfiguration 'This operation requires a complete target configuration.'
  }
  $settingNames = @()
  if ($Settings -is [Collections.IDictionary]) { $settingNames = @($Settings.Keys) }
  else { $settingNames = @($Settings.PSObject.Properties | ForEach-Object { $_.Name }) }
  if ($Action -eq 'SetUpdateConfiguration' -and 'Channel' -in $settingNames) {
    $null = New-OfficeDeploymentConfiguration -TargetProductId $target.TargetProductId -Channel $Settings.Channel -Language $target.Language
  }
  if ($Action -eq 'Remove' -and ($target -or $SourcePath -or (-not $selection.Count -and -not $RemoveMsi))) {
    Stop-PSFOfficeOperation InvalidConfiguration 'Removal requires an explicit selection and accepts no installation configuration or media.'
  }
  if ($Action -in @('AddLanguage', 'RemoveLanguage') -and -not $languages.Count) {
    Stop-PSFOfficeOperation InvalidConfiguration 'Select at least one language resource.'
  }
  if (-not $Inventory) {
    $Inventory = Get-OfficeInventory
  }
  $blockers = New-Object Collections.ArrayList
  $warnings = New-Object Collections.ArrayList
  $state = 'Clean'
  $verification = $null
  $media = $null
  if ($Inventory.Unknowns.Count) {
    $state = 'Unknown'
    [void]$blockers.Add('UnknownInventory')
  }
  elseif ($Inventory.Products.Count -or $Inventory.Msi.Count) {
    $state = 'Conflict'
  }
  if ($target) {
    if ($SourcePath) {
      $media = Test-OfficeDeploymentMedia -SourcePath $SourcePath -Configuration $target
      if (-not $media.Valid) {
        [void]$blockers.Add($media.ReasonCode)
      }
      elseif (-not $target.Version) {
        $target.Version = $media.Manifest.Version
      }
    }
    $verification = Test-OfficeDeployment -Configuration $target -Inventory $Inventory
    if ($verification.Compliant) {
      $state = 'Compliant'
    }
    elseif (@($Inventory.Products | Where-Object { $_.ProductId -eq $target.TargetProductId }).Count -and
      -not $verification.Discrepancies.Count) {
      $state = 'Incomplete'
    }
  }
  if ($Action -eq 'Install') {
    if ($state -eq 'Incomplete') {
      [void]$blockers.Add('RecoveryRequired')
    }
    elseif ($state -notin @('Clean', 'Compliant')) {
      [void]$blockers.Add('Conflict')
    }
  }
  if ($Action -eq 'Remove' -and $RemoveMsi) {
    [void]$blockers.Add('UnsupportedStandaloneMsi')
  }
  if ($Action -eq 'Remove' -and @($Inventory.Products | Where-Object { $_.ProductId -in $selection }).Count) {
    $retained = @($Inventory.Products | Where-Object { $_.ProductId -notin $selection })
    if (@($retained | Where-Object { $null -eq $_.Languages -or -not $_.PrimaryLanguage -or $null -eq $_.ExcludeApp }).Count) {
      [void]$blockers.Add('UnsupportedSharedComponentVerification')
    }
  }
  if ($Action -eq 'Migrate') {
    $others = @($Inventory.Products | Where-Object { $_.ProductId -notin $selection })
    # Complete verification includes the absence of other products, MSI and
    # unknowns. An absent original source then permits only a compliant no-op.
    if ($others.Count -and $state -ne 'Compliant') {
      [void]$blockers.Add('UnapprovedProducts')
    }
    if ($state -ne 'Compliant' -and @($selection | Where-Object { $_ -notin @($Inventory.Products | ForEach-Object { $_.ProductId }) }).Count) {
      [void]$blockers.Add('StaleRemovalSelection')
    }
    if ($Inventory.Msi.Count -and -not $RemoveMsi) {
      [void]$blockers.Add('MsiConsentRequired')
    }
    if (@($Inventory.Msi | Where-Object { $_.Version -notmatch '^(12|14|15|16)\.' -or $_.Name -match 'Lync.*2010|Access.*(Database Engine|Runtime).*2007' }).Count) {
      [void]$blockers.Add('UnsupportedMsiComponent')
    }
  }
  $maintenance = $Action -notin @('Install', 'Remove', 'Migrate')
  if ($maintenance) {
    if ($Inventory.Msi.Count -or $Inventory.Products.Count -ne 1 -or $Inventory.Products[0].ProductId -ne $target.TargetProductId) {
      [void]$blockers.Add('Conflict')
    }
    else {
      $allowedDifference = switch ($Action) {
        'Update' { 'Version' }
        'SetApplicationSelection' { 'ExcludeApp' }
        { $_ -in @('AddLanguage', 'RemoveLanguage') } { 'Languages' }
      }
      if ($verification.Unknowns.Count -or @($verification.Discrepancies | Where-Object { $_ -notin @($allowedDifference) }).Count) {
        [void]$blockers.Add('UnverifiedPreservedConfiguration')
      }
      $current = $Inventory.Products[0]
      if ($Action -eq 'Update' -and $target.Version -and [version]$target.Version -lt [version]$current.Version) {
        [void]$blockers.Add('DowngradeRequiresMigration')
      }
      if ($Action -in @('AddLanguage', 'RemoveLanguage')) {
        $expected = @($current.Languages)
        if ($Action -eq 'AddLanguage') {
          $expected += $languages
        }
        else {
          $expected = @($expected | Where-Object { $_ -notin $languages })
          if ($current.PrimaryLanguage -in $languages -or -not $expected.Count) {
            [void]$blockers.Add('PrimaryLanguageRequiresMigration')
          }
        }
        if ((@($expected | Sort-Object -Unique) -join ',') -ne (@($target.Language | Sort-Object) -join ',')) {
          [void]$blockers.Add('LanguageSelectionMismatch')
        }
      }
    }
  }
  $needsMedia = $Action -in @('Install', 'Migrate', 'Update', 'AddLanguage', 'SetApplicationSelection') -and $state -ne 'Compliant'
  $unsupportedVerification = @($Inventory.VerificationLimitations | Where-Object { $_ })
  if ($Action -in @('Install', 'Migrate') -and $state -ne 'Compliant' -and $unsupportedVerification.Count) {
    [void]$blockers.Add('UnsupportedNativeVerification')
    [void]$warnings.Add('The native inventory backend cannot verify: ' + ($Inventory.VerificationLimitations -join ', ') + '. No deployment may start until these postconditions can be verified.')
  }
  if ($needsMedia -and -not $SourcePath) {
    [void]$blockers.Add('MissingMedia')
  }
  $sourceLanguages = @($Inventory.Products | ForEach-Object { $_.Languages } | Where-Object { $_ } | Sort-Object -Unique)
  $languageKnown = ($Inventory.Msi.Count -eq 0 -and -not $Inventory.Unknowns.Count -and
    ($Inventory.Products.Count -eq 0 -or (Get-PSFOfficeSourceLocale $Inventory).Known))
  $transition = $null
  if ($target) {
    $transition = [PSCustomObject]@{
      Known         = $languageKnown
      Before        = $sourceLanguages
      After         = @($target.Language)
      Added         = @($target.Language | Where-Object { $_ -notin $sourceLanguages })
      Removed       = @($sourceLanguages | Where-Object { $_ -notin $target.Language })
      PrimaryBefore = @($Inventory.Products | ForEach-Object { $_.PrimaryLanguage } | Where-Object { $_ } | Sort-Object -Unique)
      PrimaryAfter  = $target.PrimaryLanguage
    }
    if (-not $languageKnown) {
      [void]$warnings.Add('Source languages/primary language are unknown; preservation cannot be established.')
    }
    elseif (($Inventory.Products.Count -or $Inventory.Msi.Count) -and
      ($transition.Added.Count -or $transition.Removed.Count -or ($transition.PrimaryBefore -join ',') -ne $target.PrimaryLanguage)) {
      [void]$warnings.Add("Language change: $($sourceLanguages -join ',') -> $($target.Language -join ','); removed [$($transition.Removed -join ',')]; primary -> $($target.PrimaryLanguage).")
    }
  }
  if ($Action -eq 'SetUpdateConfiguration') {
    [void]$warnings.Add('Update settings can cause future downloads/build or channel transitions. Managed policy may override them.')
  }
  if ($Action -eq 'SetApplicationPreference') {
    [void]$warnings.Add('Application preferences affect existing and future users on this machine.')
  }
  $plan = [PSCustomObject][ordered]@{
    SchemaVersion        = 1
    Action               = $Action
    MachineId            = $Inventory.MachineId
    Configuration        = $target
    SourcePath           = $SourcePath
    RemoveProductId      = $selection
    RemoveMsi            = [bool]$RemoveMsi
    Language             = $languages
    Settings             = $Settings
    Before               = $Inventory
    InventoryFingerprint = Get-PSFOfficeFingerprint $Inventory
    State                = $state
    Eligible             = ($blockers.Count -eq 0)
    Blockers             = @($blockers | Sort-Object -Unique)
    Warnings             = @($warnings)
    LanguageTransition   = $transition
    MediaFingerprint     = if ($media -and $media.Valid) { $media.Fingerprint } else { $null }
  }
  $plan
}

function Confirm-PSFOfficePlan {
  [CmdletBinding()]
  param (
    [object]
    $Plan,

    [string]
    $Action
  )

  $fields = @(
    'SchemaVersion',
    'Action',
    'MachineId',
    'Configuration',
    'SourcePath',
    'RemoveProductId',
    'RemoveMsi',
    'Language',
    'Settings',
    'Before',
    'InventoryFingerprint',
    'State',
    'Eligible',
    'Blockers',
    'Warnings',
    'LanguageTransition',
    'MediaFingerprint'
  )
  Assert-PSFOfficeField $Plan $fields $fields
  if ($Plan.SchemaVersion -ne 1 -or $Plan.Action -ne $Action -or $Plan.RemoveMsi -isnot [bool]) {
    Stop-PSFOfficeOperation InvalidAuthority 'Plan schema, action, or removal authority does not match this command.'
  }
  $current = Get-OfficeInventory
  if ($Plan.MachineId -ne $current.MachineId -or $Plan.InventoryFingerprint -ne (Get-PSFOfficeFingerprint $current)) {
    Stop-PSFOfficeOperation StalePlan 'Machine or Office inventory changed; create and review a new plan.'
  }
  $parameters = @{
    Action          = $Action
    Configuration   = $Plan.Configuration
    SourcePath      = $Plan.SourcePath
    RemoveProductId = @($Plan.RemoveProductId)
    RemoveMsi       = $Plan.RemoveMsi
    Language        = @($Plan.Language)
    Settings        = $Plan.Settings
    Inventory       = $current
  }
  $fresh = Get-OfficeDeploymentPlan @parameters
  if ($Plan.MediaFingerprint -ne $fresh.MediaFingerprint) {
    Stop-PSFOfficeOperation StaleMedia 'Prepared media changed; create a new plan.'
  }
  $fresh
}

function Resolve-OfficeDeploymentToolSource {
  <#
    .SYNOPSIS
      Returns the reviewed Microsoft ODT download location without network I/O.
    .DESCRIPTION
      Downloads are versioned. Acquisition validates Authenticode trust before
      extraction; URL metadata and hashes alone do not establish publisher trust.
    .EXAMPLE
      Resolve-OfficeDeploymentToolSource
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param ()

  [PSCustomObject]@{
    Uri          = 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20326-20112.exe'
    Version      = '16.0.20326.20112'
    Publisher    = 'Microsoft Corporation'
    Reference    = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
    ReviewedDate = '2026-09-27'
  }
}

function Test-OfficeDeploymentToolSourceAvailability {
  <#
    .SYNOPSIS
      Checks whether the reviewed Microsoft ODT download is reachable.
    .DESCRIPTION
      Sends a HEAD request to Resolve-OfficeDeploymentToolSource's URI. Downloads
      no executable and does not establish publisher trust or installation state.
      Install-OfficeDeploymentTool verifies the downloaded and extracted files.
    .EXAMPLE
      Test-OfficeDeploymentToolSourceAvailability
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param ()

  $source = Resolve-OfficeDeploymentToolSource
  try {
    $response = Invoke-WebRequest -Uri $source.Uri -Method Head -UseBasicParsing -ErrorAction Stop
    [PSCustomObject]@{
      Uri           = $source.Uri
      Version       = $source.Version
      Available     = ([int]$response.StatusCode -ge 200 -and [int]$response.StatusCode -lt 300)
      StatusCode    = [int]$response.StatusCode
      ContentLength = $response.Headers['Content-Length']
      CheckedAt     = (Get-Date).ToUniversalTime()
    }
  }
  catch {
    [PSCustomObject]@{
      Uri       = $source.Uri
      Version   = $source.Version
      Available = $false
      Error     = $_.Exception.Message
      CheckedAt = (Get-Date).ToUniversalTime()
    }
  }
}

function Test-OfficeDeploymentTool {
  <#
    .SYNOPSIS
      Verifies the signature, identity, and minimum version of ODT setup.exe.
    .DESCRIPTION
      Does not execute or download the tool. Requires a valid Microsoft
      Corporation signature and a recognized Office executable identity.
      Current ODT uses Bootstrapper.exe as its embedded original filename;
      the filename on disk is not that version-resource field.
      Invalid assessments retain UntrustedTool and explain the failed check.
    .PARAMETER OdtPath
      Existing Office Deployment Tool setup.exe.
    .EXAMPLE
      Test-OfficeDeploymentTool -OdtPath C:\Tools\ODT\setup.exe
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath
  )

  $valid = $false
  $version = $null
  $reason = 'UntrustedTool'
  $detail = $null
  $signatureStatus = $null
  $info = $null
  try {
    Assert-PSFOfficePath $OdtPath
    $file = Get-Item -LiteralPath $OdtPath -ErrorAction Stop
    if ($file.PSIsContainer -or $file.Extension -ne '.exe') {
      throw 'OdtPath must identify an existing executable file.'
    }

    $signature = Get-AuthenticodeSignature -LiteralPath $OdtPath -ErrorAction Stop
    $signatureStatus = [string]$signature.Status
    $info = $file.VersionInfo
    $version = '{0}.{1}.{2}.{3}' -f $info.FileMajorPart, $info.FileMinorPart,
    $info.FileBuildPart, $info.FilePrivatePart

    # Keep the established identity and recognize the current Microsoft ODT
    # tuple. Neither a setup.exe filename nor an Office description alone is
    # sufficient. The trusted publisher check applies to both identities.
    $legacyIdentity = $info.OriginalFilename -eq 'setup.exe' -and
    $info.FileDescription -match 'Office.*(Deployment|Click-to-Run)'
    $bootstrapperIdentity = $info.OriginalFilename -eq 'Bootstrapper.exe' -and
    $info.FileDescription -eq 'Microsoft 365 and Office' -and
    $info.ProductName -eq 'Microsoft Office' -and
    $info.CompanyName -eq 'Microsoft Corporation'

    if ($signature.Status -ne 'Valid') {
      $detail = "ODT Authenticode signature is not valid (status: $signatureStatus)."
    }
    elseif (-not $signature.SignerCertificate -or
      $signature.SignerCertificate.Subject -notmatch '(?:^|,\s*)O=Microsoft Corporation(?:,|$)') {
      $detail = 'ODT signer is not Microsoft Corporation.'
    }
    elseif (-not ($legacyIdentity -or $bootstrapperIdentity)) {
      $detail = "Unrecognized ODT identity: OriginalFilename='$($info.OriginalFilename)', FileDescription='$($info.FileDescription)', ProductName='$($info.ProductName)', CompanyName='$($info.CompanyName)'. Use the extracted Office Deployment Tool setup.exe."
    }
    elseif ([version]$version -lt [version]'16.0.12827.20258') {
      $detail = "ODT version $version is below the supported minimum 16.0.12827.20258."
    }
    else {
      $valid = $true
      $reason = $null
    }
  }
  catch {
    $detail = $_.Exception.Message
  }

  [PSCustomObject]@{
    Valid            = $valid
    Path             = $OdtPath
    Version          = $version
    ReasonCode       = $reason
    Detail           = $detail
    SignatureStatus  = $signatureStatus
    OriginalFilename = if ($info) { $info.OriginalFilename } else { $null }
    FileDescription  = if ($info) { $info.FileDescription } else { $null }
    ProductName      = if ($info) { $info.ProductName } else { $null }
    CompanyName      = if ($info) { $info.CompanyName } else { $null }
    Modes            = @('Download', 'Configure', 'Customize', 'Help')
  }
}

function Assert-PSFOfficePath {
  [CmdletBinding()]
  param (
    [string]
    $Path
  )

  if ([string]::IsNullOrWhiteSpace($Path) -or -not [IO.Path]::IsPathRooted($Path) -or $Path -match '(^|[\\/])\.\.([\\/]|$)') {
    Stop-PSFOfficeOperation UnsafePath 'An absolute filesystem path without parent traversal is required.'
  }
  $full = [IO.Path]::GetFullPath($Path)
  if ($full.Substring([IO.Path]::GetPathRoot($full).Length) -match ':') {
    Stop-PSFOfficeOperation UnsafePath 'Alternate data streams are not supported.'
  }
  $cursor = $full
  while ($cursor) {
    if (Test-Path -LiteralPath $cursor) {
      $item = Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
      if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        Stop-PSFOfficeOperation UnsafePath 'Reparse traversal is not supported for Office deployment files.'
      }
    }
    $parent = [IO.Directory]::GetParent($cursor)
    $cursor = $null
    if ($parent) {
      $cursor = $parent.FullName
    }
  }
}

function Assert-PSFOfficeProtectedPath {
  [CmdletBinding()]
  param (
    [string]
    $Path,

    [string]
    $ObjectKind = 'DeploymentFile'
  )

  $diagnostic = [ordered]@{ Stage = 'Path'; ObjectKind = $ObjectKind; Path = $Path; Sid = $null; Rights = $null; IsInherited = $null; InheritanceFlags = $null; PropagationFlags = $null }
  try {
    Assert-PSFOfficePath $Path
    $diagnostic.Stage = 'AccessControlRead'
    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
  }
  catch {
    $reason = $_.Exception.Data['OfficeReason']
    if (-not $reason) { $reason = 'UntrustedMedia' }
    Stop-PSFOfficeOperation $reason "Cannot validate $ObjectKind at '$Path' ($($diagnostic.Stage))." ([PSCustomObject]$diagnostic)
  }
  $trusted = @('S-1-5-18', 'S-1-5-32-544')
  $owner = $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value
  if ($owner -notin $trusted) {
    $diagnostic.Stage = 'Owner'
    $diagnostic.Sid = $owner
    Stop-PSFOfficeOperation UntrustedMedia "Untrusted owner SID $owner on $ObjectKind at '$Path'; Administrators or SYSTEM ownership is required." ([PSCustomObject]$diagnostic)
  }
  foreach ($rule in $acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier])) {
    $write = [Security.AccessControl.FileSystemRights]::Write -bor [Security.AccessControl.FileSystemRights]::Delete -bor
    [Security.AccessControl.FileSystemRights]::ChangePermissions -bor [Security.AccessControl.FileSystemRights]::TakeOwnership
    if ($rule.AccessControlType -eq 'Allow' -and ($rule.FileSystemRights -band $write) -and $rule.IdentityReference.Value -notin $trusted) {
      $diagnostic.Stage = 'WriteGrant'
      $diagnostic.Sid = $rule.IdentityReference.Value
      $diagnostic.Rights = [string]$rule.FileSystemRights
      $diagnostic.IsInherited = $rule.IsInherited
      $diagnostic.InheritanceFlags = [string]$rule.InheritanceFlags
      $diagnostic.PropagationFlags = [string]$rule.PropagationFlags
      Stop-PSFOfficeOperation UntrustedMedia "Untrusted write grant for SID $($diagnostic.Sid) on $ObjectKind at '$Path': $($diagnostic.Rights); inherited=$($diagnostic.IsInherited); inheritance=$($diagnostic.InheritanceFlags); propagation=$($diagnostic.PropagationFlags)." ([PSCustomObject]$diagnostic)
    }
  }
}

function New-PSFOfficeProtectedDirectory {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private primitive called only after public ShouldProcess approval.')]
  [CmdletBinding()]
  param (
    [string]
    $Path
  )

  Assert-PSFOfficePath $Path
  if (Test-Path -LiteralPath $Path) {
    Assert-PSFOfficeProtectedPath $Path
    return
  }
  $parent = Split-Path $Path -Parent
  if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
    Stop-PSFOfficeOperation UnsafePath 'The parent directory must already exist.'
  }
  $acl = New-Object Security.AccessControl.DirectorySecurity
  $acl.SetSecurityDescriptorSddlForm('O:S-1-5-32-544G:S-1-5-32-544D:P(A;OICI;FA;;;S-1-5-18)(A;OICI;FA;;;S-1-5-32-544)')
  if ($PSVersionTable.PSEdition -eq 'Core') {
    $directory = New-Object IO.DirectoryInfo($Path)
    [IO.FileSystemAclExtensions]::Create($directory, $acl)
  }
  else {
    $null = [IO.Directory]::CreateDirectory($Path, $acl)
  }
  Assert-PSFOfficeProtectedPath $Path
}

function Remove-PSFOfficeWorkDirectory {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private cleanup of the operation-owned directory after approval.')]
  [CmdletBinding()]
  param (
    [string]
    $Path,

    [string]
    $Parent
  )

  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
    return
  }
  $full = [IO.Path]::GetFullPath($Path)
  $root = [IO.Path]::GetFullPath($Parent).TrimEnd('\') + '\'
  if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or $full.TrimEnd('\') -eq $root.TrimEnd('\')) {
    Stop-PSFOfficeOperation UnsafePath 'Cleanup path escaped its operation root.'
  }
  Assert-PSFOfficeProtectedPath $full
  foreach ($item in @(Get-ChildItem -LiteralPath $full -Recurse -Force -ErrorAction Stop)) {
    Assert-PSFOfficePath $item.FullName
  }
  Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction Stop
}

function Invoke-PSFOfficeTool {
  [CmdletBinding()]
  param (
    [string]
    $OdtPath,

    [ValidateSet('/download', '/configure', '/customize', '/help')]
    [string]
    $Mode,

    [string]
    $ConfigurationPath
  )

  $tool = Test-OfficeDeploymentTool -OdtPath $OdtPath
  if (-not $tool.Valid) {
    Stop-PSFOfficeOperation UntrustedTool 'ODT signature or executable identity validation failed.'
  }
  $arguments = @($Mode)
  if ($Mode -ne '/help') {
    Assert-PSFOfficeProtectedPath $ConfigurationPath
    $arguments += $ConfigurationPath
  }
  # No timeout/cancellation kill policy: ODT can delegate work to shared services.
  Invoke-SafeProcess -FilePath $OdtPath -ArgumentList $arguments -AsResult
}

function Get-OfficeDeploymentToolHelp {
  <#
    .SYNOPSIS
      Invokes native help on a verified Office Deployment Tool executable.
    .DESCRIPTION
      This is explicit executable invocation, separate from read-only planning.
    .PARAMETER OdtPath
      Microsoft-signed ODT setup.exe.
    .EXAMPLE
      Get-OfficeDeploymentToolHelp -OdtPath C:\Tools\ODT\setup.exe
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath
  )

  if ($PSCmdlet.ShouldProcess($OdtPath, 'Run verified ODT help')) {
    Invoke-PSFOfficeTool -OdtPath $OdtPath -Mode /help
  }
}

function Install-OfficeDeploymentTool {
  <#
    .SYNOPSIS
      Acquires and provisions a reviewed Microsoft Office Deployment Tool build.
    .DESCRIPTION
      Verifies the downloaded extractor before executing it and verifies extracted
      setup.exe. Existing trusted tools are reused; untrusted files are not replaced.
    .PARAMETER Destination
      Dedicated protected tool directory whose parent already exists.
    .PARAMETER DryRun
      Describe acquisition without downloading, creating files, or executing code.
    .EXAMPLE
      Install-OfficeDeploymentTool -Destination C:\Tools\ODT -WhatIf
    .EXAMPLE
      $tool = Install-OfficeDeploymentTool -Destination C:\Tools\ODT -Confirm
      $tool | Format-List Valid, Path, Version
  #>
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Destination,

    [switch]
    $DryRun
  )

  Assert-PSFOfficePath $Destination
  $setup = Join-Path $Destination 'setup.exe'
  if (Test-Path -LiteralPath $setup) {
    $existing = Test-OfficeDeploymentTool $setup
    if (-not $existing.Valid) {
      Stop-PSFOfficeOperation UntrustedTool 'Existing ODT is untrusted; it was not replaced.'
    }
    return $existing
  }
  $source = Resolve-OfficeDeploymentToolSource
  if ($DryRun -or -not $PSCmdlet.ShouldProcess($Destination, "Acquire verified ODT $($source.Version)")) {
    return [PSCustomObject]@{
      Status = 'Preview'
      Path   = $setup
      Source = $source
    }
  }
  $parent = Split-Path $Destination -Parent
  $work = Join-Path $parent ('PSFOfficeTool-' + [guid]::NewGuid().ToString('N'))
  try {
    New-PSFOfficeProtectedDirectory $work
    $package = Join-Path $work 'odt.exe'
    Invoke-WebRequest -Uri $source.Uri -OutFile $package -UseBasicParsing -ErrorAction Stop
    $signature = Get-AuthenticodeSignature -LiteralPath $package -ErrorAction Stop
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notmatch '(?:^|,\s*)O=Microsoft Corporation(?:,|$)') {
      Stop-PSFOfficeOperation UntrustedTool 'ODT extractor did not pass Microsoft signature validation.'
    }
    $extracted = Join-Path $work 'Extracted'
    New-PSFOfficeProtectedDirectory $extracted
    $native = Invoke-SafeProcess -FilePath $package -ArgumentList @('/quiet', "/extract:$extracted") -AsResult
    if ($native.TimedOut -or $native.Cancelled -or $null -eq $native.ExitCode -or $native.ExitCode -ne 0) {
      Stop-PSFOfficeOperation ToolExtractionFailed "ODT extraction returned $($native.ExitCode)."
    }
    $candidate = Join-Path $extracted 'setup.exe'
    if (-not (Test-OfficeDeploymentTool $candidate).Valid) {
      Stop-PSFOfficeOperation UntrustedTool 'Extracted ODT verification failed.'
    }
    New-PSFOfficeProtectedDirectory $Destination
    if (Test-Path -LiteralPath $setup) {
      Stop-PSFOfficeOperation Conflict 'Destination changed during extraction.'
    }
    [IO.File]::Copy($candidate, $setup, $false)
    $result = Test-OfficeDeploymentTool $setup
    if (-not $result.Valid) {
      Stop-PSFOfficeOperation UntrustedTool 'Provisioned ODT verification failed.'
    }
    $result
  }
  finally {
    Remove-PSFOfficeWorkDirectory -Path $work -Parent $parent
  }
}

function Get-PSFOfficeMediaFile {
  [CmdletBinding()]
  param (
    [string]
    $Root
  )

  Assert-PSFOfficeProtectedPath $Root
  $prefix = [IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
  $data = Join-Path $Root 'Office\Data'
  Assert-PSFOfficePath $data
  if (-not (Test-Path -LiteralPath $data -PathType Container)) {
    Stop-PSFOfficeOperation MissingMedia 'Office\Data is missing.'
  }
  Assert-PSFOfficeProtectedPath (Join-Path $Root 'Office')
  foreach ($entry in @(Get-ChildItem -LiteralPath (Join-Path $Root 'Office') -Force -ErrorAction Stop)) {
    if ($entry.Name -ne 'Data' -or -not $entry.PSIsContainer) {
      Stop-PSFOfficeOperation InvalidMedia 'Unexpected content outside Office/Data.'
    }
  }
  # Walk one level at a time so reparse directories are rejected before recursion.
  $queue = New-Object 'Collections.Generic.Queue[string]'
  $queue.Enqueue($data)
  while ($queue.Count) {
    $directory = $queue.Dequeue()
    Assert-PSFOfficeProtectedPath $directory
    foreach ($item in @(Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop)) {
      Assert-PSFOfficeProtectedPath $item.FullName
      if ($item.PSIsContainer) {
        $queue.Enqueue($item.FullName)
      }
      else {
        if ($item.Length -eq 0) {
          Stop-PSFOfficeOperation InvalidMedia 'Empty Office payload file.'
        }
        [PSCustomObject][ordered]@{
          Path   = $item.FullName.Substring($prefix.Length).Replace('\', '/')
          Length = $item.Length
          Hash   = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
        }
      }
    }
  }
}

function Test-OfficeDeploymentMedia {
  <#
    .SYNOPSIS
      Validates an Office package, manifest, payloads, paths, and write protection.
    .DESCRIPTION
      Schema 2 separates available languages from ordered deployment languages.
      Schema 1 packages require preparation again because their language semantics
      cannot establish this contract. Validation performs no writes or downloads.
      Metadata may use the architecture's generic CAB or its exact pinned-version
      CAB. Neutral and declared language streams remain mandatory. Missing payload
      diagnostics include relative paths; validation does not prove native offline use.
    .PARAMETER SourcePath
      Absolute package directory containing psfoundation-office-media.json.
    .PARAMETER Configuration
      Optional target whose requested languages must be a subset of package languages.
    .EXAMPLE
      Test-OfficeDeploymentMedia -SourcePath C:\Media\Office -Configuration $target
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Media names the deployment payload as a collective noun; Windows PowerShell and PowerShell 7 pluralizers disagree.')]
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $SourcePath,

    [object]
    $Configuration
  )

  $manifest = $null
  $stage = 'ProtectedPath'
  try {
    Assert-PSFOfficeProtectedPath $SourcePath -ObjectKind MediaDirectory
    $path = Join-Path $SourcePath 'psfoundation-office-media.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
      Stop-PSFOfficeOperation ReprepareMedia 'No schema-2 manifest exists; prepare this package again.'
    }
    Assert-PSFOfficeProtectedPath $path -ObjectKind MediaManifest
    $stage = 'Read'
    $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
    $stage = 'Json'
    $manifest = $json | ConvertFrom-Json -ErrorAction Stop
    $stage = 'Validation'
    $fields = @('SchemaVersion', 'Product', 'Architecture', 'Channel', 'AvailableLanguages', 'Version', 'ToolVersion', 'Files')
    Assert-PSFOfficeField $manifest $fields $fields
    if ($manifest.SchemaVersion -ne 2 -or $manifest.Version -notmatch '^16\.0\.\d+\.\d+$') {
      Stop-PSFOfficeOperation ReprepareMedia 'Unsupported or incomplete media schema; prepare again.'
    }
    $available = @(ConvertTo-PSFOfficeList $manifest.AvailableLanguages -Language)
    if (-not $available.Count -or $available.Count -ne @($manifest.AvailableLanguages).Count) {
      Stop-PSFOfficeOperation InvalidMedia 'Media languages are empty or duplicated.'
    }
    $manifestTarget = New-OfficeDeploymentConfiguration -TargetProductId $manifest.Product -Architecture $manifest.Architecture -Channel $manifest.Channel -Language $available -Version $manifest.Version
    $missingRequested = @()
    if ($Configuration) {
      $target = ConvertTo-PSFOfficeConfiguration $Configuration
      if ($target.TargetProductId -ne $manifest.Product -or $target.Architecture -ne $manifest.Architecture -or
        $target.Channel -ne $manifest.Channel -or ($target.Version -and $target.Version -ne $manifest.Version)) {
        Stop-PSFOfficeOperation MediaMismatch 'Package product, architecture, channel, or build differs from target.'
      }
      $missingRequested = @($target.Language | Where-Object { $_ -notin $available })
    }
    $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $manifest.Files) {
      Assert-PSFOfficeField $file @('Path', 'Length', 'Hash') @('Path', 'Length', 'Hash')
      if ($file.Path -notmatch '^Office/Data/[A-Za-z0-9._/-]+$' -or
        $file.Path -match '(^|/)\.\.?(/|$)|//|[. ](/|$)' -or -not $seen.Add($file.Path) -or
        $file.Hash -notmatch '^[a-fA-F0-9]{64}$' -or $file.Length -le 0) {
        Stop-PSFOfficeOperation UnsafeManifest 'Invalid or colliding media file record.'
      }
    }
    $files = @(Get-PSFOfficeMediaFile $SourcePath | Sort-Object Path)
    if (-not $files.Count -or $files.Count -ne @($manifest.Files).Count) {
      Stop-PSFOfficeOperation MediaIntegrityFailed 'Media file set differs from the manifest.'
    }
    foreach ($file in $files) {
      $record = @($manifest.Files | Where-Object { $_.Path -eq $file.Path })
      if ($record.Count -ne 1 -or $record[0].Length -ne $file.Length -or $record[0].Hash -ne $file.Hash) {
        Stop-PSFOfficeOperation MediaIntegrityFailed 'Media file size or hash differs from the manifest.'
      }
    }
    $platform = 'x64'
    if ($manifestTarget.Architecture -eq '32') {
      $platform = 'x86'
    }
    $catalogs = @("Office/Data/v$($manifest.Architecture).cab", "Office/Data/v$($manifest.Architecture)_$($manifest.Version).cab")
    $missingCatalog = -not @($catalogs | Where-Object { $_ -in $files.Path }).Count
    $neutral = "Office/Data/$($manifest.Version)/stream.$platform.x-none.dat"
    $missingBase = @($neutral | Where-Object { $_ -notin $files.Path })
    $requiredLanguages = @($available) + @($missingRequested)
    $missingLanguages = @($requiredLanguages | ForEach-Object { "Office/Data/$($manifest.Version)/stream.$platform.$_.dat" } | Where-Object { $_ -notin $files.Path })
    if ($missingCatalog -or $missingBase.Count -or $missingLanguages.Count -or $missingRequested.Count) {
      $diagnostics = @()
      if ($missingCatalog) {
        $diagnostics += "Missing metadata CAB; require one of: $($catalogs -join ', ')."
      }
      if ($missingBase.Count) {
        $diagnostics += "Missing neutral payload: $($missingBase -join ', ')."
      }
      if ($missingLanguages.Count) {
        $diagnostics += "Missing language payloads: $($missingLanguages -join ', ')."
      }
      if ($missingRequested.Count) {
        $diagnostics += "Requested languages not declared in AvailableLanguages: $($missingRequested -join ', ')."
      }
      $reason = 'MissingLanguageMedia'
      if ($missingCatalog -or $missingBase.Count) {
        $reason = 'MissingMedia'
      }
      Stop-PSFOfficeOperation $reason ($diagnostics -join ' ')
    }
    [PSCustomObject]@{
      Valid       = $true
      Path        = $SourcePath
      Manifest    = $manifest
      Fingerprint = Get-PSFOfficeFingerprint $manifest
      ReasonCode  = $null
      Error       = $null
      Diagnostic  = $null
    }
  }
  catch {
    $reason = $_.Exception.Data['OfficeReason']
    $message = $_.Exception.Message
    $diagnostic = $_.Exception.Data['OfficeDiagnostic']
    if (-not $reason) {
      $reason = 'InvalidMedia'
      $message = "Office media validation failed at $stage for '$SourcePath'."
    }
    if ($reason -eq 'InvalidContract') { $message = "Office media contract is invalid at '$SourcePath'." }
    if (-not $diagnostic) {
      $diagnostic = [PSCustomObject]@{
        Stage = $stage; ObjectKind = 'MediaPackage'; Path = $SourcePath
        ExceptionType = $_.Exception.GetType().FullName
        Function = 'Test-OfficeDeploymentMedia'; ScriptPath = $PSCommandPath; Line = $_.InvocationInfo.ScriptLineNumber
      }
    }
    [PSCustomObject]@{
      Valid       = $false
      Path        = $SourcePath
      Manifest    = $null
      Fingerprint = $null
      ReasonCode  = $reason
      Error       = $message
      Diagnostic  = $diagnostic
    }
  }
}

function New-PSFOfficeXml {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory XML document.')]
  [CmdletBinding()]
  param (
    [ValidateSet(
      'Download',
      'Install',
      'Remove',
      'Migrate',
      'Update',
      'AddLanguage',
      'RemoveLanguage',
      'SetApplicationSelection',
      'SetUpdateConfiguration',
      'SetApplicationPreference'
    )]
    [string]
    $Action,

    [object]
    $Configuration,

    [string]
    $MediaPath,

    [string[]]
    $RemoveProductId = @(),

    [bool]
    $RemoveMsi = $false,

    [string[]]
    $Language = @(),

    [object]
    $Settings = @{}
  )

  if ($null -eq $RemoveProductId) { $RemoveProductId = @() }
  if ($null -eq $Language) { $Language = @() }
  if (($Action -notin @('Remove', 'Migrate') -and ($RemoveProductId.Count -or $RemoveMsi)) -or
    ($Action -eq 'Remove' -and $RemoveMsi)) {
    Stop-PSFOfficeOperation InvalidAuthority 'XML operation cannot contain the requested removal.'
  }
  Assert-PSFOfficeSetting $Action $Settings
  $target = $null
  if ($Configuration) {
    $target = ConvertTo-PSFOfficeConfiguration $Configuration
  }
  $document = New-Object Xml.XmlDocument
  $document.XmlResolver = $null
  $root = $document.CreateElement('Configuration')
  [void]$document.AppendChild($root)
  if ($Action -in @('Remove', 'RemoveLanguage')) {
    $selection = @(ConvertTo-PSFOfficeList $RemoveProductId)
    if ($Action -eq 'RemoveLanguage') {
      $selection = @($target.TargetProductId)
      $Language = @(ConvertTo-PSFOfficeList $Language -Language)
      if (-not $Language.Count) {
        Stop-PSFOfficeOperation InvalidAuthority 'Language removal requires exact languages.'
      }
    }
    if (-not $selection.Count) {
      Stop-PSFOfficeOperation InvalidAuthority 'Product removal requires exact product IDs.'
    }
    $remove = $document.CreateElement('Remove')
    $remove.SetAttribute('All', 'FALSE')
    [void]$root.AppendChild($remove)
    foreach ($id in $selection) {
      $product = $document.CreateElement('Product')
      $product.SetAttribute('ID', $id)
      [void]$remove.AppendChild($product)
      if ($Action -eq 'RemoveLanguage') {
        foreach ($locale in $Language) {
          $node = $document.CreateElement('Language')
          $node.SetAttribute('ID', $locale)
          [void]$product.AppendChild($node)
        }
      }
    }
  }
  elseif ($Action -eq 'SetUpdateConfiguration') {
    $node = $document.CreateElement('Updates')
    if ($Settings -is [Collections.IDictionary]) {
      $names = @($Settings.Keys)
    }
    else {
      $names = @($Settings.PSObject.Properties | ForEach-Object { $_.Name })
    }
    foreach ($name in $names) {
      $node.SetAttribute($name, [string]$Settings.$name)
    }
    [void]$root.AppendChild($node)
  }
  elseif ($Action -eq 'SetApplicationPreference') {
    $node = $document.CreateElement('AppSettings')
    [void]$root.AppendChild($node)
    foreach ($preference in $Settings.Preferences) {
      $user = $document.CreateElement('User')
      foreach ($name in @('Key', 'Name', 'Value', 'Type', 'App', 'Id')) {
        $user.SetAttribute($name, [string]$preference.$name)
      }
      [void]$node.AppendChild($user)
    }
  }
  else {
    if (-not $target -or -not $MediaPath -or ($Action -ne 'Download' -and -not $target.Version)) {
      Stop-PSFOfficeOperation InvalidConfiguration 'Execution requires a target, media path, and pinned build.'
    }
    $add = $document.CreateElement('Add')
    $add.SetAttribute('OfficeClientEdition', $target.Architecture)
    $add.SetAttribute('Channel', $target.Channel)
    $add.SetAttribute('SourcePath', $MediaPath)
    $add.SetAttribute('AllowCdnFallback', 'FALSE')
    if ($target.Version) {
      $add.SetAttribute('Version', $target.Version)
    }
    [void]$root.AppendChild($add)
    $product = $document.CreateElement('Product')
    $product.SetAttribute('ID', $target.TargetProductId)
    [void]$add.AppendChild($product)
    foreach ($locale in $target.Language) {
      $node = $document.CreateElement('Language')
      $node.SetAttribute('ID', $locale)
      [void]$product.AppendChild($node)
    }
    foreach ($app in $target.ExcludeApp) {
      $node = $document.CreateElement('ExcludeApp')
      $node.SetAttribute('ID', $app)
      [void]$product.AppendChild($node)
    }
    if ($Action -eq 'Migrate' -and $RemoveMsi) {
      [void]$root.AppendChild($document.CreateElement('RemoveMSI'))
    }
  }
  if ($Action -in @('Install', 'Migrate') -and $target.TargetProductId -like '*Volume') {
    $activation = $document.CreateElement('Property')
    $activation.SetAttribute('Name', 'AUTOACTIVATE')
    $activation.SetAttribute('Value', '1')
    [void]$root.AppendChild($activation)
  }
  if ($Action -notin @('Download', 'SetApplicationPreference')) {
    $display = $document.CreateElement('Display')
    $display.SetAttribute('Level', 'None')
    $display.SetAttribute('AcceptEULA', 'TRUE')
    [void]$root.AppendChild($display)
    $property = $document.CreateElement('Property')
    $property.SetAttribute('Name', 'FORCEAPPSHUTDOWN')
    $property.SetAttribute('Value', 'FALSE')
    [void]$root.AppendChild($property)
  }
  return , $document
}

function Invoke-PSFOfficeConfiguration {
  [CmdletBinding()]
  param (
    [string]
    $OdtPath,

    [xml]
    $Document,

    [string]
    $Directory,

    [ValidateSet('/download', '/configure', '/customize')]
    [string]
    $Mode = '/configure',

    [Security.SecureString]
    $ProductKey
  )

  Assert-PSFOfficeProtectedPath $Directory
  $path = Join-Path $Directory ('configuration-' + [guid]::NewGuid().ToString('N') + '.xml')
  $failure = $null
  $native = $null
  try {
    if ($ProductKey) {
      if (-not $Document.Configuration.Add.Product -or $Document.Configuration.Add.Product.ID -notlike '*Volume') {
        Stop-PSFOfficeOperation InvalidAuthority 'A product key requires a volume installation configuration.'
      }
      $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ProductKey)
      try {
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer).Trim().ToUpperInvariant()
        if ($plain -notmatch '^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$') {
          Stop-PSFOfficeOperation InvalidProductKey $script:ProductKeyFormatMessage
        }
        $Document.Configuration.Add.Product.SetAttribute('PIDKEY', $plain)
      }
      finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
        $plain = $null
      }
    }
    $Document.Save($path)
    $native = Invoke-PSFOfficeTool -OdtPath $OdtPath -Mode $Mode -ConfigurationPath $path
  }
  catch {
    $failure = $_
  }
  finally {
    if ($ProductKey -and $Document.Configuration.Add.Product) {
      $Document.Configuration.Add.Product.RemoveAttribute('PIDKEY')
    }
    if (Test-Path -LiteralPath $path) {
      try {
        Remove-Item -LiteralPath $path -Force -ErrorAction Stop
      }
      catch {
        if ($failure) {
          $failure.Exception.Data['OfficeCleanupError'] = 'Temporary configuration cleanup failed; inspect the protected working directory.'
        }
        else {
          $failure = $_
          $failure.Exception.Data['OfficeReason'] = 'ConfigurationCleanupFailed'
        }
      }
    }
  }
  if ($failure) {
    if ($native) {
      $failure.Exception.Data['OfficeExitCode'] = $native.ExitCode
    }
    throw $failure
  }
  $native
}

function Save-OfficeDeploymentMedia {
  <#
    .SYNOPSIS
      Prepares a verified Office package and publishes its manifest last.
    .DESCRIPTION
      Downloads only during explicit preparation. Existing matching packages are
      verified and reused. Incomplete or incompatible packages are never overwritten.
    .PARAMETER Configuration
      Requested products, ordered languages, channel, architecture, and optional build.
    .PARAMETER SourcePath
      New dedicated package directory; its parent must already exist.
    .PARAMETER OdtPath
      Existing Microsoft-signed Office Deployment Tool setup.exe.
    .PARAMETER DryRun
      Validate and preview without writes or downloads.
    .EXAMPLE
      Save-OfficeDeploymentMedia -Configuration $target -SourcePath C:\Media\Office -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Media names the deployment payload as a collective noun; Windows PowerShell and PowerShell 7 pluralizers disagree.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Configuration,

    [Parameter(Mandatory = $true)]
    [string]
    $SourcePath,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [switch]
    $DryRun
  )

  $target = ConvertTo-PSFOfficeConfiguration $Configuration
  Assert-PSFOfficePath $SourcePath
  if (Test-Path -LiteralPath $SourcePath) {
    $existing = Test-OfficeDeploymentMedia -SourcePath $SourcePath -Configuration $target
    if (-not $existing.Valid) {
      Stop-PSFOfficeOperation InvalidMedia 'Existing package is incomplete or incompatible; use a new destination.'
    }
    return $existing
  }
  $tool = Test-OfficeDeploymentTool $OdtPath
  if (-not $tool.Valid) {
    Stop-PSFOfficeOperation UntrustedTool 'ODT verification failed.'
  }
  if ($DryRun -or -not $PSCmdlet.ShouldProcess($SourcePath, 'Download pinned Office media and publish verified package')) {
    return [PSCustomObject]@{
      Status        = 'Preview'
      Path          = $SourcePath
      Configuration = $target
    }
  }
  $parent = Split-Path $SourcePath -Parent
  $work = Join-Path $parent ('PSFOfficePrepare-' + [guid]::NewGuid().ToString('N'))
  $published = $false
  try {
    New-PSFOfficeProtectedDirectory $work
    $setup = Join-Path $work 'setup.exe'
    Copy-Item -LiteralPath $OdtPath -Destination $setup -ErrorAction Stop
    $document = New-PSFOfficeXml -Action Download -Configuration $target -MediaPath $work
    $native = Invoke-PSFOfficeConfiguration -OdtPath $setup -Document $document -Directory $work -Mode /download
    if ($native.ExitCode -ne 0) {
      Stop-PSFOfficeOperation DownloadFailed "ODT download returned $($native.ExitCode)."
    }
    $files = @(Get-PSFOfficeMediaFile $work | Sort-Object Path)
    $builds = @(Get-ChildItem -LiteralPath (Join-Path $work 'Office\Data') -Directory -ErrorAction Stop | Where-Object { $_.Name -match '^16\.0\.\d+\.\d+$' })
    if ($builds.Count -ne 1 -or ($target.Version -and $target.Version -ne $builds[0].Name)) {
      Stop-PSFOfficeOperation InvalidMedia 'Preparation did not produce exactly the selected build.'
    }
    $manifest = [ordered]@{
      SchemaVersion      = 2
      Product            = $target.TargetProductId
      Architecture       = $target.Architecture
      Channel            = $target.Channel
      AvailableLanguages = @($target.Language)
      Version            = $builds[0].Name
      ToolVersion        = $tool.Version
      Files              = $files
    }
    $manifestPath = Join-Path $work 'psfoundation-office-media.json'
    Write-PSFOfficeJson -Path $manifestPath -Value $manifest
    $assessment = Test-OfficeDeploymentMedia -SourcePath $work -Configuration $target
    if (-not $assessment.Valid) {
      Stop-PSFOfficeOperation $assessment.ReasonCode $assessment.Error
    }
    Remove-Item -LiteralPath $setup -Force -ErrorAction Stop
    # Same-parent directory rename publishes a fully verified package, never a partial download.
    Assert-PSFOfficePath $SourcePath
    Assert-PSFOfficeProtectedPath $work
    [IO.Directory]::Move($work, $SourcePath)
    $published = $true
    Test-OfficeDeploymentMedia -SourcePath $SourcePath -Configuration $target
  }
  finally {
    if (-not $published) {
      Remove-PSFOfficeWorkDirectory -Path $work -Parent $parent
    }
  }
}

function Write-PSFOfficeJson {
  [CmdletBinding()]
  param (
    [string]
    $Path,

    [object]
    $Value
  )

  Assert-PSFOfficeProtectedPath (Split-Path $Path -Parent)
  Assert-PSFOfficePath $Path
  $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
  try {
    $json = ConvertTo-Json -InputObject $Value -Depth 30
    $encoding = New-Object Text.UTF8Encoding($false)
    $bytes = $encoding.GetBytes($json)
    $stream = New-Object IO.FileStream($temporary, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None, 4096, [IO.FileOptions]::WriteThrough)
    try {
      $stream.Write($bytes, 0, $bytes.Length)
      $stream.Flush($true)
    }
    finally {
      $stream.Dispose()
    }
    if (Test-Path -LiteralPath $Path) {
      [IO.File]::Replace($temporary, $Path, [NullString]::Value)
    }
    else {
      [IO.File]::Move($temporary, $Path)
    }
  }
  finally {
    if (Test-Path -LiteralPath $temporary) {
      Remove-Item -LiteralPath $temporary -Force -ErrorAction Stop
    }
  }
}

function Get-PSFOfficeActivity {
  [CmdletBinding()]
  param ()

  $processes = @(Get-CimInstance Win32_Process -ErrorAction Stop)
  # A running ClickToRun service alone is normal. Active clients and uncertain
  # setup/msiexec activity block; no unrelated installer is ever terminated.
  $busy = @($processes | Where-Object {
      $_.Name -match '^(setup|msiexec|OfficeC2RClient|integratedoffice)\.exe$'
    })
  $apps = @(Get-Process -Name WINWORD, EXCEL, POWERPNT, OUTLOOK, ONENOTE, ONENOTEM, MSPUB, MSACCESS, lync, VISIO, WINPROJ, INFOPATH, GROOVE, MSOHTMED, communicator -ErrorAction SilentlyContinue)
  [PSCustomObject]@{
    Busy = ($busy.Count -gt 0)
    Apps = $apps
  }
}

function Test-PSFOfficeHost {
  [CmdletBinding()]
  [OutputType([bool])]
  param ()

  $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
  $processor = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
  ($os.ProductType -eq 1 -and [int]$os.BuildNumber -ge 19045 -and $processor.Architecture -eq 9)
}

function Assert-PSFOfficeHost {
  [CmdletBinding()]
  param ()

  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = New-Object Security.Principal.WindowsPrincipal($identity)
  if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Stop-PSFOfficeOperation ElevationRequired 'Office deployment requires an elevated Windows identity.'
  }
  if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    Stop-PSFOfficeOperation Unsupported 'Use 64-bit PowerShell on 64-bit Windows.'
  }
  # Ordinary execution accepts x64 Windows 10 22H2 and later desktop hosts.
  if (-not (Test-PSFOfficeHost)) {
    Stop-PSFOfficeOperation Unsupported 'This backend requires x64 Windows 10 22H2 (build 19045) or later desktop hosts.'
  }
  if ((Test-PendingReboot).PendingReboot) {
    Stop-PSFOfficeOperation RebootRequired 'A pending reboot blocks Office mutation.'
  }
  if ((Get-PSFOfficeActivity).Busy) {
    Stop-PSFOfficeOperation DeploymentBusy 'Native deployment activity is active or uncertain.'
  }
}

function Enter-PSFOfficeLock {
  [CmdletBinding()]
  param ()

  $mutex = New-Object Threading.Mutex($false, 'Global\PSFoundation.OfficeDeployment')
  $acquired = $false
  try {
    try {
      $acquired = $mutex.WaitOne(0)
    }
    catch [Threading.AbandonedMutexException] {
      $acquired = $true
    }
    if (-not $acquired) {
      Stop-PSFOfficeOperation DeploymentBusy 'Another PSFoundation Office operation holds the deployment lock.'
    }
    return $mutex
  }
  catch {
    $mutex.Dispose()
    throw
  }
}

function Get-PSFOfficeExecutionContext {
  [CmdletBinding()]
  param ()

  $module = $ExecutionContext.SessionState.Module
  [PSCustomObject]@{
    ModuleVersion     = if ($module -and $module.Name -eq 'PSFoundation') { [string]$module.Version } else { $null }
    ModulePath        = if ($module -and $module.Name -eq 'PSFoundation') { $module.Path } else { $PSCommandPath }
    PowerShellVersion = [string]$PSVersionTable.PSVersion
    PowerShellEdition = [string]$PSVersionTable.PSEdition
    ProcessBitness    = [IntPtr]::Size * 8
  }
}

function New-PSFOfficeResult {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'In-memory result factory.')]
  [CmdletBinding()]
  param (
    [object]
    $Plan,

    [string]
    $Status = 'Planned'
  )

  $metadata = @{
    SchemaVersion      = 1
    MachineId          = $Plan.MachineId
    Phase              = 'Preflight'
    ReasonCode         = $null
    Changed            = $false
    ChangeKnown        = $true
    AlreadyCompliant   = $false
    After              = $null
    Verification       = $null
    Activation         = $null
    Configuration      = $Plan.Configuration
    LanguageTransition = $Plan.LanguageTransition
    RebootRequired     = $false
    ExitCode           = $null
    NativeResults      = @()
    WrapperExitCode    = 0
    RecoveryRequired   = $false
    RecoveryPath       = $null
    LogPaths           = @()
    Residue            = @()
    CleanupErrors      = @()
    Error              = $null
    Diagnostic         = $null
    Execution          = Get-PSFOfficeExecutionContext
    Plan               = $Plan
  }
  New-OperationResult -Target $Plan.MachineId -Source Office -Action $Plan.Action -Status $Status -RunId ([guid]::NewGuid().ToString('N')) -Before $Plan.Before -Property $metadata
}

function Get-OfficeDeploymentRecovery {
  <#
    .SYNOPSIS
      Reads a protected Office recovery record without resuming it.
    .DESCRIPTION
      Checks schema, machine identity, and configuration fingerprint. Checkpoints
      describe observed progress and never authorize extra removal or automatic replay.
    .PARAMETER RunId
      Run identifier returned by an Office operation.
    .PARAMETER LogRoot
      Protected local recovery root. Defaults to ProgramData\PSFoundation-Office.
    .EXAMPLE
      Get-OfficeDeploymentRecovery -RunId $result.RunId
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-fA-F0-9]{32}$')]
    [string]
    $RunId,

    [string]
    $LogRoot = (Join-Path $env:ProgramData 'PSFoundation-Office')
  )

  $path = Join-Path $LogRoot ($RunId + '.json')
  Assert-PSFOfficeProtectedPath $LogRoot -ObjectKind RecoveryDirectory
  Assert-PSFOfficeProtectedPath $path -ObjectKind RecoveryJournal
  $stage = 'Read'
  try {
    $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8 -ErrorAction Stop
    $stage = 'Json'
    $record = $json | ConvertFrom-Json -ErrorAction Stop
    $stage = 'Schema'
    $fields = @(
      'SchemaVersion',
      'RunId',
      'MachineId',
      'Action',
      'Plan',
      'ConfigurationFingerprint',
      'MediaFingerprint',
      'Phase',
      'PhaseCompleted',
      'NativeResults',
      'RebootRequired',
      'CreatedAt',
      'UpdatedAt',
      'Result'
    )
    Assert-PSFOfficeField $record $fields $fields
    if ($record.SchemaVersion -notin @(1, 2)) {
      Stop-PSFOfficeOperation InvalidRecoveryRecord 'Recovery schema is unsupported.'
    }
    $stage = 'Identity'
    if ($record.RunId -ne $RunId -or $record.MachineId -ne (Get-PSFOfficeMachineId)) {
      Stop-PSFOfficeOperation InvalidRecoveryRecord 'Recovery record schema, run, or machine does not match.'
    }
    $stage = 'Schema'
    $planFields = @(
      'SchemaVersion',
      'Action',
      'MachineId',
      'Configuration',
      'SourcePath',
      'RemoveProductId',
      'RemoveMsi',
      'Language',
      'Settings',
      'Before',
      'InventoryFingerprint',
      'State',
      'Eligible',
      'Blockers',
      'Warnings',
      'LanguageTransition',
      'MediaFingerprint'
    )
    if ($record.SchemaVersion -eq 2) {
      $planFields += @('PilotMigration', 'PilotResources')
      if ($record.Action -ne 'Migrate' -or $record.Plan.PilotMigration -isnot [bool] -or -not $record.Plan.PilotMigration) {
        Stop-PSFOfficeOperation InvalidRecoveryRecord 'Schema-2 journals are restricted to migration pilots.'
      }
      $resourceFields = @('UiLanguages', 'PrimaryLanguage', 'ProofingLanguages', 'Provisioning', 'Verification')
      Assert-PSFOfficeField $record.Plan.PilotResources $resourceFields $resourceFields
    }
    Assert-PSFOfficeField $record.Plan $planFields $planFields
    $stage = 'Context'
    if ($record.Plan.SchemaVersion -ne $record.SchemaVersion -or $record.Plan.Action -ne $record.Action -or
      $record.Plan.MachineId -ne $record.MachineId -or $record.Plan.RemoveMsi -isnot [bool] -or
      $record.PhaseCompleted -isnot [bool] -or $record.RebootRequired -isnot [bool] -or
      $record.MediaFingerprint -ne $record.Plan.MediaFingerprint) {
      Stop-PSFOfficeOperation InvalidRecoveryRecord 'Recovery operation context is inconsistent.'
    }
    $stage = 'InventoryFingerprint'
    if ((Get-PSFOfficeFingerprint $record.Plan.Before) -ne $record.Plan.InventoryFingerprint) {
      Stop-PSFOfficeOperation InvalidRecoveryRecord 'Recorded inventory fingerprint differs.'
    }
    $stage = 'Authority'
    if ($record.Action -eq 'Install' -and ($record.Plan.RemoveMsi -or $record.Plan.RemoveProductId.Count)) {
      Stop-PSFOfficeOperation InvalidAuthority 'Install recovery cannot contain removal authority.'
    }
    $stage = 'Settings'
    Assert-PSFOfficeSetting $record.Action $record.Plan.Settings
    $stage = 'Configuration'
    if ($record.Plan.Configuration) {
      $target = ConvertTo-PSFOfficeConfiguration $record.Plan.Configuration
      $stage = 'ConfigurationFingerprint'
      if ((Get-PSFOfficeFingerprint $target) -ne $record.ConfigurationFingerprint) {
        Stop-PSFOfficeOperation InvalidRecoveryRecord 'Recorded target fingerprint differs.'
      }
    }
    [PSCustomObject]@{
      RunId   = $RunId
      LogRoot = $LogRoot
      Path    = $path
      Record  = $record
    }
  }
  catch {
    $reason = $_.Exception.Data['OfficeReason']
    $category = 'Validation'
    if (-not $reason) {
      $reason = 'InvalidRecoveryRecord'
      $category = if ($stage -in @('Read', 'Json')) { $stage } else { 'InternalValidation' }
    }
    $diagnostic = [PSCustomObject]@{
      Stage = $stage; Category = $category; ObjectKind = 'RecoveryJournal'; Path = $path
      ExceptionType = $_.Exception.GetType().FullName
      Function = 'Get-OfficeDeploymentRecovery'; ScriptPath = $PSCommandPath; Line = $_.InvocationInfo.ScriptLineNumber
    }
    Stop-PSFOfficeOperation $reason "Recovery journal validation failed at $stage for '$path' ($category); no operation was resumed." $diagnostic
  }
}

function Set-PSFOfficeCheckpoint {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Private journal writer inside an approved operation.')]
  [CmdletBinding()]
  param (
    [object]
    $Context,

    [string]
    $Phase,

    [bool]
    $Completed = $false
  )

  $Context.Result.Phase = $Phase
  $Context.Journal.Phase = $Phase
  $Context.Journal.PhaseCompleted = $Completed
  $Context.Journal.UpdatedAt = [datetime]::UtcNow.ToString('o')
  $Context.Journal.NativeResults = @($Context.Result.NativeResults)
  $Context.Journal.RebootRequired = $Context.Result.RebootRequired
  Write-PSFOfficeJson -Path $Context.Result.RecoveryPath -Value $Context.Journal
}

function Invoke-PSFOfficePhase {
  [CmdletBinding()]
  param (
    [object]
    $Context,

    [string]
    $Action,

    [Security.SecureString]
    $ProductKey
  )

  $plan = $Context.Plan
  Set-PSFOfficeCheckpoint $Context $Action
  $parameters = @{
    Action        = $Action
    Configuration = $plan.Configuration
    MediaPath     = $Context.MediaPath
    Settings      = $plan.Settings
    Language      = @($plan.Language)
  }
  if ($Action -in @('Remove', 'Migrate')) {
    $parameters.RemoveProductId = @($plan.RemoveProductId)
  }
  if ($Action -eq 'Migrate') {
    $parameters.RemoveMsi = $plan.RemoveMsi
  }
  $document = New-PSFOfficeXml @parameters
  $mode = '/configure'
  if ($Action -eq 'SetApplicationPreference') {
    $mode = '/customize'
  }
  # From this point an error is not evidence of an unchanged machine.
  $Context.Result.Changed = $null
  $Context.Result.ChangeKnown = $false
  $Context.Result.RecoveryRequired = $true
  $native = Invoke-PSFOfficeConfiguration -OdtPath $Context.Setup -Document $document -Directory $Context.Work -Mode $mode -ProductKey $ProductKey
  # Do not copy arbitrary native stdout/stderr into durable logs; ODT may echo data.
  $Context.Result.NativeResults += [PSCustomObject]@{
    Phase    = $Action
    ExitCode = $native.ExitCode
  }
  $Context.Result.ExitCode = $native.ExitCode
  $Context.Result.RebootRequired = ($Context.Result.RebootRequired -or $native.ExitCode -eq 3010)
  Set-PSFOfficeCheckpoint $Context $Action $true
  if ($native.ExitCode -notin @(0, 3010)) {
    Stop-PSFOfficeOperation NativeFailure "ODT $Action returned $($native.ExitCode); current machine state requires inspection."
  }
}

function Invoke-PSFOfficeMigration {
  [CmdletBinding()]
  param (
    [object]
    $Context,

    [Security.SecureString]
    $ProductKey
  )

  if ($Context.Plan.RemoveProductId.Count) {
    Invoke-PSFOfficePhase -Context $Context -Action Remove
    if ($Context.Result.RebootRequired) {
      Stop-PSFOfficeOperation RebootRequired 'Removal requested a reboot. Installation was not started.'
    }
    $afterRemoval = Get-OfficeInventory
    if ($afterRemoval.Products.Count -or $afterRemoval.Unknowns.Count) {
      Stop-PSFOfficeOperation VerificationFailed 'Click-to-Run removal was not verified; installation was not started.'
    }
  }
  Invoke-PSFOfficePhase -Context $Context -Action Migrate -ProductKey $ProductKey
}

function Test-PSFOfficePostcondition {
  [CmdletBinding()]
  param (
    [object]
    $Plan,

    [object]
    $After
  )

  # Known add-ins are outside Office removal authority. Their registrations must
  # survive every workflow; this does not claim application/add-in compatibility.
  $changedAddIns = @($Plan.Before.RelatedComponents | Where-Object { $_.Role -eq 'AddIn' } | Where-Object {
      $beforeComponent = $_
      $afterComponent = @($After.RelatedComponents | Where-Object { $_.ProductCode -eq $beforeComponent.ProductCode -and $_.RegistryView -eq $beforeComponent.RegistryView })
      $afterComponent.Count -ne 1 -or (Get-PSFOfficeFingerprint $beforeComponent) -ne (Get-PSFOfficeFingerprint $afterComponent[0])
    })
  if ($changedAddIns.Count) {
    return [PSCustomObject]@{
      Compliant     = $false
      Discrepancies = @('RelatedAddInChanged')
      Unknowns      = @()
    }
  }
  if ($Plan.Action -eq 'Remove') {
    $remaining = @($After.Products | ForEach-Object { $_.ProductId })
    $retained = @($Plan.Before.Products | Where-Object { $_.ProductId -notin $Plan.RemoveProductId })
    $missing = @($retained | Where-Object { $_.ProductId -notin $remaining })
    $unexpected = @($After.Products | Where-Object { $_.ProductId -notin @($Plan.Before.Products | ForEach-Object { $_.ProductId }) })
    $changed = @($retained | Where-Object {
        $beforeProduct = $_
        $afterProduct = @($After.Products | Where-Object { $_.ProductId -eq $beforeProduct.ProductId })
        $afterProduct.Count -ne 1 -or (Get-PSFOfficeFingerprint $beforeProduct) -ne (Get-PSFOfficeFingerprint $afterProduct[0])
      })
    $valid = (-not @($Plan.RemoveProductId | Where-Object { $_ -in $remaining }).Count -and
      -not $missing.Count -and -not $changed.Count -and -not $unexpected.Count -and -not $After.Unknowns.Count -and
      (Get-PSFOfficeFingerprint @($Plan.Before.Msi)) -eq (Get-PSFOfficeFingerprint @($After.Msi)))
    return [PSCustomObject]@{
      Compliant     = $valid
      Discrepancies = @($missing | ForEach-Object { $_.ProductId }) + @($changed | ForEach-Object { $_.ProductId }) +
      @($unexpected | ForEach-Object { 'UnexpectedProduct:' + $_.ProductId })
      Unknowns      = @($After.Unknowns)
    }
  }
  if ($Plan.Action -in @('SetUpdateConfiguration', 'SetApplicationPreference')) {
    # Native success proves request submission, not effective managed policy or
    # preferences for all current/future users. Keep that distinction explicit.
    return [PSCustomObject]@{
      Compliant     = $false
      Discrepancies = @()
      Unknowns      = @('EffectiveSettings')
    }
  }
  $verification = Test-OfficeDeployment -Configuration $Plan.Configuration -Inventory $After
  $verification
}

function Invoke-PSFOfficeWorkflow {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Calls ShouldProcess on the public command PSCmdlet so its WhatIf and Confirm settings govern the operation.')]
  [CmdletBinding()]
  param (
    [object]
    $Plan,

    [string]
    $ExpectedAction,

    [System.Management.Automation.PSCmdlet]
    $Caller,

    [string]
    $OdtPath,

    [string]
    $LogRoot,

    [bool]
    $DryRun,

    [bool]
    $ForceCloseApps,

    [Security.SecureString]
    $ProductKey
  )

  $fresh = Confirm-PSFOfficePlan -Plan $Plan -Action $ExpectedAction
  $result = New-PSFOfficeResult $fresh
  foreach ($warning in $fresh.Warnings) {
    Write-Warning $warning
  }
  if (-not $fresh.Eligible) {
    $result.Status = 'Blocked'
    $result.ReasonCode = $fresh.Blockers[0]
    $result.WrapperExitCode = 1
    return $result
  }
  if ($ProductKey -and ($ExpectedAction -notin @('Install', 'Migrate') -or $fresh.Configuration.TargetProductId -notlike '*Volume')) {
    Stop-PSFOfficeOperation InvalidAuthority 'ProductKey is restricted to volume installation and migration.'
  }
  # Validate before staging/removal; never include the key in a plan or error.
  if ($ProductKey) {
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ProductKey)
    try {
      if ([Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer).Trim() -notmatch '^([A-Za-z0-9]{5}-){4}[A-Za-z0-9]{5}$') {
        Stop-PSFOfficeOperation InvalidProductKey $script:ProductKeyFormatMessage
      }
    }
    finally {
      [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
  }
  $absent = $ExpectedAction -eq 'Remove' -and -not @($fresh.RemoveProductId | Where-Object { $_ -in @($fresh.Before.Products | ForEach-Object { $_.ProductId }) }).Count
  if ($absent -or ($fresh.State -eq 'Compliant' -and $ExpectedAction -in @(
        'Install',
        'Migrate',
        'Update',
        'AddLanguage',
        'RemoveLanguage',
        'SetApplicationSelection'
      ))) {
    $result.Status = 'Completed'
    $result.AlreadyCompliant = -not $absent
    $result.ReasonCode = 'AlreadyCompliant'
    if ($absent) {
      $result.ReasonCode = 'AlreadyAbsent'
    }
    $result.After = $fresh.Before
    if ($fresh.Configuration) {
      $result.Verification = Test-OfficeDeployment -Configuration $fresh.Configuration -Inventory $fresh.Before
      $result.Activation = Get-OfficeActivationStatus $fresh.Configuration.TargetProductId
      if ($result.Activation.Status -eq 'NotVerified') {
        $result.ReasonCode = 'ActivationNotVerified'
        $result.WrapperExitCode = 1
      }
    }
    return $result
  }
  try {
    Assert-PSFOfficeHost
    if (-not (Test-OfficeDeploymentTool $OdtPath).Valid) {
      Stop-PSFOfficeOperation UntrustedTool 'ODT verification failed.'
    }
    if ($LogRoot -like '\\*') {
      Stop-PSFOfficeOperation UnsafePath 'Recovery records must be stored locally.'
    }
    Assert-PSFOfficePath $LogRoot
    $activity = Get-PSFOfficeActivity
    if ($activity.Apps.Count -and -not $ForceCloseApps) {
      Stop-PSFOfficeOperation ApplicationsRunning 'Close Office applications in all sessions or authorize ForceCloseApps.'
    }
  }
  catch {
    $result.Status = 'Blocked'
    $result.ReasonCode = $_.Exception.Data['OfficeReason']
    if (-not $result.ReasonCode) {
      $result.ReasonCode = 'PreflightFailed'
    }
    $result.Error = $_.Exception.Message
    $result.Diagnostic = $_.Exception.Data['OfficeDiagnostic']
    $result.WrapperExitCode = 1
    return $result
  }
  $actionText = "$ExpectedAction; remove [$($fresh.RemoveProductId -join ',')]; ALL supported MSI: $($fresh.RemoveMsi); force-close apps: $ForceCloseApps"
  if ($fresh.Configuration) {
    $actionText += "; target $($fresh.Configuration.TargetProductId) $($fresh.Configuration.Version); languages [$($fresh.Configuration.Language -join ',')]"
  }
  if ($DryRun -or -not $Caller.ShouldProcess($fresh.MachineId, $actionText)) {
    $result.Status = 'Preview'
    $result.ReasonCode = 'NotExecuted'
    return $result
  }
  $lock = $null
  $context = $null
  $work = $null
  try {
    Assert-PSFOfficeHost
    if (-not (Test-OfficeDeploymentTool $OdtPath).Valid) {
      Stop-PSFOfficeOperation UntrustedTool 'ODT verification failed.'
    }
    if ($LogRoot -like '\\*') {
      Stop-PSFOfficeOperation UnsafePath 'Recovery records must be stored locally.'
    }
    Assert-PSFOfficePath $LogRoot
    $lock = Enter-PSFOfficeLock
    $fresh = Confirm-PSFOfficePlan -Plan $fresh -Action $ExpectedAction
    if (-not $fresh.Eligible) {
      Stop-PSFOfficeOperation StalePlan 'Preconditions changed before lock acquisition.'
    }
    New-PSFOfficeProtectedDirectory $LogRoot
    $work = Join-Path $LogRoot ('Work-' + $result.RunId)
    New-PSFOfficeProtectedDirectory $work
    $result.RecoveryPath = Join-Path $LogRoot ($result.RunId + '.json')
    $journal = [PSCustomObject][ordered]@{
      SchemaVersion            = $fresh.SchemaVersion
      RunId                    = $result.RunId
      MachineId                = $fresh.MachineId
      Action                   = $ExpectedAction
      Plan                     = $fresh
      ConfigurationFingerprint = Get-PSFOfficeFingerprint $fresh.Configuration
      MediaFingerprint         = $fresh.MediaFingerprint
      Phase                    = 'StageMedia'
      PhaseCompleted           = $false
      NativeResults            = @()
      RebootRequired           = $false
      CreatedAt                = [datetime]::UtcNow.ToString('o')
      UpdatedAt                = [datetime]::UtcNow.ToString('o')
      Result                   = $null
    }
    $context = [PSCustomObject]@{
      Plan      = $fresh
      Result    = $result
      Journal   = $journal
      Work      = $work
      Setup     = Join-Path $work 'setup.exe'
      MediaPath = $null
    }
    Set-PSFOfficeCheckpoint $context StageMedia
    Copy-Item -LiteralPath $OdtPath -Destination $context.Setup -ErrorAction Stop
    if (-not (Test-OfficeDeploymentTool $context.Setup).Valid) {
      Stop-PSFOfficeOperation UntrustedTool 'Staged ODT verification failed.'
    }
    if ($fresh.SourcePath) {
      $media = Test-OfficeDeploymentMedia -SourcePath $fresh.SourcePath -Configuration $fresh.Configuration
      if (-not $media.Valid -or $media.Fingerprint -ne $fresh.MediaFingerprint) {
        Stop-PSFOfficeOperation StaleMedia 'Source media changed.'
      }
      $bytes = ($media.Manifest.Files | Measure-Object Length -Sum).Sum
      $drive = New-Object IO.DriveInfo([IO.Path]::GetPathRoot($work))
      if ($drive.AvailableFreeSpace -lt (2 * $bytes + 4GB)) {
        Stop-PSFOfficeOperation InsufficientSpace 'Insufficient capacity for staged media and installation.'
      }
      $context.MediaPath = Join-Path $work 'Media'
      New-PSFOfficeProtectedDirectory $context.MediaPath
      Copy-Item -LiteralPath (Join-Path $fresh.SourcePath 'Office') -Destination $context.MediaPath -Recurse -ErrorAction Stop
      Copy-Item -LiteralPath (Join-Path $fresh.SourcePath 'psfoundation-office-media.json') -Destination $context.MediaPath -ErrorAction Stop
      $staged = Test-OfficeDeploymentMedia -SourcePath $context.MediaPath -Configuration $fresh.Configuration
      if (-not $staged.Valid -or $staged.Fingerprint -ne $fresh.MediaFingerprint) {
        Stop-PSFOfficeOperation MediaIntegrityFailed 'Staged media verification failed.'
      }
    }
    Set-PSFOfficeCheckpoint $context StageMedia $true
    $null = Confirm-PSFOfficePlan -Plan $fresh -Action $ExpectedAction
    Assert-PSFOfficeHost
    $activity = Get-PSFOfficeActivity
    if ($activity.Apps.Count) {
      if (-not $ForceCloseApps) {
        Stop-PSFOfficeOperation ApplicationsRunning 'Close Office applications in all sessions or explicitly authorize ForceCloseApps.'
      }
      Set-PSFOfficeCheckpoint $context CloseApplications
      $activity.Apps | Stop-Process -Force -ErrorAction Stop
      $result.Changed = $true
      if ((Get-PSFOfficeActivity).Apps.Count) {
        Stop-PSFOfficeOperation ApplicationsRunning 'Office applications are still running.'
      }
      Set-PSFOfficeCheckpoint $context CloseApplications $true
    }
    if ($ExpectedAction -eq 'Migrate') {
      Invoke-PSFOfficeMigration -Context $context -ProductKey $ProductKey
    }
    else {
      Invoke-PSFOfficePhase -Context $context -Action $ExpectedAction -ProductKey $ProductKey
    }
    Set-PSFOfficeCheckpoint $context Verify
    $result.After = Get-OfficeInventory
    $result.Verification = Test-PSFOfficePostcondition -Plan $fresh -After $result.After
    if (-not $result.Verification.Compliant) {
      if ($ExpectedAction -in @('SetUpdateConfiguration', 'SetApplicationPreference')) {
        $result.Status = 'AppliedUnverified'
        $result.ReasonCode = 'EffectiveSettingsUnknown'
        $result.WrapperExitCode = 1
      }
      else {
        Stop-PSFOfficeOperation VerificationFailed 'ODT finished but deployment postconditions were not fully verified.'
      }
    }
    else {
      $result.Status = 'Completed'
      $result.Changed = $true
      $result.ChangeKnown = $true
      $result.RecoveryRequired = $false
      if ($fresh.Configuration) {
        $result.Activation = Get-OfficeActivationStatus $fresh.Configuration.TargetProductId
        if ($result.Activation.Status -eq 'NotVerified') {
          $result.ReasonCode = 'ActivationNotVerified'
          $result.WrapperExitCode = 1
        }
      }
    }
    Set-PSFOfficeCheckpoint $context Verify $true
  }
  catch {
    $result.Status = 'Failed'
    $result.ReasonCode = $_.Exception.Data['OfficeReason']
    if (-not $result.ReasonCode) {
      $result.ReasonCode = 'OperationFailed'
    }
    $result.Error = $_.Exception.Message
    $result.Diagnostic = $_.Exception.Data['OfficeDiagnostic']
    if ($_.Exception.Data.Contains('OfficeExitCode')) {
      $result.ExitCode = $_.Exception.Data['OfficeExitCode']
      $result.NativeResults += [PSCustomObject]@{
        Phase    = $result.Phase
        ExitCode = $result.ExitCode
      }
      $result.RebootRequired = ($result.RebootRequired -or $result.ExitCode -eq 3010)
    }
    if ($_.Exception.Data.Contains('OfficeCleanupError')) {
      $result.CleanupErrors += $_.Exception.Data['OfficeCleanupError']
    }
    $result.WrapperExitCode = 1
    if ($result.ReasonCode -eq 'RebootRequired') {
      $result.RebootRequired = $true
    }
    try {
      $result.After = Get-OfficeInventory
    }
    catch {
      $result.After = $null
    }
  }
  finally {
    try {
      if ($work) {
        Remove-PSFOfficeWorkDirectory -Path $work -Parent $LogRoot
      }
    }
    catch {
      $result.CleanupErrors += $_.Exception.Message
      $result.Residue += $work
      $result.Status = 'Failed'
      $result.WrapperExitCode = 1
    }
    if ($result.RebootRequired -and $result.WrapperExitCode -eq 0) {
      $result.WrapperExitCode = 3010
    }
    if ($context) {
      try {
        $log = Join-Path $LogRoot ($result.RunId + '.jsonl')
        $result.LogPaths = @($log)
        $context.Journal.Result = $result
        Set-PSFOfficeCheckpoint $context $result.Phase $context.Journal.PhaseCompleted
        $null = Write-OperationResultLog -Results @($result) -Path $log -ScriptName $ExpectedAction -RunId $result.RunId
      }
      catch {
        $result.CleanupErrors += $_.Exception.Message
        $result.Status = 'Failed'
        $result.WrapperExitCode = 1
      }
    }
    if ($lock) {
      $lock.ReleaseMutex()
      $lock.Dispose()
    }
  }
  $result
}

function Invoke-PSFOfficeRecovery {
  [CmdletBinding()]
  param (
    [object]
    $Recovery,

    [string]
    $OriginalAction,

    [System.Management.Automation.PSCmdlet]
    $Caller,

    [string]
    $OdtPath,

    [bool]
    $DryRun,

    [bool]
    $ForceCloseApps,

    [Security.SecureString]
    $ProductKey
  )

  Assert-PSFOfficeField $Recovery @('RunId', 'LogRoot', 'Path', 'Record') @('RunId', 'LogRoot')
  # Reopen protected on-disk state. Caller-supplied Record is never authoritative.
  $loaded = Get-OfficeDeploymentRecovery -RunId $Recovery.RunId -LogRoot $Recovery.LogRoot
  $record = $loaded.Record
  if ($record.SchemaVersion -eq 2) {
    Stop-PSFOfficeOperation UnsupportedPilotRecovery 'Pilot journals are evidence only. Review the result or restore the VM; automatic replay is not supported.'
  }
  if ($record.Action -ne $OriginalAction -or $record.Plan.Action -ne $OriginalAction) {
    Stop-PSFOfficeOperation InvalidAuthority 'This recovery command cannot resume the recorded operation.'
  }
  $target = ConvertTo-PSFOfficeConfiguration $record.Plan.Configuration
  $inventory = Get-OfficeInventory
  $verification = Test-OfficeDeployment -Configuration $target -Inventory $inventory
  $result = New-PSFOfficeResult -Plan $record.Plan
  $result.Action = 'Recover'
  $result.RecoveryPath = $loaded.Path
  $result.After = $inventory
  $result.Verification = $verification
  if ((Get-PSFOfficeActivity).Busy) {
    $result.Status = 'Blocked'
    $result.ReasonCode = 'DeploymentBusy'
    $result.WrapperExitCode = 1
    return $result
  }
  $result.RebootRequired = (Test-PendingReboot).PendingReboot
  if ($result.RebootRequired) {
    $result.Status = 'Blocked'
    $result.ReasonCode = 'RebootRequired'
    $result.WrapperExitCode = 3010
    return $result
  }
  $media = Test-OfficeDeploymentMedia -SourcePath $record.Plan.SourcePath -Configuration $target
  if (-not $media.Valid -or $media.Fingerprint -ne $record.MediaFingerprint) {
    Stop-PSFOfficeOperation StaleMedia 'Recovery media no longer matches the recorded deployment.'
  }
  if ($verification.Compliant) {
    $result.Status = 'Completed'
    $result.Phase = 'Verify'
    $result.ReasonCode = 'AlreadyCompliant'
    $result.AlreadyCompliant = $true
    $result.Activation = Get-OfficeActivationStatus $target.TargetProductId
    if ($result.Activation.Status -eq 'NotVerified') {
      $result.ReasonCode = 'ActivationNotVerified'
      $result.WrapperExitCode = 1
    }
    return $result
  }
  $currentIds = @($inventory.Products | ForEach-Object { $_.ProductId })
  $approvedIds = @($record.Plan.Before.Products | ForEach-Object { $_.ProductId })
  if ($inventory.Unknowns.Count -or @($currentIds | Where-Object { $_ -notin $approvedIds -and $_ -ne $target.TargetProductId }).Count -or
    @($inventory.Msi | Where-Object { $_.ProductCode -notin @($record.Plan.Before.Msi | ForEach-Object { $_.ProductCode }) }).Count) {
    $result.Status = 'Blocked'
    $result.ReasonCode = 'Conflict'
    $result.WrapperExitCode = 1
    return $result
  }
  # Only pre-launch continuation or verified completed C2R removal is currently
  # supported. Reapplying ODT to an uncertain partial installation needs VM evidence.
  $beforeLaunch = $record.Phase -in @('StageMedia', 'CloseApplications')
  $afterRemoval = $OriginalAction -eq 'Migrate' -and $record.Phase -eq 'Remove' -and $record.PhaseCompleted -and -not $inventory.Products.Count
  if (-not ($beforeLaunch -or $afterRemoval)) {
    $result.Status = 'Blocked'
    $result.ReasonCode = 'UnsupportedRecoveryState'
    $result.Error = 'The recorded checkpoint and current inventory do not prove a safe continuation. Review the journal and fresh inventory; do not replay an uncertain installer.'
    $result.RecoveryRequired = $true
    $result.WrapperExitCode = 1
    return $result
  }
  $parameters = @{
    Action        = $OriginalAction
    Configuration = $target
    SourcePath    = $record.Plan.SourcePath
    Inventory     = $inventory
  }
  if ($OriginalAction -eq 'Migrate') {
    $parameters.RemoveProductId = @($record.Plan.RemoveProductId | Where-Object { $_ -in $currentIds })
    $parameters.RemoveMsi = [bool]$record.Plan.RemoveMsi
  }
  $plan = Get-OfficeDeploymentPlan @parameters
  Invoke-PSFOfficeWorkflow -Plan $plan -ExpectedAction $OriginalAction -Caller $Caller -OdtPath $OdtPath -LogRoot $Recovery.LogRoot -DryRun $DryRun -ForceCloseApps $ForceCloseApps -ProductKey $ProductKey
}

function Resume-OfficeInstallation {
  <#
    .SYNOPSIS
      Verifies or continues the same interrupted Office installation.
    .DESCRIPTION
      Reopens a protected Install journal. Fully compliant targets are verified
      without reinstalling. Pre-launch work may continue. Uncertain partial ODT
      installations return UnsupportedRecoveryState pending disposable-VM validation.
      This command cannot inherit migration or removal authority.
    .PARAMETER Recovery
      Recovery descriptor from Get-OfficeDeploymentRecovery.
    .PARAMETER OdtPath
      Existing verified ODT setup.exe for supported continuation.
    .PARAMETER ForceCloseApps
      Explicitly authorize application closure during continuation.
    .PARAMETER ProductKey
      Fresh SecureString volume key when needed; never loaded from a journal.
    .PARAMETER DryRun
      Preview continuation without writes or process invocation.
    .EXAMPLE
      $recovery | Resume-OfficeInstallation -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Shared lifecycle calls the supplied PSCmdlet ShouldProcess before mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Recovery,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [switch]
    $ForceCloseApps,

    [Security.SecureString]
    $ProductKey,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Recovery       = $Recovery
      OriginalAction = 'Install'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
      ProductKey     = $ProductKey
    }
    Invoke-PSFOfficeRecovery @parameters
  }
}

function Resume-OfficeMigration {
  <#
    .SYNOPSIS
      Verifies or continues an explicitly authorized Office migration.
    .DESCRIPTION
      Reopens a protected Migrate journal, checks current products against its
      original scope, and confirms the revised outstanding plan. Never expands
      removal to newly discovered products or resumes across a pending reboot.
    .PARAMETER Recovery
      Recovery descriptor from Get-OfficeDeploymentRecovery.
    .PARAMETER OdtPath
      Existing verified ODT setup.exe for supported continuation.
    .PARAMETER ForceCloseApps
      Explicitly authorize application closure during continuation.
    .PARAMETER ProductKey
      Fresh SecureString volume key when needed; never persisted.
    .PARAMETER DryRun
      Preview continuation without writes or process invocation.
    .EXAMPLE
      $recovery | Resume-OfficeMigration -OdtPath C:\ODT\setup.exe -WhatIf
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Shared lifecycle calls the supplied PSCmdlet ShouldProcess before mutation.')]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [object]
    $Recovery,

    [Parameter(Mandatory = $true)]
    [string]
    $OdtPath,

    [switch]
    $ForceCloseApps,

    [Security.SecureString]
    $ProductKey,

    [switch]
    $DryRun
  )

  process {
    $parameters = @{
      Recovery       = $Recovery
      OriginalAction = 'Migrate'
      Caller         = $PSCmdlet
      OdtPath        = $OdtPath
      DryRun         = [bool]$DryRun
      ForceCloseApps = [bool]$ForceCloseApps
      ProductKey     = $ProductKey
    }
    Invoke-PSFOfficeRecovery @parameters
  }
}
