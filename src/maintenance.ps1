<#
  Module maintenance functions for PSFoundation.
  Provides Get-PSModule, Remove-PSModule, and Add-PSModule for inspecting,
  cleaning, and restoring PowerShell modules across PS5.1 and PS7+.
#>

#Requires -Version 5.0

# ---- Dot-source logging helper (needed when sourced standalone by tools) ----
$_logPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'src/log.ps1'
if (Test-Path -LiteralPath $_logPath) {
  . $_logPath
}

# ---- Helpers ---------------------------------------------------------------

function Resolve-ModuleDirectory {
  [CmdletBinding()]
  param(
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [string]$Path
  )

  if ($Path) { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path) }

  $isPS7 = $PSVersionTable.PSVersion.Major -ge 7

  if ($Scope -eq 'CurrentUser') {
    $docs = [Environment]::GetFolderPath('Personal')
    if ($isPS7) {
      return Join-Path -Path $docs -ChildPath 'PowerShell\Modules'
    }
    else {
      return Join-Path -Path $docs -ChildPath 'WindowsPowerShell\Modules'
    }
  }
  else {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
      Write-Warning 'AllUsers scope typically requires an elevated session.'
    }
    if ($isPS7) {
      return Join-Path -Path $env:ProgramFiles -ChildPath 'PowerShell\Modules'
    }
    else {
      return Join-Path -Path $env:ProgramFiles -ChildPath 'WindowsPowerShell\Modules'
    }
  }
}

function Test-PSResourceGetAvailable {
  $null -ne (Get-Module -Name 'Microsoft.PowerShell.PSResourceGet' -ListAvailable -ErrorAction SilentlyContinue)
}

function Ensure-PSResourceGet {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Ensure- is a descriptive verb for a private helper that guarantees PSResourceGet availability.')]
  param()
  if (-not (Test-PSResourceGetAvailable)) {
    Write-Log -Message 'Installing Microsoft.PowerShell.PSResourceGet...' -Color Yellow
    try {
      $null = Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue
      Install-Module -Name 'Microsoft.PowerShell.PSResourceGet' -Repository PSGallery -Force -Scope CurrentUser -ErrorAction Stop
    }
    catch {
      Write-Warning "Failed to install PSResourceGet: $_"
    }
  }
}

# ---- Get-PSModule ----------------------------------------------------------

function Get-PSModule {
  <#
    .SYNOPSIS
      Exports installed PowerShell modules to the pipeline or a JSON file.

    .DESCRIPTION
      Lists modules installed at the resolved scope path, filtered by name regex.
      Uses PSResourceGet on PS7+ whenever available, falling back to PowerShellGet
      on PS5.1.

    .PARAMETER Path
      JSON output file path. When omitted, modules are returned to the pipeline.

    .PARAMETER Name
      Regex filter for module name. Defaults to '*' (all).

    .PARAMETER Scope
      Module directory scope: CurrentUser (default) or AllUsers.

    .EXAMPLE
      Get-PSModule
      Returns all CurrentUser modules to the pipeline.

    .EXAMPLE
      Get-PSModule -Path ./modules.json -Name 'Pester'
      Exports matching modules to modules.json.
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletAliases', '', Justification = 'Get-PSResource is an accepted alias for Get-InstalledPSResource.')]

  [CmdletBinding()]
  param(
    [string]$Path,
    [string]$Name,
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser'
  )

  $isPS7 = $PSVersionTable.PSVersion.Major -ge 7
  $usePSResource = $false

  if ($isPS7) {
    Ensure-PSResourceGet
    $usePSResource = Test-PSResourceGetAvailable
  }

  $modulesPath = Resolve-ModuleDirectory -Scope $Scope
  $results = [System.Collections.Generic.List[PSCustomObject]]::new()

  if ($usePSResource) {
    try {
      $resources = Get-PSResource -Path $modulesPath -Scope $Scope -ErrorAction Stop
      if ($resources) {
        foreach ($r in $resources) {
          if ($Name -and $r.Name -notmatch $Name) { continue }

          $results.Add([PSCustomObject]@{
              Name = $r.Name
              Version = $r.Version.ToString()
              Repository = if ($r.Repository) { $r.Repository } else { 'Unknown' }
              Scope = $Scope
              InstalledDate = $r.InstalledDate
            })
        }
      }
    }
    catch {
      Write-Warning "Get-PSResource failed, falling back to Get-Module: $_"
      $usePSResource = $false
    }
  }

  if (-not $usePSResource) {
    if (Test-Path -LiteralPath $modulesPath -PathType Container) {
      $allModules = Get-Module -ListAvailable -ErrorAction SilentlyContinue |
        Where-Object {
          $_.ModuleType -ne 'Binary' -and (
            ($_.ModuleBase -like "$modulesPath\*") -or
            ($_.Path -like "$modulesPath\*")
          )
        }

      if ($allModules) {
        foreach ($m in $allModules) {
          if ($Name -and $m.Name -notmatch $Name) { continue }

          $installedDate = try {
            $item = Get-Item -LiteralPath $m.ModuleBase -ErrorAction SilentlyContinue
            $item.CreationTime
          }
          catch { $null }

          $results.Add([PSCustomObject]@{
              Name = $m.Name
              Version = $m.Version.ToString()
              Repository = if ($m.RepositorySourceLocation) { $m.RepositorySourceLocation } else { 'Unknown' }
              Scope = $Scope
              InstalledDate = $installedDate
            })
        }
      }
    }
  }

  $results = $results | Sort-Object Name, Version

  if ($Path) {
    $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $parent = Split-Path -Path $resolvedPath -Parent
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
      New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }
    $results | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $resolvedPath -Encoding UTF8
    Write-Log -Message "Exported $($results.Count) module(s) to $resolvedPath" -Color Green
  }
  else {
    $results
  }
}

# ---- Remove-PSModule -------------------------------------------------------

function Remove-PSModule {
  <#
    .SYNOPSIS
      Removes installed PowerShell module versions.

    .DESCRIPTION
      Removes old or all versions of modules from the resolved scope path.
      When -All is used, every matching module version is removed (requires
      confirmation unless -Force). Otherwise, only older versions beyond
      -LatestToKeep are removed.

      Uses PSResourceGet on PS7+ whenever available, falling back to
      PowerShellGet on PS5.1.

    .PARAMETER Name
      Regex filter for module name. Defaults to '*' (all modules).

    .PARAMETER LatestToKeep
      Number of newest versions to keep per module. Default: 1.

    .PARAMETER All
      Remove every version of every matching module. Overrides -LatestToKeep.
      Prompts for confirmation unless -Force is also supplied.

    .PARAMETER Scope
      Module directory scope: CurrentUser (default) or AllUsers.

    .PARAMETER Path
      Override the module directory path instead of resolving from -Scope.

    .PARAMETER Force
      Skip confirmation prompts.

    .EXAMPLE
      Remove-PSModule -WhatIf
      Preview which old module versions would be removed.

    .EXAMPLE
      Remove-PSModule -All -Force
      Remove all module versions from the CurrentUser scope without prompting.

    .EXAMPLE
      Remove-PSModule -Name 'Pester' -LatestToKeep 2
      Keep the two newest Pester versions, remove older ones.
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingCmdletAliases', '', Justification = 'Get-PSResource is an accepted alias for Get-InstalledPSResource.')]

  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [string]$Name,
    [ValidateRange(1, [int]::MaxValue)]
    [int]$LatestToKeep = 1,
    [switch]$All,
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [string]$Path,
    [switch]$Force
  )

  $isPS7 = $PSVersionTable.PSVersion.Major -ge 7
  $usePSResource = $false

  if ($isPS7) {
    Ensure-PSResourceGet
    $usePSResource = Test-PSResourceGetAvailable
  }

  $modulesPath = Resolve-ModuleDirectory -Scope $Scope -Path $Path

  # ---- Collect installed modules from the target path ----
  $allModules = [System.Collections.Generic.List[PSCustomObject]]::new()

  if ($usePSResource) {
    try {
      $resources = Get-PSResource -Path $modulesPath -Scope $Scope -ErrorAction Stop
    }
    catch {
      Write-Warning "Get-PSResource failed, falling back to Get-Module: $_"
      $usePSResource = $false
    }
  }

  if ($usePSResource -and $resources) {
    foreach ($r in $resources) {
      if ($Name -and $r.Name -notmatch $Name) { continue }
      $allModules.Add([PSCustomObject]@{
          Name = $r.Name
          Version = $r.Version
          Path = $modulesPath
        })
    }
  }
  else {
    if (Test-Path -LiteralPath $modulesPath -PathType Container) {
      Get-Module -ListAvailable -ErrorAction SilentlyContinue |
        Where-Object {
          $_.ModuleType -ne 'Binary' -and (
            ($_.ModuleBase -like "$modulesPath\*") -or
            ($_.Path -like "$modulesPath\*")
          ) -and (-not $Name -or $_.Name -match $Name)
        } |
        ForEach-Object {
          $allModules.Add([PSCustomObject]@{
              Name = $_.Name
              Version = $_.Version
              Path = $modulesPath
            })
        }
    }
  }

  if ($allModules.Count -eq 0) {
    $filterMsg = if ($Name) { "matching '$Name'" } else { '' }
    Write-Log -Message "No modules found $filterMsg in $modulesPath" -Color Yellow
    return
  }

  $grouped = $allModules | Group-Object Name

  if ($All) {
    $total = ($grouped | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
    $modNames = ($grouped | ForEach-Object { $_.Name }) -join ', '

    if (-not $Force -and -not $PSCmdlet.ShouldContinue(
        "This will remove ALL $total version(s) of [$modNames].`nPath: $modulesPath`nAre you sure?",
        'Remove ALL module versions'
      )) {
      Write-Log -Message 'Cancelled.' -Color Yellow
      return
    }

    foreach ($mod in $allModules) {
      if ($PSCmdlet.ShouldProcess("$($mod.Name) v$($mod.Version)", 'Remove')) {
        Write-Log -Message "Removing $($mod.Name) $($mod.Version)..." -Color Yellow
        try {
          if ($usePSResource) {
            Uninstall-PSResource -Name $mod.Name -Version $mod.Version.ToString() -SkipDependencyCheck -Scope $Scope -ErrorAction Stop
          }
          else {
            Uninstall-Module -Name $mod.Name -RequiredVersion $mod.Version.ToString() -Force -ErrorAction Stop
          }
          Write-Log -Message "  Removed $($mod.Name) $($mod.Version)" -Color Green
        }
        catch {
          Write-Warning "  Failed to remove $($mod.Name) v$($mod.Version): $_"
        }
      }
    }
    return
  }

  foreach ($group in $grouped) {
    $sorted = $group.Group | Sort-Object Version -Descending
    $keep = $sorted | Select-Object -First $LatestToKeep
    $remove = $sorted | Select-Object -Skip $LatestToKeep

    $latestVersion = $sorted | Select-Object -First 1 | ForEach-Object { $_.Version.ToString() }
    $keepVersions = ($keep | ForEach-Object { $_.Version.ToString() }) -join ', '

    Write-Log -Message "Latest $($group.Name): $latestVersion.  Keeping: [$keepVersions]" -Color Green

    if ($remove.Count -gt 0) {
      $oldVersions = ($remove | ForEach-Object { $_.Version.ToString() }) -join ', '
      Write-Log -Message "  Removing old versions of $($group.Name): [$oldVersions]" -Color Yellow

      foreach ($r in $remove) {
        if ($PSCmdlet.ShouldProcess("$($r.Name) v$($r.Version)", 'Remove')) {
          try {
            if ($usePSResource) {
              Uninstall-PSResource -Name $r.Name -Version $r.Version.ToString() -SkipDependencyCheck -Scope $Scope -ErrorAction Stop
            }
            else {
              Uninstall-Module -Name $r.Name -RequiredVersion $r.Version.ToString() -Force -ErrorAction Stop
            }
            Write-Log -Message "    Removed $($r.Name) $($r.Version)" -Color Green
          }
          catch {
            Write-Warning "    Failed to remove $($r.Name) v$($r.Version): $_"
          }
        }
      }
    }
  }
}

# ---- Add-PSModule ----------------------------------------------------------

function Add-PSModule {
  <#
    .SYNOPSIS
      Installs PowerShell modules.

    .DESCRIPTION
      Installs a single module by name, or restores all modules from a JSON
      file previously exported by Get-PSModule.

      Uses PSResourceGet on PS7+ whenever available, falling back to
      PowerShellGet on PS5.1.

    .PARAMETER Name
      Module name to install. Ignored when -FromFile is used.

    .PARAMETER Version
      Specific module version to install. Cannot be used with -MinimumVersion.

    .PARAMETER MinimumVersion
      Minimum acceptable module version. Installs the newest available version
      at or above this threshold. Cannot be used with -Version.

    .PARAMETER Scope
      Installation scope: CurrentUser (default) or AllUsers.

    .PARAMETER Force
      Force reinstallation even if the module is already present.

    .PARAMETER FromFile
      JSON file path from a previous Get-PSModule export. All modules in the
      file are reinstalled at their recorded versions.

    .EXAMPLE
      Add-PSModule -Name Pester -Version '5.4.0'
      Installs Pester 5.4.0 to the CurrentUser scope.

    .EXAMPLE
      Add-PSModule -FromFile ./modules.json -Force
      Restores all modules listed in modules.json, forcing reinstall.
  #>

  [CmdletBinding(SupportsShouldProcess = $true)]
  param(
    [string]$Name,
    [string]$Version,
    [string]$MinimumVersion,
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',
    [switch]$Force,
    [string]$FromFile
  )

  if ($Version -and $MinimumVersion) {
    throw '-Version and -MinimumVersion are mutually exclusive.'
  }

  $isPS7 = $PSVersionTable.PSVersion.Major -ge 7
  $usePSResource = $false

  if ($isPS7) {
    Ensure-PSResourceGet
    $usePSResource = Test-PSResourceGetAvailable
  }

  $null = Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction SilentlyContinue

  # ---- Restore from JSON file ----
  if ($FromFile) {
    $resolvedFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($FromFile)
    if (-not (Test-Path -LiteralPath $resolvedFile -PathType Leaf)) {
      throw "Module export file not found: $resolvedFile"
    }

    $moduleList = Get-Content -LiteralPath $resolvedFile -Raw | ConvertFrom-Json

    foreach ($mod in $moduleList) {
      $modName = $mod.Name
      $modVersion = $mod.Version
      $modScope = if ($mod.Scope) { $mod.Scope } else { $Scope }

      if ($PSCmdlet.ShouldProcess("$modName v$modVersion", 'Install')) {
        Write-Log -Message "Installing $modName $modVersion..." -Color Yellow
        try {
          if ($usePSResource) {
            $params = @{
              Name = $modName
              Version = $modVersion
              Scope = $modScope
            }
            Install-PSResource @params -ErrorAction Stop
          }
          else {
            $installParams = @{
              Name = $modName
              RequiredVersion = $modVersion
              Scope = $modScope
              Force = $true
              AllowClobber = $true
              SkipPublisherCheck = $true
            }
            Install-Module @installParams
          }
          Write-Log -Message "  Installed $modName $modVersion" -Color Green
        }
        catch {
          Write-Warning "  Failed to install $modName v${modVersion}: $_"
        }
      }
    }
    Write-Log -Message "Restored $($moduleList.Count) module(s) from $resolvedFile" -Color Green
    return
  }

  # ---- Install single module ----
  if (-not $Name) {
    throw '-Name is required when -FromFile is not used.'
  }

  $displayLabel = if ($Version) {
    $Version
  }
  elseif ($MinimumVersion) {
    ">= $MinimumVersion"
  }
  else {
    'latest'
  }

  if ($PSCmdlet.ShouldProcess("$Name $displayLabel", 'Install')) {
    Write-Log -Message "Installing $Name $displayLabel..." -Color Yellow
    try {
      if ($usePSResource) {
        $params = @{
          Name = $Name
          Scope = $Scope
        }
        if ($Version) {
          $params.Version = $Version
        }
        elseif ($MinimumVersion) {
          $params.Version = "[$MinimumVersion,)"
        }
        Install-PSResource @params -ErrorAction Stop
      }
      else {
        $installParams = @{
          Name = $Name
          Scope = $Scope
          Force = $Force
          AllowClobber = $true
          SkipPublisherCheck = $true
        }
        if ($Version) {
          $installParams.RequiredVersion = $Version
        }
        elseif ($MinimumVersion) {
          $installParams.MinimumVersion = $MinimumVersion
        }
        Install-Module @installParams
      }
      Write-Log -Message "  Installed $Name $displayLabel" -Color Green
    }
    catch {
      Write-Warning "  Failed to install ${Name}: $_"
    }
  }
}
