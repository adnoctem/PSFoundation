#Requires -Version 5.0

# Shared native methods - compiled once and reused by Get-SystemMemory / Get-SystemUptime
if ($null -eq ('SysInfoNative' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class SysInfoNative {
    [StructLayout(LayoutKind.Sequential)]
    public struct MEMORYSTATUSEX {
        public uint dwLength;
        public uint dwMemoryLoad;
        public ulong ullTotalPhys;
        public ulong ullAvailPhys;
        public ulong ullTotalPageFile;
        public ulong ullAvailPageFile;
        public ulong ullTotalVirtual;
        public ulong ullAvailVirtual;
        public ulong ullAvailExtendedVirtual;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GlobalMemoryStatusEx(ref MEMORYSTATUSEX lpBuffer);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern ulong GetTickCount64();
}
'@ -ErrorAction Stop
}

$script:SysInfoInvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture

function Format-SysInfoInvariant {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]
    $Format,

    [Parameter(ValueFromRemainingArguments = $true)]
    [object[]]
    $ArgumentList
  )

  return [string]::Format($script:SysInfoInvariantCulture, $Format, $ArgumentList)
}

function Resolve-WindowsProductName {
  [OutputType([string])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [AllowNull()]
    [string]
    $ProductName,

    [Parameter(Mandatory = $false)]
    [int]
    $CurrentBuild = 0
  )

  if ([string]::IsNullOrWhiteSpace($ProductName)) {
    return $ProductName
  }

  if ($CurrentBuild -ge 22000 -and $ProductName -like 'Windows 10*') {
    return $ProductName -replace '^Windows 10', 'Windows 11'
  }

  return $ProductName
}

function Get-OSBuildNumber {
  <#
    .SYNOPSIS
      Returns the Windows build number as an integer.
    .DESCRIPTION
      Reads CurrentBuild from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
      This is a single cheap registry read - far faster than Get-CimInstance
      or any WMI-based approach.  Returns e.g. 22621 (22H2), 22631 (23H2).
    .EXAMPLE
      PS> Get-OSBuildNumber
      22621
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([int])]
  [CmdletBinding()]
  param()

  $build = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'CurrentBuild' -ErrorAction Stop
  return [int]$build
}

function Get-OSDisplayVersion {
  <#
    .SYNOPSIS
      Returns the Windows feature-update display name (e.g. "22H2", "23H2").
    .DESCRIPTION
      Reads DisplayVersion from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
      Falls back to ReleaseId on older builds that lack DisplayVersion.
    .EXAMPLE
      PS> Get-OSDisplayVersion
      23H2
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param()

  try {
    $display = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'DisplayVersion' -ErrorAction Stop
    if ($display) { return $display }
  }
  catch {
    Write-Verbose "DisplayVersion was not available; falling back to ReleaseId."
  }

  $releaseId = Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'ReleaseId' -ErrorAction Stop
  return $releaseId
}

function Get-OSEdition {
  <#
    .SYNOPSIS
      Returns the Windows edition SKU (e.g. "Professional", "Enterprise").
    .DESCRIPTION
      Reads EditionID from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
    .EXAMPLE
      PS> Get-OSEdition
      Professional
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param()

  return Get-ItemPropertyValue -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'EditionID' -ErrorAction Stop
}

function Get-OSProductName {
  <#
    .SYNOPSIS
      Returns the full Windows product name string.
    .DESCRIPTION
      Reads ProductName from HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion.
      Returns e.g. "Windows 11 Pro" or "Windows 10 Enterprise".
    .EXAMPLE
      PS> Get-OSProductName
      Windows 11 Pro
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param()

  $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
  $props = Get-ItemProperty -LiteralPath $key -ErrorAction Stop
  return Resolve-WindowsProductName -ProductName $props.ProductName -CurrentBuild ([int]$props.CurrentBuild)
}

function Get-OSVersionInfo {
  <#
    .SYNOPSIS
      Returns a complete snapshot of Windows version metadata from the registry.
    .DESCRIPTION
      Performs a single Get-ItemProperty call against HKLM:\SOFTWARE\Microsoft\
      Windows NT\CurrentVersion and returns all relevant fields in one object.
      Far cheaper than calling the individual Get-OS* functions when you need
      multiple values.
    .EXAMPLE
      PS> Get-OSVersionInfo
    .EXAMPLE
      PS> Get-OSVersionInfo | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $key = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
  $props = Get-ItemProperty -LiteralPath $key -ErrorAction Stop

  $installDate = $null
  if ($props.InstallDate) {
    try {
      $unixEpochUtc = [DateTime]::SpecifyKind([DateTime]'1970-01-01T00:00:00Z', [DateTimeKind]::Utc)
      $installDate = $unixEpochUtc.AddSeconds([int64]$props.InstallDate).ToLocalTime()
    }
    catch {
      Write-Verbose "Unable to convert InstallDate registry value: $($props.InstallDate)"
    }
  }

  $currentBuild = [int]$props.CurrentBuild

  [PSCustomObject]@{
    ProductName = Resolve-WindowsProductName -ProductName $props.ProductName -CurrentBuild $currentBuild
    EditionID = $props.EditionID
    InstallationType = $props.InstallationType
    DisplayVersion = $props.DisplayVersion
    CurrentBuild = $currentBuild
    UBR = if ($null -ne $props.UBR) { [int]$props.UBR } else { 0 }
    ReleaseId = $props.ReleaseId
    BuildBranch = $props.BuildBranch
    InstallDate = $installDate
    RegisteredOwner = $props.RegisteredOwner
  }
}

function Get-SystemMemory {
  <#
    .SYNOPSIS
      Returns physical memory statistics - total, available, used, and load
      percentage.
    .DESCRIPTION
      Uses kernel32!GlobalMemoryStatusEx via P/Invoke (no CIM/WMI overhead).
      Returns an object with human-readable GiB values and the raw bytes.
    .EXAMPLE
      PS> Get-SystemMemory
    .EXAMPLE
      PS> Get-SystemMemory | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $memInfo = New-Object SysInfoNative+MEMORYSTATUSEX
  $memInfo.dwLength = [System.Runtime.InteropServices.Marshal]::SizeOf($memInfo)

  if (-not [SysInfoNative]::GlobalMemoryStatusEx([ref]$memInfo)) {
    Write-Error 'GlobalMemoryStatusEx failed.'
    return $null
  }

  $totalGiB = [math]::Round($memInfo.ullTotalPhys / 1GB, 2)
  $availableGiB = [math]::Round($memInfo.ullAvailPhys / 1GB, 2)
  $usedGiB = [math]::Round(($memInfo.ullTotalPhys - $memInfo.ullAvailPhys) / 1GB, 2)

  [PSCustomObject]@{
    TotalBytes = $memInfo.ullTotalPhys
    AvailableBytes = $memInfo.ullAvailPhys
    UsedBytes = $memInfo.ullTotalPhys - $memInfo.ullAvailPhys
    LoadPercent = $memInfo.dwMemoryLoad
    TotalGiB = Format-SysInfoInvariant '{0:0.##}' $totalGiB
    AvailableGiB = Format-SysInfoInvariant '{0:0.##}' $availableGiB
    UsedGiB = Format-SysInfoInvariant '{0:0.##}' $usedGiB
  }
}

function Get-SystemDisk {
  <#
    .SYNOPSIS
      Returns disk usage information for fixed volumes backed by physical disks.
    .DESCRIPTION
      Uses Win32_DiskDrive associations to include only volumes that resolve
      back to a physical disk. Provider-backed and cloud-mounted drives such as
      Google Drive are ignored. Filters to fixed drives by default and returns
      total size, free space, used space, and the filesystem type.
    .PARAMETER All
      Include non-fixed volumes when they are backed by a physical disk.
    .EXAMPLE
      PS> Get-SystemDisk
    .EXAMPLE
      PS> Get-SystemDisk -All
    .EXAMPLE
      PS> Get-SystemDisk | Format-Table -AutoSize
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [switch]
    $All = $false
  )

  $logicalDisksByDeviceId = @{}

  try {
    $physicalDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction Stop

    foreach ($physicalDisk in $physicalDisks) {
      $partitions = Get-CimAssociatedInstance -InputObject $physicalDisk -Association Win32_DiskDriveToDiskPartition -ErrorAction Stop
      foreach ($partition in $partitions) {
        $logicalDisks = Get-CimAssociatedInstance -InputObject $partition -Association Win32_LogicalDiskToPartition -ErrorAction Stop
        foreach ($logicalDisk in $logicalDisks) {
          if (-not $All -and [int]$logicalDisk.DriveType -ne 3) { continue }
          if ([string]::IsNullOrWhiteSpace($logicalDisk.DeviceID)) { continue }
          $logicalDisksByDeviceId[$logicalDisk.DeviceID] = $logicalDisk
        }
      }
    }
  }
  catch {
    Write-Verbose "Physical disk association inventory failed: $($_.Exception.Message)"
    $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -ErrorAction SilentlyContinue
    foreach ($logicalDisk in $logicalDisks) {
      if (-not $All -and [int]$logicalDisk.DriveType -ne 3) { continue }
      if ([string]::IsNullOrWhiteSpace($logicalDisk.DeviceID)) { continue }
      $logicalDisksByDeviceId[$logicalDisk.DeviceID] = $logicalDisk
    }
  }

  foreach ($deviceId in ($logicalDisksByDeviceId.Keys | Sort-Object)) {
    $disk = $logicalDisksByDeviceId[$deviceId]
    if ($null -eq $disk.Size -or [uint64]$disk.Size -eq 0) { continue }

    $totalBytes = [uint64]$disk.Size
    $freeBytes = [uint64]$disk.FreeSpace
    $totalGiB = [math]::Round($totalBytes / 1GB, 2)
    $freeGiB = [math]::Round($freeBytes / 1GB, 2)
    $usedGiB = [math]::Round(($totalBytes - $freeBytes) / 1GB, 2)
    $percentFree = [math]::Round($freeBytes * 100.0 / $totalBytes, 1)

    [PSCustomObject]@{
      Name = $disk.DeviceID
      Label = $disk.VolumeName
      Type = switch ([int]$disk.DriveType) {
        2 { 'Removable' }
        3 { 'Fixed' }
        4 { 'Network' }
        5 { 'CDRom' }
        6 { 'Ram' }
        default { 'Unknown' }
      }
      FileSystem = $disk.FileSystem
      TotalGiB = Format-SysInfoInvariant '{0:0.##}' $totalGiB
      FreeGiB = Format-SysInfoInvariant '{0:0.##}' $freeGiB
      UsedGiB = Format-SysInfoInvariant '{0:0.##}' $usedGiB
      PercentFree = Format-SysInfoInvariant '{0:0.#}' $percentFree
      TotalBytes = $totalBytes
      FreeBytes = $freeBytes
    }
  }
}

function Get-Hostname {
  <#
    .SYNOPSIS
      Returns the computer hostname.
    .DESCRIPTION
      Uses [System.Net.Dns]::GetHostName() and resolves it to a fully-qualified
      domain name when joined to a domain.  Returns both the short hostname and,
      if different, the FQDN.
    .EXAMPLE
      PS> Get-Hostname
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $hostname = [System.Net.Dns]::GetHostName()

  $fqdn = $null
  try {
    $entry = [System.Net.Dns]::GetHostEntry($hostname)
    $fqdn = $entry.HostName
    if ($fqdn -eq $hostname) { $fqdn = $null }
  }
  catch {
    Write-Verbose "Unable to resolve FQDN for hostname: $hostname"
  }

  [PSCustomObject]@{
    Hostname = $hostname
    FQDN = $fqdn
  }
}

function Get-SystemUptime {
  <#
    .SYNOPSIS
      Returns the system uptime (time since last boot).
    .DESCRIPTION
      Uses kernel32!GetTickCount64 via P/Invoke for a non-wrapping, high-
      precision uptime value - no CIM/WMI overhead.  Returns the raw tick
      count, total milliseconds, and a human-readable breakdown.
    .EXAMPLE
      PS> Get-SystemUptime
    .EXAMPLE
      PS> Get-SystemUptime | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $ticksMs = [SysInfoNative]::GetTickCount64()
  $span = [TimeSpan]::FromMilliseconds($ticksMs)

  [PSCustomObject]@{
    TotalMilliseconds = $ticksMs
    Days = $span.Days
    Hours = $span.Hours
    Minutes = $span.Minutes
    Seconds = $span.Seconds
    TotalHours = Format-SysInfoInvariant '{0:0.#}' ([math]::Round($span.TotalHours, 1))
    TotalDays = Format-SysInfoInvariant '{0:0.#}' ([math]::Round($span.TotalDays, 1))
    Display = Format-SysInfoInvariant '{0}d {1:D2}h {2:D2}m {3:D2}s' $span.Days $span.Hours $span.Minutes $span.Seconds
  }
}

function Get-SystemInfo {
  <#
    .SYNOPSIS
      Returns a comprehensive system information snapshot (fetch-style).
    .DESCRIPTION
      Assembles OS version, memory, disk, hostname, and uptime into a single
      structured object.  Disk data is resolved through physical disk
      associations; other data sources use registry reads, .NET APIs, and
      lightweight P/Invoke where possible.
    .EXAMPLE
      PS> Get-SystemInfo | Format-List
    .EXAMPLE
      PS> Get-SystemInfo
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param()

  $_os = Get-OSVersionInfo
  $_mem = Get-SystemMemory
  $_disks = Get-SystemDisk
  $_host = Get-Hostname
  $_up = Get-SystemUptime

  [PSCustomObject]@{
    OSProductName = $_os.ProductName
    OSEdition = $_os.EditionID
    OSVersion = $_os.DisplayVersion
    OSBuild = $_os.CurrentBuild
    OSUBRev = $_os.UBR
    Hostname = $_host.Hostname
    FQDN = $_host.FQDN
    TotalMemoryGiB = $_mem.TotalGiB
    MemoryLoadPct = $_mem.LoadPercent
    Disks = ($_disks | ForEach-Object { Format-SysInfoInvariant '{0} {1}GiB/{2}GiB ({3}% free)' $_.Name $_.FreeGiB $_.TotalGiB $_.PercentFree }) -join ' | '
    Uptime = $_up.Display
    InstallDate = $_os.InstallDate
  }
}

function Get-SystemPaths {
  <#
    .SYNOPSIS
      Returns standard winkit directory paths for tools, config, cache, data, and logs.
    .DESCRIPTION
      Provides a single structured lookup for the five canonical winkit folders.
      The -Name parameter controls the subdirectory name used under each root.
      Defaults to 'winkit' so callers can omit it for standard usage.
    .PARAMETER Name
      Subdirectory name under each root. Defaults to 'winkit'.
    .EXAMPLE
      PS> $paths = Get-SystemPaths
      PS> $paths.Data
      C:\ProgramData\winkit
    .EXAMPLE
      PS> Get-SystemPaths -Name 'myapp' | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $false)]
    [string]
    $Name = 'winkit'
  )

  return [PSCustomObject]@{
    Home = Join-Path -Path $env:USERPROFILE -ChildPath $Name
    Config = Join-Path -Path $env:APPDATA -ChildPath $Name
    Cache = Join-Path -Path $env:LOCALAPPDATA -ChildPath $Name
    Data = Join-Path -Path $env:ProgramData -ChildPath $Name
    Logs = Join-Path -Path $env:LOCALAPPDATA -ChildPath "$Name\logs"
  }
}

function Test-HostApplicability {
  <#
    .SYNOPSIS
      Returns $true if the current host meets all supplied applicability constraints.
    .DESCRIPTION
      Gates a setting or operation by build number, OS edition, and processor
      bitness.  Any constraint that is not supplied is treated as "no
      restriction."  All supplied constraints must pass — the check is a
      logical AND across them.

      Builds on the existing Get-OSBuildNumber and Get-OSEdition helpers in
      this file so it shares the same cheap registry-read path.
    .PARAMETER MinBuild
      Minimum Windows build number required (inclusive).  Example: 22000 for
      Windows 11+.
    .PARAMETER MaxBuild
      Maximum Windows build number allowed (inclusive).  Example: 22621 for
      Windows 11 22H2 and below.
    .PARAMETER Edition
      OS edition(s) that qualify.  Accepts an array of strings such as
      @('Enterprise', 'Professional').
    .PARAMETER Bitness
      Required processor architecture: x64 or x86.
    .OUTPUTS
      [bool]
    .EXAMPLE
      PS> if (-not (Test-HostApplicability -MinBuild 22000)) { Write-Log -Message 'Requires Windows 11+'; exit 0 }
    .EXAMPLE
      PS> Test-HostApplicability -Edition @('ServerStandard', 'ServerDatacenter') -Bitness x64
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([bool])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [int]
    $MinBuild,

    [Parameter(Mandatory = $false)]
    [int]
    $MaxBuild,

    [Parameter(Mandatory = $false)]
    [string[]]
    $Edition,

    [Parameter(Mandatory = $false)]
    [ValidateSet('x64', 'x86')]
    [string]
    $Bitness
  )

  $build = Get-OSBuildNumber

  if ($PSBoundParameters.ContainsKey('MinBuild') -and $build -lt $MinBuild) {
    return $false
  }

  if ($PSBoundParameters.ContainsKey('MaxBuild') -and $build -gt $MaxBuild) {
    return $false
  }

  if ($PSBoundParameters.ContainsKey('Edition') -and $Edition.Count -gt 0) {
    $currentEdition = Get-OSEdition
    if ($currentEdition -notin $Edition) {
      return $false
    }
  }

  if ($PSBoundParameters.ContainsKey('Bitness')) {
    $is64Bit = [Environment]::Is64BitOperatingSystem
    if (($Bitness -eq 'x64' -and -not $is64Bit) -or ($Bitness -eq 'x86' -and $is64Bit)) {
      return $false
    }
  }

  return $true
}

# REVIEW-DATA: .NET Framework release-number lookup table. Point-in-time mapping —
# update when Microsoft ships new .NET Framework releases. Authoritative table:
# https://support.microsoft.com/en-us/help/318785
$script:NetFrameworkReleaseMap = @(
  @{ MinimumRelease = 533509; Product = '4.8.1 Windows 11 24H2 or Windows Server 2025' }
  @{ MinimumRelease = 533320; Product = '4.8.1' }
  @{ MinimumRelease = 528449; Product = '4.8 Windows 11 or Windows Server 2022' }
  @{ MinimumRelease = 528372; Product = '4.8 Windows 10 May 2020, Oct 2020, May 2021' }
  @{ MinimumRelease = 528040; Product = '4.8 Windows 10 May 2019, Nov 2019' }
  @{ MinimumRelease = 461808; Product = '4.7.2' }
  @{ MinimumRelease = 461308; Product = '4.7.1' }
  @{ MinimumRelease = 460798; Product = '4.7 Original Release' }
  @{ ExactRelease = 394802; Product = '4.6.2 Windows 10 Anniversary Update' }
  @{ ExactRelease = 394806; Product = '4.6.2' }
  @{ ExactRelease = 394254; Product = '4.6.1 Windows 10 November Update' }
  @{ MinimumRelease = 394271; Product = '4.6.1' }
  @{ MinimumRelease = 393295; Product = '4.6 Windows 10' }
  @{ MinimumRelease = 379897; Product = '4.6 Original Release' }
  @{ MinimumRelease = 379893; Product = '4.5.2' }
  @{ MinimumRelease = 378675; Product = '4.5.1 Windows 8.1 or Windows Server 2012' }
  @{ MinimumRelease = 378758; Product = '4.5.1 Windows 8, Windows 7 SP1, or Windows Vista SP2' }
  @{ MinimumRelease = 378389; Product = '4.5 Original Release' }
)

function Resolve-NetFrameworkProductName {
  <#
    .SYNOPSIS
      Resolves a .NET Framework release number to a friendly product name.
    .DESCRIPTION
      Internal helper for Get-DotNetVersion. The release-number mapping is a
      point-in-time data table ($script:NetFrameworkReleaseMap) that must be
      reviewed when new .NET Framework versions ship.
    .PARAMETER Release
      Registry Release value for the .NET Framework installation.
    .PARAMETER ServicePack
      Optional SP value from the registry, used for the pre-4.5 fallback mapping.
    .PARAMETER ChildName
      Registry key leaf name (e.g. v4, v3.5), used for the pre-4.5 fallback mapping.
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [int]$Release = 0,

    [Parameter(Mandatory = $false)]
    [int]$ServicePack = 0,

    [Parameter(Mandatory = $false)]
    [string]$ChildName
  )

  if ($Release -gt 0) {
    foreach ($entry in $script:NetFrameworkReleaseMap) {
      if ($entry.ContainsKey('ExactRelease') -and $Release -eq $entry.ExactRelease) {
        return $entry.Product
      }
      if ($entry.ContainsKey('MinimumRelease') -and $Release -ge $entry.MinimumRelease) {
        return $entry.Product
      }
    }
  }

  switch ($ChildName) {
    'v3.5' {
      if ($ServicePack -eq 1) { return '3.5 ServicePack 1' }
      return '3.5 Original Release'
    }
    'v3.0' {
      if ($ServicePack -eq 2) { return '3.5 ServicePack 2' }
      if ($ServicePack -eq 1) { return '3.5 ServicePack 1' }
      return '3.5'
    }
    'v2.0*' {
      if ($ServicePack -eq 2) { return '2.0 ServicePack 2' }
      if ($ServicePack -eq 1) { return '2.0 ServicePack 1' }
      return '2.0'
    }
    default { return '' }
  }
}

function Get-DotNetVersion {
  <#
    .SYNOPSIS
      Returns the installed .NET runtimes, SDKs, and .NET Framework versions.
    .DESCRIPTION
      Enumerates three independent sources of .NET version information:

      * dotnet CLI runtimes and SDKs (dotnet --list-runtimes / --list-sdks) — the
        full output of each command, or $null when the dotnet CLI is not installed.
      * The classic .NET Framework version of the current runtime
        ([Environment]::Version).
      * The full .NET Framework release enumeration from the NDP registry subtree
        (HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP), with each release
        number mapped to a friendly product name.

      NOTE: the release-number-to-product mapping is a point-in-time data table
      and must be reviewed when new .NET Framework versions ship (see the
      REVIEW-DATA marker in system.ps1 and support.microsoft.com help article
      318785 for the authoritative mapping).
    .OUTPUTS
      PSCustomObject with DotnetRuntimes, DotnetSDKs, FrameworkVersion, and
      Releases (one entry per installed .NET Framework release).
    .EXAMPLE
      PS> Get-DotNetVersion | Format-List
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param ()

  $runtimes = $null
  $sdks = $null

  if (Get-Command dotnet -ErrorAction SilentlyContinue) {
    $runtimes = @(& dotnet --list-runtimes 2>$null)
    $sdks = @(& dotnet --list-sdks 2>$null)
  }

  $releases = @()
  Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse -ErrorAction SilentlyContinue |
    Get-ItemProperty -Name Version, Release, Install, PSChildName, SP -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^(?!S)\p{L}' } |
    ForEach-Object {
      $releases += [PSCustomObject]@{
        Name = $_.PSChildName
        Version = $_.Version
        Release = $_.Release
        Installed = $_.Install
        Product = Resolve-NetFrameworkProductName -Release ([int]$_.Release) -ServicePack ([int]$_.SP) -ChildName $_.PSChildName
      }
    }

  [PSCustomObject]@{
    DotnetRuntimes = $runtimes
    DotnetSDKs = $sdks
    FrameworkVersion = [Environment]::Version.ToString()
    Releases = $releases
  }
}

function New-DriveMapping {
  <#
    .SYNOPSIS
      New-DriveMapping - Persistently maps a local folder to a drive letter.
    .DESCRIPTION
      Creates a persistent mapping of a local folder (e.g. C:\Development) to a
      drive letter (e.g. D:) via the DOS Devices registry key, plus a matching
      Explorer volume-label entry under DriveIcons.

      The mapping only takes effect for the current session once the registry
      value is read at logon; a reboot is required for it to become visible to
      Explorer and most applications. Unlike the reference implementation, this
      function NEVER reboots by default - a restart only happens when -Restart
      is explicitly supplied.

      Requires elevation (writes HKLM registry values).
    .PARAMETER DriveLetter
      Drive letter to map to Path. Must not be a physical volume, and must not
      already be mapped (unless -Force).
    .PARAMETER Path
      Folder path to map to DriveLetter. Must exist and must not be a root-level
      folder.
    .PARAMETER SourceDriveLabel
      Label to apply to the source drive in Explorer. Default: current volume
      label (or the existing DriveIcons label when the volume has none).
    .PARAMETER DriveLabel
      Volume label for the mapped drive. Default: leaf folder name of Path.
    .PARAMETER Restart
      When supplied, restarts the machine after the mapping is created. A
      library function never reboots the machine silently - this is opt-in.
    .PARAMETER Force
      Override an existing mapping on DriveLetter, and override the source
      drive's DriveIcons label.
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result.
    .EXAMPLE
      PS> New-DriveMapping -DriveLetter 'D' -Path 'C:\Development'
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateLength(1, 1)]
    [ValidatePattern('[A-Z]')]
    [ValidateScript({
        if (($null -ne (Get-Volume $_ -ErrorAction SilentlyContinue))) {
          throw 'DriveLetter cannot be a physical volume'
        }
        $true
      })]
    [string]
    $DriveLetter,

    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateScript({
        if ((Test-Path -LiteralPath $_) -and ((Split-Path -Path $_ -Leaf) -ne $_)) {
          return $true
        }
        throw 'Path does not exist or is a root level folder'
      })]
    [string]
    $Path,

    [Parameter(Mandatory = $false)]
    [string]
    $SourceDriveLabel,

    [Parameter(Mandatory = $false)]
    [string]
    $DriveLabel,

    [Parameter(Mandatory = $false)]
    [switch]
    $Restart,

    [Parameter(Mandatory = $false)]
    [switch]
    $Force
  )

  if (-not (Test-Elevation)) {
    throw 'New-DriveMapping requires an elevated process (administrator token).'
  }

  $driveIconsKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons'
  $dosDevicesKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\DOS Devices'

  $DriveLetter = $DriveLetter.ToUpper()

  if (-not $Force) {
    $existing = Get-ItemProperty -LiteralPath $dosDevicesKey -Name "$DriveLetter`:" -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
      throw 'Drive letter is already mapped; use -Force to override'
    }
  }

  $Path = [System.IO.Path]::GetFullPath($Path)
  if ($Path.EndsWith('\')) {
    $Path = $Path.Substring(0, $Path.Length - 1)
  }

  $sourceLetter = ([System.IO.Path]::GetPathRoot($Path))[0]

  if (-not $SourceDriveLabel) {
    $sourceVolume = Get-Volume -DriveLetter $sourceLetter -ErrorAction SilentlyContinue
    $SourceDriveLabel = if ($null -ne $sourceVolume) { $sourceVolume.FileSystemLabel } else { '' }
    if ([string]::IsNullOrEmpty($SourceDriveLabel)) {
      $SourceDriveLabel = Get-ItemPropertyValue -LiteralPath "$driveIconsKey\$sourceLetter\DefaultLabel" -Name '(Default)' -ErrorAction SilentlyContinue
    }
  }

  if (-not $DriveLabel) {
    $DriveLabel = Split-Path -Path $Path -Leaf
  }

  if (-not $PSCmdlet.ShouldProcess("$DriveLetter`:", "Map folder '$Path' to drive letter")) {
    if ($WhatIfPreference) {
      New-OperationResult -Target "$DriveLetter`:" -Source 'DOS Devices' -Action 'CreateMapping' -Status 'DryRun' -Detail "Would map '$Path' to drive letter '$DriveLetter`:'."
    }
    return
  }

  # Clear the source drive's volume name so the DriveIcons label can take over.
  Get-Volume -DriveLetter $sourceLetter -ErrorAction SilentlyContinue |
    Where-Object { $_.FileSystemLabel -ne '' } |
    Set-Volume -NewFileSystemLabel ''

  if ($Force -or [string]::IsNullOrEmpty(
      (Get-ItemPropertyValue -LiteralPath "$driveIconsKey\$sourceLetter\DefaultLabel" -Name '(Default)' -ErrorAction SilentlyContinue))) {
    $sourceLabelPath = "$driveIconsKey\$sourceLetter\DefaultLabel"
    if (-not (Test-Path -LiteralPath $sourceLabelPath)) {
      $null = New-Item -Path $sourceLabelPath -Force
    }
    Set-ItemProperty -LiteralPath $sourceLabelPath -Name '(Default)' -Value $SourceDriveLabel -Force
  }

  $mappedLabelPath = "$driveIconsKey\$DriveLetter\DefaultLabel"
  if (-not (Test-Path -LiteralPath $mappedLabelPath)) {
    $null = New-Item -Path $mappedLabelPath -Force
  }
  Set-ItemProperty -LiteralPath $mappedLabelPath -Name '(Default)' -Value $DriveLabel -Force

  Set-ItemProperty -LiteralPath $dosDevicesKey -Name "$DriveLetter`:" -Value "\??\$Path" -Force

  if ($Restart) {
    Write-Verbose 'Restart requested by -Restart; restarting the machine.'
    Restart-Computer -Force
  }

  New-OperationResult -Target "$DriveLetter`:" -Source 'DOS Devices' -Action 'CreateMapping' -Status 'Completed' -Detail "Mapped '$Path' to drive letter '$DriveLetter`:'." -Property @{ Path = $Path; DriveLabel = $DriveLabel }
}

function Remove-DriveMapping {
  <#
    .SYNOPSIS
      Remove-DriveMapping - Removes a persistent folder-to-drive-letter mapping.
    .DESCRIPTION
      Removes a mapping created by New-DriveMapping: deletes the DOS Devices
      value, removes the mapped drive's DriveIcons label, and - when the removed
      mapping was the last one to its source drive - restores the source
      drive's volume label.

      Never reboots by default; a restart only happens when -Restart is
      explicitly supplied.

      Requires elevation (writes HKLM registry values).
    .PARAMETER DriveLetter
      The mapped drive letter to remove. Must currently be mapped.
    .PARAMETER SourceDriveLabel
      Label to restore on the source drive. Default: the label previously
      stored in DriveIcons, or 'System'/'Data' based on the drive role.
    .PARAMETER Restart
      When supplied, restarts the machine after the mapping is removed.
    .PARAMETER Force
      Force restoring the source drive label even when other mappings to the
      source drive remain.
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result.
    .EXAMPLE
      PS> Remove-DriveMapping -DriveLetter 'D'
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateLength(1, 1)]
    [ValidatePattern('[A-Z]')]
    [ValidateScript({
        $value = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\DOS Devices' -Name "$_`:" -ErrorAction SilentlyContinue
        if ($null -eq $value) {
          throw 'Drive letter not found or does not represent a mapped drive'
        }
        $true
      })]
    [string]
    $DriveLetter,

    [Parameter(Mandatory = $false)]
    [string]
    $SourceDriveLabel,

    [Parameter(Mandatory = $false)]
    [switch]
    $Restart,

    [Parameter(Mandatory = $false)]
    [switch]
    $Force
  )

  if (-not (Test-Elevation)) {
    throw 'Remove-DriveMapping requires an elevated process (administrator token).'
  }

  $driveIconsKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\DriveIcons'
  $dosDevicesKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\DOS Devices'

  $DriveLetter = $DriveLetter.ToUpper()

  $source = Get-ItemProperty -LiteralPath $dosDevicesKey -Name "$DriveLetter`:" -ErrorAction SilentlyContinue
  if ($null -eq $source) {
    throw "Drive letter '$DriveLetter`:' is not a mapped drive"
  }

  # Value format is something like \??\C:\foobar - the 5th character is the source drive letter.
  $sourceValue = $source.PSObject.Properties[$DriveLetter + ':'].Value
  $sourceLetter = $sourceValue[4]

  if (-not $SourceDriveLabel) {
    $SourceDriveLabel = Get-ItemPropertyValue -LiteralPath "$driveIconsKey\$sourceLetter\DefaultLabel" -Name '(Default)' -ErrorAction SilentlyContinue
  }

  if (-not $SourceDriveLabel) {
    $SourceDriveLabel = if ($sourceLetter -eq ($env:HOMEDRIVE)[0]) { 'System' } else { 'Data' }
  }

  if (-not $PSCmdlet.ShouldProcess("$DriveLetter`:", 'Remove drive letter mapping')) {
    if ($WhatIfPreference) {
      New-OperationResult -Target "$DriveLetter`:" -Source 'DOS Devices' -Action 'RemoveMapping' -Status 'DryRun' -Detail "Would remove mapping for drive letter '$DriveLetter`:'."
    }
    return
  }

  Remove-ItemProperty -LiteralPath $dosDevicesKey -Name "$DriveLetter`:" -Force

  if (Test-Path -LiteralPath "$driveIconsKey\$DriveLetter") {
    Remove-Item -LiteralPath "$driveIconsKey\$DriveLetter" -Recurse -Force
  }

  $relatedMappings = (Get-ItemProperty -LiteralPath $dosDevicesKey).PSObject.Properties |
    Where-Object { $_.Name -match '^[A-Z]:$' -and $_.Value -match "^\\\?\?\\$sourceLetter`:\\" }

  if ($Force -or $relatedMappings.Count -eq 0) {
    if (Test-Path -LiteralPath "$driveIconsKey\$sourceLetter") {
      Remove-Item -LiteralPath "$driveIconsKey\$sourceLetter" -Recurse -Force
    }
    Set-Volume -DriveLetter $sourceLetter -NewFileSystemLabel $SourceDriveLabel -ErrorAction SilentlyContinue
  }

  if ($Restart) {
    Write-Verbose 'Restart requested by -Restart; restarting the machine.'
    Restart-Computer -Force
  }

  New-OperationResult -Target "$DriveLetter`:" -Source 'DOS Devices' -Action 'RemoveMapping' -Status 'Completed' -Detail "Removed mapping for drive letter '$DriveLetter`:'." -Property @{ SourceDriveLabel = $SourceDriveLabel }
}

function Test-RegistryValuePresent {
  <#
    .SYNOPSIS
      Returns whether a named value exists under a registry key.
    .DESCRIPTION
      Internal helper for Test-PendingReboot. Safe under StrictMode: missing
      properties and missing keys both yield $false rather than throwing.
  #>

  [OutputType([bool])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$KeyPath,

    [Parameter(Mandatory = $true)]
    [string]$ValueName
  )

  try {
    $props = Get-ItemProperty -LiteralPath $KeyPath -ErrorAction Stop
    return $null -ne ($props.PSObject.Properties[$ValueName])
  }
  catch {
    return $false
  }
}

function Test-PendingReboot {
  <#
    .SYNOPSIS
      Returns whether the machine has a pending reboot, and which indicators
      triggered.
    .DESCRIPTION
      Checks ten independent read-only registry/WMI indicators of a pending
      reboot:

        1. CBS RebootPending key presence
        2. CBS servicing state (RebootInProgress / PackagesPending)
        3. PendingFileRenameOperations value
        4. PendingFileRenameOperations2 value
        5. DVDRebootSignal value under RunOnce
        6. Netlogon JoinDomain / AvoidSpnSet values (pending domain join)
        7. Windows Update RebootRequired / PostRebootReporting values
        8. SCCM/MECM client SDK DetermineIfRebootPending (no-op when the
           client is not installed)
        9. Computer rename queued (ActiveComputerName differs from ComputerName)
        10. Win32_ComputerSystem.PendingSystemReboot

      No elevation required - all checks are read-only. Note: the ADDS
      provisioning scripts in winkit embed a self-contained copy of this
      indicator set for remote execution on freshly-provisioned targets that do
      not have PSFoundation installed; keep the two in sync.
    .OUTPUTS
      PSCustomObject with PendingReboot (bool) and Indicators (string[] of the
      indicator names that triggered).
    .EXAMPLE
      PS> Test-PendingReboot
      PendingReboot   Indicators
      -------------   ----------
              False   {}
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param ()

  $indicators = [System.Collections.Generic.List[string]]::new()

  $cbsKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing'

  if (Test-Path -LiteralPath "$cbsKey\RebootPending" -ErrorAction SilentlyContinue) {
    $indicators.Add('CBSRebootPending')
  }

  if ((Test-RegistryValuePresent -KeyPath $cbsKey -ValueName 'RebootInProgress') -or (Test-RegistryValuePresent -KeyPath $cbsKey -ValueName 'PackagesPending')) {
    $indicators.Add('CBSServicingState')
  }

  $sessionManagerKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'

  if (Test-RegistryValuePresent -KeyPath $sessionManagerKey -ValueName 'PendingFileRenameOperations') {
    $indicators.Add('PendingFileRenameOperations')
  }

  if (Test-RegistryValuePresent -KeyPath $sessionManagerKey -ValueName 'PendingFileRenameOperations2') {
    $indicators.Add('PendingFileRenameOperations2')
  }

  if (Test-RegistryValuePresent -KeyPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce' -ValueName 'DVDRebootSignal') {
    $indicators.Add('DVDRebootSignal')
  }

  $netlogonKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon'
  if ((Test-RegistryValuePresent -KeyPath $netlogonKey -ValueName 'JoinDomain') -or (Test-RegistryValuePresent -KeyPath $netlogonKey -ValueName 'AvoidSpnSet')) {
    $indicators.Add('NetlogonJoin')
  }

  $autoUpdateKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update'
  if ((Test-RegistryValuePresent -KeyPath $autoUpdateKey -ValueName 'RebootRequired') -or (Test-RegistryValuePresent -KeyPath $autoUpdateKey -ValueName 'PostRebootReporting')) {
    $indicators.Add('WindowsUpdate')
  }

  try {
    $sccmClient = Get-CimInstance -Namespace 'root\ccm\clientsdk' -ClassName CCM_ClientUtilities -ErrorAction Stop
    if ($null -ne $sccmClient -and $sccmClient.DetermineIfRebootPending().RebootPending) {
      $indicators.Add('SCCMClient')
    }
  }
  catch {
    Write-Verbose 'SCCM/MECM client is not installed; skipping the SCCM pending-reboot probe.'
  }

  try {
    $activeComputerName = (Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -ErrorAction Stop).ComputerName
    $computerName = (Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -ErrorAction Stop).ComputerName
    if ((-not [string]::IsNullOrWhiteSpace($activeComputerName)) -and (-not [string]::IsNullOrWhiteSpace($computerName)) -and ($activeComputerName -ne $computerName)) {
      $indicators.Add('ComputerRename')
    }
  }
  catch {
    Write-Verbose 'Computer rename comparison failed; skipping the rename pending-reboot probe.'
  }

  try {
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
    if ($computerSystem.PendingSystemReboot) {
      $indicators.Add('CIMPendingReboot')
    }
  }
  catch {
    Write-Verbose 'CIM pending-reboot probe failed; skipping.'
  }

  [PSCustomObject]@{
    PendingReboot = $indicators.Count -gt 0
    Indicators = $indicators.ToArray()
  }
}

function Resolve-SystemFileIntegrityVerdict {
  <#
    .SYNOPSIS
      Synthesizes a system-file-integrity verdict from sfc output and the CBS log.
    .DESCRIPTION
      Internal helper for Test-SystemFileIntegrity. Matches the documented sfc
      verdict markers against the combined sfc output and CBS log excerpt, and
      falls back to the raw exit code mapping (0 = clean, 1 = violations,
      2 = could not run).
    .OUTPUTS
      PSCustomObject with Status (Completed | Failed) and Detail (string).
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [string[]]$SfcOutput,

    [Parameter(Mandatory = $false)]
    [int]$ExitCode = 0,

    [Parameter(Mandatory = $false)]
    [string[]]$LogExcerpt
  )

  $haystack = (@($SfcOutput) + @($LogExcerpt)) -join [Environment]::NewLine

  if ($haystack -match 'found corrupt files') {
    return [PSCustomObject]@{ Status = 'Failed'; Detail = 'integrity violations found' }
  }
  if ($haystack -match 'did not find any integrity violations') {
    return [PSCustomObject]@{ Status = 'Completed'; Detail = 'no violations' }
  }
  if ($haystack -match 'could not perform the requested operation') {
    return [PSCustomObject]@{ Status = 'Failed'; Detail = 'could not run' }
  }

  switch ($ExitCode) {
    0 { return [PSCustomObject]@{ Status = 'Completed'; Detail = 'no violations' } }
    1 { return [PSCustomObject]@{ Status = 'Failed'; Detail = 'integrity violations found' } }
    2 { return [PSCustomObject]@{ Status = 'Failed'; Detail = 'could not run' } }
    default { return [PSCustomObject]@{ Status = 'Failed'; Detail = "sfc exited with code $ExitCode" } }
  }
}

function Test-SystemFileIntegrity {
  <#
    .SYNOPSIS
      Runs sfc /verifyOnly and returns a structured integrity verdict.
    .DESCRIPTION
      Runs a synchronous, read-only system file integrity check
      (sfc.exe /verifyOnly) and synthesizes a structured verdict from the
      output, the exit code, and the tail of the CBS log (the source of truth).

      The function intentionally runs /verifyOnly - a full scan-and-repair
      pass (/scannow) has significant runtime and disruption cost and must
      only ever run as a deliberate, separately-named action.

      The check requires an elevated process for sfc itself to run; the
      function does not perform any elevation.
    .OUTPUTS
      PSCustomObject with Status (Completed | Failed), Detail (verdict
      string), LogExcerpt (tail of the CBS log), and ExitCode.
    .EXAMPLE
      PS> Test-SystemFileIntegrity
      Status  : Completed
      Detail  : no integrity violations
      ExitCode: 0
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param ()

  $sfcOutput = & "$env:SystemRoot\System32\sfc.exe" /verifyOnly 2>&1
  $sfcExitCode = $LASTEXITCODE

  $cbsLogPath = Join-Path -Path $env:SystemRoot -ChildPath 'Logs\CBS\CBS.log'
  $logExcerpt = @()
  if (Test-Path -LiteralPath $cbsLogPath -ErrorAction SilentlyContinue) {
    $logExcerpt = Get-Content -LiteralPath $cbsLogPath -Tail 30 -ErrorAction SilentlyContinue
  }

  $verdict = Resolve-SystemFileIntegrityVerdict -SfcOutput $sfcOutput -ExitCode $sfcExitCode -LogExcerpt $logExcerpt

  [PSCustomObject]@{
    Status = $verdict.Status
    Detail = $verdict.Detail
    ExitCode = $sfcExitCode
    LogExcerpt = $logExcerpt
  }
}

function Convert-RobocopyExitCode {
  <#
    .SYNOPSIS
      Decodes a Robocopy exit code into its success/failure meaning.
    .DESCRIPTION
      Robocopy uses bitmask exit codes where 0-7 are all success variants and
      8+ indicates failure. Decodes a raw exit code into the corresponding
      description.
    .PARAMETER ExitCode
      The raw exit code from robocopy.exe.
    .OUTPUTS
      string
    .EXAMPLE
      PS> Convert-RobocopyExitCode -ExitCode 3
      OKCOPY + XTRA
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [int]
    $ExitCode
  )

  switch ($ExitCode) {
    16 { '***FATAL ERROR***' }
    15 { 'OKCOPY + FAIL + MISMATCHES + XTRA' }
    14 { 'FAIL + MISMATCHES + XTRA' }
    13 { 'OKCOPY + FAIL + MISMATCHES' }
    12 { 'FAIL + MISMATCHES' }
    11 { 'OKCOPY + FAIL + XTRA' }
    10 { 'FAIL + XTRA' }
    9 { 'OKCOPY + FAIL' }
    8 { 'FAIL' }
    7 { 'OKCOPY + MISMATCHES + XTRA' }
    6 { 'MISMATCHES + XTRA' }
    5 { 'OKCOPY + MISMATCHES' }
    4 { 'MISMATCHES' }
    3 { 'OKCOPY + XTRA' }
    2 { 'XTRA' }
    1 { 'OKCOPY' }
    0 { 'No Change' }
    default { 'Unknown' }
  }
}

function Find-ServiceAccountUsage {
  <#
    .SYNOPSIS
      Finds where a service account is used across services and scheduled tasks.
    .DESCRIPTION
      Given an account name (e.g. 'DOMAIN\svc-backup') and one or more
      computers, searches Windows services (Win32_Service.StartName) and
      scheduled tasks (schtasks, 'Run As User') for matches. Use this before
      rotating a service account's password - it reports every place a
      manually-run credential could be hiding.

      -Credential is a plain optional parameter with no default value, so
      non-interactive callers are never forced into a Get-Credential prompt.
    .PARAMETER Name
      Account name(s) to search for (wildcards supported).
    .PARAMETER ComputerName
      Computer(s) to search. Defaults to the local computer.
    .PARAMETER Credential
      Optional credential for remote computers.
    .OUTPUTS
      One PSCustomObject per searched account: Name, Services, and SchTasks
      collections.
    .EXAMPLE
      PS> Find-ServiceAccountUsage -Name 'svc-backup' -ComputerName 'SRV01','SRV02'
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $false)]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [string[]]
    $Name,

    [Parameter(Mandatory = $false)]
    [string[]]
    $ComputerName = @('localhost'),

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    $Credential
  )

  process {
    foreach ($account in $Name) {
      $services = @()
      $tasks = @()

      foreach ($computer in $ComputerName) {
        $cimParams = @{ ClassName = 'Win32_Service'; ComputerName = $computer; ErrorAction = 'SilentlyContinue' }
        if ($Credential) { $cimParams.Credential = $Credential }

        foreach ($service in @(Get-CimInstance @cimParams)) {
          if ($service.StartName -like "*$account*") {
            $services += [PSCustomObject]@{
              ComputerName = $computer
              DisplayName = $service.DisplayName
              StartName = $service.StartName
              State = $service.State
              ProcessId = $service.ProcessId
            }
          }
        }

        $invokeParams = @{ ComputerName = $computer; ScriptBlock = { schtasks.exe /query /s localhost /V /FO CSV | ConvertFrom-Csv }; ErrorAction = 'SilentlyContinue' }
        if ($Credential) { $invokeParams.Credential = $Credential }

        try {
          foreach ($task in @(Invoke-Command @invokeParams)) {
            if ($task.'Run As User' -like "*$account*") {
              $tasks += [PSCustomObject]@{
                ComputerName = $computer
                TaskName = $task.TaskName
                Status = $task.Status
                RunAsUser = $task.'Run As User'
              }
            }
          }
        }
        catch {
          Write-Verbose "Scheduled task query failed on $computer : $($_.Exception.Message)"
        }
      }

      [PSCustomObject]@{
        Name = $account
        Services = $services
        SchTasks = $tasks
      }
    }
  }
}
# Shared Restart Manager binding - compiled once, reused by Get-FileLockProcess.
# Technique sourced from pldmgg/misc-powershell (Get-FileLockProcess.ps1),
# itself derived from the well-established Restart Manager C# pattern.
if ($null -eq ('MyCore.Utils.FileLockUtil' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace MyCore.Utils
{
    static public class FileLockUtil
    {
        [StructLayout(LayoutKind.Sequential)]
        struct RM_UNIQUE_PROCESS
        {
            public int dwProcessId;
            public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
        }
        const int RmRebootReasonNone = 0;
        const int CCH_RM_MAX_APP_NAME = 255;
        const int CCH_RM_MAX_SVC_NAME = 63;
        enum RM_APP_TYPE
        {
            RmUnknownApp = 0,
            RmMainWindow = 1,
            RmOtherWindow = 2,
            RmService = 3,
            RmExplorer = 4,
            RmConsole = 5,
            RmCritical = 1000
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct RM_PROCESS_INFO
        {
            public RM_UNIQUE_PROCESS Process;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)]
            public string strAppName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)]
            public string strServiceShortName;
            public RM_APP_TYPE ApplicationType;
            public uint AppStatus;
            public uint TSSessionId;
            [MarshalAs(UnmanagedType.Bool)]
            public bool bRestartable;
        }
        [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
        static extern int RmRegisterResources(uint pSessionHandle,
            UInt32 nFiles,
            string[] rgsFilenames,
            UInt32 nApplications,
            [In] RM_UNIQUE_PROCESS[] rgApplications,
            UInt32 nServices,
            string[] rgsServiceNames);
        [DllImport("rstrtmgr.dll", CharSet = CharSet.Auto)]
        static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);
        [DllImport("rstrtmgr.dll")]
        static extern int RmEndSession(uint pSessionHandle);
        [DllImport("rstrtmgr.dll")]
        static extern int RmGetList(uint dwSessionHandle,
            out uint onProcInfoNeeded,
            ref uint onProcInfo,
            [In, Out] RM_PROCESS_INFO[] rgAffectedApps,
            ref uint lpdwRebootReasons);

        public class FileLockInfo
        {
            public int ProcessId;
            public string AppName;
            public string ServiceShortName;
            public int ApplicationType;
            public uint TSSessionId;
            public bool Restartable;
        }

        /// <summary>Find out what process(es) have a lock on the specified file.</summary>
        /// <param name="path">Path of the file.</param>
        /// <returns>Processes locking the file.</returns>
        static public List<FileLockInfo> WhoIsLocking(string path)
        {
            uint handle;
            string key = Guid.NewGuid().ToString();
            List<FileLockInfo> processes = new List<FileLockInfo>();
            int res = RmStartSession(out handle, 0, key);
            if (res != 0) throw new Exception("Could not begin restart session. Unable to determine file locker.");
            try
            {
                const int ERROR_MORE_DATA = 234;
                uint onProcInfoNeeded = 0,
                onProcInfo = 0,
                lpdwRebootReasons = RmRebootReasonNone;
                string[] resources = new string[] { path };
                res = RmRegisterResources(handle, (uint)resources.Length, resources, 0, null, 0, null);
                if (res != 0) throw new Exception("Could not register resource.");
                res = RmGetList(handle, out onProcInfoNeeded, ref onProcInfo, null, ref lpdwRebootReasons);
                if (res == ERROR_MORE_DATA)
                {
                    RM_PROCESS_INFO[] processInfo = new RM_PROCESS_INFO[onProcInfoNeeded];
                    onProcInfo = onProcInfoNeeded;
                    res = RmGetList(handle, out onProcInfoNeeded, ref onProcInfo, processInfo, ref lpdwRebootReasons);
                    if (res == 0)
                    {
                        processes = new List<FileLockInfo>((int)onProcInfo);
                        for (int i = 0; i < onProcInfo; i++)
                        {
                            FileLockInfo info = new FileLockInfo();
                            info.ProcessId = processInfo[i].Process.dwProcessId;
                            info.AppName = processInfo[i].strAppName;
                            info.ServiceShortName = processInfo[i].strServiceShortName;
                            info.ApplicationType = (int)processInfo[i].ApplicationType;
                            info.TSSessionId = processInfo[i].TSSessionId;
                            info.Restartable = processInfo[i].bRestartable;
                            processes.Add(info);
                        }
                    }
                    else throw new Exception("Could not list processes locking resource.");
                }
                else if (res != 0) throw new Exception("Could not list processes locking resource. Failed to get size of result.");
            }
            finally
            {
                RmEndSession(handle);
            }
            return processes;
        }
    }
}
'@ -ErrorAction Stop
}

function Get-FileLockProcess {
  <#
    .SYNOPSIS
      Get-FileLockProcess - Finds the process(es) holding a file open.
    .DESCRIPTION
      Uses the Windows Restart Manager API (RmStartSession, RmRegisterResources,
      RmGetList - the same mechanism Windows itself uses for "this file is in
      use by X" during updates/uninstalls) to determine which running
      process(es) hold a lock on the specified file.

      Returns one structured result per file with a Lockers collection carrying
      the process ID, process name, the Restart Manager application name and
      type, session ID, and restartable flag.
    .PARAMETER FilePath
      Path(s) of the file(s) to check.
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result with a Lockers
      property.
    .EXAMPLE
      PS> Get-FileLockProcess -FilePath 'C:\Temp\locked.dat'
    .LINK
      https://github.com/adnoctem/winkit/lib/system.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
    [string[]]
    $FilePath
  )

  process {
    foreach ($path in $FilePath) {
      $resolvedPath = $null
      try {
        $resolvedPath = (Resolve-Path -LiteralPath $path -ErrorAction Stop).Path
      }
      catch {
        New-OperationResult -Target $path -Source 'RestartManager' -Action 'FindLockers' -Status 'Failed' -ErrorMessage "Path does not exist: $path"
        continue
      }

      try {
        $lockInfos = [MyCore.Utils.FileLockUtil]::WhoIsLocking($resolvedPath)

        $lockers = @()
        foreach ($info in $lockInfos) {
          $processName = $null
          try {
            $processName = (Get-Process -Id $info.ProcessId -ErrorAction Stop).ProcessName
          }
          catch {
            Write-Verbose "Process $($info.ProcessId) is no longer running."
          }

          $lockers += [PSCustomObject]@{
            ProcessId = $info.ProcessId
            ProcessName = $processName
            AppName = $info.AppName
            SessionId = $info.TSSessionId
            ApplicationType = switch ([int]$info.ApplicationType) {
              0 { 'RmUnknownApp' }
              1 { 'RmMainWindow' }
              2 { 'RmOtherWindow' }
              3 { 'RmService' }
              4 { 'RmExplorer' }
              5 { 'RmConsole' }
              1000 { 'RmCritical' }
              default { "Unknown($($info.ApplicationType))" }
            }
            Restartable = $info.Restartable
          }
        }

        New-OperationResult -Target $resolvedPath -Source 'RestartManager' -Action 'FindLockers' -Status 'Completed' -Detail "$($lockers.Count) process(es) hold this file open." -Property @{ Lockers = $lockers }
      }
      catch {
        New-OperationResult -Target $resolvedPath -Source 'RestartManager' -Action 'FindLockers' -Status 'Failed' -ErrorMessage $_.Exception.Message
      }
    }
  }
}
