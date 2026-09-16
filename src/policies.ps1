#Requires -Version 5.0

function Resolve-LGPOSource {
  <#
    .SYNOPSIS
      Returns metadata for a known LGPO source location.
    .DESCRIPTION
      Pure data lookup. No I/O. Centralises the "where do we get LGPO from"
      decision so that every other function in this module can reference it
      without duplicating URLs or hashes.

      When Microsoft moves the file or publishes a new SCT release, update
      the table below and have a human review the diff. This function is the
      single point of change for supply-chain trust.
    .PARAMETER Source
      Source identifier. Currently only 'SCT-LGPO-Standalone' is recognised.
    .EXAMPLE
      PS> Resolve-LGPOSource
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('SCT-LGPO-Standalone')]
    [string]
    $Source = 'SCT-LGPO-Standalone'
  )

  $sources = @{
    'SCT-LGPO-Standalone' = [PSCustomObject]@{
      Name = 'Security Compliance Toolkit - LGPO standalone'
      Url = 'https://download.microsoft.com/download/8/5/C/85C25433-A1B0-4FFA-9429-7E023E7DA8D8/LGPO.zip'
      Sha256 = 'PLACEHOLDER_REPLACE_ON_FIRST_VENDORING'
      ExpectedBinaryPath = 'LGPO_30/LGPO.exe'
      LastVerified = '2026-06-14'
    }
  }

  if (-not $sources.ContainsKey($Source)) {
    throw "Unknown LGPO source '$Source'."
  }
  return $sources[$Source]
}

function Test-LGPOSourceAvailability {
  <#
    .SYNOPSIS
      Verifies the LGPO download URL is still reachable.
    .DESCRIPTION
      Issues a HEAD request to the URL returned by Resolve-LGPOSource.
      Does NOT download or verify the content - that's Install-LGPO's job.

      Intended for weekly CI runs. The output object is suitable for
      serialising to JSON and archiving as ISO supply-chain evidence
      ("we monitor external dependencies weekly").
    .PARAMETER Source
      Source identifier forwarded to Resolve-LGPOSource.
    .EXAMPLE
      PS> Test-LGPOSourceAvailability
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $false)]
    [ValidateSet('SCT-LGPO-Standalone')]
    [string]
    $Source = 'SCT-LGPO-Standalone'
  )

  $info = Resolve-LGPOSource -Source $Source
  try {
    $response = Invoke-WebRequest -Uri $info.Url -Method Head -UseBasicParsing -ErrorAction Stop
    [PSCustomObject]@{
      Source = $Source
      Url = $info.Url
      Available = $true
      StatusCode = [int]$response.StatusCode
      ContentLength = $response.Headers['Content-Length']
      CheckedAt = (Get-Date).ToUniversalTime()
    }
  }
  catch {
    [PSCustomObject]@{
      Source = $Source
      Url = $info.Url
      Available = $false
      Error = $_.Exception.Message
      CheckedAt = (Get-Date).ToUniversalTime()
    }
  }
}

function Install-LGPO {
  <#
    .SYNOPSIS
      Downloads, verifies, and installs LGPO.exe to a controllable path.
    .DESCRIPTION
      Hash-verified install. Refuses to proceed if the SHA-256 of the
      downloaded archive does not match Resolve-LGPOSource's recorded hash.

      Idempotent: re-running when LGPO.exe is already present at the
      destination returns the existing path unless -Force is specified.

      Writing to the default destination (%ProgramData%) requires
      administrator elevation. A non-elevated session can specify an
      alternate -Destination within the user's writeable scope.
    .PARAMETER Destination
      Directory where LGPO.exe ends up. Defaults to %ProgramData%\winkit\tools.
      Non-elevated callers should supply a user-writeable path.
    .PARAMETER Source
      Source identifier forwarded to Resolve-LGPOSource.
    .PARAMETER Force
      Re-download and re-install even if LGPO.exe is already present.
    .EXAMPLE
      PS> Install-LGPO
    .EXAMPLE
      PS> Install-LGPO -Force -Verbose
    .EXAMPLE
      PS> Install-LGPO -Destination "$env:LOCALAPPDATA\winkit\tools"
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([string])]
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $Destination = (Join-Path -Path $env:ProgramData -ChildPath 'winkit\tools'),

    [Parameter(Mandatory = $false)]
    [ValidateSet('SCT-LGPO-Standalone')]
    [string]
    $Source = 'SCT-LGPO-Standalone',

    [Parameter(Mandatory = $false)]
    [switch]
    $Force
  )

  $info = Resolve-LGPOSource -Source $Source
  $exePath = Join-Path -Path $Destination -ChildPath 'LGPO.exe'

  if (Test-Path -LiteralPath $exePath -PathType Leaf) {
    if (-not $Force) {
      Write-Verbose "LGPO.exe already present at '$exePath'; skipping download."
      return $exePath
    }
    Write-Verbose "LGPO.exe already present at '$exePath'; -Force supplied, re-downloading."
  }

  if (-not $PSCmdlet.ShouldProcess($Destination, "Install LGPO from $($info.Url)")) {
    return
  }

  if (-not (Read-ProcessElevation)) {
    Write-Error "Writing to '$Destination' requires administrator rights. Run the session elevated or supply a user-writeable -Destination such as '$(Join-Path -Path $env:LOCALAPPDATA -ChildPath 'winkit\tools')'."
    return
  }

  $null = New-Item -Path $Destination -ItemType Directory -Force -ErrorAction SilentlyContinue
  $zipPath = Join-Path -Path $env:TEMP -ChildPath "LGPO_$(New-Guid).zip"
  $extractDir = Join-Path -Path $env:TEMP -ChildPath "LGPO_extract_$(New-Guid)"

  try {
    Write-Verbose "Downloading LGPO from $($info.Url)..."
    Invoke-WebRequest -Uri $info.Url -OutFile $zipPath -UseBasicParsing -ErrorAction Stop

    $actualHash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash
    $expectedHash = $info.Sha256.ToUpperInvariant()
    if ($actualHash -ne $expectedHash) {
      throw "SHA-256 mismatch for $($info.Url). Expected '$expectedHash', got '$actualHash'. Refusing to install."
    }

    Write-Verbose 'SHA-256 hash matches. Extracting archive...'
    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
    $found = Get-ChildItem -Path $extractDir -Recurse -Filter 'LGPO.exe' | Select-Object -First 1
    if (-not $found) {
      throw "LGPO.exe not found in archive at $($info.Url)."
    }
    Copy-Item -Path $found.FullName -Destination $exePath -Force
    Write-Verbose "Installed LGPO.exe to '$exePath'."
  }
  finally {
    if (Test-Path -LiteralPath $zipPath -PathType Leaf) {
      Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $extractDir -PathType Container) {
      Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  return $exePath
}

function Test-LGPOInstalled {
  <#
    .SYNOPSIS
      Returns $true if LGPO.exe is present at the expected path.
    .PARAMETER Path
      Full path to LGPO.exe. Defaults to %ProgramData%\winkit\tools\LGPO.exe.
    .EXAMPLE
      PS> if (Test-LGPOInstalled) { Invoke-LGPO -PolicyPath .\policy.txt }
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  [OutputType([bool])]
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $Path = (Join-Path -Path $env:ProgramData -ChildPath 'winkit\tools\LGPO.exe')
  )

  return (Test-Path -LiteralPath $Path -PathType Leaf)
}

function Invoke-LGPO {
  <#
    .SYNOPSIS
      Applies a policy file or GPO backup directory using LGPO.exe.
    .DESCRIPTION
      Accepts a path to either a policy text file (applied via /t) or a
      directory containing a GPO backup (applied via /g). The argument is
      auto-detected based on whether PolicyPath is a file or directory.

      Returns a structured object describing the apply, suitable for
      logging as audit evidence.
    .PARAMETER PolicyPath
      Path to a policy text file or a directory containing a GPO backup.
    .PARAMETER LgpoExe
      Path to LGPO.exe. Defaults to %ProgramData%\winkit\tools\LGPO.exe.
    .EXAMPLE
      PS> Invoke-LGPO -PolicyPath .\resources\policies\01-telemetry.txt
    .LINK
      https://github.com/adnoctem/winkit/lib/policies.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding(SupportsShouldProcess)]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [string]
    $PolicyPath,

    [Parameter(Mandatory = $false)]
    [string]
    $LgpoExe = (Join-Path -Path $env:ProgramData -ChildPath 'winkit\tools\LGPO.exe')
  )

  if (-not (Test-Path -LiteralPath $LgpoExe -PathType Leaf)) {
    throw "LGPO.exe not found at '$LgpoExe'. Run Install-LGPO first."
  }

  $isDirectory = Test-Path -LiteralPath $PolicyPath -PathType Container
  $arg = if ($isDirectory) { '/g' } else { '/t' }

  if (-not $PSCmdlet.ShouldProcess($PolicyPath, "Apply via LGPO.exe ($arg)")) {
    return
  }

  $stdoutFile = Join-Path -Path $env:TEMP -ChildPath "lgpo_stdout_$(New-Guid).log"
  $stderrFile = Join-Path -Path $env:TEMP -ChildPath "lgpo_stderr_$(New-Guid).log"
  try {
    $proc = Start-Process -FilePath $LgpoExe `
      -ArgumentList @($arg, $PolicyPath) `
      -RedirectStandardOutput $stdoutFile `
      -RedirectStandardError $stderrFile `
      -Wait -NoNewWindow -PassThru

    [PSCustomObject]@{
      PolicyPath = $PolicyPath
      Mode = if ($isDirectory) { 'GpoBackup' } else { 'TextSource' }
      ExitCode = $proc.ExitCode
      StdOut = (Get-Content -LiteralPath $stdoutFile -Raw -ErrorAction SilentlyContinue)
      StdErr = (Get-Content -LiteralPath $stderrFile -Raw -ErrorAction SilentlyContinue)
      AppliedAt = (Get-Date).ToUniversalTime()
      Success = ($proc.ExitCode -eq 0)
    }
  }
  finally {
    if (Test-Path -LiteralPath $stdoutFile -PathType Leaf) {
      Remove-Item -LiteralPath $stdoutFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $stderrFile -PathType Leaf) {
      Remove-Item -LiteralPath $stderrFile -Force -ErrorAction SilentlyContinue
    }
  }
}

# Binary codec helpers are private; only the two converters are exported.
function Read-PSFPolicyString {
  param ([System.IO.BinaryReader]$Reader)
  $text = [Text.StringBuilder]::new()
  while ($true) {
    $character = $Reader.ReadUInt16()
    if ($character -eq 0) { return $text.ToString() }
    if ($text.Length -ge 32767) { throw 'Policy identifier exceeds 32767 characters.' }
    [void]$text.Append([char]$character)
  }
}

function Assert-PSFPolicyDelimiter {
  param ([System.IO.BinaryReader]$Reader, [char]$Expected)
  if ($Reader.ReadUInt16() -ne [uint16]$Expected) {
    throw "Expected policy delimiter '$Expected'."
  }
}

function ConvertFrom-PSFPolicyPayload {
  param ([uint32]$Type, [byte[]]$Bytes)
  if ($Bytes.Length -eq 0) { return $null }
  switch ($Type) {
    { $_ -in 1, 2, 7 } {
      if ($Bytes.Length % 2 -ne 0) { throw 'String payload has an odd byte count.' }
      $encoding = [Text.UnicodeEncoding]::new($false, $false, $true)
      $text = $encoding.GetString($Bytes)
      if (-not $text.EndsWith([string][char]0)) { throw 'String payload is not null terminated.' }
      if ($Type -eq 7) {
        if (-not $text.EndsWith(([string][char]0) * 2)) { throw 'MULTI_SZ requires two terminating nulls.' }
        $text = $text.Substring(0, $text.Length - 2)
        if ($text.Length -eq 0) { return , ([string[]]@()) }
        return , ([string[]]$text.Split([char]0))
      }
      return $text.Substring(0, $text.Length - 1)
    }
    4 {
      if ($Bytes.Length -ne 4) { throw 'DWORD payload must contain four bytes.' }
      return [BitConverter]::ToUInt32($Bytes, 0)
    }
    5 {
      if ($Bytes.Length -ne 4) { throw 'DWORD_BIG_ENDIAN payload must contain four bytes.' }
      $copy = [byte[]]$Bytes.Clone()
      [array]::Reverse($copy)
      return [BitConverter]::ToUInt32($copy, 0)
    }
    11 {
      if ($Bytes.Length -ne 8) { throw 'QWORD payload must contain eight bytes.' }
      return [BitConverter]::ToUInt64($Bytes, 0)
    }
    default { return , $Bytes }
  }
}

function ConvertTo-PSFPolicyPayload {
  param ([uint32]$Type, [AllowNull()][object]$Data, [switch]$Raw)
  if ($null -eq $Data) { return , ([byte[]]@()) }
  if ($Raw -or $Type -notin 1, 2, 4, 5, 7, 11) {
    if ($Data -isnot [byte[]]) { throw 'Raw and opaque policy data must be a byte array.' }
    return , $Data
  }
  switch ($Type) {
    { $_ -in 1, 2 } {
      if ($Data -isnot [string] -or $Data.Contains([string][char]0)) {
        throw 'SZ and EXPAND_SZ data must be a string without embedded nulls.'
      }
      return , ([Text.Encoding]::Unicode.GetBytes($Data + [char]0))
    }
    7 {
      foreach ($item in $Data) {
        if ($item -isnot [string] -or $item.Length -eq 0 -or $item.Contains([string][char]0)) {
          throw 'MULTI_SZ data must contain nonempty strings without embedded nulls.'
        }
      }
      return , ([Text.Encoding]::Unicode.GetBytes(($Data -join [char]0) + ([string][char]0) * 2))
    }
    { $_ -in 4, 5, 11 } {
      # Avoid PowerShell silently rounding fractional numbers during a cast.
      $integerText = [Convert]::ToString($Data, [Globalization.CultureInfo]::InvariantCulture)
      if ($integerText -notmatch '^\d+$') { throw 'Integer policy data must be an unsigned whole number.' }
      if ($Type -eq 11) { return , ([BitConverter]::GetBytes([uint64]::Parse($integerText))) }
      $bytes = [BitConverter]::GetBytes([uint32]::Parse($integerText))
      if ($Type -eq 5) { [array]::Reverse($bytes) }
      return , $bytes
    }
  }
}

function ConvertFrom-RegistryPolicy {
  <#
    .SYNOPSIS
      Reads registry.pol records without applying policy or requiring LGPO.
    .DESCRIPTION
      Validates a PReg version 1 binary file and emits records in file order,
      retaining duplicates and special policy directives. Keys are relative to
      a hive; the file's Machine/User location determines that hive.
      Type is a numeric registry type. Data is a string, string array, UInt32,
      UInt64, or byte array. Zero-length data is null. Unknown types remain
      opaque bytes. EXPAND_SZ variables are not expanded. Files are limited to
      64 MiB and individual payloads to 65535 bytes. Malformed files terminate
      with path and byte-offset information; no partial records are emitted.
    .PARAMETER Path
      Literal filesystem path to a registry.pol file.
    .PARAMETER Raw
      Return every payload as bytes, including empty arrays, tagged with the
      PSFoundation.RegistryPolicy.RawEntry type name. The writer recognizes
      this type for lossless binary round trips, including opaque data.
    .EXAMPLE
      PS> ConvertFrom-RegistryPolicy -Path '.\Machine\registry.pol'
    .EXAMPLE
      PS> ConvertFrom-RegistryPolicy -Path '.\source.pol' -Raw | ConvertTo-RegistryPolicy -Path '.\copy.pol'
    .OUTPUTS
      PSCustomObject with Key, ValueName, Type, and Data properties.
    .LINK
      https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-gpreg/5c092c22-bf6b-4e7f-b180-b20743d368f5
  #>
  [CmdletBinding()]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,
    [switch]$Raw
  )

  $filePath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
  $stream = [IO.File]::OpenRead($filePath)
  $reader = [IO.BinaryReader]::new($stream)
  $entries = [Collections.Generic.List[object]]::new()
  try {
    if ($stream.Length -gt 64MB) { throw 'Policy file exceeds the 64 MiB limit.' }
    if ($reader.ReadUInt32() -ne 0x67655250) { throw 'Invalid PReg signature.' }
    if ($reader.ReadUInt32() -ne 1) { throw 'Unsupported PReg version (expected 1).' }
    while ($stream.Position -lt $stream.Length) {
      Assert-PSFPolicyDelimiter -Reader $reader -Expected '['
      $key = Read-PSFPolicyString -Reader $reader
      if ([string]::IsNullOrEmpty($key)) { throw 'Policy key must not be empty.' }
      Assert-PSFPolicyDelimiter -Reader $reader -Expected ';'
      $name = Read-PSFPolicyString -Reader $reader
      Assert-PSFPolicyDelimiter -Reader $reader -Expected ';'
      $type = $reader.ReadUInt32()
      Assert-PSFPolicyDelimiter -Reader $reader -Expected ';'
      $size = $reader.ReadUInt32()
      Assert-PSFPolicyDelimiter -Reader $reader -Expected ';'
      if ($size -gt 65535 -or $size -gt ($stream.Length - $stream.Position - 2)) {
        throw 'Invalid or truncated policy payload size.'
      }
      $bytes = $reader.ReadBytes([int]$size)
      Assert-PSFPolicyDelimiter -Reader $reader -Expected ']'
      $entry = [PSCustomObject]@{ Key = $key; ValueName = $name; Type = $type; Data = $bytes }
      if ($Raw) {
        $entry.PSObject.TypeNames.Insert(0, 'PSFoundation.RegistryPolicy.RawEntry')
      }
      else {
        $entry.Data = ConvertFrom-PSFPolicyPayload -Type $type -Bytes $bytes
        $entry.PSObject.TypeNames.Insert(0, 'PSFoundation.RegistryPolicy.Entry')
      }
      $entries.Add($entry)
    }
  }
  catch {
    throw [IO.InvalidDataException]::new("Invalid registry policy '$filePath' at byte $($stream.Position): $($_.Exception.Message)", $_.Exception)
  }
  finally { $reader.Dispose() }
  $entries.ToArray()
}

function ConvertTo-RegistryPolicy {
  <#
    .SYNOPSIS
      Writes ordered records to a registry.pol file without applying policy.
    .DESCRIPTION
      Accepts records with Key, ValueName, numeric Type, and Data. Supports
      decoded records and RawEntry records from ConvertFrom-RegistryPolicy.
      Preserves order, duplicates, and directive names. Raw entries retain
      payload bytes; other records are encoded by registry type. Null data
      encodes a zero-byte payload. Unknown types require byte arrays.
      Validates and serializes the entire input before creating a temporary
      sibling file and atomically replacing the destination. Empty input writes
      a header-only file. Parent directories must already exist. Limits match
      the reader: 64 MiB per file, 65535 bytes per payload.
    .PARAMETER InputObject
      One record or an array of records, also accepted from the pipeline.
    .PARAMETER Path
      Literal filesystem destination. Machine/User scope is chosen by the caller.
    .PARAMETER Force
      Allow replacement of an existing destination file.
    .EXAMPLE
      PS> $entries | ConvertTo-RegistryPolicy -Path '.\registry.pol' -Force -WhatIf
    .EXAMPLE
      PS> ConvertTo-RegistryPolicy -InputObject @() -Path '.\empty.pol'
    .OUTPUTS
      None. Writes the destination only after validation and ShouldProcess.
  #>
  [CmdletBinding(SupportsShouldProcess = $true)]
  [OutputType([void])]
  param (
    [Parameter(ValueFromPipeline = $true)]
    [AllowEmptyCollection()]
    [object[]]$InputObject = @(),
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [switch]$Force
  )

  begin { $records = [Collections.Generic.List[object]]::new() }
  process {
    foreach ($record in $InputObject) { $records.Add($record) }
  }
  end {
    $filePath = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)
    if ([IO.File]::Exists($filePath) -and -not $Force) {
      throw "Destination '$filePath' exists. Use -Force to replace it."
    }
    $buffer = [IO.MemoryStream]::new()
    $writer = [IO.BinaryWriter]::new($buffer)
    $temporaryPath = $null
    try {
      $writer.Write([uint32]0x67655250)
      $writer.Write([uint32]1)
      foreach ($record in $records) {
        if ($null -eq $record) { throw 'Policy records must not be null.' }
        $entry = [PSCustomObject]$record
        foreach ($property in @('Key', 'ValueName', 'Type', 'Data')) {
          if ($null -eq $entry.PSObject.Properties[$property]) { throw "Policy record is missing '$property'." }
        }
        if ($entry.Key -isnot [string] -or [string]::IsNullOrEmpty($entry.Key) -or
          $entry.Key -match '^(HKLM|HKCU|HKEY_LOCAL_MACHINE|HKEY_CURRENT_USER)(:|\\|$)') {
          throw 'Policy Key must be a nonempty path relative to the registry hive.'
        }
        if ($entry.ValueName -isnot [string]) { throw 'Policy ValueName must be a string (empty is allowed).' }
        foreach ($identifier in @($entry.Key, $entry.ValueName)) {
          if ($identifier.Length -gt 32767 -or $identifier.Contains([string][char]0)) {
            throw 'Policy identifiers must not contain nulls or exceed 32767 characters.'
          }
        }
        if ([string]$entry.Type -notmatch '^\d+$') { throw 'Policy Type must be an unsigned integer.' }
        $type = [uint32]$entry.Type
        $isRaw = $entry.PSObject.TypeNames -contains 'PSFoundation.RegistryPolicy.RawEntry'
        $bytes = ConvertTo-PSFPolicyPayload -Type $type -Data $entry.Data -Raw:$isRaw
        if ($bytes.Length -gt 65535) { throw 'Policy payload exceeds 65535 bytes.' }
        $recordSize = 24L + 2L * ($entry.Key.Length + $entry.ValueName.Length) + $bytes.Length
        if ($buffer.Length + $recordSize -gt 64MB) { throw 'Policy file exceeds the 64 MiB limit.' }
        $writer.Write([uint16][char]'[')
        $writer.Write([Text.Encoding]::Unicode.GetBytes($entry.Key + [char]0))
        $writer.Write([uint16][char]';')
        $writer.Write([Text.Encoding]::Unicode.GetBytes($entry.ValueName + [char]0))
        $writer.Write([uint16][char]';')
        $writer.Write($type)
        $writer.Write([uint16][char]';')
        $writer.Write([uint32]$bytes.Length)
        $writer.Write([uint16][char]';')
        $writer.Write([byte[]]$bytes)
        $writer.Write([uint16][char]']')
      }
      $writer.Flush()
      if (-not $PSCmdlet.ShouldProcess($filePath, 'Write registry policy file')) { return }
      $temporaryPath = Join-Path ([IO.Path]::GetDirectoryName($filePath)) ([IO.Path]::GetRandomFileName())
      $output = [IO.File]::Open($temporaryPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
      try {
        $buffer.Position = 0
        $buffer.CopyTo($output)
        $output.Flush()
      }
      finally { $output.Dispose() }
      if ($Force -and [IO.File]::Exists($filePath)) {
        # PowerShell converts $null to an empty string for this .NET overload.
        [IO.File]::Replace($temporaryPath, $filePath, [NullString]::Value)
      }
      else { [IO.File]::Move($temporaryPath, $filePath) }
    }
    finally {
      $writer.Dispose()
      if ($temporaryPath -and [IO.File]::Exists($temporaryPath)) { [IO.File]::Delete($temporaryPath) }
    }
  }
}
