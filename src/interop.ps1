#Requires -Version 5.0

function Remove-ComObject {
  <#
    .SYNOPSIS
      Releases one or more COM objects.
    .DESCRIPTION
      Wraps Marshal.ReleaseComObject with null checks and best-effort error
      handling so scripts can safely clean up Outlook and Office interop
      objects from finally blocks.
    .EXAMPLE
      PS> Remove-ComObject $items $folder
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'COM cleanup must be best-effort during script teardown.')]
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Releases local COM references only; it does not change external system state.')]
  [CmdletBinding()]
  param (
    [Parameter(ValueFromRemainingArguments = $true)]
    [AllowNull()]
    [object[]]
    $InputObject
  )

  foreach ($_object in $InputObject) {
    if ($null -eq $_object) { continue }

    try {
      if ([System.Runtime.InteropServices.Marshal]::IsComObject($_object)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($_object)
      }
    }
    catch { }
  }
}

function Invoke-ComGarbageCollection {
  <#
    .SYNOPSIS
      Runs final COM cleanup garbage collection passes.
    .DESCRIPTION
      Forces garbage collection and waits for pending finalizers. This is useful
      after releasing Office COM references so Outlook can close PST files and
      exit cleanly when requested.
    .EXAMPLE
      PS> Invoke-ComGarbageCollection
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param ()

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Get-OutlookInstallation {
  <#
    .SYNOPSIS
      Finds local Outlook installation directories.
    .DESCRIPTION
      Discovers common Microsoft Office and Microsoft 365 installation roots
      that may contain Outlook.exe or Outlook data-file repair tools such as
      ScanPST.exe. The function checks App Paths registry
      entries first, then common Office directory layouts under Program Files.
    .EXAMPLE
      PS> Get-OutlookInstallation
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param ()

  $_candidateDirectories = New-Object System.Collections.Generic.List[string]
  $_registryPaths = @(
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
    'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
  )

  foreach ($_registryPath in $_registryPaths) {
    if (-not (Test-Path -LiteralPath $_registryPath)) { continue }

    $_property = Get-ItemProperty -LiteralPath $_registryPath -ErrorAction SilentlyContinue
    if (-not $_property) { continue }

    $_outlookPath = $_property.'(default)'
    if ([string]::IsNullOrWhiteSpace($_outlookPath) -and ($_property.PSObject.Properties.Name -contains 'Path')) {
      $_outlookPath = Join-Path -Path $_property.Path -ChildPath 'OUTLOOK.EXE'
    }

    if (-not [string]::IsNullOrWhiteSpace($_outlookPath)) {
      $_directory = Split-Path -Path ([Environment]::ExpandEnvironmentVariables($_outlookPath)) -Parent
      if (-not [string]::IsNullOrWhiteSpace($_directory)) {
        $_candidateDirectories.Add($_directory)
      }
    }
  }

  $_programRoots = @(
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)}
  ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

  $_officeVersions = @('Office16', 'Office15', 'Office14', 'Office12', 'Office11')
  foreach ($_programRoot in $_programRoots) {
    $_officeRoot = Join-Path -Path $_programRoot -ChildPath 'Microsoft Office'

    foreach ($_version in $_officeVersions) {
      $_candidateDirectories.Add((Join-Path -Path $_officeRoot -ChildPath $_version))
      $_candidateDirectories.Add((Join-Path -Path $_officeRoot -ChildPath "root\$_version"))
    }
  }

  $_seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($_directory in $_candidateDirectories) {
    if ([string]::IsNullOrWhiteSpace($_directory)) { continue }
    if (-not (Test-Path -LiteralPath $_directory -PathType Container)) { continue }

    $_resolvedDirectory = Resolve-LongPath -LiteralPath $_directory
    if (-not $_seen.Add($_resolvedDirectory)) { continue }

    $_outlookPath = Join-Path -Path $_resolvedDirectory -ChildPath 'OUTLOOK.EXE'
    $_scanPstPath = Join-Path -Path $_resolvedDirectory -ChildPath 'SCANPST.EXE'

    if (
      -not (Test-Path -LiteralPath $_outlookPath -PathType Leaf) -and
      -not (Test-Path -LiteralPath $_scanPstPath -PathType Leaf)
    ) {
      continue
    }

    [PSCustomObject]@{
      Path        = $_resolvedDirectory
      OutlookPath = if (Test-Path -LiteralPath $_outlookPath -PathType Leaf) { $_outlookPath } else { $null }
      ScanPstPath = if (Test-Path -LiteralPath $_scanPstPath -PathType Leaf) { $_scanPstPath } else { $null }
    }
  }
}

function Get-OutlookRepairToolInfo {
  <#
    .SYNOPSIS
      Reads repair-tool metadata and identifies supported ScanPST file targeting.
    .DESCRIPTION
      Inspects an existing executable without launching it. Only ScanPST from
      the Office 16 family at build 16.0.10325.20082 or later is classified as
      supporting the documented file argument and rescan execution mode.
      Older, unknown, and explicitly selected other executables remain interactive.
      SupportsFileArgument is a conservative inference from the executable's
      version resource, not a runtime capability probe. MSI and Click-to-Run
      file versions can differ; a false value selects the interactive fallback.
      The directory name alone is not evidence of command-line support.
    .PARAMETER LiteralPath
      Literal path to the repair executable, including a legacy explicit override.
    .EXAMPLE
      PS> Get-OutlookRepairToolInfo -LiteralPath 'C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE'
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $LiteralPath
  )

  $_path = Resolve-LongPath -LiteralPath $LiteralPath
  $_file = Get-Item -LiteralPath $_path -ErrorAction Stop
  if ($_file.PSIsContainer -or $_file.Extension -ne '.exe') {
    throw 'The repair tool must be an existing .exe file.'
  }

  $_name = [IO.Path]::GetFileNameWithoutExtension($_path)
  $_version = $null
  $_versionInfo = $_file.VersionInfo
  if ($_versionInfo -and $_versionInfo.FileMajorPart -gt 0) {
    $_version = [version]('{0}.{1}.{2}.{3}' -f $_versionInfo.FileMajorPart,
      $_versionInfo.FileMinorPart, $_versionInfo.FileBuildPart, $_versionInfo.FilePrivatePart)
  }
  $_supportsFileArgument = $_name -ieq 'SCANPST' -and $null -ne $_version -and
  $_version.Major -eq 16 -and $_version -ge [version]'16.0.10325.20082'

  [PSCustomObject]@{
    Name                 = $_name
    Path                 = $_path
    InstallationPath     = $_file.DirectoryName
    FileVersion          = $_version
    SupportsFileArgument = [bool]$_supportsFileArgument
  }
}

function Find-OutlookRepairTool {
  <#
    .SYNOPSIS
      Finds installed ScanPST repair tools and their targeting capabilities.
    .DESCRIPTION
      Searches Outlook installations, then application executables on PATH.
      Returned paths use long names and include executable version metadata.
      Legacy alternatives can be inspected explicitly with Get-OutlookRepairToolInfo.
    .PARAMETER Name
      ScanPST is the only automatically discovered repair tool.
    .EXAMPLE
      PS> Find-OutlookRepairTool -Name ScanPST
  #>

  [OutputType([PSCustomObject[]])]
  [CmdletBinding()]
  param (
    [ValidateSet('ScanPST')]
    [string]
    $Name = 'ScanPST'
  )

  $_seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($_installation in Get-OutlookInstallation) {
    $_path = $_installation.ScanPstPath
    if ([string]::IsNullOrWhiteSpace($_path)) {
      continue
    }
    $_tool = Get-OutlookRepairToolInfo -LiteralPath $_path
    if ($_seen.Add($_tool.Path)) {
      $_tool
    }
  }

  foreach ($_command in @(Get-Command -Name ($Name + '.exe') -CommandType Application -ErrorAction SilentlyContinue)) {
    $_tool = Get-OutlookRepairToolInfo -LiteralPath $_command.Source
    if ($_seen.Add($_tool.Path)) {
      $_tool
    }
  }
}

function Get-TransportMessageId {
  <#
    .SYNOPSIS
      Extracts the RFC Message-ID from transport header text.
    .DESCRIPTION
      Parses a raw Outlook transport header block and returns the first RFC 5322
      Message-ID field value, or $null when none is present. Headers may be
      supplied in Unicode or ANSI form; field matching is case-insensitive and
      line-based so embedded Received headers do not interfere.
    .PARAMETER HeaderText
      Raw transport header text as returned by Outlook (PR_TRANSPORT_MESSAGE_HEADERS).
    .EXAMPLE
      PS> Get-TransportMessageId -HeaderText "Message-ID: <abc123@example.com>`r`nReceived: ..."
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string]
    $HeaderText
  )

  if ([string]::IsNullOrWhiteSpace($HeaderText)) { return $null }

  $_match = [regex]::Match($HeaderText, '(?im)^Message-ID:\s*(<[^>]+>)')
  if ($_match.Success) {
    return $_match.Groups[1].Value.Trim()
  }

  return $null
}

function Connect-Outlook {
  <#
    .SYNOPSIS
      Connects to an Outlook COM application and MAPI namespace.
    .DESCRIPTION
      Reuses a running Outlook instance when available, otherwise starts one,
      then logs on to the MAPI namespace without prompting.
    .EXAMPLE
      PS> $context = Connect-Outlook
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingEmptyCatchBlock', '', Justification = 'Falling back to a new Outlook COM instance is intentional.')]
  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param ()

  try {
    $_application = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
  }
  catch {
    $_application = New-Object -ComObject Outlook.Application
  }

  $_namespace = $_application.GetNamespace('MAPI')
  $_namespace.Logon($null, $null, $false, $false)

  [PSCustomObject]@{
    App       = $_application
    Namespace = $_namespace
  }
}

function Get-OutlookStoreRoot {
  <#
    .SYNOPSIS
      Gets the root folder for an Outlook store.
    .DESCRIPTION
      Resolves a named Outlook store by DisplayName, or the default delivery
      store when no name is supplied. Store COM objects are released as they are
      inspected; the returned root folder is owned by the caller.
    .PARAMETER Namespace
      Outlook MAPI namespace returned by Connect-Outlook.
    .PARAMETER Name
      Optional store display name.
    .EXAMPLE
      PS> Get-OutlookStoreRoot -Namespace $context.Namespace -Name 'user@example.com'
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Namespace,

    [string]
    $Name
  )

  $_stores = $Namespace.Stores
  try {
    Write-Verbose 'Available Outlook stores:'
    for ($_index = 1; $_index -le $_stores.Count; $_index++) {
      $_store = $_stores.Item($_index)
      try {
        Write-Verbose ("  - {0}" -f $_store.DisplayName)
        if ([string]::IsNullOrWhiteSpace($Name) -and $_store.IsDefault) {
          return $_store.GetRootFolder()
        }

        if ($_store.DisplayName -eq $Name) {
          return $_store.GetRootFolder()
        }
      }
      finally {
        Remove-ComObject $_store
      }
    }
  }
  finally {
    Remove-ComObject $_stores
  }

  throw "Outlook store '$Name' not found. Run with -Verbose to list available stores."
}

function Add-OutlookStoreRoot {
  <#
    .SYNOPSIS
      Adds a Unicode PST store and returns its root folder.
    .DESCRIPTION
      Calls Outlook Namespace.AddStoreEx with OlStoreType.olStoreUnicode and
      locates the newly attached store by FilePath. The returned root folder is
      owned by the caller.
    .PARAMETER Namespace
      Outlook MAPI namespace returned by Connect-Outlook.
    .PARAMETER Path
      Full path to the PST file to attach or create.
    .EXAMPLE
      PS> Add-OutlookStoreRoot -Namespace $context.Namespace -Path 'D:\Archive\mail.pst'
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Namespace,

    [Parameter(Mandatory = $true)]
    [string]
    $Path
  )

  $_olStoreUnicode = 2
  $_resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
  $Namespace.AddStoreEx($_resolvedPath, $_olStoreUnicode)

  $_stores = $Namespace.Stores
  try {
    for ($_index = 1; $_index -le $_stores.Count; $_index++) {
      $_store = $_stores.Item($_index)
      try {
        if ($_store.FilePath -eq $_resolvedPath) {
          return $_store.GetRootFolder()
        }
      }
      finally {
        Remove-ComObject $_store
      }
    }
  }
  finally {
    Remove-ComObject $_stores
  }

  throw "Outlook store was added but could not be located by path: $_resolvedPath"
}

function Get-PSFOutlookPathDriveType {
  [CmdletBinding()]
  param ([string]$Path)

  (New-Object IO.DriveInfo([IO.Path]::GetPathRoot($Path))).DriveType
}

function Resolve-PSFOutlookFilePath {
  [CmdletBinding()]
  param (
    [string]$LiteralPath,
    [switch]$SourcePst
  )

  if ([string]::IsNullOrWhiteSpace($LiteralPath)) { throw 'A nonempty local file path is required.' }
  $_provider = $null
  $_drive = $null
  $_path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($LiteralPath, [ref]$_provider, [ref]$_drive)
  if ($_provider.Name -ne 'FileSystem' -or $_path -notmatch '^[A-Za-z]:\\') {
    throw "A local filesystem path is required: '$LiteralPath'."
  }
  if ((Get-PSFOutlookPathDriveType $_path) -notin @([IO.DriveType]::Fixed, [IO.DriveType]::Removable)) {
    throw "Network or unavailable drives are not supported: '$_path'."
  }
  $_file = Get-Item -LiteralPath $_path -Force -ErrorAction Stop
  if ($_file.PSIsContainer) { throw "Expected an existing file: '$_path'." }
  if ($SourcePst -and ($_file.Extension -ne '.pst' -or ($_file.Attributes -band [IO.FileAttributes]::ReadOnly))) {
    throw "The source must be an existing writable .pst file: '$_path'."
  }
  # Reject traversal rather than claiming to resolve symlink/junction identities.
  $_ancestor = $_path
  while ($_ancestor) {
    $_item = Get-Item -LiteralPath $_ancestor -Force -ErrorAction Stop
    if ($_item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
      throw "Reparse-point paths are not supported for Outlook file identity: '$_ancestor'."
    }
    $_ancestor = Split-Path -Path $_ancestor -Parent
  }
  Resolve-LongPath -LiteralPath $_path
}

function Get-PSFOutlookProperty {
  [CmdletBinding()]
  param ([object]$InputObject, [string]$Name)

  if ($null -eq $InputObject) { throw "Cannot read Outlook property '$Name' from a null reference." }
  $_property = $InputObject.PSObject.Properties[$Name]
  if ($null -eq $_property) { throw "Outlook property '$Name' is unavailable." }
  # Ordinary PowerShell property syntax can silently turn a throwing getter into
  # null. Calling the accessor as a method preserves the inspection failure.
  return , ($_property.get_Value())
}

function Get-PSFOutlookPstSnapshot {
  [CmdletBinding()]
  param ([object]$Namespace, [string]$Path)

  $_ids = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $_matches = New-Object 'Collections.Generic.List[object]'
  $_stores = $null
  try {
    $_stores = Get-PSFOutlookProperty $Namespace Stores
    $_count = Get-PSFOutlookProperty $_stores Count
    if ($null -eq $_count -or $_count -isnot [int] -or $_count -lt 0) { throw 'Outlook returned an invalid store count.' }
    for ($_index = 1; $_index -le $_count; $_index++) {
      $_store = $null
      try {
        $_store = $_stores.Item($_index)
        $_id = [string](Get-PSFOutlookProperty $_store StoreID)
        if ([string]::IsNullOrWhiteSpace($_id) -or -not $_ids.Add($_id)) { throw 'Outlook returned missing or duplicate StoreIDs.' }
        $_filePath = [string](Get-PSFOutlookProperty $_store FilePath)
        if ([string]::IsNullOrWhiteSpace($_filePath)) { continue }
        $_resolved = Resolve-PSFOutlookFilePath $_filePath
        if ([StringComparer]::OrdinalIgnoreCase.Equals($_resolved, $Path)) {
          $_matches.Add([PSCustomObject]@{ StoreId = $_id; DisplayName = [string](Get-PSFOutlookProperty $_store DisplayName) })
        }
      }
      finally { Remove-ComObject $_store }
    }
    [PSCustomObject]@{ Ids = $_ids; Matches = $_matches.ToArray() }
  }
  catch { throw (New-Object InvalidOperationException("Cannot inspect Outlook stores for PST '$Path': $($_.Exception.Message)", $_.Exception)) }
  finally { Remove-ComObject $_stores }
}

function Get-PSFOutlookPstRoot {
  [CmdletBinding()]
  param ([object]$Namespace, [string]$Path, [string]$StoreId)

  $_store = $null
  $_root = $null
  try {
    $_store = $Namespace.GetStoreFromID($StoreId)
    if ((Get-PSFOutlookProperty $_store StoreID) -ne $StoreId -or (Resolve-PSFOutlookFilePath (Get-PSFOutlookProperty $_store FilePath)) -ne $Path) {
      throw "Outlook store identity changed for PST '$Path'."
    }
    $_root = $_store.GetRootFolder()
    if ($null -eq $_root -or (Get-PSFOutlookProperty $_root StoreID) -ne $StoreId) { throw "Outlook returned no matching root for PST '$Path'." }
    $_owned = $_root
    $_root = $null
    return , $_owned
  }
  finally { Remove-ComObject $_root $_store }
}

function Remove-PSFOutlookPstAttachment {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Mandatory cleanup of a uniquely identified attachment created by Open; not a new optional user operation.')]
  [CmdletBinding()]
  param ([object]$Namespace, [string]$Path, [string]$StoreId, [object]$Root)

  $_snapshot = Get-PSFOutlookPstSnapshot -Namespace $Namespace -Path $Path
  if (-not $_snapshot.Ids.Contains($StoreId)) { return }
  if (@($_snapshot.Matches | Where-Object { $_.StoreId -ieq $StoreId }).Count -ne 1) {
    throw "The recorded Outlook StoreID no longer matches PST '$Path'."
  }
  $_temporaryRoot = $null
  try {
    if ($null -eq $Root) {
      $_temporaryRoot = Get-PSFOutlookPstRoot -Namespace $Namespace -Path $Path -StoreId $StoreId
      $Root = $_temporaryRoot
    }
    if ((Get-PSFOutlookProperty $Root StoreID) -ne $StoreId) { throw "The owned Outlook root no longer matches PST '$Path'." }
    $null = $Namespace.RemoveStore($Root)
  }
  finally { Remove-ComObject $_temporaryRoot }
}

function New-PSFOutlookPstContext {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-process lifetime object only.')]
  [CmdletBinding()]
  param ([string]$Path, [object]$Match, [object]$Root, [object]$Namespace, [bool]$AttachedByCall)

  # Keep the cleanup identity separate from the caller's informational fields.
  $_state = [PSCustomObject]@{ Path = $Path; StoreId = $Match.StoreId; Root = $Root; Namespace = $Namespace; AttachedByCall = $AttachedByCall; Closed = $false }
  [PSCustomObject]@{
    PSTypeName = 'PSFoundation.OutlookPstStoreContext'
    Path = $Path; StoreId = $Match.StoreId; DisplayName = $Match.DisplayName
    Root = $Root; Namespace = $Namespace; AttachedByCall = $AttachedByCall; Closed = $false
    _State     = $_state
  }
}

function Open-OutlookPstStore {
  <#
    .SYNOPSIS
      Opens an existing local PST and returns an ownership-aware lifetime context.
    .DESCRIPTION
      Reuses a uniquely matching profile attachment or adds the existing file with
      Namespace.AddStore. Does not create destinations, rename stores or choose a
      format. ANSI compatibility needs native validation with the installed Outlook.
      Requires a writable local PST; UNC, network drives and reparse traversal are
      rejected. Relative paths use PowerShell's current filesystem location and
      short names are expanded. Hard-link aliases are not detected.

      The returned PSFoundation.OutlookPstStoreContext owns Root and borrows Namespace.
      Path, StoreId and DisplayName describe the observed store; AttachedByCall marks
      attachment ownership and Closed tracks cleanup. Private _State is implementation
      state. Do not edit, serialize or reuse it across processes. Release child COM
      references before Close-OutlookPstStore, and keep the creating context alive
      until overlapping consumers finish; contexts are not reference-counted leases.

      Standalone WhatIf/declined confirmation returns no context when attachment is
      needed. An explicitly documented inspection preview may override WhatIf for
      this call only, then must close in finally. Outlook may update PST metadata.
      Existence is rechecked before AddStore, but Outlook has no atomic existing-only
      open. Do not remove/replace the file or change profile attachments concurrently.
    .PARAMETER Namespace
      Borrowed MAPI namespace, normally returned by Connect-Outlook. Never released here.
    .PARAMETER LiteralPath
      Existing writable local .pst file. Wildcards are literal; directories are rejected.
    .EXAMPLE
      $source = $null
      try {
        $source = Open-OutlookPstStore -Namespace $context.Namespace -LiteralPath '.\Archive.pst'
        if ($null -ne $source) { $source | Select-Object Path, StoreId, DisplayName, AttachedByCall }
      }
      finally { if ($null -ne $source) { Close-OutlookPstStore -Context $source } }
    .EXAMPLE
      Open-OutlookPstStore -Namespace $context.Namespace -LiteralPath 'D:\Archives\Old mail.pst' -WhatIf
      Previews a required attachment without adding it. A pre-existing match is returned normally.
    .LINK
      https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.addstore
  #>
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [object]$Namespace,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LiteralPath
  )

  $_path = Resolve-PSFOutlookFilePath -LiteralPath $LiteralPath -SourcePst
  $_snapshot = Get-PSFOutlookPstSnapshot -Namespace $Namespace -Path $_path
  if ($_snapshot.Matches.Count -gt 1) { throw "Multiple Outlook stores match PST '$_path'." }
  if (-not $_snapshot.Matches.Count) {
    if (-not $PSCmdlet.ShouldProcess($_path, 'Attach existing PST to the current Outlook profile')) { return }
    # Account for attachments that appeared while approval was pending.
    $_snapshot = Get-PSFOutlookPstSnapshot -Namespace $Namespace -Path $_path
    if ($_snapshot.Matches.Count -gt 1) { throw "Multiple Outlook stores match PST '$_path'." }
  }
  $_initialIds = $_snapshot.Ids
  $_attempted = $false
  $_ownedId = $null
  $_root = $null
  try {
    if (-not $_snapshot.Matches.Count) {
      if ((Resolve-PSFOutlookFilePath -LiteralPath $_path -SourcePst) -ne $_path) { throw "PST path changed before attachment: '$_path'." }
      $_attempted = $true
      $null = $Namespace.AddStore($_path)
      $_snapshot = Get-PSFOutlookPstSnapshot -Namespace $Namespace -Path $_path
      if ($_snapshot.Matches.Count -ne 1 -or $_initialIds.Contains($_snapshot.Matches[0].StoreId)) {
        throw "No uniquely identified new Outlook attachment for PST '$_path'."
      }
      $_ownedId = $_snapshot.Matches[0].StoreId
    }
    $_match = $_snapshot.Matches[0]
    $_root = Get-PSFOutlookPstRoot -Namespace $Namespace -Path $_path -StoreId $_match.StoreId
    $_context = New-PSFOutlookPstContext -Path $_path -Match $_match -Root $_root -Namespace $Namespace -AttachedByCall ([bool]$_ownedId)
    $_root = $null
    return $_context
  }
  catch {
    $_failure = $_
    if ($_attempted) {
      try {
        if (-not $_ownedId) {
          $_current = Get-PSFOutlookPstSnapshot -Namespace $Namespace -Path $_path
          if ($_current.Matches.Count -gt 1) { throw 'Cannot identify a unique attachment to clean up.' }
          if ($_current.Matches.Count -eq 1 -and -not $_initialIds.Contains($_current.Matches[0].StoreId)) {
            $_ownedId = $_current.Matches[0].StoreId
          }
        }
        if ($_ownedId) { Remove-PSFOutlookPstAttachment -Namespace $Namespace -Path $_path -StoreId $_ownedId -Root $_root }
      }
      catch {
        $_cleanup = "Cleanup of PST '$_path' could not be completed: $($_.Exception.Message)"
        $_failure.Exception.Data['OutlookPstCleanupError'] = $_cleanup
        $_failure.ErrorDetails = New-Object Management.Automation.ErrorDetails("$($_failure.Exception.Message) $_cleanup")
      }
    }
    throw $_failure
  }
  finally { Remove-ComObject $_root }
}

function Close-OutlookPstStore {
  <#
    .SYNOPSIS
      Releases an existing-PST context and detaches only its owned attachment.
    .DESCRIPTION
      Checks the recorded StoreID and path before detaching an attachment created by
      Open-OutlookPstStore. Pre-existing stores and replacement StoreIDs are left alone.
      Always releases the owned root, clears Root and sets Closed, even on failure.
      Subsequent calls do nothing. The namespace/application remain borrowed and live.
      Never deletes, renames, replaces, repairs or compacts files, or quits Outlook.

      This is mandatory finally cleanup, not a second optional profile operation:
      it does not prompt or honor inherited WhatIf suppression. It can only undo the
      attachment identified by its context. Detach/inspection failure is terminating
      and identifies the PST path; the context is closed but the attachment may remain.
      Callers must release all child COM references first and report cleanup failure
      without hiding any original processing error. Concurrent profile changes and
      unobservable reuse of the same store identity cannot be made transactional.
    .PARAMETER Context
      Live context returned by Open-OutlookPstStore. Guard null in the caller.
    .EXAMPLE
      if ($null -ne $source) { Close-OutlookPstStore -Context $source }
      Releases a context in finally without releasing its borrowed namespace.
  #>
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Required lifetime cleanup must undo the context-owned attachment even under inherited WhatIf; no new operation is authorized.')]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNull()]
    [object]$Context
  )

  if ($Context.PSObject.TypeNames -notcontains 'PSFoundation.OutlookPstStoreContext' -or -not $Context.PSObject.Properties['_State']) {
    throw 'Context must be a live object returned by Open-OutlookPstStore.'
  }
  $_state = $Context._State
  if ($_state.Closed) { return }
  try {
    if ($_state.AttachedByCall) {
      Remove-PSFOutlookPstAttachment -Namespace $_state.Namespace -Path $_state.Path -StoreId $_state.StoreId -Root $_state.Root
    }
  }
  catch { throw (New-Object InvalidOperationException("Cannot detach PST '$($_state.Path)': $($_.Exception.Message)", $_.Exception)) }
  finally {
    Remove-ComObject $_state.Root
    $_state.Root = $null
    $_state.Closed = $true
    $Context.Root = $null
    $Context.Closed = $true
  }
}

function Get-OutlookSubFolder {
  <#
    .SYNOPSIS
      Gets or creates an Outlook child folder.
    .DESCRIPTION
      Searches a parent folder's Folders collection by display name. When
      -Create is set, the folder is created if missing. The returned folder is
      owned by the caller.
    .PARAMETER ParentFolder
      Outlook parent folder.
    .PARAMETER Name
      Child folder display name.
    .PARAMETER Create
      Create the folder when it does not exist.
    .EXAMPLE
      PS> Get-OutlookSubFolder -ParentFolder $root -Name '_Review' -Create
    .LINK
      https://github.com/adnoctem/winkit/lib/interop.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $ParentFolder,

    [Parameter(Mandatory = $true)]
    [string]
    $Name,

    [switch]
    $Create
  )

  $_folders = $ParentFolder.Folders
  try {
    for ($_index = 1; $_index -le $_folders.Count; $_index++) {
      $_folder = $_folders.Item($_index)
      if ($_folder.Name -eq $Name) {
        return $_folder
      }

      Remove-ComObject $_folder
    }

    if ($Create) {
      return $_folders.Add($Name)
    }
  }
  finally {
    Remove-ComObject $_folders
  }

  return $null
}

function Get-OutlookStandardFolderIdentity {
  <#
    .SYNOPSIS
      Reads locale-independent standard folder identities for one Outlook store.
    .DESCRIPTION
      Returns plain records keyed by StoreID and EntryID. Outlook 2010 and later
      use Store.GetDefaultFolder. Outlook 2007 uses read-only MAPI properties,
      with the default Inbox as an additional property source. No localized
      folder names are used. Missing optional properties are reported separately
      from unexpected provider errors. Does not create optional default folders.
    .PARAMETER Namespace
      Connected Outlook MAPI namespace.
    .PARAMETER StoreRoot
      Root folder of the selected store, owned by the caller.
    .EXAMPLE
      PS> Get-OutlookStandardFolderIdentity -Namespace $context.Namespace -StoreRoot $root
  #>

  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Namespace,

    [Parameter(Mandatory = $true)]
    [object]
    $StoreRoot
  )

  $_definitions = [ordered]@{
    DeletedItems      = 3
    Outbox            = 4
    SentItems         = 5
    Inbox             = 6
    Calendar          = 9
    Contacts          = 10
    Journal           = 11
    Notes             = 12
    Tasks             = 13
    Drafts            = 16
    AllPublicFolders  = 18
    Conflicts         = 19
    SyncIssues        = 20
    LocalFailures     = 21
    ServerFailures    = 22
    Junk              = 23
    RssFeeds          = 25
    ToDo              = 28
    ManagedEmail      = 29
    SuggestedContacts = 30
  }

  $_store = $null
  $_default = $null
  $_inbox = $null
  $_accessors = New-Object Collections.ArrayList
  try {
    $_store = $StoreRoot.Store
    $_application = $Namespace.Application
    try {
      $_major = [int](($_application.Version -split '\.')[0])
    }
    finally {
      Remove-ComObject $_application
    }

    if ($_major -ge 14) {
      foreach ($_definition in $_definitions.GetEnumerator()) {
        $_folder = $null
        try {
          $_folder = $_store.GetDefaultFolder($_definition.Value)
          if ($_folder -and ([string]::IsNullOrWhiteSpace([string]$_folder.EntryID) -or $_folder.StoreID -ne $StoreRoot.StoreID)) {
            throw 'Standard folder identity is empty or belongs to another store.'
          }
          [PSCustomObject]@{
            Kind     = $_definition.Key
            StoreID  = [string]$StoreRoot.StoreID
            EntryID  = if ($_folder) { [string]$_folder.EntryID } else { $null }
            State    = if ($_folder) { 'Resolved' } else { 'Absent' }
            Evidence = 'Store.GetDefaultFolder'
          }
        }
        catch {
          $_exception = $_.Exception
          while ($_exception.InnerException) {
            $_exception = $_exception.InnerException
          }
          # MAPI_E_NOT_FOUND and MAPI_E_NO_SUPPORT describe unavailable optional
          # folders. Access denied, disconnected providers, and other errors fail.
          if ($_exception.HResult -notin @(-2147221233, -2147221246) -and
            -not ($_exception.HResult -eq -2147024809 -and $_definition.Key -in @('AllPublicFolders', 'ManagedEmail', 'SuggestedContacts', 'ToDo', 'RssFeeds'))) {
            throw "Cannot determine standard folder '$($_definition.Key)' in the selected store: $($_.Exception.Message)"
          }
          [PSCustomObject]@{
            Kind     = $_definition.Key
            StoreID  = [string]$StoreRoot.StoreID
            EntryID  = $null
            State    = 'Unavailable'
            Evidence = 'Store.GetDefaultFolder'
          }
        }
        finally {
          Remove-ComObject $_folder
        }
      }
      return
    }

    $_known = @{}
    if (-not $_store.IsDataFileStore -or [IO.Path]::GetExtension([string]$_store.FilePath) -ne '.pst') {
      throw 'Outlook 2007 standard-folder discovery currently supports PST stores only. Use a newer classic Outlook client for Exchange or OST stores.'
    }
    $_default = $Namespace.DefaultStore
    if ($_default.StoreID -eq $StoreRoot.StoreID) {
      # Inbox is mandatory in a default delivery store. Do not request optional
      # folders through Namespace.GetDefaultFolder, which can create them.
      $_inbox = $Namespace.GetDefaultFolder(6)
      if ($_inbox.StoreID -ne $StoreRoot.StoreID) {
        throw 'Default Inbox belongs to a different store.'
      }
      $_known.Inbox = [string]$_inbox.EntryID
      $null = $_accessors.Add($_inbox.PropertyAccessor)
    }
    $null = $_accessors.Add($StoreRoot.PropertyAccessor)
    $null = $_accessors.Add($_store.PropertyAccessor)

    $_tags = [ordered]@{
      Outbox       = '35E2'
      DeletedItems = '35E3'
      SentItems    = '35E4'
      Calendar     = '36D0'
      Contacts     = '36D1'
      Journal      = '36D2'
      Notes        = '36D3'
      Tasks        = '36D4'
      Drafts       = '36D7'
    }

    foreach ($_accessor in $_accessors) {
      foreach ($_tag in $_tags.GetEnumerator()) {
        if ($_known.ContainsKey($_tag.Key)) {
          continue
        }
        try {
          $_binary = $_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x$($_tag.Value)0102")
          if ($_binary -and $_binary.Length -gt 0) {
            $_known[$_tag.Key] = $_accessor.BinaryToString($_binary)
          }
        }
        catch {
          $_exception = $_.Exception
          while ($_exception.InnerException) {
            $_exception = $_exception.InnerException
          }
          if ($_exception.HResult -ne -2147221233) {
            throw
          }
        }
      }

      try {
        $_additional = $_accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36D81102')
        $_kinds = @('Conflicts', 'SyncIssues', 'LocalFailures', 'ServerFailures', 'Junk')
        for ($_index = 0; $_index -lt [math]::Min($_additional.Length, $_kinds.Count); $_index++) {
          if ($_additional[$_index] -and -not $_known.ContainsKey($_kinds[$_index])) {
            $_known[$_kinds[$_index]] = $_accessor.BinaryToString($_additional[$_index])
          }
        }
      }
      catch {
        $_exception = $_.Exception
        while ($_exception.InnerException) {
          $_exception = $_exception.InnerException
        }
        if ($_exception.HResult -ne -2147221233) {
          throw
        }
      }

      # PR_ADDITIONAL_REN_ENTRYIDS_EX contains bounded PersistData blocks.
      try {
        [byte[]]$_data = $_accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36D90102')
        $_offset = 0
        while ($_offset -lt $_data.Length) {
          if ($_data.Length - $_offset -lt 4) {
            throw 'Truncated PersistData header.'
          }
          $_id = [BitConverter]::ToUInt16($_data, $_offset)
          $_size = [BitConverter]::ToUInt16($_data, $_offset + 2)
          $_offset += 4
          if ($_id -eq 0) {
            break
          }
          $_end = $_offset + $_size
          if ($_end -gt $_data.Length) {
            throw 'PersistData exceeds property bounds.'
          }
          $_kind = switch ($_id) {
            0x8001 { 'RssFeeds' }
            0x8004 { 'ToDo' }
            0x8008 { 'SuggestedContacts' }
          }
          if (-not $_kind) {
            $_offset = $_end
            continue
          }
          while ($_offset -lt $_end) {
            if ($_end - $_offset -lt 4) {
              throw 'Truncated PersistElement header.'
            }
            $_elementId = [BitConverter]::ToUInt16($_data, $_offset)
            $_length = [BitConverter]::ToUInt16($_data, $_offset + 2)
            $_offset += 4
            if ($_offset + $_length -gt $_end) {
              throw 'PersistElement exceeds block bounds.'
            }
            if ($_elementId -eq 0) {
              if ($_length -ne 0) {
                throw 'ELEMENT_SENTINEL must have zero length.'
              }
              $_offset = $_end
              break
            }
            if ($_elementId -eq 1 -and $_length -gt 0 -and -not $_known.ContainsKey($_kind)) {
              [byte[]]$_entry = $_data[$_offset..($_offset + $_length - 1)]
              $_known[$_kind] = $_accessor.BinaryToString($_entry)
            }
            $_offset += $_length
          }
        }
      }
      catch {
        $_exception = $_.Exception
        while ($_exception.InnerException) {
          $_exception = $_exception.InnerException
        }
        if ($_exception.HResult -ne -2147221233) {
          throw
        }
      }
    }

    foreach ($_definition in $_definitions.GetEnumerator()) {
      [PSCustomObject]@{
        Kind     = $_definition.Key
        StoreID  = [string]$StoreRoot.StoreID
        EntryID  = $_known[$_definition.Key]
        State    = if ($_known.ContainsKey($_definition.Key)) {
          'Resolved'
        }
        elseif ($_definition.Key -eq 'Inbox') {
          'Unresolved'
        }
        else {
          'Absent'
        }
        Evidence = 'Outlook2007.MAPI'
      }
    }
  }
  finally {
    foreach ($_accessor in $_accessors) {
      Remove-ComObject $_accessor
    }
    Remove-ComObject $_inbox $_default $_store
  }
}

function Get-OutlookFolderPlan {
  <#
    .SYNOPSIS
      Builds a read-only, locale-independent Outlook folder processing plan.
    .DESCRIPTION
      Resolves one exact store-relative path and optionally its descendants.
      Standard folders require inclusion by identity, custom exclusions win,
      and search folders are always excluded. Only mail items are eligible;
      included non-mail containers permit traversal to mail subfolders.
      Returns plain metadata; callers reopen selected folders by EntryID and
      StoreID. The entire plan must be collected successfully before mutation.
    .PARAMETER Namespace
      Connected Outlook MAPI namespace.
    .PARAMETER StoreRoot
      Root folder of the selected store, owned by the caller.
    .PARAMETER FolderName
      Exact store-relative path when explicitly supplied. Empty selects the root.
      When omitted, selects the store's Inbox by identity regardless of its name.
      If that identity cannot be resolved, supply an explicit path; no fallback
      to a name or the store root is attempted.
    .PARAMETER Recurse
      Visit descendants of the selected folder.
    .PARAMETER Include
      Standard folder kinds permitted within the selected scope. Default Inbox.
    .PARAMETER Exclusions
      Exact store-relative folder paths to exclude with their descendants.
    .PARAMETER ProgressId
      Progress record identifier used during enumeration.
    .EXAMPLE
      PS> Get-OutlookFolderPlan -Namespace $context.Namespace -StoreRoot $root -FolderName Posteingang -Recurse
    .EXAMPLE
      PS> Get-OutlookFolderPlan -Namespace $context.Namespace -StoreRoot $root
      Selects the Inbox by identity, including when localized or renamed.
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Nested traversal functions read Recurse and ProgressId from the parent scope.')]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [object]
    $Namespace,

    [Parameter(Mandatory = $true)]
    [object]
    $StoreRoot,

    [AllowEmptyString()]
    [string]
    $FolderName = 'Inbox',

    [switch]
    $Recurse,

    [ValidateSet('DeletedItems', 'Outbox', 'SentItems', 'Inbox', 'Calendar', 'Contacts', 'Journal', 'Notes', 'Tasks', 'Drafts', 'AllPublicFolders', 'Conflicts', 'SyncIssues', 'LocalFailures', 'ServerFailures', 'Junk', 'RssFeeds', 'ToDo', 'ManagedEmail', 'SuggestedContacts')]
    [string[]]
    $Include = @('Inbox'),

    [string[]]
    $Exclusions = @(),

    [int]
    $ProgressId = 0
  )

  if (@($Exclusions | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
    throw 'Exclusions must contain nonblank store-relative paths.'
  }

  foreach ($_path in @($FolderName) + @($Exclusions)) {
    if ($_path -eq '' -and $_path -eq $FolderName) {
      continue
    }
    foreach ($_segment in $_path.Split('\')) {
      if ([string]::IsNullOrWhiteSpace($_segment) -or $_segment -in @('.', '..')) {
        throw 'Folder paths must be exact, nonblank store-relative paths. Only FolderName may be empty for the store root.'
      }
    }
  }

  $_identities = @(Get-OutlookStandardFolderIdentity -Namespace $Namespace -StoreRoot $StoreRoot)
  $_identityById = @{}
  foreach ($_identity in $_identities) {
    if ($_identity.State -eq 'Unresolved' -and $_identity.Kind -notin $Include) {
      throw "Cannot enforce exclusion of '$($_identity.Kind)' on this store and Outlook version. Select a supported store or explicitly include this kind."
    }
    if ($_identity.EntryID) {
      $_identityById[$_identity.EntryID] = $_identity.Kind
    }
  }

  $_implicitInboxId = $null
  if (-not $PSBoundParameters.ContainsKey('FolderName')) {
    $_inboxIdentities = @($_identities | Where-Object { $_.Kind -eq 'Inbox' })
    if ($_inboxIdentities.Count -ne 1 -or $_inboxIdentities[0].State -ne 'Resolved' -or
      [string]::IsNullOrWhiteSpace([string]$_inboxIdentities[0].EntryID) -or
      $_inboxIdentities[0].StoreID -ne $StoreRoot.StoreID) {
      throw 'Cannot resolve the selected store Inbox identity. Supply FolderName explicitly.'
    }
    $_implicitInboxId = [string]$_inboxIdentities[0].EntryID
    $_inbox = $null
    try {
      $_inbox = $Namespace.GetFolderFromID($_implicitInboxId, [string]$StoreRoot.StoreID)
      if (-not $_inbox -or $_inbox.StoreID -ne $StoreRoot.StoreID -or $_inbox.EntryID -ne $_implicitInboxId) {
        throw 'Resolved Inbox does not match the selected store and folder identity.'
      }
      $_rootPrefix = ([string]$StoreRoot.FolderPath).TrimEnd('\') + '\'
      $_inboxPath = [string]$_inbox.FolderPath
      if (-not $_inboxPath.StartsWith($_rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Resolved Inbox is outside the selected store root.'
      }
      $FolderName = $_inboxPath.Substring($_rootPrefix.Length)
      foreach ($_segment in $FolderName.Split('\')) {
        if ([string]::IsNullOrWhiteSpace($_segment) -or $_segment -in @('.', '..')) {
          throw 'Resolved Inbox does not have an exact store-relative folder path.'
        }
      }
    }
    finally {
      Remove-ComObject $_inbox
    }
    # Walk the resolved path below so exclusions on every ancestor still apply.
  }

  function Get-FolderDecision {
    param (
      [object]
      $Folder,

      [string]
      $RelativePath
    )

    $_kind = $_identityById[[string]$Folder.EntryID]
    $_reason = $null
    foreach ($_excluded in $Exclusions) {
      if ($RelativePath -ieq $_excluded -or $RelativePath.StartsWith($_excluded + '\', [StringComparison]::OrdinalIgnoreCase)) {
        $_reason = 'CustomExclusion'
        break
      }
    }
    if (-not $_reason -and $_kind -and $_kind -notin $Include) {
      $_reason = "StandardFolder:$($_kind):Include$($_kind) required"
    }

    $_accessor = $Folder.PropertyAccessor
    try {
      if ($_accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x36010003') -eq 2) {
        $_reason = 'SearchFolder'
      }
    }
    finally {
      Remove-ComObject $_accessor
    }

    [PSCustomObject]@{
      EntryID      = [string]$Folder.EntryID
      StoreID      = [string]$Folder.StoreID
      FolderPath   = [string]$Folder.FolderPath
      RelativePath = $RelativePath
      StandardKind = $_kind
      Process      = -not [bool]$_reason -and $Folder.DefaultItemType -eq 0
      Traverse     = -not [bool]$_reason
      Reason       = if ($_reason) { $_reason } elseif ($Folder.DefaultItemType -ne 0) { 'NonMailContainer' } else { 'Included' }
    }
  }

  function Get-FolderTreePlan {
    param (
      [object]
      $Folder,

      [string]
      $RelativePath
    )

    Write-Progress -Id $ProgressId -Activity 'Outlook folders' -Status 'Reading folder identities and applying exclusions' -CurrentOperation $Folder.FolderPath
    $_decision = Get-FolderDecision -Folder $Folder -RelativePath $RelativePath
    $_decision
    if (-not $Recurse -or -not $_decision.Traverse) {
      return
    }

    $_folders = $Folder.Folders
    try {
      for ($_index = 1; $_index -le $_folders.Count; $_index++) {
        $_child = $_folders.Item($_index)
        try {
          $_relative = if ($RelativePath) { $RelativePath + '\' + $_child.Name } else { [string]$_child.Name }
          Get-FolderTreePlan -Folder $_child -RelativePath $_relative
        }
        finally {
          Remove-ComObject $_child
        }
      }
    }
    finally {
      Remove-ComObject $_folders
    }
  }

  $_selected = $StoreRoot
  $_relative = ''
  try {
    if ($FolderName -ne '') {
      foreach ($_segment in $FolderName.Split('\')) {
        $_next = Get-OutlookSubFolder -ParentFolder $_selected -Name $_segment
        if (-not $_next) {
          throw "FolderName '$FolderName' was not found. Use the displayed folder path, for example Posteingang."
        }
        if ($_selected -ne $StoreRoot) {
          Remove-ComObject $_selected
        }
        $_selected = $_next
        $_relative = if ($_relative) { $_relative + '\' + $_selected.Name } else { [string]$_selected.Name }
        $_decision = Get-FolderDecision -Folder $_selected -RelativePath $_relative
        if (-not $_decision.Traverse) {
          throw "Selected folder is excluded by '$($_decision.Reason)' at '$_relative'."
        }
      }
    }

    if ($_implicitInboxId -and $_selected.EntryID -ne $_implicitInboxId) {
      throw 'Inbox identity changed while resolving its path. Request a new folder plan.'
    }
    Get-FolderTreePlan -Folder $_selected -RelativePath $_relative
  }
  finally {
    if ($_selected -ne $StoreRoot) {
      Remove-ComObject $_selected
    }
  }
}
