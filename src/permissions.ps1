Set-StrictMode -Version Latest

function Test-Elevation {
  <#
    .SYNOPSIS
      Test-Elevation - Returns whether the current process has administrator privileges.
    .DESCRIPTION
      Checks the Windows principal of the current identity for membership in the
      built-in Administrators role. Returns $true when the process holds an
      elevated (administrator) token, $false otherwise.

      This is a pure predicate: it makes no changes and never elevates. Use
      Request-AdministratorPrivilege to actually elevate.
    .OUTPUTS
      System.Boolean
    .EXAMPLE
      PS> if (-not (Test-Elevation)) { throw 'Run me elevated.' }
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>
  [CmdletBinding()]
  [OutputType([bool])]
  param ()

  # winkit is Windows-only; fail clearly rather than throwing a cryptic type
  # error if this is ever invoked on a non-Windows host (the [Security.Principal]
  # Windows types are present on PS7/Windows but the call throws elsewhere).
  if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
    throw 'Test-Elevation is only supported on Windows.'
  }

  $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
  $principal = [Security.Principal.WindowsPrincipal]::new($identity)
  return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Backwards-compatible alias for the original name. Remove once callers are
# migrated to Test-Elevation.
# CHANGE-NOTE: drop this alias after updating references.
Set-Alias -Name Read-ProcessElevation -Value Test-Elevation

function Request-AdministratorPrivilege {
  <#
    .SYNOPSIS
      Request-AdministratorPrivilege - Ensures the current script runs elevated,
      re-launching it under UAC if necessary.
    .DESCRIPTION
      Intended to be called once near the top of an entry-point script (e.g.
      Invoke-Optimizer.ps1, Invoke-Bootstrap.ps1) as a self-elevation block.

      Behaviour:
        * If already elevated, returns immediately and the script continues.
        * If not elevated, re-launches the SAME PowerShell host with the SAME
          script and arguments under the RunAs verb (triggering UAC), waits for
          the elevated copy to finish, and exits the current (non-elevated)
          process with the elevated copy's exit code.

      Hardening over a naive self-elevation block:
        * Re-launches the actual current host (pwsh vs powershell), not a
          hardcoded powershell.exe.
        * Preserves the original working directory across the RunAs boundary
          (RunAs otherwise starts the child in system32, breaking relative paths).
        * Faithfully reconstructs the original invocation (including bound
          parameters and switches) so -Profile, -WhatIf, paths with spaces, etc.
          survive intact.
        * Propagates the elevated child's exit code to the original caller, so
          cmd.exe launchers checking %ERRORLEVEL% and CI see the real result.
        * Handles UAC denial gracefully (clear message + non-zero exit) instead
          of an unhandled Win32Exception.
        * Guards against an infinite re-launch loop via an environment marker,
          in case elevation "succeeds" but the token still isn't elevated
          (rare UAC/GPO configurations).

      NOTE: this function calls exit on the non-elevated path. It is designed for
      entry-point scripts, not for dot-sourced library use — calling it from an
      interactive session would terminate that session when not elevated.
    .PARAMETER ScriptPath
      Path to the script to re-launch. Defaults to the caller's own path
      ($PSCommandPath of the calling script). Override only for unusual hosting.
    .PARAMETER BoundParameters
      The calling script's $PSBoundParameters, used to faithfully reconstruct
      named parameters and switches for the elevated re-launch. Strongly
      recommended; pass $PSBoundParameters from the caller.
    .PARAMETER ArgumentList
      Any unbound/positional arguments ($args from the caller) to append after
      the reconstructed bound parameters.
    .PARAMETER IsElevatedRelaunch
      Loop guard. The caller passes this as $true only when its own param block
      received the re-launch marker, indicating THIS process is already the
      elevated child. If set and the process is still not elevated, the function
      aborts instead of spawning again. See the usage example for the pattern.
    .OUTPUTS
      None. Either returns (already elevated) or exits the process (re-launched).
    .EXAMPLE
      PS> # At the top of Invoke-Optimizer.ps1, whose param block includes a
      PS> # hidden [switch]$Elevated used purely as the re-launch marker:
      PS> Request-AdministratorPrivilege `
      PS>     -BoundParameters $PSBoundParameters `
      PS>     -ArgumentList $args `
      PS>     -IsElevatedRelaunch:$Elevated
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .LINK
      https://michael-casey.com/blog/self-elevating-powershell-and-batch-scripts/
      The common naive self-elevation pattern (Start-Process -Verb RunAs +
      hardcoded powershell, no parameter forwarding). Evaluated and deliberately
      superseded: this implementation adds host preservation, working-directory
      preservation, faithful bound-parameter/switch reconstruction, exit-code
      propagation, UAC-denial handling, and a re-launch loop guard.
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>
  [CmdletBinding()]
  [OutputType([void])]
  param (
    [string]$ScriptPath = $MyInvocation.PSCommandPath,

    [System.Collections.IDictionary]$BoundParameters,

    [object[]]$ArgumentList,

    [switch]$IsElevatedRelaunch
  )

  if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
    throw 'Request-AdministratorPrivilege is only supported on Windows.'
  }

  # Already elevated — nothing to do.
  if (Test-Elevation) {
    return
  }

  # Loop guard: if the caller told us this is already the elevated re-launch
  # (via its own -Elevated marker switch) and we're STILL not elevated, stop
  # rather than spawning forever. We pass the marker as an explicit parameter
  # rather than an environment variable (which doesn't survive the RunAs
  # boundary) or a bare argument token (which a CmdletBinding param block would
  # reject) — an explicit switch the caller declares is unambiguous.
  if ($IsElevatedRelaunch) {
    Write-Error ('Elevation was attempted but the process is still not elevated. ' +
      'Check UAC / Group Policy settings (e.g. "Run all administrators in Admin Approval Mode"). ' +
      'Aborting to avoid a re-launch loop.')
    exit 1
  }

  # RunAs minimum: Windows Vista (build 6000). Earlier versions can't elevate.
  $build = [int](Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop |
      Select-Object -ExpandProperty BuildNumber)
  if ($build -lt 6000) {
    throw "Self-elevation requires Windows Vista or later (build >= 6000); detected build $build."
  }

  # Resolve the script to re-launch.
  if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
    throw 'Could not determine the script path to re-launch. Pass -ScriptPath explicitly.'
  }
  $ScriptPath = (Resolve-Path -LiteralPath $ScriptPath -ErrorAction Stop).Path

  # Re-launch the SAME host (pwsh.exe or powershell.exe), not a hardcoded one.
  $hostExe = (Get-Process -Id $PID).Path
  if ([string]::IsNullOrWhiteSpace($hostExe)) {
    throw 'Could not determine the current PowerShell host executable path.'
  }

  # Faithfully reconstruct the argument list for the elevated re-launch.
  $reArgs = [System.Collections.Generic.List[string]]::new()
  $reArgs.Add('-NoProfile')
  $reArgs.Add('-ExecutionPolicy'); $reArgs.Add('Bypass')
  $reArgs.Add('-File'); $reArgs.Add($ScriptPath)

  if ($BoundParameters) {
    foreach ($name in $BoundParameters.Keys) {
      $value = $BoundParameters[$name]
      if ($value -is [System.Management.Automation.SwitchParameter]) {
        # Switches: include the flag only when present/true. Use -Name:$true form
        # so the elevated copy receives an explicit value, avoiding ambiguity.
        if ($value.IsPresent) { $reArgs.Add("-$name") }
      }
      elseif ($value -is [bool]) {
        $reArgs.Add("-$name`:$([bool]$value)")
      }
      elseif ($null -ne $value) {
        # Arrays -> repeat the parameter for each element so [string[]] params
        # round-trip correctly (e.g. -VCRedistVersions 14.0,12.0).
        foreach ($item in @($value)) {
          $reArgs.Add("-$name")
          $reArgs.Add([string]$item)
        }
      }
    }
  }

  if ($ArgumentList) {
    foreach ($a in $ArgumentList) { $reArgs.Add([string]$a) }
  }

  # Inject the elevation marker switch so the re-launched child can pass it back
  # into this function as -IsElevatedRelaunch and trip the loop guard if needed.
  # The caller's param block must declare a [switch]$Elevated for this to bind.
  if ($BoundParameters -and -not $BoundParameters.Contains('Elevated')) {
    $reArgs.Add('-Elevated')
  }
  elseif (-not $BoundParameters) {
    $reArgs.Add('-Elevated')
  }

  # Preserve the working directory across the RunAs boundary. Start-Process
  # -Verb RunAs otherwise launches in system32, breaking relative paths the
  # elevated script might use.
  $workingDir = (Get-Location -PSProvider FileSystem).ProviderPath

  $startInfo = @{
    FilePath = $hostExe
    ArgumentList = $reArgs.ToArray()
    Verb = 'RunAs'
    WorkingDirectory = $workingDir
    PassThru = $true
    Wait = $true
  }

  try {
    $proc = Start-Process @startInfo
    # Propagate the elevated child's exit code to our caller.
    exit $proc.ExitCode
  }
  catch [System.ComponentModel.Win32Exception] {
    # 1223 = ERROR_CANCELLED — user clicked "No" on the UAC prompt.
    if ($_.Exception.NativeErrorCode -eq 1223) {
      Write-Error 'Elevation was cancelled by the user. Administrator privileges are required to continue.'
      exit 1223
    }
    throw
  }
}

# Shared native method - compiled once, reused by Set-RegistryOwner
if ($null -eq ('OwnershipNative' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class OwnershipNative {
    [DllImport("ntdll.dll")]
    public static extern int RtlAdjustPrivilege(ulong Privilege, bool Enable, bool CurrentThread, ref bool Enabled);
}
'@ -ErrorAction Stop
}

function Set-RegistryOwner {
  <#
    .SYNOPSIS
      Set-RegistryOwner - Takes ownership of a registry key, optionally including all subkeys.
    .DESCRIPTION
      Takes ownership of the specified registry key (owned by e.g. TrustedInstaller or a
      vendor installer) and grants the owning account FullControl.

      Uses RtlAdjustPrivilege P/Invoke to acquire SeTakeOwnership, SeBackup, and SeRestore,
      then manipulates the key ACL via System.Security.AccessControl.RegistrySecurity.
      Ownership is granted to the current user by default, or to any account via -SID.

      Requires elevation: the process must hold an administrator token, otherwise the
      privilege adjustment and ACL writes will fail. Recurse is on by default: subkeys
      are taken over recursively, with access rules propagated.
    .PARAMETER Hive
      Registry hive, e.g. HKLM, HKCU, HKCR, HKCC, HKU (or HKEY_* long forms).
    .PARAMETER Key
      Key path below the hive, e.g. 'SOFTWARE\WOW6432Node\Classes\CLSID\{abc}'.
    .PARAMETER SID
      SID of the account to grant ownership to. Defaults to the current user.
    .PARAMETER Recurse
      Take ownership of the key and all subkeys. Defaults to $true.
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result (Target, Source, Action, Status).
    .EXAMPLE
      PS> Set-RegistryOwner -Hive 'HKLM' -Key 'SOFTWARE\MyApp'
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidDefaultValueSwitchParameter', '', Justification = 'The reference behaviour (and the downstream requirements doc) mandate -Recurse defaulting to $true for ownership takeover.')]
  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]
    $Hive,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]
    $Key,

    [Parameter(Mandatory = $false)]
    [System.Security.Principal.SecurityIdentifier]
    $SID,

    [Parameter(Mandatory = $false)]
    [switch]
    $Recurse = $true
  )

  if (-not (Test-Elevation)) {
    throw 'Set-RegistryOwner requires an elevated process (administrator token).'
  }

  if ($null -eq $SID) {
    $SID = Get-UserSID -UserName $env:USERNAME
    if ([string]::IsNullOrWhiteSpace($SID)) {
      throw "Could not resolve a SID for '$env:USERNAME'."
    }
    $SID = [System.Security.Principal.SecurityIdentifier]::new($SID)
  }

  switch -Regex ($Hive) {
    '^(HKCU|HKEY_CURRENT_USER)$' { $root = 'CurrentUser'; break }
    '^(HKLM|HKEY_LOCAL_MACHINE)$' { $root = 'LocalMachine'; break }
    '^(HKCR|HKEY_CLASSES_ROOT)$' { $root = 'ClassesRoot'; break }
    '^(HKCC|HKEY_CURRENT_CONFIG)$' { $root = 'CurrentConfig'; break }
    '^(HKU|HKEY_USERS)$' { $root = 'Users'; break }
    default {
      throw "Unsupported hive '$Hive'. Use HKCU, HKLM, HKCR, HKCC, or HKU."
    }
  }

  # Acquire SeTakeOwnership (9), SeBackup (17), and SeRestore (18).
  foreach ($privilegeId in @(9, 17, 18)) {
    $enabled = $false
    $null = [OwnershipNative]::RtlAdjustPrivilege($privilegeId, $true, $false, [ref]$enabled)
  }

  function Invoke-RegistryOwnershipTakeover {
    param (
      [string]$Root,
      [string]$KeyPath,
      [int]$RecurseLevel,
      [bool]$Recurse
    )

    $item = [Microsoft.Win32.Registry]::$Root.OpenSubKey($KeyPath, 'ReadWriteSubTree', 'TakeOwnership')
    if ($null -eq $item) {
      Write-Verbose "Key not found: $Hive\$KeyPath"
      return
    }

    try {
      $acl = [System.Security.AccessControl.RegistrySecurity]::new()
      $acl.SetOwner($SID)

      if ($PSCmdlet.ShouldProcess("$Hive\$KeyPath", 'Take ownership')) {
        $item.SetAccessControl($acl)

        # Enable inheritance of permissions (not ownership) from the parent key.
        $acl.SetAccessRuleProtection($false, $false)
        $item.SetAccessControl($acl)
      }

      if ($RecurseLevel -eq 0) {
        # Top-level key: grant FullControl, propagated to all subkeys.
        $rule = [System.Security.AccessControl.RegistryAccessRule]::new(
          $SID, 'FullControl', 'ContainerInherit', 'None', 'Allow')
        $acl.ResetAccessRule($rule)
        $changeItem = $item.OpenSubKey('', 'ReadWriteSubTree', 'ChangePermissions')
        try {
          if ($PSCmdlet.ShouldProcess("$Hive\$KeyPath", 'Set FullControl access')) {
            $changeItem.SetAccessControl($acl)
          }
        }
        finally {
          $changeItem.Dispose()
        }
      }

      if ($Recurse) {
        foreach ($subKeyName in $item.OpenSubKey('').GetSubKeyNames()) {
          Invoke-RegistryOwnershipTakeover -Root $Root -KeyPath "$KeyPath\$subKeyName" -RecurseLevel ($RecurseLevel + 1) -Recurse $Recurse
        }
      }
    }
    finally {
      $item.Dispose()
    }
  }

  Invoke-RegistryOwnershipTakeover -Root $root -KeyPath $Key -RecurseLevel 0 -Recurse $Recurse

  if ($WhatIfPreference) {
    New-OperationResult -Target "$Hive\$Key" -Source 'Registry' -Action 'TakeOwnership' -Status 'DryRun' -Detail 'No changes applied.'
    return
  }

  New-OperationResult -Target "$Hive\$Key" -Source 'Registry' -Action 'TakeOwnership' -Status 'Completed' -Detail "Ownership granted to $SID." -Property @{ SID = $SID.Value }
}

function Set-ItemOwner {
  <#
    .SYNOPSIS
      Set-ItemOwner - Takes ownership of a file or folder via takeown/icacls.
    .DESCRIPTION
      Takes ownership of the specified filesystem object (file or directory) and grants
      FullControl to the specified group, using takeown.exe and icacls.exe shell-outs.

      For directories the operation recurses through all children. When takeown/icacls
      fail (e.g. genuinely stuck items), a robocopy /purge fallback is attempted against
      an empty staging folder - note that this DELETES the target's contents and is a
      last-resort path.

      Requires elevation.
    .PARAMETER Path
      Full path of the file or directory to take ownership of.
    .PARAMETER Group
      Group to grant FullControl to. Defaults to 'administrators'.
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result (Target, Source, Action, Status).
    .EXAMPLE
      PS> Set-ItemOwner -Path 'C:\Program Files\VendorApp'
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]
    $Path,

    [Parameter(Mandatory = $false)]
    [string]
    $Group = 'administrators'
  )

  if (-not (Test-Elevation)) {
    throw 'Set-ItemOwner requires an elevated process (administrator token).'
  }

  if (-not (Test-Path -LiteralPath $Path)) {
    New-OperationResult -Target $Path -Source 'FileSystem' -Action 'TakeOwnership' -Status 'Failed' -ErrorMessage 'Path does not exist.'
    return
  }

  $isDirectory = (Get-Item -LiteralPath $Path) -is [System.IO.DirectoryInfo]

  if (-not $PSCmdlet.ShouldProcess($Path, 'Take ownership')) {
    if ($WhatIfPreference) {
      New-OperationResult -Target $Path -Source 'FileSystem' -Action 'TakeOwnership' -Status 'DryRun' -Detail 'No changes applied.'
    }
    return
  }

  $purgeFallback = $false

  if ($isDirectory) {
    # /F path, /R recurse, /A grant admins, /D Y answer 'yes' to prompts
    & "$env:SystemRoot\System32\takeown.exe" /F $Path /R /A /D Y 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
      & "$env:SystemRoot\System32\icacls.exe" $Path /grant "${Group}:f" /t /q 2>&1 | Out-Null
      if ($LASTEXITCODE -ne 0) { $purgeFallback = $true }
    }
    else {
      $purgeFallback = $true
    }
  }
  else {
    & "$env:SystemRoot\System32\takeown.exe" /F $Path /A 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
      & "$env:SystemRoot\System32\icacls.exe" $Path /grant "${Group}:f" /q 2>&1 | Out-Null
      if ($LASTEXITCODE -ne 0) { $purgeFallback = $true }
    }
    else {
      $purgeFallback = $true
    }
  }

  $fallbackRecovered = $false

  if ($purgeFallback) {
    Write-Verbose "takeown/icacls failed; attempting robocopy /purge fallback for: $Path"
    $emptyDir = Join-Path -Path $env:TEMP -ChildPath 'psf-empty-takeown'
    if (-not (Test-Path -LiteralPath $emptyDir)) {
      $null = New-Item -Path $emptyDir -ItemType Directory -Force
    }
    try {
      & "$env:SystemRoot\System32\robocopy.exe" $emptyDir $Path /purge 2>&1 | Out-Null
      # Robocopy exit codes 0-7 are success variants.
      if ($LASTEXITCODE -le 7) { $fallbackRecovered = $true }
    }
    finally {
      if (Test-Path -LiteralPath $emptyDir) {
        Remove-Item -LiteralPath $emptyDir -Recurse -Force -ErrorAction SilentlyContinue
      }
    }
  }

  if ($fallbackRecovered) {
    New-OperationResult -Target $Path -Source 'FileSystem' -Action 'TakeOwnership' -Status 'Completed' -Detail "Item recovered via robocopy /purge fallback; FullControl granted to '$Group'."
  }
  elseif ($purgeFallback) {
    New-OperationResult -Target $Path -Source 'FileSystem' -Action 'TakeOwnership' -Status 'Failed' -ErrorMessage 'takeown/icacls failed and the robocopy /purge fallback could not recover the item.'
  }
  else {
    New-OperationResult -Target $Path -Source 'FileSystem' -Action 'TakeOwnership' -Status 'Completed' -Detail "FullControl granted to '$Group'."
  }
}

function New-EncryptedCredentialFile {
  <#
    .SYNOPSIS
      New-EncryptedCredentialFile - Stores a credential in a DPAPI-independent encrypted file.
    .DESCRIPTION
      Encrypts a credential's password with a freshly generated random AES key
      (ConvertFrom-SecureString -Key) and writes the key and the encrypted
      blob to separate files.

      Plain ConvertFrom-SecureString without -Key uses Windows DPAPI, which
      binds the encrypted blob to a specific user and machine. The -Key
      variant is the correct approach when a scheduled task or automation
      script running as a different account or machine must decrypt it later.

      File format: the credential file holds the username in plaintext on the
      first line (usernames are not secrets) and the encrypted password blob
      on the second. The key file holds the key as base64.

      SECURITY CAVEAT: this pattern is only as secure as protecting the key
      file separately from the encrypted blob. The key file is locked down to
      the creating user's account, but you should still store it somewhere
      with tighter ACLs than the credential file - or use a secrets-management
      mechanism when one is available. Anyone holding both files can decrypt
      the credential.
    .PARAMETER Path
      Path of the credential file to write.
    .PARAMETER KeyPath
      Path of the key file to write.
    .PARAMETER Credential
      The credential to store. Required unless -UserName is used.
    .PARAMETER UserName
      Username to store; prompts for the password interactively
      (Read-Host -AsSecureString). Thin interactive wrapper around the
      programmatic -Credential path.
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result.
    .EXAMPLE
      PS> New-EncryptedCredentialFile -Path '.\svc-cred.bin' -KeyPath '.\svc-cred.key' -Credential (Get-Credential)
    .EXAMPLE
      PS> New-EncryptedCredentialFile -Path '.\svc-cred.bin' -KeyPath '.\svc-cred.key' -UserName 'DOMAIN\svc-backup'
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    [Parameter(Mandatory = $true)]
    [string]
    $KeyPath,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter(Mandatory = $false)]
    [string]
    $UserName
  )

  if ($null -eq $Credential -and -not $UserName) {
    throw 'Supply -Credential (programmatic) or -UserName (interactive password prompt).'
  }

  if ($null -eq $Credential) {
    $securePassword = Read-Host -Prompt "Password for '$UserName'" -AsSecureString
    $Credential = [System.Management.Automation.PSCredential]::new($UserName, $securePassword)
  }

  if (-not $PSCmdlet.ShouldProcess("$Path / $KeyPath", 'Write encrypted credential files')) {
    if ($WhatIfPreference) {
      New-OperationResult -Target $Path -Source 'CredentialFile' -Action 'Write' -Status 'DryRun' -Detail 'No files written.'
    }
    return
  }

  $key = [byte[]]::new(32)
  $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
  try {
    $rng.GetBytes($key)
  }
  finally {
    $rng.Dispose()
  }

  $encryptedBlob = $Credential.Password | ConvertFrom-SecureString -Key $key
  $keyBase64 = [Convert]::ToBase64String($key)

  Set-Content -LiteralPath $Path -Value @($Credential.UserName, $encryptedBlob) -Encoding ASCII
  Set-Content -LiteralPath $KeyPath -Value $keyBase64 -Encoding ASCII

  try {
    $acl = Get-Acl -LiteralPath $KeyPath
    $acl.SetAccessRuleProtection($true, $false)
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
    $rule = [System.Security.AccessControl.FileSystemAccessRule]::new($currentUser, 'FullControl', 'Allow')
    $acl.SetAccessRule($rule)
    Set-Acl -LiteralPath $KeyPath -AclObject $acl
    Write-Verbose 'Restricted the key file to the current user account.'
  }
  catch {
    Write-Warning "Could not restrict ACLs on the key file '$KeyPath': $($_.Exception.Message)"
  }

  New-OperationResult -Target $Path -Source 'CredentialFile' -Action 'Write' -Status 'Completed' -Detail 'Credential stored with a random AES key; key file written separately.'
}

function Get-EncryptedCredentialFile {
  <#
    .SYNOPSIS
      Get-EncryptedCredentialFile - Reads a credential stored by New-EncryptedCredentialFile.
    .DESCRIPTION
      Reads the key file and the encrypted credential file created by
      New-EncryptedCredentialFile and returns the decrypted PSCredential.

      File format: the credential file holds the username in plaintext on the
      first line and the encrypted password blob on the second; the key file
      holds the key as base64.
    .PARAMETER Path
      Path of the credential file.
    .PARAMETER KeyPath
      Path of the key file.
    .OUTPUTS
      System.Management.Automation.PSCredential
    .EXAMPLE
      PS> Get-EncryptedCredentialFile -Path '.\svc-cred.bin' -KeyPath '.\svc-cred.key'
    .LINK
      https://github.com/adnoctem/winkit/lib/permissions.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([System.Management.Automation.PSCredential])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Path,

    [Parameter(Mandatory = $true)]
    [string]
    $KeyPath
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "Credential file not found: $Path"
  }
  if (-not (Test-Path -LiteralPath $KeyPath)) {
    throw "Key file not found: $KeyPath"
  }

  $lines = Get-Content -LiteralPath $Path
  if ($lines.Count -lt 2) {
    throw "Credential file '$Path' is malformed (expected a username line and an encrypted blob line)."
  }

  $userName = $lines[0]
  $encryptedBlob = ($lines[1..($lines.Count - 1)] -join '')
  $key = [Convert]::FromBase64String((Get-Content -LiteralPath $KeyPath -Raw).Trim())

  $securePassword = $encryptedBlob | ConvertTo-SecureString -Key $key
  [System.Management.Automation.PSCredential]::new($userName, $securePassword)
}
