#Requires -Version 5.0

function ConvertTo-PSFErrorCode {
  [CmdletBinding()]
  [OutputType([string])]
  param ([object]$Code)

  $text = [Convert]::ToString($Code, [Globalization.CultureInfo]::InvariantCulture).Trim()
  if ($text -match '^0[xX][0-9a-fA-F]{1,8}$') {
    $number = [Convert]::ToInt64($text.Substring(2), 16)
  }
  elseif ($text -match '^-?\d{1,10}$') { $number = [long]::Parse($text, [Globalization.CultureInfo]::InvariantCulture) }
  else { throw 'Error code must be a 32-bit integer or hexadecimal code.' }
  if ($number -lt [int]::MinValue -or $number -gt [uint32]::MaxValue) {
    throw 'Error code is outside the signed/unsigned 32-bit range.'
  }
  if ($number -lt 0) { $number += 4294967296L }
  return $number.ToString('X8', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ErrorTranslation {
  <#
    .SYNOPSIS
      Translates known deployment error codes into operator guidance.
    .DESCRIPTION
      Uses separate Appx, Winget, Dism, and Msi tables. Returns Code, Domain,
      Benign, Detail, and Matched, or no result for an unknown/ambiguous error.
      Code is an uppercase hexadecimal HRESULT, or a decimal MSI exit code.
      Recognized exception HRESULTs take precedence (outer to inner), followed
      by codes in ErrorDetails and exception messages. Multiple distinct known
      codes in text are ambiguous and produce no result. Text matching accepts
      hexadecimal HRESULTs and signed/unsigned decimal HRESULTs, never short
      decimal substrings that could be counts or identifiers.
      Benign means an unconditional success code, not permission to suppress
      an error merely because it is recognized. Conditional cases, including
      already-installed versions, remain false; the caller decides whether its
      intended state is satisfied. Reboot success codes retain guidance about
      completing the restart. This function does not log, retry, or change state.
    .PARAMETER ErrorRecord
      A caught or pipeline ErrorRecord. Defaults to searching all domains.
    .PARAMETER Code
      A native exit code or HRESULT, as an integer or hexadecimal string.
      A domain is required for this parameter set to disambiguate native codes.
    .PARAMETER Domain
      Restrict lookup to Appx, Winget, Dism, or Msi. WinGet's table describes
      its own codes; translate wrapped installer failures in their own domain.
    .EXAMPLE
      PS> Get-ErrorTranslation -ErrorRecord $_ -Domain Appx
    .EXAMPLE
      PS> Get-ErrorTranslation -Code 3010 -Domain Msi
    .OUTPUTS
      PSCustomObject, or no output when a translation cannot be selected.
    .LINK
      https://learn.microsoft.com/en-us/windows/win32/appxpkg/troubleshooting
    .LINK
      https://learn.microsoft.com/en-us/windows/win32/msi/error-codes
    .LINK
      https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md
    .LINK
      https://learn.microsoft.com/en-us/troubleshoot/windows-client/application-management/dotnet-framework-35-installation-error
  #>
  [CmdletBinding(DefaultParameterSetName = 'Record')]
  [OutputType([PSCustomObject])]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'Record')]
    [System.Management.Automation.ErrorRecord]$ErrorRecord,
    [Parameter(Mandatory = $true, ParameterSetName = 'Code')]
    [ValidateNotNullOrEmpty()]
    [object]$Code,
    [Parameter(Mandatory = $true, ParameterSetName = 'Code')]
    [Parameter(ParameterSetName = 'Record')]
    [ValidateSet('Appx', 'Winget', 'Dism', 'Msi')]
    [string]$Domain
  )

  begin {
    # Small, intentionally curated tables. Source references are in public help.
    $tables = [ordered]@{
      Appx   = @{
        '80073CF0' = @($false, 'The package could not be opened. Check its path, access, signature, and the AppxPackagingOM log.')
        '80073CF3' = @($false, 'Dependency or conflict validation failed. Check dependencies, architecture, and AppXDeployment-Server logs.')
        '80073CF9' = @($false, 'Package installation failed. Inspect AppXDeployment-Server logs for the specific cause.')
        '80073CFB' = @($false, 'A package is already installed and blocks reinstalling this package. Verify identity and version before deciding to skip.')
        '80073D02' = @($false, 'Resources needed by the package are in use. Close affected applications and retry when the locks are released.')
        '80073D06' = @($false, 'A newer package version is installed. Skip only if that version satisfies the requested state; downgrades remain failures.')
      }
      Winget = @{
        '8A150011' = @($false, 'The installer hash differs from the manifest. Refresh the source and investigate the mismatch; do not bypass verification.')
        '8A15002B' = @($false, 'No applicable update was found. Verify the installed version and requested target before deciding to skip.')
        '8A15002C' = @($false, 'One or more upgrades failed. Review the individual package results and WinGet logs.')
      }
      Dism   = @{
        '800F081F' = @($false, 'Required source files were not found. Provide a repair source matching the target Windows image and inspect DISM/CBS logs.')
        '800F0906' = @($false, 'Required source files could not be downloaded. Check connectivity and servicing-source policy, or provide a matching local source.')
      }
      Msi    = @{
        '00000000' = @($true, 'The installer completed successfully.')
        '00000642' = @($false, 'The user cancelled installation. Retry only when installation is still intended.')
        '00000643' = @($false, 'Installation failed. Enable verbose MSI logging and inspect the failure before retrying.')
        '00000652' = @($false, 'Another installation is running. Wait for it to finish before retrying.')
        '00000666' = @($false, 'Another version of the product is installed. Check the requested version and upgrade path.')
        '00000669' = @($true, 'Installation succeeded and initiated a restart. Resume dependent work after restart.')
        '00000BC2' = @($true, 'Installation succeeded; a restart is required to complete it. Schedule a restart before dependent work.')
      }
    }
  }
  process {
    $domains = @($tables.Keys)
    if ($Domain) { $domains = @($tables.Keys | Where-Object { $_ -eq $Domain }) }
    $structured = [Collections.Generic.List[string]]::new()
    $textCodes = [Collections.Generic.List[string]]::new()
    if ($PSCmdlet.ParameterSetName -eq 'Code') {
      $structured.Add((ConvertTo-PSFErrorCode -Code $Code))
    }
    else {
      $messages = [Collections.Generic.List[string]]::new()
      if ($ErrorRecord.ErrorDetails) { $messages.Add($ErrorRecord.ErrorDetails.Message) }
      $exception = $ErrorRecord.Exception
      # Bound traversal even for unusual exception implementations.
      for ($depth = 0; $null -ne $exception -and $depth -lt 32; $depth++) {
        $structured.Add((ConvertTo-PSFErrorCode -Code $exception.HResult))
        $messages.Add($exception.Message)
        $exception = $exception.InnerException
      }
      foreach ($message in $messages) {
        foreach ($match in [regex]::Matches($message, '(?i)(?<![\w])(?:0x[0-9a-f]{8}|-\d{10}|[23]\d{9})(?![\w])')) {
          try { $textCodes.Add((ConvertTo-PSFErrorCode -Code $match.Value)) }
          catch { continue }
        }
      }
    }

    $selected = $null
    foreach ($candidate in $structured) {
      $hits = @($domains | Where-Object { $tables[$_].ContainsKey($candidate) })
      if ($hits.Count -eq 1) { $selected = @($candidate, $hits[0]); break }
      if ($hits.Count -gt 1) { return }
    }
    if ($null -eq $selected) {
      $matchesFound = @(
        foreach ($candidate in @($textCodes | Select-Object -Unique)) {
          foreach ($candidateDomain in $domains) {
            if ($tables[$candidateDomain].ContainsKey($candidate)) {
              , @($candidate, $candidateDomain)
            }
          }
        }
      )
      if ($matchesFound.Count -ne 1) { return }
      $selected = $matchesFound[0]
    }
    $translation = $tables[$selected[1]][$selected[0]]
    $displayCode = '0x' + $selected[0]
    if ($selected[1] -eq 'Msi') { $displayCode = [Convert]::ToUInt32($selected[0], 16).ToString([Globalization.CultureInfo]::InvariantCulture) }
    [PSCustomObject]@{
      Code    = $displayCode
      Domain  = $selected[1]
      Benign  = [bool]$translation[0]
      Detail  = $translation[1]
      Matched = $true
    }
  }
}
