#Requires -Version 5.1

# Runs in an isolated, unelevated test process. Never invokes RunAs.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidDefaultValueSwitchParameter', '', Justification = 'Exercises forwarding an explicit false over a true default.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Isolated subprocess stubs prevent actual UAC and OS queries.')]
param (
  [string]$Source,
  [string]$Report,
  [string]$LaunchReport,
  [string]$Kind,
  [string]$Target = (Join-Path $PSScriptRoot 'ElevationTarget.ps1')
)

$ErrorActionPreference = 'Stop'
$script:FixtureLaunchReport = $LaunchReport
$script:FixtureTarget = $Target
. $Source

function Test-Elevation { return $Kind -eq 'Already' }
function Get-CimInstance { [PSCustomObject]@{ BuildNumber = '19045' } }
function Start-Process {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Test stub captures a launch or runs only the harmless fixture without elevation.')]
  [CmdletBinding()]
  param ($FilePath, $ArgumentList, $WorkingDirectory, $Verb, [switch]$PassThru, [switch]$Wait)
  [PSCustomObject]@{
    HostPath = $FilePath
    Arguments = $ArgumentList
    Directory = $WorkingDirectory
    Verb = $Verb
    PassThru = [bool]$PassThru
    Wait = [bool]$Wait
  } | Export-Clixml -LiteralPath $script:FixtureLaunchReport
  if ($Kind -eq 'Cancelled') { throw [ComponentModel.Win32Exception]::new(1223) }
  if ($Kind -eq 'LaunchFailure') { throw [ComponentModel.Win32Exception]::new(2) }
  if ($Kind -eq 'Execute') {
    return Microsoft.PowerShell.Management\Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -WorkingDirectory $WorkingDirectory -WindowStyle Hidden -Wait -PassThru
  }
  return [PSCustomObject]@{ ExitCode = 37 }
}

function Invoke-BoundFixture {
  param (
    [string]$Report,
    [string]$Payload,
    [string[]]$Items,
    [switch]$Enabled = $true,
    [bool]$Flag,
    [AllowNull()][object]$Optional,
    [switch]$Elevated
  )
  Request-AdministratorPrivilege -ScriptPath $script:FixtureTarget -BoundParameters $PSBoundParameters -ArgumentList @('loose value', "quote'`"`$()", '')
}

switch ($Kind) {
  'Ordered' { Request-AdministratorPrivilege -ScriptPath $target -BoundParameters ([ordered]@{ Report = $Report; Elevated = $false }) }
  'Hashtable' { Request-AdministratorPrivilege -ScriptPath $target -BoundParameters @{ Report = $Report; Elevated = $true } }
  'Empty' { Request-AdministratorPrivilege -ScriptPath $target -BoundParameters @{} }
  'Null' { Request-AdministratorPrivilege -ScriptPath $target }
  'Loop' { Request-AdministratorPrivilege -ScriptPath $target -IsElevatedRelaunch }
  'Unsupported' { Request-AdministratorPrivilege -ScriptPath $target -BoundParameters @{ Payload = [PSCustomObject]@{ Name = 'unsupported' } } }
  default {
    $smartQuotes = 'smart ' + [char]0x2019 + "; throw 'untrusted' # " + [char]0x2018 + [char]0x201a + [char]0x201b
    Invoke-BoundFixture -Report $Report -Payload "spaces ' quotes `" and `$([throw]) ; #" -Items @('one', 'two words', '', $smartQuotes) -Enabled:$false -Flag:$false -Optional $null -Elevated:$false
  }
}
exit 0
