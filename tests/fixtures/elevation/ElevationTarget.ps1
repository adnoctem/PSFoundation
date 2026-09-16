#Requires -Version 5.1
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidDefaultValueSwitchParameter', '', Justification = 'Regression fixture verifies explicit false switches override true defaults.')]
param (
  [string]$Report,
  [string]$Payload,
  [string[]]$Items,
  [switch]$Enabled = $true,
  [bool]$Flag,
  [AllowNull()][object]$Optional,
  [switch]$Elevated
)

[PSCustomObject]@{
  Bound = $PSBoundParameters
  Directory = (Get-Location).Path
  ScriptPath = $PSCommandPath
  Extra = @($args)
} | Export-Clixml -LiteralPath $Report
exit 37
