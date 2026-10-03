# PSFoundation source acknowledgements

This records attributions present in the source and distinguishes adapted code from technical references.

## Adapted work

- `Install-Font` in [system.ps1](../src/system.ps1) credits
  [LeDragoX/Win-Debloat-Tools, Install-Font.psm1](https://github.com/LeDragoX/Win-Debloat-Tools/blob/main/src/lib/Install-Font.psm1). Its
  recorded attribution also credits [anthonyeden](https://github.com/anthonyeden) for the earlier font-installation gist. PSFoundation's
  implementation adds its own validation, structured outcomes and explicit source-removal consent.
- Domain helpers retain links to [adnoctem/winkit](https://github.com/adnoctem/winkit), the related project from which this shared module
  evolved. Those links record project provenance rather than an additional third-party dependency.

## Evaluated patterns and reference data

- `Request-AdministratorPrivilege` in [permissions.ps1](../src/permissions.ps1) records
  [Michael Casey's self-elevating script article](https://michael-casey.com/blog/self-elevating-powershell-and-batch-scripts/). The source
  explicitly identifies that pattern as evaluated and superseded; the current implementation preserves host, parameters and working
  directory and adds a relaunch guard. This is a reference acknowledgement, not a claim that the present implementation was copied.
- WinGet error mappings in [errors.ps1](../src/errors.ps1) reference Microsoft's
  [Windows Package Manager return codes](https://github.com/microsoft/winget-cli/blob/master/doc/windows/package-manager/winget/returnCodes.md).
  Windows, Office and PowerShell documentation links elsewhere in the module identify the relevant API contracts.

Source comments and upstream license notices remain authoritative for the attributed material. Similar function names alone are not evidence
of shared provenance.
