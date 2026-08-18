#Requires -Version 5.0

# Native offline domain join binding - compiled once, reused by
# New-OfflineDomainJoinBlob. P/Invoke port of the well-established technique
# from lazywinadmin/PowerShell (AD-COMPUTER-New-ADDomainJoin), which attributes
# its own source (WinPENanoDomainJoin). The 2012+ API path
# (NetCreateProvisioningPackage) is used exclusively - this module targets
# Windows 10/11/Server 2016 and newer.
if ($null -eq ('PSFOfflineJoin.Native' -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace PSFOfflineJoin
{
    public class Native
    {
        [StructLayout(LayoutKind.Sequential)]
        public class NetSetupProvisioningParams
        {
            public uint dwVersion;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpDomain;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpHostName;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpMachineAccountOU;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpDcName;
            public uint dwProvisionOptions;
            public IntPtr aCertTemplateNames;
            public uint cCertTemplateNames;
            public IntPtr aMachinePolicyNames;
            public uint cMachinePolicyNames;
            public IntPtr aMachinePolicyPaths;
            public uint cMachinePolicyPaths;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpNetbiosName;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpSiteName;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpPrimaryDNSDomain;
        }

        [DllImport("netapi32.dll", EntryPoint = "NetCreateProvisioningPackage", SetLastError = true, ExactSpelling = true, CharSet = CharSet.Unicode)]
        public static extern int NetCreateProvisioningPackage(
            NetSetupProvisioningParams pProvisioningParams,
            IntPtr ppPackageBinData,
            IntPtr pdwPackageBinDataSize,
            [MarshalAs(UnmanagedType.LPWStr)] out string ppPackageText);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LogonUser(string username, string domain, string password, int logonType, int logonProvider, out IntPtr token);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool DuplicateToken(IntPtr existingToken, int impersonationLevel, out IntPtr duplicateToken);

        [DllImport("advapi32.dll")]
        public static extern bool ImpersonateLoggedOnUser(IntPtr token);

        [DllImport("advapi32.dll")]
        public static extern bool RevertToSelf();

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr handle);

        public const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
        public const int LOGON32_PROVIDER_WINNT50 = 3;
        public const int SECURITY_IMPERSONATION = 2;

        public static string CreateProvisioningPackage(string domain, string machine, string ou, string dc, uint options, string username, string password)
        {
            if (!string.IsNullOrEmpty(username))
            {
                string userDomain = username;
                string user = username;
                int idx = username.IndexOf('\\');
                if (idx >= 0)
                {
                    userDomain = username.Substring(0, idx);
                    user = username.Substring(idx + 1);
                }

                IntPtr token = IntPtr.Zero;
                IntPtr dupToken = IntPtr.Zero;
                try
                {
                    if (!LogonUser(user, userDomain, password, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_WINNT50, out token))
                        throw new Exception("LogonUser failed with Win32 error " + Marshal.GetLastWin32Error());
                    if (!DuplicateToken(token, SECURITY_IMPERSONATION, out dupToken))
                        throw new Exception("DuplicateToken failed with Win32 error " + Marshal.GetLastWin32Error());
                    if (!ImpersonateLoggedOnUser(dupToken))
                        throw new Exception("ImpersonateLoggedOnUser failed with Win32 error " + Marshal.GetLastWin32Error());
                    return CallNetCreateProvisioningPackage(domain, machine, ou, dc, options);
                }
                finally
                {
                    RevertToSelf();
                    if (dupToken != IntPtr.Zero) CloseHandle(dupToken);
                    if (token != IntPtr.Zero) CloseHandle(token);
                }
            }
            return CallNetCreateProvisioningPackage(domain, machine, ou, dc, options);
        }

        private static string CallNetCreateProvisioningPackage(string domain, string machine, string ou, string dc, uint options)
        {
            NetSetupProvisioningParams p = new NetSetupProvisioningParams();
            p.dwVersion = 1;
            p.lpDomain = domain;
            p.lpHostName = machine;
            p.lpMachineAccountOU = ou;
            p.lpDcName = dc;
            p.dwProvisionOptions = options;
            string blob = string.Empty;
            int res = NetCreateProvisioningPackage(p, IntPtr.Zero, IntPtr.Zero, out blob);
            if (res != 0) throw new Exception("NetCreateProvisioningPackage failed with error " + res);
            return blob;
        }
    }
}
'@ -ErrorAction Stop
}

function New-OfflineDomainJoinBlob {
  <#
    .SYNOPSIS
      New-OfflineDomainJoinBlob - Provisions an AD computer account and returns the offline-join blob.
    .DESCRIPTION
      Fully native offline domain join: provisions the computer account in
      Active Directory and generates the join blob via the
      NetCreateProvisioningPackage Win32 API (P/Invoke) - no djoin.exe
      dependency. When -Credential is supplied, the API runs under that
      account's impersonated context; otherwise the current context is used.

      The returned blob is consumed by New-DjoinFile (or djoin.exe /loadfile)
      on the target machine - join it during unattended setup or first boot,
      before it has live DC connectivity.
    .PARAMETER ComputerName
      Computer account name to provision.
    .PARAMETER Domain
      Domain (DNS or NetBIOS) to join.
    .PARAMETER Credential
      Optional credential with rights to create computer accounts. When
      omitted, the current context is used.
    .PARAMETER MachineAccountOU
      Optional distinguished name of the OU for the machine account, e.g.
      'OU=Computers,DC=contoso,DC=com'.
    .PARAMETER DCName
      Optional target domain controller.
    .PARAMETER ProvisionOptions
      NetCreateProvisioningPackage provisioning options. Defaults to 2
      (NETSETUP_PROVISION_REUSE_ACCOUNT, matching the reference behaviour).
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result with a Blob property.
    .EXAMPLE
      # AD-side: provision the computer account and generate the join blob.
      PS> $result = New-OfflineDomainJoinBlob -ComputerName 'SRV001' -Domain 'contoso.com' -Credential (Get-Credential) -MachineAccountOU 'OU=Servers,DC=contoso,DC=com'

      # Target-side: write the djoin-format file, then copy it to the new machine.
      PS> New-DjoinFile -Blob $result.Blob -DestinationFile 'C:\join\SRV001.djoin'

      # On the target machine, join offline during unattended setup / first boot:
      PS> djoin.exe /loadfile SRV001.djoin /reuse
    .LINK
      https://github.com/adnoctem/winkit/lib/provisioning.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
  param (
    [Parameter(Mandatory = $true)]
    [Alias('MachineName')]
    [string]
    $ComputerName,

    [Parameter(Mandatory = $true)]
    [string]
    $Domain,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter(Mandatory = $false)]
    [string]
    $MachineAccountOU,

    [Parameter(Mandatory = $false)]
    [string]
    $DCName,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 4096)]
    [uint32]
    $ProvisionOptions = 2
  )

  if (-not $PSCmdlet.ShouldProcess("$ComputerName in $Domain", 'Provision computer account and generate offline-join blob')) {
    if ($WhatIfPreference) {
      New-OperationResult -Target $ComputerName -Source 'NetAPI32' -Action 'ProvisionAccount' -Status 'DryRun' -Detail "Would provision '$ComputerName' in '$Domain'."
    }
    return
  }

  try {
    $blob = [PSFOfflineJoin.Native]::CreateProvisioningPackage(
      $Domain,
      $ComputerName,
      $MachineAccountOU,
      $DCName,
      $ProvisionOptions,
      $(if ($Credential) { $Credential.UserName } else { $null }),
      $(if ($Credential) { $Credential.GetNetworkCredential().Password } else { $null }))

    if ([string]::IsNullOrWhiteSpace($blob)) {
      New-OperationResult -Target $ComputerName -Source 'NetAPI32' -Action 'ProvisionAccount' -Status 'Failed' -ErrorMessage 'The API returned an empty blob.'
    }
    else {
      New-OperationResult -Target $ComputerName -Source 'NetAPI32' -Action 'ProvisionAccount' -Status 'Completed' -Detail "Provisioned '$ComputerName' in '$Domain'." -Property @{ Blob = $blob }
    }
  }
  catch {
    New-OperationResult -Target $ComputerName -Source 'NetAPI32' -Action 'ProvisionAccount' -Status 'Failed' -ErrorMessage $_.Exception.Message
  }
}

function New-DjoinFile {
  <#
    .SYNOPSIS
      New-DjoinFile - Writes an offline-join blob in the format djoin.exe consumes.
    .DESCRIPTION
      Writes the blob produced by New-OfflineDomainJoinBlob (or djoin.exe
      /requestjoindomain /saveold) to a file in the format Windows Setup and
      djoin.exe /loadfile expect: UTF-16LE with a BOM, blob data, and two
      terminating null bytes.
    .PARAMETER Blob
      The offline-join blob string.
    .PARAMETER DestinationFile
      Full path of the file to write. Defaults to C:\temp\djoin.tmp.
    .OUTPUTS
      PSCustomObject - New-OperationResult-shaped result.
    .EXAMPLE
      # AD-side: provision the computer account and generate the join blob.
      PS> $result = New-OfflineDomainJoinBlob -ComputerName 'SRV001' -Domain 'contoso.com' -Credential (Get-Credential) -MachineAccountOU 'OU=Servers,DC=contoso,DC=com'

      # Target-side: write the djoin-format file, then copy it to the new machine.
      PS> New-DjoinFile -Blob $result.Blob -DestinationFile 'C:\join\SRV001.djoin'

      # On the target machine, join offline during unattended setup / first boot:
      PS> djoin.exe /loadfile SRV001.djoin /reuse

      # New-DjoinFile also accepts blobs obtained elsewhere, e.g. from a
      # reference machine: djoin.exe /requestjoindomain /saveold C:\join\REF.djoin
    .LINK
      https://github.com/adnoctem/winkit/lib/provisioning.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
  param (
    [Parameter(Mandatory = $true)]
    [string]
    $Blob,

    [Parameter(Mandatory = $false)]
    [System.IO.FileInfo]
    $DestinationFile = 'C:\temp\djoin.tmp'
  )

  if (-not $PSCmdlet.ShouldProcess($DestinationFile.FullName, 'Write offline-join blob file')) {
    if ($WhatIfPreference) {
      New-OperationResult -Target $DestinationFile.FullName -Source 'DJoin' -Action 'WriteBlob' -Status 'DryRun' -Detail 'No file written.'
    }
    return
  }

  try {
    $byteChain = [byte[]]::new(2)
    $byteChain[0] = 255
    $byteChain[1] = 254

    $fileStream = $DestinationFile.OpenWrite()
    try {
      $byteChain += [System.Text.Encoding]::Unicode.GetBytes($Blob)
      $byteChain += 0
      $byteChain += 0
      $fileStream.Write($byteChain, 0, $byteChain.Length)
    }
    finally {
      $fileStream.Close()
    }

    New-OperationResult -Target $DestinationFile.FullName -Source 'DJoin' -Action 'WriteBlob' -Status 'Completed' -Detail "Wrote $($byteChain.Length) bytes."
  }
  catch {
    New-OperationResult -Target $DestinationFile.FullName -Source 'DJoin' -Action 'WriteBlob' -Status 'Failed' -ErrorMessage $_.Exception.Message
  }
}
