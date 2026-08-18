function Get-UserInfo {
  <#
    .SYNOPSIS
      Get-UserInfo - Retrieves information about the current user, including their username, administrator status, and SID.
    .DESCRIPTION
      This function gathers details about the currently logged-in user by accessing the WindowsIdentity class. It checks if the user has administrator privileges and returns a hashtable containing the username, administrator status, and SID of the user.
    .EXAMPLE
      PS> Get-UserInfo
    .LINK
      https://github.com/adnoctem/winkit/lib/user.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([hashtable])]
  param ()

  $user = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  $isAdmin = (New-Object System.Security.Principal.WindowsPrincipal($user)).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

  return @{
    UserName = $user.Name
    IsAdministrator = $isAdmin
    SID = $user.User.Value
  }
}

function Get-UserSID {
  <#
    .SYNOPSIS
      Get-UserSID - Retrieves the Security Identifier (SID) for a specified user.
    .DESCRIPTION
      This function takes a username as input and attempts to retrieve the corresponding SID by creating an NTAccount object and translating it to a SecurityIdentifier. If the user is not found, it returns null and logs an error message.
    .PARAMETER UserName
      The username for which to retrieve the SID (e.g., "DOMAIN\Username").
    .EXAMPLE
      PS> Get-UserSID -UserName "DOMAIN\Username"
    .LINK
      https://github.com/adnoctem/winkit/lib/user.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([string])]
  param (
    [Parameter(Mandatory = $true)]
    [string]$UserName
  )

  try {
    $user = New-Object System.Security.Principal.NTAccount($UserName)
    $sid = $user.Translate([System.Security.Principal.SecurityIdentifier])
    return $sid.Value
  }
  catch {
    Write-Error "Could not find SID for user '$UserName'. $_"
    return $null
  }
}

function Test-ADCredential {
  <#
    .SYNOPSIS
      Test-ADCredential - Validates a username/password pair against AD.
    .DESCRIPTION
      Simulates an authentication request against the domain using
      System.DirectoryServices.AccountManagement.PrincipalContext.
      ValidateCredentials, without opening a session. Returns $true only when
      both the username and password are valid.

      The domain is auto-detected from the credential when not specified.
      Accepts pipeline input of one or more credentials.
    .PARAMETER Credential
      The credential to validate.
    .PARAMETER Domain
      Domain to validate against. Defaults to the credential's own domain.
    .OUTPUTS
      bool
    .EXAMPLE
      PS> Test-ADCredential -Credential (Get-Credential)
    .LINK
      https://github.com/adnoctem/winkit/lib/user.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([bool])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('PSCredential')]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter(Mandatory = $false)]
    [string]
    $Domain
  )

  begin {
    try {
      Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
    }
    catch {
      throw 'Test-ADCredential requires the System.DirectoryServices.AccountManagement assembly.'
    }

    if ([string]::IsNullOrWhiteSpace($Domain)) {
      $Domain = $Credential.GetNetworkCredential().Domain
    }

    $principalContext = [System.DirectoryServices.AccountManagement.PrincipalContext]::new(
      [System.DirectoryServices.AccountManagement.ContextType]::Domain, $Domain)
  }

  process {
    foreach ($item in $Credential) {
      $networkCredential = $item.GetNetworkCredential()
      try {
        [bool]$principalContext.ValidateCredentials($networkCredential.UserName, $networkCredential.Password)
      }
      catch {
        Write-Verbose "Credential validation failed: $($_.Exception.Message)"
        $false
      }
    }
  }

  end {
    $principalContext.Dispose()
  }
}

function Get-ADAccountLockoutSource {
  <#
    .SYNOPSIS
      Get-ADAccountLockoutSource - Reports the source of account lockouts from the PDC.
    .DESCRIPTION
      Dynamically resolves the PDC emulator for the domain (never hardcoded)
      and queries its Security event log for Event ID 4740 (account locked
      out), reporting which account was locked, when, and the ClientName of
      the machine that triggered the lockout. The ClientName pinpoints the
      device with stale cached credentials that is repeatedly failing auth.
    .PARAMETER DomainName
      Domain to query. Defaults to the current user's domain.
    .PARAMETER UserName
      Account filter (wildcards supported). Defaults to '*'.
    .PARAMETER StartTime
      Only report lockouts at or after this time. Defaults to the last 24
      hours.
    .PARAMETER Credential
      Optional alternative credential for the PDC query.
    .OUTPUTS
      One object per lockout event: TimeCreated, UserName, ClientName.
    .EXAMPLE
      PS> Get-ADAccountLockoutSource -UserName 'svc-backup' -StartTime (Get-Date).AddDays(-7)
    .LINK
      https://github.com/adnoctem/winkit/lib/user.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'UserName and StartTime are consumed by the remote scriptblock via the $Using: scope modifier.')]
  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [string]
    $DomainName = $env:USERDOMAIN,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]
    $UserName = '*',

    [Parameter(Mandatory = $false)]
    [datetime]
    $StartTime = (Get-Date).AddDays(-1),

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    $Credential
  )

  begin {
    function Resolve-PDCServer {
      param (
        [string]$Domain = $env:USERDOMAIN,
        [System.Management.Automation.PSCredential]$Credential
      )

      if ($Credential) {
        $context = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::new(
          'Domain', $Domain, $Credential.UserName, $Credential.GetNetworkCredential().Password)
      }
      else {
        $context = [System.DirectoryServices.ActiveDirectory.DirectoryContext]::new('Domain', $Domain)
      }

      [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context).PdcRoleOwner.Name
    }

    try {
      $pdc = Resolve-PDCServer -Domain $DomainName -Credential $Credential
      Write-Verbose "PDC emulator resolved: $pdc"
    }
    catch {
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }

  process {
    $properties = @(
      'TimeCreated',
      @{ Label = 'UserName'; Expression = { $_.Properties[0].Value } },
      @{ Label = 'ClientName'; Expression = { $_.Properties[1].Value } }
    )

    $invokeParams = @{
      ComputerName = $pdc
      ScriptBlock = {
        $props = @(
          'TimeCreated',
          @{ Label = 'UserName'; Expression = { $_.Properties[0].Value } },
          @{ Label = 'ClientName'; Expression = { $_.Properties[1].Value } }
        )
        Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4740; StartTime = $Using:StartTime } |
          Where-Object { $_.Properties[0].Value -like $Using:UserName } |
          Select-Object -Property $props
      }
      ErrorAction = 'Stop'
    }
    if ($Credential) {
      $invokeParams.Credential = $Credential
    }

    try {
      Invoke-Command @invokeParams | Select-Object -Property $properties
    }
    catch {
      $PSCmdlet.ThrowTerminatingError($_)
    }
  }
}

function Get-ADFSMORoleHolder {
  <#
    .SYNOPSIS
      Get-ADFSMORoleHolder - Reports all five FSMO role holders in one call.
    .DESCRIPTION
      Queries the forest and domain for the Schema Master, Domain Naming
      Master, Infrastructure Master, RID Master, and PDC Emulator role
      holders. Loads the ActiveDirectory module if it is not already loaded.
    .PARAMETER Credential
      Optional alternative credential for the queries.
    .OUTPUTS
      PSCustomObject with SchemaMaster, DomainNamingMaster,
      InfrastructureMaster, RIDMaster, and PDCEmulator.
    .EXAMPLE
      PS> Get-ADFSMORoleHolder
    .LINK
      https://github.com/adnoctem/winkit/lib/user.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $false)]
    [Alias('RunAs')]
    [System.Management.Automation.PSCredential]
    $Credential
  )

  try {
    if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
      Import-Module -Name ActiveDirectory -ErrorAction Stop
    }

    if ($Credential) {
      $forestRoles = Get-ADForest -Credential $Credential -ErrorAction Stop
      $domainRoles = Get-ADDomain -Credential $Credential -ErrorAction Stop
    }
    else {
      $forestRoles = Get-ADForest -ErrorAction Stop
      $domainRoles = Get-ADDomain -ErrorAction Stop
    }

    [PSCustomObject]@{
      SchemaMaster = $forestRoles.SchemaMaster
      DomainNamingMaster = $forestRoles.DomainNamingMaster
      InfrastructureMaster = $domainRoles.InfrastructureMaster
      RIDMaster = $domainRoles.RIDMaster
      PDCEmulator = $domainRoles.PDCEmulator
    }
  }
  catch {
    $PSCmdlet.ThrowTerminatingError($_)
  }
}
