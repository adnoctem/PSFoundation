# --- Internal ----------------------------------------------------------------

function Resolve-IPv6PrefixData {
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Data is a singular mass noun; there is no other option here.')]

  # Extracts the IPv6 address, masked prefix, and prefix length from a resolved
  # adapter's CIM parallel IPAddress/IPSubnet arrays. Shared by Get-NetworkPrefix
  # and Get-NetworkPrefixCIDR to avoid duplicating the masking logic.
  param([Parameter(Mandatory)][PSCustomObject]$Adapter)

  $addresses = $Adapter.CimConfig.IPAddress
  $subnets = $Adapter.CimConfig.IPSubnet

  for ($i = 0; $i -lt $addresses.Count; $i++) {
    if ($addresses[$i] -notmatch ':') { continue }

    $ipv6 = $addresses[$i]
    $sub = $subnets[$i]

    $prefixLength = if ($sub -match '^\d+$') {
      [int]$sub
    }
    else {
      $bits = 0
      foreach ($byte in [System.Net.IPAddress]::Parse($sub).GetAddressBytes()) {
        if ($byte -eq 0xFF) { $bits += 8; continue }
        for ($b = 7; $b -ge 0; $b--) { if (($byte -shr $b) -band 1) { $bits++ } else { break } }
        break
      }
      $bits
    }

    $ipBytes = [System.Net.IPAddress]::Parse($ipv6).GetAddressBytes()
    $prefixBytes = [byte[]]::new(16)
    $remaining = $prefixLength

    for ($j = 0; $j -lt 16; $j++) {
      if ($remaining -ge 8) { $prefixBytes[$j] = $ipBytes[$j]; $remaining -= 8 }
      elseif ($remaining -gt 0) { $prefixBytes[$j] = $ipBytes[$j] -band ([byte](0xFF -shl (8 - $remaining))); $remaining = 0 }
      else { $prefixBytes[$j] = 0 }
    }

    return [PSCustomObject]@{
      Address = $ipv6
      Prefix = [System.Net.IPAddress]::new($prefixBytes).ToString()
      PrefixLength = $prefixLength
    }
  }

  return $null
}

# --- Adapter resolution -------------------------------------------------------

function Get-DefaultNetworkAdapter {
  <#
    .SYNOPSIS
      Resolves the default network adapter via the IPv4 routing table.
    .DESCRIPTION
      Selects the lowest-metric 0.0.0.0/0 route and resolves the corresponding
      adapter. Physical type is identified via NDIS PhysicalMediaType, making the
      result locale-independent. Returns a structured object consumed by all
      downstream networking functions.
    .PARAMETER Type
      Filters by physical adapter type: WiFi (802.11), Ethernet (802.3),
      VPN (Unspecified), or Any. Defaults to Any.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [PSCustomObject] Name, ifIndex, PhysicalMedia, NetAdapter, CimConfig
    .EXAMPLE
      PS> Get-DefaultNetworkAdapter -Type Ethernet -Required
  #>
  param(
    [ValidateSet('Any', 'WiFi', 'Ethernet', 'VPN')]
    [string]$Type = 'Any',

    [switch]$Required
  )

  $defaultRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
    Where-Object { $_.NextHop -ne '0.0.0.0' } |
    Sort-Object -Property RouteMetric |
    Select-Object -First 1

  if ($null -eq $defaultRoute) {
    $message = "No default IPv4 route found."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  $netAdapter = Get-NetAdapter -InterfaceIndex $defaultRoute.ifIndex -ErrorAction SilentlyContinue

  if ($null -eq $netAdapter) {
    $message = "Could not resolve adapter for interface index $($defaultRoute.ifIndex)."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  if ($Type -ne 'Any') {
    $isMatch = switch ($Type) {
      'WiFi' { $netAdapter.PhysicalMediaType -like '*802.11*' }
      'Ethernet' { $netAdapter.PhysicalMediaType -eq '802.3' }
      'VPN' { $netAdapter.PhysicalMediaType -eq 'Unspecified' }
    }

    if (-not $isMatch) {
      $message = "Default adapter '$($netAdapter.Name)' (type: $($netAdapter.PhysicalMediaType)) does not match the requested type '$Type'."
      if ($Required) { throw $message }
      Write-Log -Message $message -Color Red
      return $null
    }
  }

  $cimConfig = Get-CimInstance -Class Win32_NetworkAdapterConfiguration `
    -Filter "InterfaceIndex = $($defaultRoute.ifIndex)"

  if ($null -eq $cimConfig) {
    $message = "Could not retrieve CIM configuration for adapter '$($netAdapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return [PSCustomObject]@{
    Name = $netAdapter.Name
    ifIndex = $defaultRoute.ifIndex
    PhysicalMedia = $netAdapter.PhysicalMediaType
    NetAdapter = $netAdapter
    CimConfig = $cimConfig
  }
}

# --- Address retrieval --------------------------------------------------------

function Get-IPAddress {
  <#
    .SYNOPSIS
      Returns the IP address of the default network adapter.
    .DESCRIPTION
      For IPv6, global/unique-local addresses are preferred over link-local.
      Falls back to link-local if no routable address is assigned.
    .PARAMETER AddressFamily
      IPv4 or IPv6. Defaults to IPv4.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-IPAddress -AddressFamily IPv6
  #>
  param(
    [ValidateSet('IPv4', 'IPv6')]
    [string]$AddressFamily = 'IPv4',

    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  $ip = if ($AddressFamily -eq 'IPv4') {
    $Adapter.CimConfig.IPAddress | Where-Object { $_ -notmatch ':' } | Select-Object -First 1
  }
  else {
    $global = $Adapter.CimConfig.IPAddress |
      Where-Object { $_ -match ':' -and $_ -notmatch '^fe80' } |
      Select-Object -First 1

    if ($null -ne $global) {
      $global
    }
    else {
      $Adapter.CimConfig.IPAddress | Where-Object { $_ -match ':' } | Select-Object -First 1
    }
  }

  if ($null -eq $ip) {
    $message = "No $AddressFamily address found on adapter '$($Adapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return $ip
}

function Get-SubnetMask {
  <#
    .SYNOPSIS
      Returns the IPv4 subnet mask of the default network adapter.
    .DESCRIPTION
      Resolves the mask from the CIM IPSubnet array at the index corresponding
      to the IPv4 entry in the parallel IPAddress array.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-SubnetMask
  #>
  param(
    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  $addresses = $Adapter.CimConfig.IPAddress
  $subnets = $Adapter.CimConfig.IPSubnet
  $mask = $null

  for ($i = 0; $i -lt $addresses.Count; $i++) {
    if ($addresses[$i] -notmatch ':') { $mask = $subnets[$i]; break }
  }

  if ($null -eq $mask) {
    $message = "No IPv4 subnet mask found on adapter '$($Adapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return $mask
}

function Get-DefaultGateway {
  <#
    .SYNOPSIS
      Returns the default gateway of the default network adapter.
    .PARAMETER AddressFamily
      IPv4 or IPv6. Defaults to IPv4.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-DefaultGateway -AddressFamily IPv4
  #>
  param(
    [ValidateSet('IPv4', 'IPv6')]
    [string]$AddressFamily = 'IPv4',

    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  $gateway = if ($AddressFamily -eq 'IPv4') {
    $Adapter.CimConfig.DefaultIPGateway | Where-Object { $_ -notmatch ':' } | Select-Object -First 1
  }
  else {
    $Adapter.CimConfig.DefaultIPGateway | Where-Object { $_ -match ':' } | Select-Object -First 1
  }

  if ($null -eq $gateway) {
    $message = "No $AddressFamily default gateway found on adapter '$($Adapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return $gateway
}

function Get-DNSServer {
  <#
    .SYNOPSIS
      Returns the configured DNS servers of the default network adapter.
    .PARAMETER AddressFamily
      IPv4 or IPv6. Defaults to IPv4.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string[]]
    .EXAMPLE
      PS> Get-DNSServer
  #>
  param(
    [ValidateSet('IPv4', 'IPv6')]
    [string]$AddressFamily = 'IPv4',

    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  $servers = if ($AddressFamily -eq 'IPv4') {
    $Adapter.CimConfig.DNSServerSearchOrder | Where-Object { $_ -notmatch ':' }
  }
  else {
    $Adapter.CimConfig.DNSServerSearchOrder | Where-Object { $_ -match ':' }
  }

  if ($null -eq $servers -or @($servers).Count -eq 0) {
    $message = "No $AddressFamily DNS servers found on adapter '$($Adapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return $servers
}

function Get-MACAddress {
  <#
    .SYNOPSIS
      Returns the MAC address of the default network adapter.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-MACAddress
  #>
  param(
    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  $mac = $Adapter.CimConfig.MACAddress

  if ($null -eq $mac) {
    $message = "No MAC address found on adapter '$($Adapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return $mac
}

# --- Network / prefix ---------------------------------------------------------

function Get-NetworkPrefix {
  <#
    .SYNOPSIS
      Returns the network address (IPv4) or prefix address (IPv6) of the default adapter.
    .PARAMETER AddressFamily
      IPv4 or IPv6. Defaults to IPv4.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-NetworkPrefix
    .EXAMPLE
      PS> Get-Prefix -AddressFamily IPv6
  #>
  [Alias('Get-Network', 'Get-Prefix')]
  param(
    [ValidateSet('IPv4', 'IPv6')]
    [string]$AddressFamily = 'IPv4',

    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  if ($AddressFamily -eq 'IPv4') {
    $ip = Get-IPAddress -AddressFamily IPv4 -Adapter $Adapter -Required:$Required
    $mask = Get-SubnetMask -Adapter $Adapter -Required:$Required

    if ($null -eq $ip -or $null -eq $mask) { return $null }

    $ipBytes = [System.Net.IPAddress]::Parse($ip).GetAddressBytes()
    $maskBytes = [System.Net.IPAddress]::Parse($mask).GetAddressBytes()
    $netBytes = [byte[]](0..3 | ForEach-Object { $ipBytes[$_] -band $maskBytes[$_] })

    return [System.Net.IPAddress]::new($netBytes).ToString()
  }

  $data = Resolve-IPv6PrefixData -Adapter $Adapter

  if ($null -eq $data) {
    $message = "No IPv6 address with a valid prefix length found on adapter '$($Adapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return $data.Prefix
}

function Get-NetworkPrefixCIDR {
  <#
    .SYNOPSIS
      Returns the network prefix in CIDR notation for the default adapter.
    .PARAMETER AddressFamily
      IPv4 or IPv6. Defaults to IPv4.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-NetworkPrefixCIDR
    .EXAMPLE
      PS> Get-PrefixCIDR -AddressFamily IPv6
  #>
  [Alias('Get-NetworkCIDR', 'Get-PrefixCIDR')]
  param(
    [ValidateSet('IPv4', 'IPv6')]
    [string]$AddressFamily = 'IPv4',

    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  if ($AddressFamily -eq 'IPv4') {
    $prefix = Get-NetworkPrefix -AddressFamily IPv4 -Adapter $Adapter -Required:$Required
    $mask = Get-SubnetMask -Adapter $Adapter -Required:$Required

    if ($null -eq $prefix -or $null -eq $mask) { return $null }

    $cidr = (([System.Net.IPAddress]::Parse($mask).GetAddressBytes() |
          ForEach-Object { [Convert]::ToString($_, 2) }) -join '').Replace('0', '').Length

    return "$prefix/$cidr"
  }

  $data = Resolve-IPv6PrefixData -Adapter $Adapter

  if ($null -eq $data) {
    $message = "No IPv6 address with a valid prefix length found on adapter '$($Adapter.Name)'."
    if ($Required) { throw $message }
    Write-Log -Message $message -Color Red
    return $null
  }

  return "$($data.Prefix)/$($data.PrefixLength)"
}

function Get-BroadcastAddress {
  <#
    .SYNOPSIS
      Returns the IPv4 broadcast address of the default network adapter.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-BroadcastAddress
  #>
  param(
    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  $ip = Get-IPAddress -AddressFamily IPv4 -Adapter $Adapter -Required:$Required
  $mask = Get-SubnetMask -Adapter $Adapter -Required:$Required

  if ($null -eq $ip -or $null -eq $mask) { return $null }

  $ipBytes = [System.Net.IPAddress]::Parse($ip).GetAddressBytes()
  $maskBytes = [System.Net.IPAddress]::Parse($mask).GetAddressBytes()
  $broadcastBytes = [byte[]](0..3 | ForEach-Object { $ipBytes[$_] -bor (-bnot $maskBytes[$_] -band 0xFF) })

  return [System.Net.IPAddress]::new($broadcastBytes).ToString()
}

function Get-MulticastAddress {
  <#
    .SYNOPSIS
      Returns the solicited-node multicast address for the default adapter's IPv6 address.
    .DESCRIPTION
      Computes the solicited-node multicast address per RFC 4291 section 2.7.1 by combining
      the ff02::1:ff00:0/104 prefix with the lower 24 bits of the unicast address.
      Used by Neighbor Discovery as the IPv6 replacement for ARP.
    .PARAMETER Adapter
      Pre-resolved adapter from Get-DefaultNetworkAdapter. Resolved automatically if omitted.
    .PARAMETER Required
      Throws on failure instead of returning $null.
    .OUTPUTS
      [string]
    .EXAMPLE
      PS> Get-MulticastAddress
  #>
  param(
    [PSCustomObject]$Adapter,
    [switch]$Required
  )

  if ($null -eq $Adapter) { $Adapter = Get-DefaultNetworkAdapter -Required:$Required }
  if ($null -eq $Adapter) { return $null }

  $ip = Get-IPAddress -AddressFamily IPv6 -Adapter $Adapter -Required:$Required

  if ($null -eq $ip) { return $null }

  $ipBytes = [System.Net.IPAddress]::Parse($ip).GetAddressBytes()
  $multicastBytes = [byte[]]@(
    0xFF, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x01, 0xFF,
    $ipBytes[13], $ipBytes[14], $ipBytes[15]
  )

  return [System.Net.IPAddress]::new($multicastBytes).ToString()
}

# --- Validation ---------------------------------------------------------------

function Test-IPv4Address {
  <#
    .SYNOPSIS
      Returns $true if the input is a valid IPv4 address per RFC 791.
    .DESCRIPTION
      Validates via arithmetic decomposition: four dot-separated decimal octets,
      each in 0-255, no leading zeros. Does not use regex.
    .PARAMETER Address
      The string to validate.
    .OUTPUTS
      [bool]
    .EXAMPLE
      PS> Test-IPv4Address -Address '192.168.1.1'
  #>
  param(
    [Parameter(Mandatory)]
    [string]$Address
  )

  if ([string]::IsNullOrEmpty($Address)) { return $false }

  $octets = $Address -split '\.'
  if ($octets.Count -ne 4) { return $false }

  foreach ($octet in $octets) {
    if ($octet.Length -eq 0) { return $false }
    if ($octet.Length -gt 1 -and $octet[0] -eq '0') { return $false }
    $value = 0
    if (-not [int]::TryParse($octet, [ref]$value)) { return $false }
    if ($value -lt 0 -or $value -gt 255) { return $false }
  }

  return $true
}

function Test-IPv6Address {
  <#
    .SYNOPSIS
      Returns $true if the input is a valid IPv6 address per RFC 4291.
    .DESCRIPTION
      Validates via structural decomposition: handles full form, :: compression,
      and IPv4-mapped addresses (::ffff:x.x.x.x). The embedded IPv4 portion of
      mapped addresses is delegated to Test-IPv4Address.
    .PARAMETER Address
      The string to validate.
    .OUTPUTS
      [bool]
    .EXAMPLE
      PS> Test-IPv6Address -Address '2001:db8::ff00:42:8329'
  #>
  param(
    [Parameter(Mandatory)]
    [string]$Address
  )

  if ([string]::IsNullOrEmpty($Address)) { return $false }

  # Delegate embedded IPv4 portion of mapped addresses (e.g. ::ffff:192.168.1.1)
  $lastColon = $Address.LastIndexOf(':')
  if ($lastColon -ge 0) {
    $lastSegment = $Address.Substring($lastColon + 1)
    if ($lastSegment.Contains('.')) {
      if (-not (Test-IPv4Address -Address $lastSegment)) { return $false }
      $Address = $Address.Substring(0, $lastColon + 1) + '0'
    }
  }

  # :: must appear at most once
  $pos = 0; $doubleColonCount = 0
  while (($pos = $Address.IndexOf('::', $pos)) -ge 0) { $doubleColonCount++; $pos += 2 }
  if ($doubleColonCount -gt 1) { return $false }

  $hasCompression = $doubleColonCount -eq 1

  $groupsToValidate = if ($hasCompression) {
    $parts = $Address -split '::', 2
    $leftGroups = if ($parts[0]) { $parts[0] -split ':' | Where-Object { $_ } } else { @() }
    $rightGroups = if ($parts[1]) { $parts[1] -split ':' | Where-Object { $_ } } else { @() }
    $explicit = @($leftGroups) + @($rightGroups)
    if ($explicit.Count -gt 7) { return $false }
    $explicit
  }
  else {
    $all = $Address -split ':'
    if ($all.Count -ne 8) { return $false }
    $all
  }

  $hexDigits = [char[]]'0123456789abcdefABCDEF'

  foreach ($group in $groupsToValidate) {
    if ($group.Length -lt 1 -or $group.Length -gt 4) { return $false }
    foreach ($char in $group.ToCharArray()) { if ($char -notin $hexDigits) { return $false } }
    if ([Convert]::ToInt32($group, 16) -gt 0xFFFF) { return $false }
  }

  return $true
}

# --- Remote reachability ------------------------------------------------------

function Test-TcpPortReachable {
  <#
    .SYNOPSIS
      Bounded TCP port probe with an explicit timeout.
    .DESCRIPTION
      Internal helper for Test-RemoteHostReachability. Attempts a TCP
      connection to the given port and gives up after TimeoutSeconds. Never
      blocks indefinitely - every caller is network-bound and must stay
      bounded.
    .OUTPUTS
      PSCustomObject with Reachable (bool), Detail (string), LatencyMs (long).
  #>
  param(
    [Parameter(Mandatory)][string]$ComputerName,
    [Parameter(Mandatory)][int]$Port,
    [Parameter(Mandatory)][int]$TimeoutSeconds
  )

  $client = [System.Net.Sockets.TcpClient]::new()
  $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  try {
    $connectTask = $client.ConnectAsync($ComputerName, $Port)
    if (-not $connectTask.Wait([int]($TimeoutSeconds * 1000))) {
      return [PSCustomObject]@{ Reachable = $false; Detail = "TCP port $Port timed out"; LatencyMs = $null }
    }
    if ($client.Connected) {
      return [PSCustomObject]@{ Reachable = $true; Detail = "TCP port $Port is open"; LatencyMs = $stopwatch.ElapsedMilliseconds }
    }
    return [PSCustomObject]@{ Reachable = $false; Detail = "TCP port $Port refused or filtered"; LatencyMs = $stopwatch.ElapsedMilliseconds }
  }
  catch {
    return [PSCustomObject]@{ Reachable = $false; Detail = $_.Exception.Message; LatencyMs = $stopwatch.ElapsedMilliseconds }
  }
  finally {
    $client.Dispose()
  }
}

function Test-RemoteHostReachability {
  <#
    .SYNOPSIS
      Tests which remote-management channels a host answers on.
    .DESCRIPTION
      Given a computer name, probes DNS resolution, ICMP ping, WSMAN
      availability (Test-WSMan), the remote registry service (CIM probe), and
      WMI/RPC connectivity (TCP port 135). The RDP channel (TCP 3389) is only
      probed when -IncludeRdp is supplied - not every managed host needs RDP
      checked. An optional CredSSP authentication attempt runs when both
      -Credential and -CredSSP are supplied (cross-domain / workgroup
      scenarios).

      Every network-touching check has an explicit timeout (-TimeoutSeconds,
      default 5s) - a filtered or unreachable host can never hang the call.

      Returns one structured result object per host, with a convenience
      boolean per channel plus a Channels collection carrying per-channel
      detail and latency.
    .PARAMETER ComputerName
      One or more host names or IP addresses to test.
    .PARAMETER Credential
      Optional credential used for the CredSSP authentication attempt.
    .PARAMETER IncludeRdp
      Also probe RDP (TCP port 3389). Off by default.
    .PARAMETER CredSSP
      Attempt a CredSSP WSMAN authentication probe. Requires -Credential.
    .PARAMETER TimeoutSeconds
      Timeout ceiling for each network-touching check. Defaults to 5.
    .OUTPUTS
      PSCustomObject per host: ComputerName, boolean per channel, and Channels
      (New-OperationResult-shaped per-channel results).
    .EXAMPLE
      PS> Test-RemoteHostReachability -ComputerName 'DC01' -IncludeRdp
    .LINK
      https://github.com/adnoctem/winkit/lib/networking.ps1
    .NOTES
      Author: MVProwess <info@mvprowess.com>
      License: MIT
  #>

  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '', Justification = 'Start-Job scriptblocks receive their values via -ArgumentList, which the analyzer does not recognise.')]
  [OutputType([PSCustomObject])]
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
    [Alias('Computer', 'Host', 'Server')]
    [string[]]
    $ComputerName,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter(Mandatory = $false)]
    [switch]
    $IncludeRdp,

    [Parameter(Mandatory = $false)]
    [switch]
    $CredSSP,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 120)]
    [int]
    $TimeoutSeconds = 5
  )

  process {
    foreach ($name in $ComputerName) {
      $timeoutMs = $TimeoutSeconds * 1000
      $channels = @()

      # DNS resolution
      $dnsOk = $false
      $dnsDetail = 'Failed'
      $dnsLatency = $null
      $dnsWatch = [System.Diagnostics.Stopwatch]::StartNew()
      try {
        $dnsTask = [System.Net.Dns]::GetHostAddressesAsync($name)
        if ($dnsTask.Wait($timeoutMs) -and $dnsTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion -and $dnsTask.Result.Count -gt 0) {
          $dnsOk = $true
          $dnsDetail = "Resolved to $($dnsTask.Result[0]) (+$($dnsTask.Result.Count - 1) more)"
        }
        else {
          $dnsDetail = 'DNS lookup failed or timed out'
        }
      }
      catch {
        $dnsDetail = "DNS lookup failed: $($_.Exception.Message)"
      }
      $dnsLatency = $dnsWatch.ElapsedMilliseconds
      $channels += New-OperationResult -Target $name -Source 'DNS' -Action 'Resolve' -Status $(if ($dnsOk) { 'Reachable' } else { 'Unreachable' }) -Detail $dnsDetail -Property @{ LatencyMs = $dnsLatency }

      # ICMP ping
      $icmpOk = $false
      $icmpDetail = 'Failed'
      $icmpLatency = $null
      $icmpWatch = [System.Diagnostics.Stopwatch]::StartNew()
      try {
        $ping = [System.Net.NetworkInformation.Ping]::new()
        $pingTask = $ping.SendPingAsync($name, $timeoutMs)
        if ($pingTask.Wait($timeoutMs) -and $pingTask.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion -and $pingTask.Result.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
          $icmpOk = $true
          $icmpDetail = "ICMP reply received"
          $icmpLatency = $pingTask.Result.RoundtripTime
        }
        else {
          $icmpDetail = 'ICMP timed out or unreachable'
        }
      }
      catch {
        $icmpDetail = "ICMP failed: $($_.Exception.Message)"
      }
      $icmpLatency = if ($null -eq $icmpLatency) { $icmpWatch.ElapsedMilliseconds } else { $icmpLatency }
      $channels += New-OperationResult -Target $name -Source 'ICMP' -Action 'Ping' -Status $(if ($icmpOk) { 'Reachable' } else { 'Unreachable' }) -Detail $icmpDetail -Property @{ LatencyMs = $icmpLatency }

      # WSMAN availability (Test-WSMan, job-wrapped so it stays bounded)
      $wsmanOk = $false
      $wsmanDetail = 'Failed'
      try {
        $wsmanJob = Start-Job -ScriptBlock { param($hostName) Test-WSMan -ComputerName $hostName -ErrorAction Stop } -ArgumentList $name
        $completed = $wsmanJob | Wait-Job -Timeout $TimeoutSeconds
        if ($null -ne $completed -and $wsmanJob.State -eq 'Completed') {
          $wsmanJobOutput = Receive-Job -Job $wsmanJob -ErrorAction SilentlyContinue
          if ($null -ne $wsmanJobOutput) {
            $wsmanOk = $true
            $wsmanDetail = 'WSMan identity query succeeded'
          }
          else {
            $wsmanDetail = 'WSMan query returned no identity'
          }
        }
        else {
          $wsmanDetail = 'WSMan timed out'
        }
      }
      catch {
        $wsmanDetail = "WSMan failed: $($_.Exception.Message)"
      }
      finally {
        if ($null -ne $wsmanJob) {
          Remove-Job -Job $wsmanJob -Force -ErrorAction SilentlyContinue
        }
      }
      $channels += New-OperationResult -Target $name -Source 'WSMan' -Action 'Query' -Status $(if ($wsmanOk) { 'Reachable' } else { 'Unreachable' }) -Detail $wsmanDetail

      # Remote registry service (CIM probe with operation timeout)
      $regOk = $false
      $regDetail = 'Failed'
      $regState = $null
      try {
        $regService = Get-CimInstance -ClassName Win32_Service -Filter "Name = 'RemoteRegistry'" -ComputerName $name -OperationTimeoutSec $TimeoutSeconds -ErrorAction Stop
        if ($null -ne $regService) {
          $regOk = $true
          $regState = $regService.State
          $regDetail = "RemoteRegistry service state: $regState"
        }
        else {
          $regDetail = 'RemoteRegistry service not found on target'
        }
      }
      catch {
        $regDetail = "CIM probe failed: $($_.Exception.Message)"
      }
      $channels += New-OperationResult -Target $name -Source 'RemoteRegistry' -Action 'Probe' -Status $(if ($regOk) { 'Reachable' } else { 'Unreachable' }) -Detail $regDetail -Property @{ ServiceState = $regState }

      # WMI/RPC connectivity (TCP port 135, the RPC endpoint mapper)
      $wmiProbe = Test-TcpPortReachable -ComputerName $name -Port 135 -TimeoutSeconds $TimeoutSeconds
      $wmiOk = $wmiProbe.Reachable
      $channels += New-OperationResult -Target $name -Source 'WMI/RPC' -Action 'Probe' -Status $(if ($wmiOk) { 'Reachable' } else { 'Unreachable' }) -Detail $wmiProbe.Detail -Property @{ LatencyMs = $wmiProbe.LatencyMs }

      # RDP port - only when explicitly requested
      if ($IncludeRdp) {
        $rdpProbe = Test-TcpPortReachable -ComputerName $name -Port 3389 -TimeoutSeconds $TimeoutSeconds
        $channels += New-OperationResult -Target $name -Source 'RDP' -Action 'Probe' -Status $(if ($rdpProbe.Reachable) { 'Reachable' } else { 'Unreachable' }) -Detail $rdpProbe.Detail -Property @{ LatencyMs = $rdpProbe.LatencyMs }
      }

      # CredSSP authentication attempt - only when a credential is supplied
      if ($CredSSP -and $null -ne $Credential) {
        $credOk = $false
        $credDetail = 'Failed'
        try {
          $credJob = Start-Job -ScriptBlock { param($hostName, [System.Management.Automation.PSCredential]$cred) Test-WSMan -ComputerName $hostName -Authentication CredSSP -Credential $cred -ErrorAction Stop } -ArgumentList $name, $Credential
          $credCompleted = $credJob | Wait-Job -Timeout $TimeoutSeconds
          if ($null -ne $credCompleted -and $credJob.State -eq 'Completed') {
            $credOutput = Receive-Job -Job $credJob -ErrorAction SilentlyContinue
            if ($null -ne $credOutput) {
              $credOk = $true
              $credDetail = 'CredSSP authentication succeeded'
            }
            else {
              $credDetail = 'CredSSP authentication rejected'
            }
          }
          else {
            $credDetail = 'CredSSP attempt timed out'
          }
        }
        catch {
          $credDetail = "CredSSP failed: $($_.Exception.Message)"
        }
        finally {
          if ($null -ne $credJob) {
            Remove-Job -Job $credJob -Force -ErrorAction SilentlyContinue
          }
        }
        $channels += New-OperationResult -Target $name -Source 'CredSSP' -Action 'Authenticate' -Status $(if ($credOk) { 'Reachable' } else { 'Unreachable' }) -Detail $credDetail
      }
      elseif ($CredSSP) {
        Write-Warning "-CredSSP was supplied without -Credential; skipping the CredSSP authentication probe for '$name'."
      }

      $hostResult = [ordered]@{
        ComputerName = $name
        DNSResolved = $dnsOk
        ICMP = $icmpOk
        WSMAN = $wsmanOk
        RemoteRegistry = $regOk
        WMIRPC = $wmiOk
      }

      if ($IncludeRdp) {
        $hostResult.RDP = $rdpProbe.Reachable
      }

      if ($CredSSP -and $null -ne $Credential) {
        $hostResult.CredSSP = $credOk
      }

      $hostResult.Channels = $channels

      [PSCustomObject]$hostResult
    }
  }
}
