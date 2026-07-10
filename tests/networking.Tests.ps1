#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/common.ps1
  . $PSScriptRoot/../src/networking.ps1
  . $PSScriptRoot/../src/log.ps1
}

Describe 'Test-IPv4Address' {
  Context 'valid addresses' {
    It 'accepts standard private address' {
      Test-IPv4Address '192.168.1.1' | Should -BeTrue
    }

    It 'accepts loopback address' {
      Test-IPv4Address '127.0.0.1' | Should -BeTrue
    }

    It 'accepts lowest address' {
      Test-IPv4Address '0.0.0.0' | Should -BeTrue
    }

    It 'accepts highest address' {
      Test-IPv4Address '255.255.255.255' | Should -BeTrue
    }

    It 'accepts private network address' {
      Test-IPv4Address '10.0.0.0' | Should -BeTrue
    }
  }

  Context 'invalid addresses' {
    It 'rejects too few octets' {
      Test-IPv4Address '192.168.1' | Should -BeFalse
    }

    It 'rejects too many octets' {
      Test-IPv4Address '192.168.1.1.1' | Should -BeFalse
    }

    It 'rejects octet above 255' {
      Test-IPv4Address '192.168.1.256' | Should -BeFalse
    }

    It 'rejects negative octet' {
      Test-IPv4Address '192.168.-1.1' | Should -BeFalse
    }

    It 'rejects leading zero in octet' {
      Test-IPv4Address '192.168.01.1' | Should -BeFalse
    }

    It 'rejects non-numeric octet' {
      Test-IPv4Address '192.168.abc.1' | Should -BeFalse
    }
  }
}

Describe 'Test-IPv6Address' {
  Context 'valid addresses' {
    It 'accepts full uncompressed form' {
      Test-IPv6Address '2001:0db8:85a3:0000:0000:8a2e:0370:7334' | Should -BeTrue
    }

    It 'accepts compressed middle groups' {
      Test-IPv6Address '2001:db8::ff00:42:8329' | Should -BeTrue
    }

    It 'accepts loopback' {
      Test-IPv6Address '::1' | Should -BeTrue
    }

    It 'accepts unspecified address' {
      Test-IPv6Address '::' | Should -BeTrue
    }

    It 'accepts link-local address' {
      Test-IPv6Address 'fe80::1' | Should -BeTrue
    }

    It 'accepts compressed at end' {
      Test-IPv6Address '2001:db8:1:2:3:4:5::' | Should -BeTrue
    }

    It 'accepts compressed at start' {
      Test-IPv6Address '::ff00:42:8329' | Should -BeTrue
    }

    It 'accepts IPv4-mapped address' {
      Test-IPv6Address '::ffff:192.168.1.1' | Should -BeTrue
    }

    It 'rejects IPv4-mapped with invalid embedded IPv4' {
      Test-IPv6Address '::ffff:192.168.1.256' | Should -BeFalse
    }
  }

  Context 'invalid addresses' {
    It 'rejects multiple :: compressions' {
      Test-IPv6Address '2001::db8::1' | Should -BeFalse
    }

    It 'rejects invalid hex characters' {
      Test-IPv6Address '2001:db8:xxxx::1' | Should -BeFalse
    }

    It 'rejects group longer than 4 hex digits' {
      Test-IPv6Address '2001:db8:12345::1' | Should -BeFalse
    }

    It 'rejects too many groups without compression' {
      Test-IPv6Address '2001:db8:1:2:3:4:5:6:7' | Should -BeFalse
    }
  }
}

Describe 'Get-DefaultNetworkAdapter' {
  It 'returns an adapter object with expected properties' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100'); IPSubnet = @('255.255.255.0'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4') }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }

    $result = Get-DefaultNetworkAdapter
    $result | Should -Not -BeNullOrEmpty
    $result.Name | Should -Be 'Ethernet'
    $result.ifIndex | Should -Be 1
  }

  It 'returns null when no default route exists' {
    Mock Get-NetRoute { @() }

    $result = Get-DefaultNetworkAdapter
    $result | Should -BeNullOrEmpty
  }

  It 'throws when Required is set and no route exists' {
    Mock Get-NetRoute { @() }

    { Get-DefaultNetworkAdapter -Required -ErrorAction Stop } | Should -Throw
  }
}

Describe 'Get-IPAddress' {
  It 'returns an IP address from the default adapter' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }
    Mock Get-NetIPAddress {
      @([PSCustomObject]@{ IPAddress = '192.168.1.100'; InterfaceIndex = 1; AddressFamily = 'IPv4' })
    }

    $result = Get-IPAddress
    $result | Should -Be '192.168.1.100'
  }

  It 'returns null when adapter is missing' {
    Mock Get-NetRoute { @() }

    $result = Get-IPAddress
    $result | Should -BeNullOrEmpty
  }
}

Describe 'Get-SubnetMask' {
  It 'returns a subnet mask string' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }
    Mock Get-NetIPAddress {
      @([PSCustomObject]@{ IPAddress = '192.168.1.100'; InterfaceIndex = 1; AddressFamily = 'IPv4' })
    }

    $result = Get-SubnetMask
    $result | Should -Not -BeNullOrEmpty
    $result | Should -Match '^(\d{1,3}\.){3}\d{1,3}$'
  }
}

Describe 'Get-DefaultGateway' {
  It 'returns the gateway IP from the default route' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }

    $result = Get-DefaultGateway
    $result | Should -Be '192.168.1.1'
  }

  It 'returns null when no default route exists' {
    Mock Get-NetRoute { @() }

    $result = Get-DefaultGateway
    $result | Should -BeNullOrEmpty
  }
}

Describe 'Get-DNSServer' {
  It 'returns DNS server addresses' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }
    Mock Get-DnsClientServerAddress {
      @([PSCustomObject]@{ InterfaceIndex = 1; ServerAddresses = @('8.8.8.8', '8.8.4.4') })
    }

    $result = Get-DNSServer
    $result | Should -Contain '8.8.8.8'
    @($result).Count | Should -Be 2
  }

  It 'returns null when adapter is missing' {
    Mock Get-NetRoute { @() }

    $result = Get-DNSServer
    $result | Should -BeNullOrEmpty
  }
}

Describe 'Get-MACAddress' {
  It 'returns a formatted MAC address' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up'; MacAddress = 'AA-BB-CC-DD-EE-FF' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }

    $result = Get-MACAddress
    $result | Should -Match '^([0-9A-F]{2}[:-]){5}[0-9A-F]{2}$'
  }
}

Describe 'Get-NetworkPrefix' {
  It 'returns a network prefix string' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }
    Mock Get-NetIPAddress {
      @([PSCustomObject]@{ IPAddress = '192.168.1.100'; PrefixLength = 24; AddressFamily = 'IPv4' })
    }

    $result = Get-NetworkPrefix
    $result | Should -Not -BeNullOrEmpty
    $result | Should -Match '^(\d{1,3}\.){3}\d{1,3}$'
  }
}

Describe 'Get-NetworkPrefixCIDR' {
  It 'returns a CIDR notation string' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }
    Mock Get-NetIPAddress {
      @([PSCustomObject]@{ IPAddress = '192.168.1.100'; PrefixLength = 24; AddressFamily = 'IPv4' })
    }

    $result = Get-NetworkPrefixCIDR
    $result | Should -Match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/\d{1,2}$'
  }
}

Describe 'Get-BroadcastAddress' {
  It 'returns a broadcast address string' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }
    Mock Get-NetIPAddress {
      @([PSCustomObject]@{ IPAddress = '192.168.1.100'; PrefixLength = 24; AddressFamily = 'IPv4' })
    }

    $result = Get-BroadcastAddress
    $result | Should -Not -BeNullOrEmpty
    $result | Should -Match '^(\d{1,3}\.){3}\d{1,3}$'
  }
}

Describe 'Get-MulticastAddress' {
  It 'returns a multicast address string' {
    Mock Get-NetRoute {
      @([PSCustomObject]@{ DestinationPrefix = '0.0.0.0/0'; NextHop = '192.168.1.1'; ifIndex = 1; RouteMetric = 25 })
    }
    Mock Get-NetAdapter {
      @([PSCustomObject]@{ Name = 'Ethernet'; ifIndex = 1; InterfaceDescription = 'Test Adapter'; Status = 'Up' })
    }
    Mock Get-CimInstance {
      [PSCustomObject]@{ InterfaceIndex = 1; IPAddress = @('192.168.1.100', '2001:db8::1'); IPSubnet = @('255.255.255.0', '64'); DefaultIPGateway = @('192.168.1.1'); DNSServerSearchOrder = @('8.8.8.8', '8.8.4.4'); MACAddress = 'AA:BB:CC:DD:EE:FF' }
    } -ParameterFilter { $ClassName -eq 'Win32_NetworkAdapterConfiguration' }
    Mock Get-NetIPAddress {
      @([PSCustomObject]@{ IPAddress = '192.168.1.100'; PrefixLength = 24; AddressFamily = 'IPv4' })
    }

    $result = Get-MulticastAddress
    $result | Should -Not -BeNullOrEmpty
    $result | Should -Match '^[0-9a-f:]+$'
  }
}
