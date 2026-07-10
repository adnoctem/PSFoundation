#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/log.ps1
}

Describe 'Show-Color' {
  It 'does not throw and produces console color output' {
    { Show-Color } | Should -Not -Throw
  }
}

Describe 'Write-Log' {
  It 'writes a message with default color without timestamps' {
    Mock Write-Host { }
    Write-Log -Message 'Hello'
    Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
      $Object -eq 'Hello' -and $ForegroundColor -eq 'White'
    }
  }

  It 'writes a message with a specified color' {
    Mock Write-Host { }
    Write-Log -Message 'Warning' -Color Yellow
    Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
      $Object -eq 'Warning' -and $ForegroundColor -eq 'Yellow'
    }
  }

  It 'prepends a timestamp when -Timestamps is set' {
    Mock Write-Host { }
    Write-Log -Message 'Timed' -Color Green -Timestamps
    Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter {
      $Object -match '^\[' -and $Object -match '\]: Timed$'
    }
  }
}
