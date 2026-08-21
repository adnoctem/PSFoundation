#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  . $PSScriptRoot/../src/interop.ps1
}

Describe 'Remove-ComObject' {
  It 'does not throw when passed null' {
    { Remove-ComObject $null } | Should -Not -Throw
  }

  It 'does not throw when passed multiple values including null' {
    { Remove-ComObject $null, 'non-com-object', $null } | Should -Not -Throw
  }

  It 'does not throw when no arguments are passed' {
    { Remove-ComObject } | Should -Not -Throw
  }
}

Describe 'Invoke-ComGarbageCollection' {
  It 'does not throw' {
    { Invoke-ComGarbageCollection } | Should -Not -Throw
  }
}

Describe 'Get-TransportMessageId' {
  It 'extracts a Message-ID from a Unicode-style header block' {
    $_headers = "Received: from smtp.local (10.0.0.1)`r`nMessage-ID: <abc123@example.com>`r`nSubject: test"
    Get-TransportMessageId -HeaderText $_headers | Should -Be '<abc123@example.com>'
  }

  It 'extracts a Message-ID from an ANSI-style header block' {
    $_headers = "Message-ID: <xyz789@example.com>`nDate: Mon, 1 Jan 2024 00:00:00 +0000"
    Get-TransportMessageId -HeaderText $_headers | Should -Be '<xyz789@example.com>'
  }

  It 'matches Message-ID case-insensitively' {
    Get-TransportMessageId -HeaderText 'message-id: <CaseTest@example.com>' | Should -Be '<CaseTest@example.com>'
  }

  It 'ignores Received headers that embed a Message-ID reference' {
    $_headers = "Received: from a (b) by c; with Message-ID <wrong@example.com>`r`nMessage-ID: <right@example.com>"
    Get-TransportMessageId -HeaderText $_headers | Should -Be '<right@example.com>'
  }

  It 'returns the first Message-ID when several are present' {
    $_headers = "Message-ID: <first@example.com>`r`nMessage-ID: <second@example.com>"
    Get-TransportMessageId -HeaderText $_headers | Should -Be '<first@example.com>'
  }

  It 'returns $null when no Message-ID is present' {
    Get-TransportMessageId -HeaderText 'Received: from smtp.local`r`nSubject: none here' | Should -Be $null
  }

  It 'returns $null for empty or whitespace input' {
    Get-TransportMessageId -HeaderText '' | Should -Be $null
    Get-TransportMessageId -HeaderText '   ' | Should -Be $null
  }
}
