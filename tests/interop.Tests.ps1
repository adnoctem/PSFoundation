#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  Remove-Module PSFoundation -Force -ErrorAction SilentlyContinue
  Import-Module "$PSScriptRoot/../src/PSFoundation.psd1" -Force
}

AfterAll {
  Remove-Module PSFoundation -Force -ErrorAction SilentlyContinue
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
  BeforeAll {
    # This helper is private; the public Outlook tests use the imported module.
    . $PSScriptRoot/../src/interop.ps1
  }

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

Describe 'Outlook standard folder identities' {
  BeforeAll {
    function New-IdentityAccessor {
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only in-memory fixtures.')]
      [CmdletBinding()]
      param ([hashtable]$Properties = @{})
      $accessor = [PSCustomObject]@{ Properties = $Properties }
      $accessor | Add-Member ScriptMethod GetProperty {
        param($Tag)
        if (-not $this.Properties.ContainsKey($Tag)) {
          throw (New-Object Runtime.InteropServices.COMException('Property not found', -2147221233))
        }
        return , $this.Properties[$Tag]
      }
      $accessor | Add-Member ScriptMethod BinaryToString {
        param($Bytes)
        [BitConverter]::ToString([byte[]]$Bytes).Replace('-', '')
      }
      $accessor
    }
  }

  BeforeEach {
    Mock Remove-ComObject { } -ModuleName PSFoundation
    $accessor = New-IdentityAccessor
    $store = [PSCustomObject]@{
      StoreID          = 'source'
      IsDataFileStore  = $true
      FilePath         = 'C:\fixture.pst'
      PropertyAccessor = $accessor
      Defaults         = @{}
    }
    $store | Add-Member ScriptMethod GetDefaultFolder { param($Kind) $this.Defaults[$Kind] }
    $script:FolderFixtureRoot = [PSCustomObject]@{ StoreID = 'source'; Store = $store; PropertyAccessor = $accessor }
    $inbox = [PSCustomObject]@{ StoreID = 'source'; EntryID = 'inbox'; PropertyAccessor = $accessor }
    $namespace = [PSCustomObject]@{ Application = [PSCustomObject]@{ Version = '12.0' }; DefaultStore = $store; Inbox = $inbox }
    $namespace | Add-Member ScriptMethod GetDefaultFolder {
      param($Kind)
      if ($Kind -ne 6) { throw 'Optional folder creation must never be requested' }
      $this.Inbox
    }
  }

  It 'reads German or renamed standard folders by binary identity on Outlook 2007' {
    $accessor.Properties['http://schemas.microsoft.com/mapi/proptag/0x35E30102'] = [byte[]](1, 2)
    $accessor.Properties['http://schemas.microsoft.com/mapi/proptag/0x35E40102'] = [byte[]](3, 4)
    $accessor.Properties['http://schemas.microsoft.com/mapi/proptag/0x36D81102'] = @([byte[]]@(), [byte[]]@(), [byte[]]@(), [byte[]]@(), [byte[]](5, 6))
    $identities = @(Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot)
    ($identities | Where-Object Kind -EQ DeletedItems).EntryID | Should -Be '0102'
    ($identities | Where-Object Kind -EQ SentItems).EntryID | Should -Be '0304'
    ($identities | Where-Object Kind -EQ Junk).EntryID | Should -Be '0506'
    ($identities | Where-Object Kind -EQ Inbox).EntryID | Should -Be inbox
    $identities.Count | Should -Be 20
  }

  It 'parses RSS extended entry identifiers and rejects corrupt block lengths' {
    $tag = 'http://schemas.microsoft.com/mapi/proptag/0x36D90102'
    $accessor.Properties[$tag] = [byte[]](1, 128, 6, 0, 1, 0, 2, 0, 222, 173)
    $identities = @(Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot)
    ($identities | Where-Object Kind -EQ RssFeeds).EntryID | Should -Be DEAD
    $accessor.Properties[$tag] = [byte[]](1, 128, 255, 0)
    { Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot } | Should -Throw '*bounds*'
  }

  It 'does not confuse another stores Inbox with the selected store' {
    $namespace.DefaultStore = [PSCustomObject]@{ StoreID = 'other' }
    $identities = @(Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot)
    ($identities | Where-Object Kind -EQ Inbox).EntryID | Should -BeNullOrEmpty
    ($identities | Where-Object Kind -EQ Inbox).State | Should -Be Unresolved
  }

  It 'honors extended entry-id boundaries for <Label>' -ForEach @(
    @{ Label = 'unknown block'; Bytes = [byte[]](255, 143, 1, 0, 170, 1, 128, 5, 0, 1, 0, 1, 0, 171); Rss = 'AB'; ToDo = $null }
    @{ Label = 'element sentinel and next block'; Bytes = [byte[]](1, 128, 9, 0, 0, 0, 0, 0, 1, 0, 1, 0, 171, 4, 128, 5, 0, 1, 0, 1, 0, 205); Rss = $null; ToDo = 'CD' }
    @{ Label = 'persist sentinel'; Bytes = [byte[]](0, 0, 0, 0, 1, 128, 5, 0, 1, 0, 1, 0, 171); Rss = $null; ToDo = $null }
    @{ Label = 'unknown element'; Bytes = [byte[]](1, 128, 10, 0, 255, 0, 1, 0, 222, 1, 0, 1, 0, 171); Rss = 'AB'; ToDo = $null }
  ) {
    $accessor.Properties['http://schemas.microsoft.com/mapi/proptag/0x36D90102'] = $Bytes
    $identities = @(Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot)
    ($identities | Where-Object Kind -EQ RssFeeds).EntryID | Should -Be $Rss
    ($identities | Where-Object Kind -EQ ToDo).EntryID | Should -Be $ToDo
  }

  It 'rejects malformed extended entry-id data: <Label>' -ForEach @(
    @{ Label = 'truncated block header'; Bytes = [byte[]](1, 128, 5); ErrorText = '*Truncated PersistData*' }
    @{ Label = 'unknown block out of bounds'; Bytes = [byte[]](255, 143, 5, 0, 170); ErrorText = '*property bounds*' }
    @{ Label = 'truncated element header'; Bytes = [byte[]](1, 128, 1, 0, 1); ErrorText = '*Truncated PersistElement*' }
    @{ Label = 'element out of bounds'; Bytes = [byte[]](1, 128, 4, 0, 1, 0, 2, 0); ErrorText = '*block bounds*' }
    @{ Label = 'nonempty element sentinel'; Bytes = [byte[]](1, 128, 5, 0, 0, 0, 1, 0, 171); ErrorText = '*zero length*' }
  ) {
    $accessor.Properties['http://schemas.microsoft.com/mapi/proptag/0x36D90102'] = $Bytes
    { Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot } | Should -Throw $ErrorText
  }

  It 'uses per-store defaults on modern Outlook without localized names' {
    $namespace.Application.Version = '16.0'
    $store.Defaults[23] = [PSCustomObject]@{ StoreID = 'source'; EntryID = 'renamed-junk' }
    $identities = @(Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot)
    ($identities | Where-Object Kind -EQ Junk).EntryID | Should -Be 'renamed-junk'
    ($identities | Where-Object Kind -EQ Junk).Evidence | Should -Be 'Store.GetDefaultFolder'
  }

  It 'propagates provider access failures instead of treating folders as ordinary mail' {
    $accessor | Add-Member ScriptMethod GetProperty { param($Tag) throw ('Access denied: ' + $Tag) } -Force
    { Get-OutlookStandardFolderIdentity -Namespace $namespace -StoreRoot $script:FolderFixtureRoot } | Should -Throw '*Access denied*'
  }
}

Describe 'Outlook folder selection plan' {
  BeforeAll {
    function New-PlanFolder {
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates only in-memory fixtures.')]
      [CmdletBinding()]
      param ([string]$Name, [string]$Id = $Name, [int]$Type = 1, [int]$ItemType = 0)
      $accessor = [PSCustomObject]@{ Type = $Type }
      $accessor | Add-Member ScriptMethod GetProperty { param($Tag) if ($Tag) { $this.Type } }
      $collection = [PSCustomObject]@{ Values = (New-Object Collections.ArrayList) }
      $collection | Add-Member ScriptProperty Count { $this.Values.Count }
      $collection | Add-Member ScriptMethod Item { param($Index) $this.Values[$Index - 1] }
      [PSCustomObject]@{
        Name             = $Name
        EntryID          = $Id
        StoreID          = 'source'
        FolderPath       = '\\Source\' + $Name
        DefaultItemType  = $ItemType
        PropertyAccessor = $accessor
        Folders          = $collection
      }
    }
  }

  BeforeEach {
    Mock Remove-ComObject { } -ModuleName PSFoundation
    Mock Write-Progress { } -ModuleName PSFoundation
    Mock Get-OutlookStandardFolderIdentity {
      [PSCustomObject]@{ Kind = 'Inbox'; StoreID = 'source'; EntryID = 'inbox'; State = 'Resolved' }
      [PSCustomObject]@{ Kind = 'Junk'; StoreID = 'source'; EntryID = 'junk'; State = 'Resolved' }
      [PSCustomObject]@{ Kind = 'SentItems'; StoreID = 'source'; EntryID = 'sent'; State = 'Resolved' }
    } -ModuleName PSFoundation
    $script:FolderFixtureRoot = New-PlanFolder Root
    $script:FolderFixtureRoot.FolderPath = '\\Source'
    $inbox = New-PlanFolder Posteingang inbox
    $junk = New-PlanFolder 'Renamed junk' junk
    $sent = New-PlanFolder 'Gesendete Objekte' sent
    $nested = New-PlanFolder Nested
    $search = New-PlanFolder Search -Type 2
    $null = $inbox.Folders.Values.Add($nested)
    foreach ($folder in @($inbox, $junk, $sent, $search)) {
      $null = $script:FolderFixtureRoot.Folders.Values.Add($folder)
    }
    $namespace = [PSCustomObject]@{ Inbox = $inbox }
    $namespace | Add-Member ScriptMethod GetFolderFromID {
      param($EntryId, $StoreId)
      if ($EntryId -ne $this.Inbox.EntryID -or $StoreId -ne $this.Inbox.StoreID) {
        throw 'Unexpected folder identity lookup.'
      }
      $this.Inbox
    }
  }

  It 'selects one localized folder and only recurses explicitly' {
    $plan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName Posteingang)
    $plan.Count | Should -Be 1
    $plan[0].EntryID | Should -Be inbox
    $plan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName Posteingang -Recurse)
    $plan.Count | Should -Be 2
    $plan[1].RelativePath | Should -Be 'Posteingang\Nested'
  }

  It 'selects the implicit Inbox by identity when named <Name>' -ForEach @(
    @{ Name = 'Posteingang' }, @{ Name = 'Renamed Inbox' }
  ) {
    $inbox.Name = $Name
    $inbox.FolderPath = '\\Source\' + $Name
    $plan = @(Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot)
    $plan.Count | Should -Be 1
    $plan[0].EntryID | Should -Be inbox
    $plan[0].RelativePath | Should -Be $Name
    $plan[0].StandardKind | Should -Be Inbox
    $plan[0].Process | Should -BeTrue
    $recursive = @(Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot -Recurse -ProgressId 7)
    $recursive.Count | Should -Be 2
    Should -Invoke Write-Progress -ModuleName PSFoundation -ParameterFilter { $Id -eq 7 }
  }

  It 'keeps explicit Inbox paths literal and explicit empty paths at the root' {
    $customInbox = New-PlanFolder Inbox custom-inbox
    $null = $script:FolderFixtureRoot.Folders.Values.Add($customInbox)
    $defaultPlan = @(Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot)
    $literalPlan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName Inbox)
    $rootPlan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName '')
    $defaultPlan[0].EntryID | Should -Be inbox
    $literalPlan[0].EntryID | Should -Be custom-inbox
    $rootPlan[0].EntryID | Should -Be Root
  }

  It 'requires an explicit path when the Inbox identity is <State>' -ForEach @(
    @{ State = 'Absent' }, @{ State = 'Unavailable' }, @{ State = 'Unresolved' }
  ) {
    Mock Get-OutlookStandardFolderIdentity {
      [PSCustomObject]@{ Kind = 'Inbox'; StoreID = 'source'; EntryID = $null; State = $State }
    } -ModuleName PSFoundation
    { Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot } | Should -Throw '*Supply FolderName explicitly*'
    $explicit = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName Posteingang)
    $explicit[0].EntryID | Should -Be inbox
  }

  It 'enforces inclusion and custom exclusions for the implicit Inbox' {
    { Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot -Include @() } | Should -Throw '*IncludeInbox*'
    { Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot -Exclusions Posteingang } | Should -Throw '*CustomExclusion*'
    $script:FolderFixtureRoot.Folders.Values.Remove($inbox)
    $parent = New-PlanFolder Parent
    $null = $parent.Folders.Values.Add($inbox)
    $null = $script:FolderFixtureRoot.Folders.Values.Add($parent)
    $inbox.FolderPath = '\\Source\Parent\Posteingang'
    { Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot -Exclusions Parent } | Should -Throw '*CustomExclusion*'
  }

  It 'rejects mismatched or out-of-root Inbox objects before planning' {
    $namespace | Add-Member ScriptMethod GetFolderFromID { $this.Inbox } -Force
    $inbox.StoreID = 'other'
    { Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot } | Should -Throw '*does not match*'
    $inbox.StoreID = 'source'
    $inbox.FolderPath = '\\Other\Posteingang'
    { Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot } | Should -Throw '*outside*'
  }

  It 'rejects an Inbox path that now points at a different identity' {
    $replacement = New-PlanFolder Posteingang replacement
    $script:FolderFixtureRoot.Folders.Values.Remove($inbox)
    $null = $script:FolderFixtureRoot.Folders.Values.Add($replacement)
    { Get-OutlookFolderPlan -Namespace $namespace -StoreRoot $script:FolderFixtureRoot } | Should -Throw '*identity changed*'
  }

  It 'rejects blank exclusions even when the selected folder is the root' {
    { Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName '' -Exclusions '' } | Should -Throw '*nonblank*'
  }

  It 'requires explicit inclusion of standard folders while skipping search views' {
    $plan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName '' -Recurse)
    @($plan | Where-Object Process).Count | Should -Be 3
    ($plan | Where-Object EntryID -EQ junk).Process | Should -BeFalse
    ($plan | Where-Object EntryID -EQ sent).Process | Should -BeFalse
    ($plan | Where-Object EntryID -EQ Search).Reason | Should -Be SearchFolder
    $plan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName '' -Recurse -Include Inbox, SentItems)
    ($plan | Where-Object EntryID -EQ sent).Process | Should -BeTrue
  }

  It 'never bypasses excluded ancestors by selecting a child directly' {
    { Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName 'Posteingang\Nested' -Exclusions Posteingang } | Should -Throw '*CustomExclusion*'
    $null = $junk.Folders.Values.Add($nested)
    { Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName 'Renamed junk\Nested' } | Should -Throw '*IncludeJunk*'
  }

  It 'gives custom exclusions precedence over identity inclusions without wildcard matching' {
    $plan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName '' -Recurse -Include Inbox, Junk -Exclusions 'Renamed junk')
    ($plan | Where-Object EntryID -EQ junk).Reason | Should -Be CustomExclusion
    $plan = @(Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName '' -Recurse -Include Inbox, Junk -Exclusions '*junk*')
    ($plan | Where-Object EntryID -EQ junk).Process | Should -BeTrue
  }

  It 'does not broaden missing or malformed selections to the store root' {
    { Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName Inbox } | Should -Throw '*not found*'
    { Get-OutlookFolderPlan -Namespace @{} -StoreRoot $script:FolderFixtureRoot -FolderName 'Posteingang\..' } | Should -Throw '*exact*'
  }
}
