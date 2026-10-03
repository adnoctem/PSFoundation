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


Describe 'Existing Outlook PST lifetime' {
  BeforeAll {
    function New-PstTestStore {
      [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'In-memory COM substitute.')]
      [CmdletBinding()]
      param ([string]$Path, [string]$Id = 'pst-id', [string]$Name = 'Same display name')
      $root = [PSCustomObject]@{ StoreID = $Id; Kind = 'Root'; Acquires = 0; Releases = 0 }
      $store = [PSCustomObject]@{ Path = $Path; StoreID = $Id; DisplayName = $Name; Root = $root; Kind = 'Store'; Acquires = 0; Releases = 0; RootFailures = 0; FailPath = $false }
      $store | Add-Member ScriptProperty FilePath { if ($this.FailPath) { throw 'Synthetic FilePath inspection failure' }; $this.Path }
      $store | Add-Member ScriptMethod GetRootFolder {
        if ($this.RootFailures -gt 0) { $this.RootFailures--; throw 'Synthetic root acquisition failure' }
        $this.Root.Acquires++
        $this.Root
      }
      $store
    }
  }

  BeforeEach {
    $script:pstPath = Join-Path $TestDrive ('Archive ä [' + [guid]::NewGuid().ToString('N') + '].pst')
    [IO.File]::WriteAllText($script:pstPath, 'Synthetic file, never opened by Outlook.')
    $script:pstStore = New-PstTestStore $script:pstPath
    $script:pstDuplicate = New-PstTestStore $script:pstPath 'second-id'
    $script:pstProfile = [PSCustomObject]@{
      Items = (New-Object Collections.ArrayList); Collections = (New-Object Collections.ArrayList)
      Added = $script:pstStore; Duplicate = $script:pstDuplicate; Mode = 'Success'
      Adds = 0; Removes = 0; RemovedId = $null; Accesses = 0; FailItem = 0; FailCount = $false; FailStores = $false; FailRemove = $false
      DeleteOnAccess = 0; AttachOnAccess = 0; SourcePath = $script:pstPath; LastAddedPath = $null
    }
    $script:pstNamespace = [PSCustomObject]@{ Profile = $script:pstProfile; Kind = 'Namespace'; Releases = 0 }
    $script:pstNamespace | Add-Member ScriptProperty Stores {
      $this.Profile.Accesses++
      if ($this.Profile.FailStores) { throw 'Synthetic Stores lookup failure' }
      if ($this.Profile.DeleteOnAccess -eq $this.Profile.Accesses) { [IO.File]::Delete($this.Profile.SourcePath) }
      if ($this.Profile.AttachOnAccess -eq $this.Profile.Accesses) { [void]$this.Profile.Items.Add($this.Profile.Added) }
      $collection = [PSCustomObject]@{ Profile = $this.Profile; Kind = 'Stores'; Acquires = 1; Releases = 0 }
      $collection | Add-Member ScriptProperty Count {
        if ($this.Profile.FailCount) { throw 'Synthetic Stores count failure' }
        $this.Profile.Items.Count
      }
      $collection | Add-Member ScriptMethod Item {
        param($Index)
        if ($this.Profile.FailItem -eq $Index) { throw 'Synthetic Store inspection failure' }
        $store = $this.Profile.Items[$Index - 1]
        $store.Acquires++
        $store
      }
      [void]$this.Profile.Collections.Add($collection)
      $collection
    }
    $script:pstNamespace | Add-Member ScriptMethod GetStoreFromID {
      param($Id)
      $stores = @($this.Profile.Items | Where-Object { $_.StoreID -eq $Id })
      if ($stores.Count -ne 1) { throw 'Synthetic store identity not found' }
      $stores[0].Acquires++
      $stores[0]
    }
    $script:pstNamespace | Add-Member ScriptMethod AddStore {
      param($Path)
      $this.Profile.Adds++
      $this.Profile.LastAddedPath = $Path
      if ($this.Profile.Mode -eq 'BeforeThrow') { throw 'Synthetic attach failed before adding' }
      if ($this.Profile.Mode -eq 'ExistingId') { $this.Profile.Added.Path = $Path; return }
      if ($this.Profile.Mode -ne 'Invisible') { [void]$this.Profile.Items.Add($this.Profile.Added) }
      if ($this.Profile.Mode -eq 'InspectAfter') { $this.Profile.FailStores = $true }
      if ($this.Profile.Mode -eq 'Ambiguous') { [void]$this.Profile.Items.Add($this.Profile.Duplicate) }
      if ($this.Profile.Mode -eq 'AfterThrow') { throw 'Synthetic attach failed after adding' }
      'Unexpected AddStore output'
    }
    $script:pstNamespace | Add-Member ScriptMethod RemoveStore {
      param($Root)
      $this.Profile.Removes++
      $this.Profile.RemovedId = $Root.StoreID
      if ($this.Profile.FailRemove) { throw 'Synthetic detach failure' }
      $match = @($this.Profile.Items | Where-Object { $_.StoreID -eq $Root.StoreID })
      foreach ($store in $match) { $this.Profile.Items.Remove($store) }
      'Unexpected RemoveStore output'
    }
    Mock Remove-ComObject {
      foreach ($value in $InputObject) { if ($null -ne $value) { $value.Releases++ } }
    } -ModuleName PSFoundation
    Mock Invoke-ComGarbageCollection { throw 'No global teardown permitted' } -ModuleName PSFoundation
  }

  AfterEach {
    $script:pstNamespace.Releases | Should -Be 0
    foreach ($collection in $script:pstProfile.Collections) { $collection.Releases | Should -Be $collection.Acquires }
    foreach ($store in @($script:pstStore, $script:pstDuplicate)) {
      $store.Releases | Should -Be $store.Acquires
      $store.Root.Releases | Should -Be $store.Root.Acquires
    }
    Should -Invoke Invoke-ComGarbageCollection -ModuleName PSFoundation -Times 0
  }

  It 'reuses an existing attachment by path and preserves its name and borrowed namespace' {
    [void]$script:pstProfile.Items.Add($script:pstDuplicate)
    $script:pstDuplicate.Path = Join-Path $TestDrive 'different.pst'
    [IO.File]::WriteAllText($script:pstDuplicate.Path, 'Synthetic other archive')
    [void]$script:pstProfile.Items.Add($script:pstStore)
    $contexts = @(Open-OutlookPstStore $script:pstNamespace $script:pstPath -WhatIf)
    $contexts.Count | Should -Be 1
    $context = $contexts[0]
    $context.PSObject.TypeNames | Should -Contain 'PSFoundation.OutlookPstStoreContext'
    $context.StoreId | Should -Be 'pst-id'
    $context.AttachedByCall | Should -BeFalse
    [object]::ReferenceEquals($context.Namespace, $script:pstNamespace) | Should -BeTrue
    @(Close-OutlookPstStore $context).Count | Should -Be 0
    Close-OutlookPstStore $context
    $context.Root | Should -BeNullOrEmpty
    $context.Closed | Should -BeTrue
    $script:pstProfile.Adds | Should -Be 0
    $script:pstProfile.Removes | Should -Be 0
    $script:pstStore.DisplayName | Should -Be 'Same display name'
    $script:pstStore.Root.Releases | Should -Be 1
  }

  It 'attaches and detaches once without streaming COM method results' {
    $contexts = @(Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false)
    $contexts.Count | Should -Be 1
    $contexts[0].Path | Should -Be $script:pstPath
    $contexts[0].AttachedByCall | Should -BeTrue
    $contexts[0].Closed | Should -BeFalse
    @(Close-OutlookPstStore $contexts[0]).Count | Should -Be 0
    Close-OutlookPstStore $contexts[0]
    $script:pstProfile.Adds | Should -Be 1
    $script:pstProfile.Removes | Should -Be 1
    $script:pstProfile.LastAddedPath | Should -Be $script:pstPath
    $script:pstProfile.RemovedId | Should -Be 'pst-id'
    Test-Path -LiteralPath $script:pstPath | Should -BeTrue
  }

  It 'normalizes relative, case, literal brackets, Unicode and short spelling' {
    [void]$script:pstProfile.Items.Add($script:pstStore)
    $script:shortPstPath = Join-Path $TestDrive 'SOURCE~1.PST'
    $script:pstStore.Path = $script:shortPstPath
    $script:pstGetItem = Get-Command Microsoft.PowerShell.Management\Get-Item
    $script:pstResolveLongPath = (Get-Command PSFoundation\Resolve-LongPath).ScriptBlock
    Mock Get-Item {
      if ($LiteralPath -eq $script:shortPstPath) { $LiteralPath = $script:pstPath }
      & $script:pstGetItem -LiteralPath $LiteralPath -Force -ErrorAction Stop
    } -ModuleName PSFoundation
    Mock Resolve-LongPath {
      if ($LiteralPath -eq $script:shortPstPath) { return $script:pstPath }
      & $script:pstResolveLongPath -LiteralPath $LiteralPath
    } -ModuleName PSFoundation
    Push-Location $TestDrive
    try {
      $context = Open-OutlookPstStore $script:pstNamespace ([IO.Path]::GetFileName($script:pstPath).ToUpperInvariant())
      $context.Path | Should -Be $script:pstPath
      Close-OutlookPstStore $context
    }
    finally { Pop-Location }
    $script:pstProfile.Adds | Should -Be 0
  }

  It 'rejects <Case> sources before any Outlook inspection' -ForEach @(
    @{ Case = 'missing' }, @{ Case = 'directory' }, @{ Case = 'extension' }, @{ Case = 'readonly' },
    @{ Case = 'provider' }, @{ Case = 'UNC' }, @{ Case = 'mapped network' }, @{ Case = 'empty' }, @{ Case = 'reparse' }
  ) {
    $path = $script:pstPath
    switch ($Case) {
      missing { $path = Join-Path $TestDrive 'absent.pst' }
      directory { $path = $TestDrive }
      extension { $path = Join-Path $TestDrive 'source.ost'; [IO.File]::WriteAllText($path, 'Synthetic OST') }
      readonly { [IO.File]::SetAttributes($path, [IO.FileAttributes]::ReadOnly) }
      provider { $path = 'Env:PATH' }
      UNC { $path = '\\unreachable.invalid\share\archive.pst' }
      'mapped network' { Mock Get-PSFOutlookPathDriveType { [IO.DriveType]::Network } -ModuleName PSFoundation }
      empty { $path = ' ' }
      reparse {
        Mock Get-Item { [PSCustomObject]@{ PSIsContainer = $false; Extension = '.pst'; Attributes = [IO.FileAttributes]::ReparsePoint } } -ModuleName PSFoundation
      }
    }
    try { { Open-OutlookPstStore $script:pstNamespace $path -Confirm:$false } | Should -Throw }
    finally { if ($Case -eq 'readonly') { [IO.File]::SetAttributes($path, [IO.FileAttributes]::Normal) } }
    $script:pstProfile.Accesses | Should -Be 0
    $script:pstProfile.Adds | Should -Be 0
  }

  It 'rejects ambiguous matches before attachment or root acquisition' {
    [void]$script:pstProfile.Items.Add($script:pstStore)
    [void]$script:pstProfile.Items.Add($script:pstDuplicate)
    { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } | Should -Throw '*Multiple*'
    $script:pstProfile.Adds | Should -Be 0
    $script:pstStore.Root.Acquires | Should -Be 0
  }

  It 'tolerates blank unrelated store paths' {
    $script:pstDuplicate.Path = ''
    [void]$script:pstProfile.Items.Add($script:pstDuplicate)
    $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false
    Close-OutlookPstStore $context
    $script:pstProfile.Items.Count | Should -Be 1
  }

  It 'surfaces <Failure> inspection failure and releases acquired references' -ForEach @(
    @{ Failure = 'Stores' }, @{ Failure = 'Count' }, @{ Failure = 'Item' }, @{ Failure = 'FilePath' }, @{ Failure = 'unreadable path' }
  ) {
    [void]$script:pstProfile.Items.Add($script:pstStore)
    switch ($Failure) {
      Stores { $script:pstProfile.FailStores = $true }
      Count { $script:pstProfile.FailCount = $true }
      Item { $script:pstProfile.FailItem = 1 }
      FilePath { $script:pstStore.FailPath = $true }
      'unreadable path' { $script:pstStore.Path = Join-Path $TestDrive 'missing-attached.pst' }
    }
    { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } | Should -Throw '*Cannot inspect*'
    $script:pstProfile.Adds | Should -Be 0
  }

  It 'returns no context or attachment during standalone WhatIf' {
    @(Open-OutlookPstStore $script:pstNamespace $script:pstPath -WhatIf).Count | Should -Be 0
    $script:pstProfile.Adds | Should -Be 0
  }

  It 'cleans up explicit inspection under inherited WhatIf' {
    $WhatIfPreference = $true
    $context = $null
    try { $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath -WhatIf:$false -Confirm:$false }
    finally { if ($null -ne $context) { Close-OutlookPstStore $context } }
    $script:pstProfile.Adds | Should -Be 1
    $script:pstProfile.Removes | Should -Be 1
  }

  It 'rechecks file existence immediately before attaching' {
    $script:pstProfile.DeleteOnAccess = 2
    { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } | Should -Throw
    $script:pstProfile.Adds | Should -Be 0
  }

  It 'reuses an attachment that appeared before the attachment call without claiming it' {
    $script:pstProfile.AttachOnAccess = 2
    $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false
    $context.AttachedByCall | Should -BeFalse
    Close-OutlookPstStore $context
    $script:pstProfile.Adds | Should -Be 0
    $script:pstProfile.Removes | Should -Be 0
  }

  It 'handles <Failure> during opening without detaching unrelated stores' -ForEach @(
    @{ Failure = 'BeforeThrow'; Removes = 0 }, @{ Failure = 'AfterThrow'; Removes = 1 },
    @{ Failure = 'Ambiguous'; Removes = 0 }, @{ Failure = 'Invisible'; Removes = 0 },
    @{ Failure = 'Root'; Removes = 1 }, @{ Failure = 'Context'; Removes = 1 }, @{ Failure = 'InspectAfter'; Removes = 0 }
  ) {
    $script:pstProfile.Mode = $Failure
    if ($Failure -eq 'Root') { $script:pstStore.RootFailures = 1 }
    if ($Failure -eq 'Context') { Mock New-PSFOutlookPstContext { throw 'Synthetic context construction failure' } -ModuleName PSFoundation }
    { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } | Should -Throw
    $script:pstProfile.Removes | Should -Be $Removes
    if ($Removes) { $script:pstProfile.RemovedId | Should -Be 'pst-id' }
  }

  It 'retains the opening error and reports rollback failure' {
    $script:pstProfile.Mode = 'AfterThrow'
    $script:pstProfile.FailRemove = $true
    $failure = $null
    try { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } catch { $failure = $_ }
    $failure.Exception.Message | Should -Match 'attach failed after adding'
    $failure.Exception.Data['OutlookPstCleanupError'] | Should -Match 'detach failure'
    $failure.ErrorDetails.Message | Should -Match ([regex]::Escape($script:pstPath))
    $script:pstProfile.Items.Count | Should -Be 1
  }

  It 'never claims a previously observed StoreID after AddStore' {
    $script:pstStore.Path = ''
    [void]$script:pstProfile.Items.Add($script:pstStore)
    $script:pstProfile.Mode = 'ExistingId'
    { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } | Should -Throw '*uniquely identified new*'
    $script:pstProfile.Removes | Should -Be 0
    $script:pstStore.Root.Acquires | Should -Be 0
  }

  It 'rejects <Identity> StoreIDs without mutation' -ForEach @(@{ Identity = 'missing' }, @{ Identity = 'duplicate' }) {
    [void]$script:pstProfile.Items.Add($script:pstStore)
    if ($Identity -eq 'missing') { $script:pstStore.StoreID = '' }
    else {
      $script:pstDuplicate.StoreID = $script:pstStore.StoreID
      [void]$script:pstProfile.Items.Add($script:pstDuplicate)
    }
    { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } | Should -Throw '*StoreIDs*'
    $script:pstProfile.Adds | Should -Be 0
    $script:pstProfile.Removes | Should -Be 0
  }

  It 'leaves a pre-existing store attached after <Failure> lookup failure' -ForEach @(@{ Failure = 'root' }, @{ Failure = 'context' }) {
    [void]$script:pstProfile.Items.Add($script:pstStore)
    if ($Failure -eq 'root') { $script:pstStore.RootFailures = 1 }
    else { Mock New-PSFOutlookPstContext { throw 'Synthetic context construction failure' } -ModuleName PSFoundation }
    { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } | Should -Throw
    $script:pstProfile.Adds | Should -Be 0
    $script:pstProfile.Removes | Should -Be 0
  }

  It 'reports a mismatched root identity without detaching it' {
    $script:pstStore.Root.StoreID = 'unrelated-root'
    $failure = $null
    try { Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false } catch { $failure = $_ }
    $failure.Exception.Message | Should -Match 'no matching root'
    $failure.Exception.Data['OutlookPstCleanupError'] | Should -Match 'no matching root'
    $script:pstProfile.Removes | Should -Be 0
    $script:pstStore.Root.Releases | Should -Be 2
  }

  It 'never detaches an externally <State> attachment' -ForEach @(@{ State = 'removed' }, @{ State = 'replaced' }) {
    $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false
    $script:pstProfile.Items.Clear()
    if ($State -eq 'replaced') { [void]$script:pstProfile.Items.Add($script:pstDuplicate) }
    Close-OutlookPstStore $context
    $context.Closed | Should -BeTrue
    $script:pstProfile.Removes | Should -Be 0
  }

  It 'closes and releases once after <Failure> cleanup failure' -ForEach @(@{ Failure = 'detach' }, @{ Failure = 'inspection' }) {
    $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false
    if ($Failure -eq 'detach') { $script:pstProfile.FailRemove = $true }
    else { $script:pstProfile.FailStores = $true }
    $failureRecord = $null
    try { Close-OutlookPstStore $context } catch { $failureRecord = $_ }
    $failureRecord.Exception.Message | Should -Match ([regex]::Escape($script:pstPath))
    $context.Closed | Should -BeTrue
    $context.Root | Should -BeNullOrEmpty
    $attempts = $script:pstProfile.Removes
    Close-OutlookPstStore $context
    $script:pstProfile.Removes | Should -Be $attempts
    $script:pstStore.Root.Releases | Should -Be 1
  }

  It 'refuses detachment when the recorded StoreID now identifies another path' {
    $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false
    $script:pstStore.Path = Join-Path $TestDrive 'replacement.pst'
    [IO.File]::WriteAllText($script:pstStore.Path, 'Synthetic replacement')
    { Close-OutlookPstStore $context } | Should -Throw '*no longer matches*'
    $context.Closed | Should -BeTrue
    $script:pstProfile.Removes | Should -Be 0
  }

  It 'keeps legacy named-store lookup and Unicode destination creation usable' {
    [void]$script:pstProfile.Items.Add($script:pstStore)
    $root = Get-OutlookStoreRoot -Namespace $script:pstNamespace -Name 'Same display name'
    [object]::ReferenceEquals($root, $script:pstStore.Root) | Should -BeTrue
    & (Get-Module PSFoundation) { param($value) Remove-ComObject $value } $root
    $script:pstProfile.Items.Clear()
    $script:pstNamespace | Add-Member NoteProperty Format $null
    $script:pstNamespace | Add-Member ScriptMethod AddStoreEx {
      param($Path, $Type)
      $this.Format = $Type
      $this.Profile.Added.Path = $Path
      [void]$this.Profile.Items.Add($this.Profile.Added)
    }
    $destination = Join-Path $TestDrive 'new-destination.pst'
    $root = Add-OutlookStoreRoot -Namespace $script:pstNamespace -Path $destination
    $script:pstNamespace.Format | Should -Be 2
    $script:pstStore.FilePath | Should -Be $destination
    [object]::ReferenceEquals($root, $script:pstStore.Root) | Should -BeTrue
    & (Get-Module PSFoundation) { param($value) Remove-ComObject $value } $root
  }

  It 'does not use edited public metadata to acquire detachment authority' {
    [void]$script:pstProfile.Items.Add($script:pstStore)
    $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath
    $context.AttachedByCall = $true
    $context.StoreId = 'second-id'
    Close-OutlookPstStore $context
    $script:pstProfile.Removes | Should -Be 0
  }

  It 'rejects ordinary and deserialized contexts' {
    { Close-OutlookPstStore ([PSCustomObject]@{ Root = $script:pstStore.Root }) } | Should -Throw '*live object*'
    $context = Open-OutlookPstStore $script:pstNamespace $script:pstPath -Confirm:$false
    try {
      # Do not serialize live references: even inspection of COM getters could
      # acquire new objects. Reproduce only the deserialized type marker.
      $copy = [PSCustomObject]@{ PSTypeName = 'Deserialized.PSFoundation.OutlookPstStoreContext'; Root = $null; _State = $null }
      { Close-OutlookPstStore $copy } | Should -Throw '*live object*'
    }
    finally { Close-OutlookPstStore $context }
  }
}

Describe 'Get-OutlookRepairToolInfo' {
  It 'classifies executable version <Version> as targeted=<Supported>' -ForEach @(
    @{ Version = '12.0.6650.5000'; Supported = $false }
    @{ Version = '16.0.10325.20081'; Supported = $false }
    @{ Version = '16.0.10325.20082'; Supported = $true }
    @{ Version = '16.0.19000.20000'; Supported = $true }
    @{ Version = '17.0.20000.20000'; Supported = $false }
  ) {
    InModuleScope PSFoundation -Parameters @{ TestVersion = [version]$Version; Expected = $Supported } {
      param($TestVersion, $Expected)
      Mock Resolve-LongPath { 'C:\Long directory\SCANPST.EXE' }
      Mock Get-Item {
        [PSCustomObject]@{
          PSIsContainer = $false
          Extension     = '.exe'
          DirectoryName = 'C:\Long directory'
          VersionInfo   = [PSCustomObject]@{
            FileMajorPart   = $TestVersion.Major
            FileMinorPart   = $TestVersion.Minor
            FileBuildPart   = $TestVersion.Build
            FilePrivatePart = $TestVersion.Revision
          }
        }
      }
      $info = Get-OutlookRepairToolInfo -LiteralPath 'C:\LONGDI~1\SCANPST.EXE'
      $info.Path | Should -Be 'C:\Long directory\SCANPST.EXE'
      $info.FileVersion | Should -Be $TestVersion
      $info.SupportsFileArgument | Should -Be $Expected
    }
  }

  It 'keeps a versionless explicit executable interactive' {
    $path = Join-Path $TestDrive 'legacy.exe'
    [IO.File]::WriteAllText($path, 'not a versioned PE image')
    $info = Get-OutlookRepairToolInfo -LiteralPath $path
    $info.SupportsFileArgument | Should -BeFalse
  }
}
