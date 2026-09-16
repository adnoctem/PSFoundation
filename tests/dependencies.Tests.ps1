#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
  $script:DependencyTool = Join-Path $PSScriptRoot '../tools/dependencies.ps1'
  . $DependencyTool
}

Describe 'Dependency version updates' {
  BeforeEach {
    $script:RuntimeUpdate = [pscustomobject]@{
      Name = 'Example'
      Source = 'runtime'
      VersionField = 'ModuleVersion'
      DeclaredVersion = '1.2.0'
      LatestVersion = '1.25.0'
    }
    $script:DevUpdate = [pscustomobject]@{
      Name = 'Example'
      Source = 'dev'
      VersionField = 'MinimumVersion'
      DeclaredVersion = '1.2.0'
      LatestVersion = '1.25.0'
    }
    $script:ManifestContent = "@{ RequiredModules = @(@{ ModuleName = 'Example'; ModuleVersion = '1.2.0' }) }"
    $script:DevContent = '[{"Name":"Example","MinimumVersion":"1.2.0"}]'
  }

  It 'updates only the intended runtime dependency with <Layout>' -ForEach @(
    @{ Layout = 'inline fields'; Declaration = "@{ ModuleName = 'Example'; ModuleVersion = '1.2.0' }" },
    @{ Layout = 'reordered fields'; Declaration = "@{ ModuleVersion = '1.2.0'; ModuleName = 'Example' }" },
    @{ Layout = 'double quotes'; Declaration = '@{ ModuleName = "Example"; ModuleVersion = "1.2.0" }' },
    @{ Layout = 'distant fields'; Declaration = @'
@{
  ModuleName = 'Example'
  # Versions can be separated from names by comments and whitespace.



  ModuleVersion = '1.2.0'
}
'@
    }
  ) {
    $manifest = @"
@{
  ModuleVersion = '8.0.0'
  RequiredModules = @(
    # @{ ModuleName = 'Example'; ModuleVersion = '7.0.0' }
    $Declaration
    @{ ModuleName = 'Neighbor'; ModuleVersion = '9.0.0' }
  )
  PrivateData = @{ ModuleName = 'Example'; ModuleVersion = '6.0.0' }
}
"@
    $result = Get-DependencyUpdate -ManifestContent $manifest -DevContent '[]' -Updates $RuntimeUpdate
    $result.ManifestContent | Should -BeExactly ($manifest.Replace('1.2.0', '1.25.0'))
    $result.DevContent | Should -BeExactly '[]'
  }

  It 'reads and updates runtime RequiredVersion pins' {
    $manifest = "@{ RequiredModules = @(@{ RequiredVersion = '1.2.0'; ModuleName = 'Example' }, 'Unpinned') }"
    $document = Get-DependencyDocument -ManifestContent $manifest -DevContent '[]'
    $document.Dependencies[0].VersionField | Should -Be 'RequiredVersion'
    $document.Dependencies[0].DeclaredVersion | Should -Be '1.2.0'
    $document.Dependencies[1].DeclaredVersion | Should -BeNullOrEmpty
    $RuntimeUpdate.VersionField = 'RequiredVersion'
    $result = Get-DependencyUpdate -ManifestContent $manifest -DevContent '[]' -Updates $RuntimeUpdate
    $result.ManifestContent | Should -BeExactly ($manifest.Replace('1.2.0', '1.25.0'))
  }

  It 'updates only the intended dev dependency with <Layout>' -ForEach @(
    @{ Layout = 'inline fields'; Declaration = '{"Name":"Example","MinimumVersion":"1.2.0"}' },
    @{ Layout = 'reordered fields'; Declaration = '{"MinimumVersion":"1.2.0","Name":"Example"}' },
    @{ Layout = 'distant fields'; Declaration = @'
{
  "Name": "Example",



  "MinimumVersion": "1.2.0"
}
'@
    }
  ) {
    $json = '[' + $Declaration + ',{"Name":"Neighbor","MinimumVersion":"9.0.0"}]'
    $result = Get-DependencyUpdate -ManifestContent '@{}' -DevContent $json -Updates $DevUpdate
    $modules = $result.DevContent | ConvertFrom-Json
    $modules.Count | Should -Be 2
    $modules[0].MinimumVersion | Should -Be '1.25.0'
    $modules[1].Name | Should -Be 'Neighbor'
    $modules[1].MinimumVersion | Should -Be '9.0.0'
    $result.ManifestContent | Should -BeExactly '@{}'
  }

  It 'uses RequiredVersion consistently when both dev version fields exist in <Order>' -ForEach @(
    @{ Order = 'minimum first'; Declaration = '{"Name":"Example","MinimumVersion":"1.0.0","RequiredVersion":"1.2.0"}' },
    @{ Order = 'required first'; Declaration = '{"RequiredVersion":"1.2.0","Name":"Example","MinimumVersion":"1.0.0"}' }
  ) {
    $json = '[' + $Declaration + ']'
    $document = Get-DependencyDocument -ManifestContent '@{}' -DevContent $json
    $document.Dependencies[0].VersionField | Should -Be 'RequiredVersion'
    $document.Dependencies[0].DeclaredVersion | Should -Be '1.2.0'
    $DevUpdate.VersionField = 'RequiredVersion'
    $result = Get-DependencyUpdate -ManifestContent '@{}' -DevContent $json -Updates $DevUpdate
    $result.DevContent.TrimStart().StartsWith('[') | Should -BeTrue
    $module = $result.DevContent | ConvertFrom-Json
    $module.RequiredVersion | Should -Be '1.25.0'
    $module.MinimumVersion | Should -Be '1.0.0'
  }

  It 'matches module names literally and preserves unrelated JSON properties' {
    $DevUpdate.Name = 'Example.Core[Tools]'
    $json = '[{"Name":"Example.Core[Tools]","MinimumVersion":"1.2.0","Metadata":{"MinimumVersion":"7.0.0"}},{"Name":"ExampleXCoreT","MinimumVersion":"9.0.0"}]'
    $result = Get-DependencyUpdate -ManifestContent '@{}' -DevContent $json -Updates $DevUpdate
    $modules = $result.DevContent | ConvertFrom-Json
    $modules[0].MinimumVersion | Should -Be '1.25.0'
    $modules[0].Metadata.MinimumVersion | Should -Be '7.0.0'
    $modules[1].MinimumVersion | Should -Be '9.0.0'
  }

  It 'prepares multiple updates together without corrupting later offsets' {
    $manifest = "@{ RequiredModules = @(@{ ModuleName = 'Example'; ModuleVersion = '1.2.0' }, @{ ModuleName = 'Second'; RequiredVersion = '2.0.0' }) }"
    $second = [pscustomobject]@{ Name = 'Second'; Source = 'runtime'; VersionField = 'RequiredVersion'; DeclaredVersion = '2.0.0'; LatestVersion = '20.12.10' }
    $result = Get-DependencyUpdate -ManifestContent $manifest -DevContent $DevContent -Updates @($RuntimeUpdate, $second, $DevUpdate)
    $result.ManifestContent | Should -BeExactly ($manifest.Replace('1.2.0', '1.25.0').Replace('2.0.0', '20.12.10'))
    ($result.DevContent | ConvertFrom-Json).MinimumVersion | Should -Be '1.25.0'
  }

  It 'rejects a proposed version of <Version>' -ForEach @(
    @{ Version = '1.2.0' },
    @{ Version = '1.1.0' },
    @{ Version = 'not-a-version' },
    @{ Version = "2.0.0'" }
  ) {
    $RuntimeUpdate.LatestVersion = $Version
    { Get-DependencyUpdate -ManifestContent $ManifestContent -DevContent '[]' -Updates $RuntimeUpdate } | Should -Throw
  }

  It 'rejects a missing dependency instead of silently succeeding' {
    $RuntimeUpdate.Name = 'Missing'
    { Get-DependencyUpdate -ManifestContent $ManifestContent -DevContent '[]' -Updates $RuntimeUpdate } | Should -Throw '*exactly one*'
  }

  It 'rejects ambiguous duplicate declarations' {
    $json = '[{"Name":"Example","MinimumVersion":"1.2.0"},{"Name":"Example","MinimumVersion":"1.2.0"}]'
    { Get-DependencyUpdate -ManifestContent '@{}' -DevContent $json -Updates $DevUpdate } | Should -Throw '*exactly one*'
  }

  It 'rejects a stale declared version' {
    $RuntimeUpdate.DeclaredVersion = '1.1.0'
    { Get-DependencyUpdate -ManifestContent $ManifestContent -DevContent '[]' -Updates $RuntimeUpdate } | Should -Throw '*does not match*'
  }

  It 'rejects a request to bump the inactive version field' {
    $json = '[{"Name":"Example","MinimumVersion":"1.2.0","RequiredVersion":"1.3.0"}]'
    { Get-DependencyUpdate -ManifestContent '@{}' -DevContent $json -Updates $DevUpdate } | Should -Throw '*does not match*'
  }

  It 'rejects duplicate updates' {
    { Get-DependencyUpdate -ManifestContent $ManifestContent -DevContent '[]' -Updates @($RuntimeUpdate, $RuntimeUpdate) } | Should -Throw '*Duplicate update*'
  }

  It 'rejects malformed input before preparing updates' -ForEach @(
    @{ Manifest = '@{'; Json = '[]' },
    @{ Manifest = '@{}'; Json = '[{' },
    @{ Manifest = '@{}'; Json = '{}' }
  ) {
    { Get-DependencyUpdate -ManifestContent $Manifest -DevContent $Json -Updates $RuntimeUpdate } | Should -Throw
  }

  It 'rejects serialized output that did not apply the requested bump' {
    Mock ConvertTo-Json { '[{"Name":"Example","MinimumVersion":"1.2.0"}]' }
    { Get-DependencyUpdate -ManifestContent '@{}' -DevContent $DevContent -Updates $DevUpdate } | Should -Throw '*verification failed*'
  }

  It 'rejects malformed serialized output' {
    Mock ConvertTo-Json { '[{' }
    { Get-DependencyUpdate -ManifestContent '@{}' -DevContent $DevContent -Updates $DevUpdate } | Should -Throw
  }
}

Describe 'Dependency command validation' {
  BeforeAll {
    # Make publishing impossible even if a regression passes the validation gate.
    function git { throw 'Unexpected git invocation in dependency tests.' }
    function gh { throw 'Unexpected gh invocation in dependency tests.' }
  }

  BeforeEach {
    $script:FixtureRoot = Join-Path $TestDrive 'repo'
    $fixtureTools = Join-Path $FixtureRoot 'tools'
    $fixtureSource = Join-Path $FixtureRoot 'src'
    New-Item -ItemType Directory -Path $fixtureTools, $fixtureSource -Force | Out-Null
    $script:FixtureTool = Join-Path $fixtureTools 'dependencies.ps1'
    $script:FixtureManifest = Join-Path $fixtureSource 'PSFoundation.psd1'
    $script:FixtureDev = Join-Path $fixtureTools 'dev-dependencies.json'
    Copy-Item -LiteralPath $DependencyTool -Destination $FixtureTool
    Set-Content -LiteralPath $FixtureManifest -Value "@{ RequiredModules = @(@{ ModuleName = 'Example'; RequiredVersion = '1.2.0' }) }" -Encoding UTF8
    Set-Content -LiteralPath $FixtureDev -Value '[]' -Encoding UTF8
    $script:PreviousGitHubToken = $env:GH_TOKEN
    $env:GH_TOKEN = 'synthetic-dependency-test-token'
    Mock Invoke-WebRequest {
      [pscustomobject]@{ Content = @"
<feed xmlns="http://www.w3.org/2005/Atom" xmlns:m="http://schemas.microsoft.com/ado/2007/08/dataservices/metadata" xmlns:d="http://schemas.microsoft.com/ado/2007/08/dataservices">
  <entry><m:properties><d:Version>1.25.0</d:Version><d:IsPrerelease>false</d:IsPrerelease></m:properties></entry>
</feed>
"@
      }
    }
    Mock git { throw 'Unexpected git invocation in dependency tests.' }
    Mock gh { throw 'Unexpected gh invocation in dependency tests.' }
  }

  AfterEach {
    $env:GH_TOKEN = $PreviousGitHubToken
  }

  It 'returns <ExitCode> in check mode for an exact runtime pin of <Declared>' -ForEach @(
    @{ Declared = '1.2.0'; ExitCode = 1 },
    @{ Declared = '1.25.0'; ExitCode = 0 }
  ) {
    Set-Content -LiteralPath $FixtureManifest -Value "@{ RequiredModules = @(@{ ModuleName = 'Example'; RequiredVersion = '$Declared' }) }" -Encoding UTF8
    $manifestBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureManifest))
    $devBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureDev))
    & $FixtureTool -Check
    $LASTEXITCODE | Should -Be $ExitCode
    [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureManifest)) | Should -BeExactly $manifestBytes
    [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureDev)) | Should -BeExactly $devBytes
    Should -Invoke git -Times 0 -Exactly
    Should -Invoke gh -Times 0 -Exactly
  }

  It 'leaves both files untouched and never stages a partially valid batch' {
    Set-Content -LiteralPath $FixtureDev -Value '[{"Name":"DevExample","MinimumVersion":"1.2.0"},{"Name":"DevExample","MinimumVersion":"1.2.0"}]' -Encoding UTF8
    $manifestBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureManifest))
    $devBytes = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureDev))
    { & $FixtureTool } | Should -Throw '*exactly one*'
    [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureManifest)) | Should -BeExactly $manifestBytes
    [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($FixtureDev)) | Should -BeExactly $devBytes
    Should -Invoke git -Times 0 -Exactly
    Should -Invoke gh -Times 0 -Exactly
  }
}
