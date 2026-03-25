# Requires -Version 5.1
# Pester test suite for Set-BoxToOneDriveItemPermission.ps1.
#
# Test design goals:
# - Keep tests isolated from tenant dependencies by mocking all PnP and file I/O commands.
# - Verify high-risk behavior added during recent refactors.
# - Use explicit Arrange-Act-Assert sections in each test for maintainability.

Set-StrictMode -Version Latest

Describe 'Set-BoxToOneDriveItemPermission.ps1' {

    BeforeAll {
        $OriginalScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Set-BoxToOneDriveItemPermission.ps1'
        $script:ScriptUnderTest = Join-Path -Path $TestDrive -ChildPath 'Set-BoxToOneDriveItemPermission.NoRequires.ps1'
        [void] $script:ScriptUnderTest

        # Remove #Requires lines in the test copy so module-loading issues do not block unit tests.
        $ScriptWithoutRequires = Get-Content -LiteralPath $OriginalScriptPath | Where-Object {
            $_ -notmatch '^\s*#Requires\b'
        }
        Set-Content -LiteralPath $script:ScriptUnderTest -Value $ScriptWithoutRequires -Encoding UTF8

        $script:DefaultUser = 'user@contoso.onmicrosoft.com'
        $script:DefaultPersonalSiteUrl = 'https://contoso-my.sharepoint.com/personal/user_contoso_onmicrosoft_com'
        [void] $script:DefaultPersonalSiteUrl

        function New-BoxRow {
            param(
                [string] $CollaboratorPermission = 'Editor',
                [string] $ItemName = 'Doc1.txt',
                [string] $Path = 'Documents',
                [string] $ItemType = 'File',
                [string] $CollaboratorLogin = 'collab@contoso.com'
            )

            [PSCustomObject]@{
                'Owner Login' = $script:DefaultUser
                'Path' = $Path
                'Item Name' = $ItemName
                'Item Type' = $ItemType
                'Collaborator Login' = $CollaboratorLogin
                'Collaborator Permission' = $CollaboratorPermission
            }
        }

        # Create lightweight stubs so Pester can reliably mock wrapper commands
        # before the script under test executes and conditionally defines wrappers.
        foreach ($wrapperName in @(
            'Invoke-SboGetPnPConnection',
            'Invoke-SboConnectPnPOnline',
            'Invoke-SboGetPnPUserProfileProperty',
            'Invoke-SboGetPnPLists',
            'Invoke-SboGetPnPListByIdentity',
            'Invoke-SboGetPnPProperty',
            'Invoke-SboGetPnPFileAsListItem',
            'Invoke-SboGetPnPFolderAsListItem',
            'Invoke-SboSetPnPListItemPermission',
            'Invoke-SboDisconnectPnPOnline'
        )) {
            if (-not (Get-Command -Name $wrapperName -ErrorAction SilentlyContinue)) {
                Set-Item -Path ("Function:" + $wrapperName) -Value { throw 'Wrapper stub should be mocked in tests.' }
            }
        }
    }

    BeforeEach {
        # Shared mock state that each test can override in Arrange as needed.
        $global:TestCsvRows = @(New-BoxRow)
        $global:TestGetConnectionQueue = @(
            [PSCustomObject]@{ Url = 'https://contoso-my.sharepoint.com' },
            $null
        )

        $global:TestHostConnection = [PSCustomObject]@{ Url = 'https://contoso-my.sharepoint.com'; Name = 'HostConnection' }
        $global:TestPersonalConnection = [PSCustomObject]@{ Url = $script:DefaultPersonalSiteUrl; Name = 'PersonalConnection' }

        $global:TestResolvedLibrary = [PSCustomObject]@{
            Title = 'Documents'
            BaseTemplate = 101
            Hidden = $false
            RootFolder = [PSCustomObject]@{
                ServerRelativeUrl = '/personal/user_contoso_onmicrosoft_com/Documents'
            }
        }

        $global:CapturedGetFileUrls = @()
        $global:CapturedPermissionCalls = @()

        Mock Test-Path { $true }
        Mock New-Item {}
        Mock Write-Verbose {}
        Mock Write-Warning {}
        Mock Add-Content {}

        Mock Invoke-SboGetPnPConnection {
            if ($global:TestGetConnectionQueue.Count -eq 0) {
                return $null
            }

            $next = $global:TestGetConnectionQueue[0]
            if ($global:TestGetConnectionQueue.Count -gt 1) {
                $global:TestGetConnectionQueue = $global:TestGetConnectionQueue[1..($global:TestGetConnectionQueue.Count - 1)]
            }
            else {
                $global:TestGetConnectionQueue = @()
            }

            return $next
        }

        Mock Invoke-SboConnectPnPOnline {
            param(
                [string] $Url,
                [string] $ClientId
            )

            if ($Url -notlike '*/personal/*') {
                return $global:TestHostConnection
            }

            return $global:TestPersonalConnection
        }

        Mock Invoke-SboGetPnPUserProfileProperty {
            [PSCustomObject]@{
                PersonalSiteUrl = 'https://contoso-my.sharepoint.com/personal/user_contoso_onmicrosoft_com'
            }
        }

        Mock Import-Csv { $global:TestCsvRows }

        Mock Invoke-SboGetPnPLists {
            @($global:TestResolvedLibrary)
        }

        Mock Invoke-SboGetPnPListByIdentity {
            param(
                [string] $Identity
            )

            if ($Identity -eq 'Documents' -or $Identity -eq 'IgnoredByAutoDiscover') {
                return $global:TestResolvedLibrary
            }

            throw "Unknown list identity in test mock: $Identity"
        }

        Mock Invoke-SboGetPnPProperty {}
        Mock Invoke-SboGetPnPFileAsListItem {
            param(
                [string] $Url,
                [object] $Connection
            )

            $global:CapturedGetFileUrls += $Url
            [PSCustomObject]@{ Id = 101 }
        }
        Mock Invoke-SboGetPnPFolderAsListItem { [PSCustomObject]@{ Id = 202 } }
        Mock Invoke-SboSetPnPListItemPermission {
            param(
                [string] $List,
                [int] $Identity,
                [string] $User,
                [string] $AddRole,
                [object] $Connection
            )

            $global:CapturedPermissionCalls += [PSCustomObject]@{
                List = $List
                Identity = $Identity
                User = $User
                AddRole = $AddRole
            }
        }
        Mock Invoke-SboDisconnectPnPOnline {}

        # Keep these mocks for non-wrapper helpers.
        Mock Import-Csv { $global:TestCsvRows }
    }

    It 'fails unknown collaborator role values by default' {
        # Arrange
        # Strict mode is default now. This row intentionally includes an unmapped role.
        $global:TestCsvRows = @(
            New-BoxRow -CollaboratorPermission 'Custom-Role'
        )
        $errors = @()
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
            ErrorAction = 'SilentlyContinue'
            ErrorVariable = 'errors'
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        $errors | Should -Not -BeNullOrEmpty
        ($errors[0].ToString()) | Should -Match 'Unknown Box role'
        @($result).Count | Should -Be 0
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 0
    }

    It 'allows unmapped role values only when AllowUnknownRole is used' {
        # Arrange
        $global:TestCsvRows = @(
            New-BoxRow -CollaboratorPermission 'Custom-Role'
        )

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            AllowUnknownRole = $true
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        @($result).Count | Should -Be 1
        @($result)[0].PermissionLevel | Should -Be 'Custom-Role'
        @($result)[0].PermissionChangeStatus | Should -Be 'Applied'
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
        $global:CapturedPermissionCalls[0].AddRole | Should -Be 'Custom-Role'
    }

    It 'deduplicates identical CSV rows before attempting permission updates' {
        # Arrange
        $duplicate = New-BoxRow -CollaboratorPermission 'Viewer' -ItemName 'DocA.txt' -Path 'Documents/FolderA' -CollaboratorLogin 'dup@contoso.com'
        $global:TestCsvRows = @($duplicate, $duplicate)

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        @($result).Count | Should -Be 1
        Assert-MockCalled Invoke-SboGetPnPFileAsListItem -Times 1
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
    }

    It 'builds path-aware encoded item URLs from Path plus Item Name' {
        # Arrange
        $global:TestCsvRows = @(
            New-BoxRow -CollaboratorPermission 'Viewer' -Path 'Documents/Folder A/Sub#Folder' -ItemName 'File #%1.txt'
        )
        $expectedServerRelativeUrl = '/personal/user_contoso_onmicrosoft_com/Documents/Folder%20A/Sub%23Folder/File%20%23%251.txt'
        $expectedAbsoluteUrl = "https://contoso-my.sharepoint.com$expectedServerRelativeUrl"
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Invoke-SboGetPnPFileAsListItem -Times 1 -ParameterFilter {
            $true
        }
        $global:CapturedGetFileUrls[0] | Should -Be $expectedServerRelativeUrl
        @($result)[0].ItemUrl.AbsoluteUri | Should -Be $expectedAbsoluteUrl
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
        $global:CapturedPermissionCalls[0].AddRole | Should -Be 'Read'
    }

    It 'does not disconnect reused connections that were not created by this run' {
        # Arrange
        # First Get-PnPConnection call returns tenant connection and second returns personal connection.
        # This simulates full reuse and no new Connect-PnPOnline calls should occur.
        $global:TestGetConnectionQueue = @(
            [PSCustomObject]@{ Url = 'https://contoso-my.sharepoint.com' },
            [PSCustomObject]@{ Url = $script:DefaultPersonalSiteUrl }
        )

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $null = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Invoke-SboConnectPnPOnline -Times 0
        Assert-MockCalled Invoke-SboDisconnectPnPOnline -Times 0
    }

    It 'uses auto-discovered document library when AutoDiscoverDefaultLibrary is provided' {
        # Arrange
        $global:TestResolvedLibrary = [PSCustomObject]@{
            Title = 'Shared Documents'
            BaseTemplate = 101
            Hidden = $false
            RootFolder = [PSCustomObject]@{
                ServerRelativeUrl = '/personal/user_contoso_onmicrosoft_com/Shared Documents'
            }
        }
        $global:TestCsvRows = @(
            New-BoxRow -CollaboratorPermission 'Editor' -Path 'Shared Documents/SubFolder' -ItemName 'Doc2.txt'
        )
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            AutoDiscoverDefaultLibrary = $true
            TargetLibraryTitle = 'IgnoredByAutoDiscover'
            LogFolder = $TestDrive
        }

        # Act
        $null = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Invoke-SboGetPnPLists -Times 1
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
        $global:CapturedPermissionCalls[0].List | Should -Be 'Shared Documents'
    }

    It 'falls back to warning stream when log writes fail' {
        # Arrange
        Mock Add-Content {
            throw 'Simulated disk write failure'
        }

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $null = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Write-Warning -ParameterFilter {
            $Message -like 'Log write failed*'
        }
    }
}
