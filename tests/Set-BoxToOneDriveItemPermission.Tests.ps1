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
            'Resolve-SboPersonalSiteUrl',
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
        if (-not (Get-Command -Name Resolve-SboPersonalSiteUrl -ErrorAction SilentlyContinue)) {
            Set-Item -Path Function:Resolve-SboPersonalSiteUrl -Value { param([Alias('Profile')][object] $InputObject) $null }
        }

        # Shared mock state that each test can override in Arrange as needed.
        $global:TestCsvRows = @(New-BoxRow)
        $global:TestGetConnectionQueue = @(
            [PSCustomObject]@{ Url = 'https://contoso-admin.sharepoint.com/' },
            $null
        )

        $global:TestAdminConnection = [PSCustomObject]@{ Url = 'https://contoso-admin.sharepoint.com/'; Name = 'AdminConnection' }
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
        $global:CapturedLogLines = @()

        Mock Test-Path { $true }
        Mock New-Item {}
        Mock Write-Verbose {}
        Mock Write-Warning {}
        Mock Add-Content {
            param(
                [string] $LiteralPath,
                [string] $Value
            )

            $global:CapturedLogLines += $Value
        }

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
                return $global:TestAdminConnection
            }

            return $global:TestPersonalConnection
        }

        Mock Invoke-SboGetPnPUserProfileProperty {
            [PSCustomObject]@{
                PersonalSiteUrl = 'https://contoso-my.sharepoint.com/personal/user_contoso_onmicrosoft_com'
            }
        }

        Mock Resolve-SboPersonalSiteUrl {
            param(
                [Alias('Profile')]
                [object] $InputObject
            )

            if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
                $entry = @($InputObject | Where-Object {
                    $_ -and $_.PSObject.Properties['Key'] -and $_.Key -eq 'PersonalUrl'
                } | Select-Object -First 1)

                if ($entry.Count -gt 0) {
                    return [string] $entry[0].Value
                }

                $hostEntry = @($InputObject | Where-Object {
                    $_ -and $_.PSObject.Properties['Key'] -and $_.Key -eq 'PersonalSiteHostUrl'
                } | Select-Object -First 1)
                $spaceEntry = @($InputObject | Where-Object {
                    $_ -and $_.PSObject.Properties['Key'] -and $_.Key -eq 'PersonalSpace'
                } | Select-Object -First 1)

                if ($hostEntry.Count -gt 0 -and $spaceEntry.Count -gt 0) {
                    $hostAuthority = ([Uri] ([string] $hostEntry[0].Value)).GetLeftPart([System.UriPartial]::Authority)
                    $space = [string] $spaceEntry[0].Value
                    if (-not $space.StartsWith('/')) {
                        $space = '/' + $space
                    }
                    return ($hostAuthority + $space)
                }
            }

            if ($InputObject.PSObject.Properties.Name -contains 'PersonalSiteUrl') {
                return [string] $InputObject.PersonalSiteUrl
            }
            if ($InputObject.PSObject.Properties.Name -contains 'PersonalUrl') {
                return [string] $InputObject.PersonalUrl
            }

            return $null
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
            [PSCustomObject]@{ Url = 'https://contoso-admin.sharepoint.com/' },
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

    It 'uses existing PNPConnection and processes items successfully' -Tag 'PNPConnection' {
        # Arrange
        # Connection queue returns existing host then existing personal connection.
        $global:TestGetConnectionQueue = @(
            [PSCustomObject]@{ Url = 'https://contoso-admin.sharepoint.com/' },
            [PSCustomObject]@{ Url = $script:DefaultPersonalSiteUrl }
        )

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        @($result).Count | Should -Be 1
        @($result)[0].PermissionChangeStatus | Should -Be 'Applied'
        Assert-MockCalled Invoke-SboConnectPnPOnline -Times 0
        Assert-MockCalled Invoke-SboDisconnectPnPOnline -Times 0
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
    }

    It 'calls profile and identity-list wrapper functions in default library mode' {
        # Arrange
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $null = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Invoke-SboGetPnPUserProfileProperty -Times 1
        Assert-MockCalled Resolve-SboPersonalSiteUrl -Times 1
        Assert-MockCalled Invoke-SboGetPnPListByIdentity -Times 1
        Assert-MockCalled Invoke-SboGetPnPLists -Times 0
        Assert-MockCalled Invoke-SboGetPnPProperty -Times 0
    }

    It 'logs extracted profile properties before resolving personal site URL' {
        # Arrange
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        @($result).Count | Should -Be 1
        ($global:CapturedLogLines -join "`n") | Should -Match "Profile properties for '$($script:DefaultUser)': PersonalSiteHostUrl='<null>', PersonalSpace='<null>', PersonalUrl='<null>'"
    }

    It 'calls list discovery and property-load wrappers in auto-discover mode' {
        # Arrange
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            AutoDiscoverDefaultLibrary = $true
            LogFolder = $TestDrive
        }

        # Act
        $null = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Invoke-SboGetPnPLists -Times 1
        Assert-MockCalled Invoke-SboGetPnPProperty -Times 1
        Assert-MockCalled Invoke-SboGetPnPListByIdentity -Times 0
    }

    It 'uses folder wrapper for folder items and still applies permissions' {
        # Arrange
        $global:TestCsvRows = @(
            New-BoxRow -CollaboratorPermission 'Editor' -ItemType 'Folder' -Path 'Documents/FolderA' -ItemName 'FolderA'
        )

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        @($result).Count | Should -Be 1
        Assert-MockCalled Invoke-SboGetPnPFolderAsListItem -Times 1
        Assert-MockCalled Invoke-SboGetPnPFileAsListItem -Times 0
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
    }

    It 'marks item as failed and continues when file or folder lookup throws' {
        # Arrange
        $global:TestCsvRows = @(
            New-BoxRow -CollaboratorPermission 'Editor' -ItemType 'File' -Path 'Documents/FolderA' -ItemName 'Missing.docx' -CollaboratorLogin 'first@contoso.com'
            New-BoxRow -CollaboratorPermission 'Viewer' -ItemType 'File' -Path 'Documents/FolderA' -ItemName 'Exists.docx' -CollaboratorLogin 'second@contoso.com'
        )

        Mock Invoke-SboGetPnPFileAsListItem {
            param(
                [string] $Url,
                [object] $Connection
            )

            if ($Url -like '*Missing.docx') {
                throw 'File Not Found.'
            }

            $global:CapturedGetFileUrls += $Url
            [PSCustomObject]@{ Id = 101 }
        }

        $errors = @()
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
            ErrorAction = 'SilentlyContinue'
            ErrorVariable = 'errors'
        }

        # Act
        $result = @(& $script:ScriptUnderTest @invokeParams)

        # Assert
        @($errors | ForEach-Object { $_.ToString() } | Where-Object { $_ -match 'Unhandled error while processing user|does not have a PersonalSiteUrl' }) | Should -BeNullOrEmpty
        $result.Count | Should -Be 2
        ($result | Where-Object { $_.'Item Name' -eq 'Missing.docx' }).PermissionChangeStatus | Should -Be 'Failed'
        ($result | Where-Object { $_.'Item Name' -eq 'Missing.docx' }).PermissionChangeError | Should -Match '^Lookup failed:'
        ($result | Where-Object { $_.'Item Name' -eq 'Exists.docx' }).PermissionChangeStatus | Should -Be 'Applied'
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
    }

    It 'creates and disconnects both admin and personal connections when none exist' {
        # Arrange
        # First connection check returns null (host connection required),
        # second check for personal connection also returns null.
        $global:TestGetConnectionQueue = @($null, $null)

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $null = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Invoke-SboConnectPnPOnline -Times 2
        Assert-MockCalled Invoke-SboDisconnectPnPOnline -Times 2
    }

    It 'uses SharePointOnlineAdminUrl when creating the admin-scoped connection' {
        # Arrange
        $customAdminUrl = 'https://fabrikam-admin.sharepoint.com/'

        Mock Invoke-SboGetPnPConnection {
            $null
        }

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            SharePointOnlineAdminUrl = $customAdminUrl
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        @($result).Count | Should -Be 1
        @($result)[0].PermissionChangeStatus | Should -Be 'Applied'
    }

    It 'writes log lines through the logging helper during normal execution' {
        # Arrange
        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $null = & $script:ScriptUnderTest @invokeParams

        # Assert
        Assert-MockCalled Add-Content
        Assert-MockCalled Write-Verbose
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

    It 'resolves personal site URL from PersonalSiteHostUrl and PersonalSpace when PersonalUrl is absent' {
        # Arrange
        Mock Invoke-SboGetPnPUserProfileProperty {
            @(
                [PSCustomObject]@{ Key = 'PersonalSiteHostUrl'; Value = 'https://contoso-my.sharepoint.com:443/' },
                [PSCustomObject]@{ Key = 'PersonalSpace'; Value = '/personal/user_contoso_onmicrosoft_com/' }
            )
        }

        $invokeParams = @{
            InputFile = (Join-Path -Path $TestDrive -ChildPath 'input.csv')
            UserToProcess = $script:DefaultUser
            LogFolder = $TestDrive
        }

        # Act
        $result = & $script:ScriptUnderTest @invokeParams

        # Assert
        @($result).Count | Should -Be 1
        @($result)[0].ItemUrl.AbsoluteUri | Should -Match '^https://contoso-my\.sharepoint\.com/personal/'
        ($global:CapturedLogLines -join "`n") | Should -Match "Profile properties for '$($script:DefaultUser)': PersonalSiteHostUrl='https://contoso-my\.sharepoint\.com:443/', PersonalSpace='/personal/user_contoso_onmicrosoft_com/', PersonalUrl='<null>'"
        Assert-MockCalled Resolve-SboPersonalSiteUrl -Times 1
        Assert-MockCalled Invoke-SboSetPnPListItemPermission -Times 1
    }

    It 'writes an error when personal site URL cannot be resolved from profile properties' {
        # Arrange
        Mock Invoke-SboGetPnPUserProfileProperty {
            [PSCustomObject]@{}
        }
        Mock Resolve-SboPersonalSiteUrl {
            param(
                [Alias('Profile')]
                [object] $InputObject
            )

            return $null
        }

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
        @($result).Count | Should -Be 0
        $errors | Should -Not -BeNullOrEmpty
        ($errors[0].ToString()) | Should -Match 'does not have a PersonalSiteUrl'
    }
}
