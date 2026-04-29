#Requires -Version 7.0
# Pester test suite for Test-AzureAppRegistration.ps1.
#
# Test design goals:
# - Keep tests isolated from Graph dependencies by mocking all external commands.
# - Verify output object structure, permission checking logic, and error handling.

Set-StrictMode -Version Latest

Describe 'Test-AzureAppRegistration.ps1' {

    BeforeAll {
        $OriginalScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Test-AzureAppRegistration.ps1'
        $script:ScriptUnderTest = Join-Path -Path $TestDrive -ChildPath 'Test-AzureAppRegistration.NoRequires.ps1'

        # Remove #Requires lines in the test copy so module-loading issues do not block unit tests.
        $ScriptWithoutRequires = Get-Content -LiteralPath $OriginalScriptPath | Where-Object {
            $_ -notmatch '^\s*#Requires\b'
        }
        Set-Content -LiteralPath $script:ScriptUnderTest -Value $ScriptWithoutRequires -Encoding UTF8

        $script:DefaultThumbprint = 'AABBCCDDEE1122334455AABBCCDDEE1122334455'
        $script:DefaultTenantId   = '92075952-90f3-4613-833b-d2e19ec649e4'
        $script:DefaultClientId   = '912696b9-1374-4110-893d-545fc17c3371'

        # Stubs for script-internal functions.
        function Assert-RequiredModules { }
        function Connect-Graph { }

        # Module cmdlet stubs — only create if not already available.
        foreach ($cmdletName in @(
            'Connect-MgGraph',
            'Disconnect-MgGraph',
            'Get-MgContext',
            'Get-MgServicePrincipal',
            'Get-MgServicePrincipalAppRoleAssignment'
        ))
        {
            if (-not (Get-Command -Name $cmdletName -ErrorAction SilentlyContinue))
            {
                New-Item -Path "Function:\$cmdletName" -Value {} -Force | Out-Null
            }
        }

        # Helper: build a Graph AppRole definition object.
        function New-AppRoleDef {
            param(
                [string] $Value,
                [string] $Id = [guid]::NewGuid().ToString()
            )
            [PSCustomObject]@{ Value = $Value; Id = $Id }
        }

        # Helper: build an app role assignment object.
        function New-AppRoleAssignment {
            param(
                [string] $AppRoleId,
                [datetime] $CreatedDateTime = (Get-Date)
            )
            [PSCustomObject]@{ AppRoleId = $AppRoleId; CreatedDateTime = $CreatedDateTime }
        }
    }

    Context 'All Permissions Granted' {
        BeforeEach {
            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }

            $script:FilesRoleId = [guid]::NewGuid().ToString()
            $script:UserRoleId  = [guid]::NewGuid().ToString()

            Mock -CommandName 'Get-MgServicePrincipal' -MockWith {
                param($Filter)
                if ($Filter -match 'appId') {
                    return [PSCustomObject]@{ Id = 'sp-app-id'; DisplayName = 'NPSBox App' }
                }
                if ($Filter -match 'Microsoft Graph') {
                    return [PSCustomObject]@{
                        Id       = 'sp-graph-id'
                        AppRoles = @(
                            (New-AppRoleDef -Value 'Files.ReadWrite.All' -Id $script:FilesRoleId),
                            (New-AppRoleDef -Value 'User.Read.All'      -Id $script:UserRoleId),
                            (New-AppRoleDef -Value 'Mail.Read'          -Id ([guid]::NewGuid().ToString()))
                        )
                    }
                }
            }

            Mock -CommandName 'Get-MgServicePrincipalAppRoleAssignment' -MockWith {
                @(
                    (New-AppRoleAssignment -AppRoleId $script:FilesRoleId),
                    (New-AppRoleAssignment -AppRoleId $script:UserRoleId)
                )
            }
        }

        It 'should output one result per required permission' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $results.Count | Should -Be 2
        }

        It 'should mark all permissions as granted' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $results | ForEach-Object { $_.IsGranted | Should -Be $true }
        }

        It 'should include all expected output properties' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $result = $results | Select-Object -First 1
            $result.PSObject.Properties.Name | Should -Contain 'Permission'
            $result.PSObject.Properties.Name | Should -Contain 'Type'
            $result.PSObject.Properties.Name | Should -Contain 'IsGranted'
            $result.PSObject.Properties.Name | Should -Contain 'RoleId'
            $result.PSObject.Properties.Name | Should -Contain 'GrantedOn'
            $result.PSObject.Properties.Name | Should -Contain 'AppId'
            $result.PSObject.Properties.Name | Should -Contain 'TenantId'
            $result.PSObject.Properties.Name | Should -Contain 'DisplayName'
        }

        It 'should set Type to Application for all results' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $results | ForEach-Object { $_.Type | Should -Be 'Application' }
        }

        It 'should populate AppId and TenantId from parameters' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $results | ForEach-Object {
                $_.AppId    | Should -Be $script:DefaultClientId
                $_.TenantId | Should -Be $script:DefaultTenantId
            }
        }

        It 'should populate DisplayName from the service principal' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $results | ForEach-Object { $_.DisplayName | Should -Be 'NPSBox App' }
        }
    }

    Context 'Missing Permissions' {
        BeforeEach {
            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }

            $script:FilesRoleId = [guid]::NewGuid().ToString()
            $script:UserRoleId  = [guid]::NewGuid().ToString()

            Mock -CommandName 'Get-MgServicePrincipal' -MockWith {
                param($Filter)
                if ($Filter -match 'appId') {
                    return [PSCustomObject]@{ Id = 'sp-app-id'; DisplayName = 'NPSBox App' }
                }
                if ($Filter -match 'Microsoft Graph') {
                    return [PSCustomObject]@{
                        Id       = 'sp-graph-id'
                        AppRoles = @(
                            (New-AppRoleDef -Value 'Files.ReadWrite.All' -Id $script:FilesRoleId),
                            (New-AppRoleDef -Value 'User.Read.All'      -Id $script:UserRoleId)
                        )
                    }
                }
            }

            # Only Files.ReadWrite.All is granted — User.Read.All is missing.
            Mock -CommandName 'Get-MgServicePrincipalAppRoleAssignment' -MockWith {
                @(
                    (New-AppRoleAssignment -AppRoleId $script:FilesRoleId)
                )
            }
        }

        It 'should mark Files.ReadWrite.All as granted' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $filesResult = $results | Where-Object { $_.Permission -eq 'Files.ReadWrite.All' }
            $filesResult.IsGranted | Should -Be $true
        }

        It 'should mark User.Read.All as not granted' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $userResult = $results | Where-Object { $_.Permission -eq 'User.Read.All' }
            $userResult.IsGranted | Should -Be $false
            $userResult.GrantedOn | Should -BeNullOrEmpty
        }
    }

    Context 'No Permissions Granted' {
        BeforeEach {
            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }

            Mock -CommandName 'Get-MgServicePrincipal' -MockWith {
                param($Filter)
                if ($Filter -match 'appId') {
                    return [PSCustomObject]@{ Id = 'sp-app-id'; DisplayName = 'NPSBox App' }
                }
                if ($Filter -match 'Microsoft Graph') {
                    return [PSCustomObject]@{
                        Id       = 'sp-graph-id'
                        AppRoles = @(
                            (New-AppRoleDef -Value 'Files.ReadWrite.All'),
                            (New-AppRoleDef -Value 'User.Read.All')
                        )
                    }
                }
            }

            Mock -CommandName 'Get-MgServicePrincipalAppRoleAssignment' -MockWith { @() }
        }

        It 'should mark all permissions as not granted' {
            $results = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            $results | ForEach-Object { $_.IsGranted | Should -Be $false }
        }
    }

    Context 'Error Handling' {
        BeforeEach {
            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
        }

        It 'should throw when service principal is not found' {
            Mock -CommandName 'Get-MgServicePrincipal' -MockWith { return $null }

            { & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } } | Should -Throw '*Service principal not found*'
        }

        It 'should disconnect from Graph after processing' {
            Mock -CommandName 'Get-MgServicePrincipal' -MockWith {
                param($Filter)
                if ($Filter -match 'appId') {
                    return [PSCustomObject]@{ Id = 'sp-app-id'; DisplayName = 'NPSBox App' }
                }
                if ($Filter -match 'Microsoft Graph') {
                    return [PSCustomObject]@{
                        Id       = 'sp-graph-id'
                        AppRoles = @((New-AppRoleDef -Value 'Files.ReadWrite.All'), (New-AppRoleDef -Value 'User.Read.All'))
                    }
                }
            }
            Mock -CommandName 'Get-MgServicePrincipalAppRoleAssignment' -MockWith { @() }

            $null = & {
                . $script:ScriptUnderTest -CertificateThumbprint $script:DefaultThumbprint -Verbose:$false
            } 6>&1

            Assert-MockCalled -CommandName 'Disconnect-MgGraph' -Scope It
        }
    }
}
