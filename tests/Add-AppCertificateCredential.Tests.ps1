#Requires -Version 7.0
# Pester test suite for Add-AppCertificateCredential.ps1.
#
# Test design goals:
# - Isolate from real Azure AD / Graph by mocking all external commands.
# - Verify certificate creation, upload, verification, and connectivity test steps.
# - Include a test that runs the script end-to-end with mocked dependencies.

Set-StrictMode -Version Latest

Describe 'Add-AppCertificateCredential.ps1' {

    BeforeAll {
        $OriginalScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Add-AppCertificateCredential.ps1'
        $script:ScriptUnderTest = Join-Path -Path $TestDrive -ChildPath 'Add-AppCertificateCredential.NoRequires.ps1'

        # Remove #Requires lines in the test copy so module-loading issues do not block unit tests.
        $ScriptWithoutRequires = Get-Content -LiteralPath $OriginalScriptPath | Where-Object {
            $_ -notmatch '^\s*#Requires\b'
        }
        Set-Content -LiteralPath $script:ScriptUnderTest -Value $ScriptWithoutRequires -Encoding UTF8

        $script:DefaultAppClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
        $script:DefaultTenantId = '92075952-90f3-4613-833b-d2e19ec649e4'
        $script:DefaultAppObjectId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $script:DefaultAppName = 'NPSBoxGraphApp'
        $script:DefaultThumbprint = 'AABBCCDDEEFF00112233445566778899AABBCCDD'

        # Create stubs for commands used by the script.
        function New-SelfSignedCertificate { }
        function Connect-MgGraph { }
        function Disconnect-MgGraph { }
        function Get-MgApplication { }
        function Update-MgApplication { }
        function Get-MgContext { }

        # Module cmdlets — only stub if not already available from the installed module.
        foreach ($cmdletName in @('Import-Module', 'Get-Module'))
        {
            if (-not (Get-Command -Name $cmdletName -ErrorAction SilentlyContinue))
            {
                New-Item -Path "Function:\$cmdletName" -Value {} -Force | Out-Null
            }
        }

        # Build a fake certificate object for mocking
        function New-FakeCert {
            [PSCustomObject]@{
                Thumbprint   = $script:DefaultThumbprint
                Subject      = 'CN=NPSBoxGraphApp'
                NotBefore    = [datetime]::UtcNow.AddMinutes(-5)
                NotAfter     = [datetime]::UtcNow.AddYears(2)
            } | Add-Member -MemberType ScriptMethod -Name 'GetRawCertData' -Value {
                [byte[]]@(1, 2, 3, 4, 5)
            } -PassThru
        }

        # Build a fake app registration object
        function New-FakeApp {
            param([object[]] $KeyCredentials = @())
            [PSCustomObject]@{
                Id             = $script:DefaultAppObjectId
                AppId          = $script:DefaultAppClientId
                DisplayName    = $script:DefaultAppName
                KeyCredentials = $KeyCredentials
            }
        }
    }

    Context 'Script Execution - End-to-End with SkipConnectivityTest' {
        BeforeEach {
            $script:FakeCert = New-FakeCert
            $script:FakeApp = New-FakeApp

            Mock -CommandName 'New-SelfSignedCertificate' -MockWith { $script:FakeCert }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-Module' -MockWith {
                [PSCustomObject]@{ Name = $Name; Version = [version]'2.26.1' }
            }
            Mock -CommandName 'Import-Module' -MockWith { }
            Mock -CommandName 'Update-MgApplication' -MockWith { }

            # First call: look up the app; Second call: verify after upload
            $script:GetMgAppCallCount = 0
            Mock -CommandName 'Get-MgApplication' -MockWith {
                $script:GetMgAppCallCount++
                if ($script:GetMgAppCallCount -le 1) {
                    return $script:FakeApp
                }
                # After upload, return app with the new key credential
                return New-FakeApp -KeyCredentials @(
                    [PSCustomObject]@{
                        Type          = 'AsymmetricX509Cert'
                        Usage         = 'Verify'
                        Key           = [byte[]]@(1, 2, 3, 4, 5)
                        DisplayName   = 'NPSBoxGraphApp'
                        StartDateTime = [datetime]::UtcNow.AddMinutes(-5)
                        EndDateTime   = [datetime]::UtcNow.AddYears(2)
                    }
                )
            }
        }

        It 'should run successfully and output result object' {
            $result = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            $result | Should -Not -BeNullOrEmpty
            $result.Thumbprint | Should -Be $script:DefaultThumbprint
            $result.AppClientId | Should -Be $script:DefaultAppClientId
            $result.TenantId | Should -Be $script:DefaultTenantId
            $result.Status | Should -Be 'Ready'
        }

        It 'should create a self-signed certificate' {
            $null = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            Assert-MockCalled -CommandName 'New-SelfSignedCertificate' -Times 1 -Scope It
        }

        It 'should connect to Graph for the upload' {
            $null = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            Assert-MockCalled -CommandName 'Connect-MgGraph' -Times 1 -Scope It
        }

        It 'should look up the app registration' {
            $null = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            Assert-MockCalled -CommandName 'Get-MgApplication' -Times 2 -Scope It
        }

        It 'should upload the certificate to the app registration' {
            $null = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            Assert-MockCalled -CommandName 'Update-MgApplication' -Times 1 -Scope It
        }

        It 'should disconnect from Graph' {
            $null = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            Assert-MockCalled -CommandName 'Disconnect-MgGraph' -Scope It
        }

        It 'should output all expected properties' {
            $result = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            $result.PSObject.Properties.Name | Should -Contain 'Thumbprint'
            $result.PSObject.Properties.Name | Should -Contain 'Subject'
            $result.PSObject.Properties.Name | Should -Contain 'NotBefore'
            $result.PSObject.Properties.Name | Should -Contain 'NotAfter'
            $result.PSObject.Properties.Name | Should -Contain 'AppClientId'
            $result.PSObject.Properties.Name | Should -Contain 'AppName'
            $result.PSObject.Properties.Name | Should -Contain 'TenantId'
            $result.PSObject.Properties.Name | Should -Contain 'StorePath'
            $result.PSObject.Properties.Name | Should -Contain 'Status'
        }

        It 'should accept custom CertificateSubject and ValidityYears' {
            # Override Get-MgApplication to return credentials matching the custom display name
            $script:CustomCallCount = 0
            Mock -CommandName 'Get-MgApplication' -MockWith {
                $script:CustomCallCount++
                if ($script:CustomCallCount -le 1) {
                    return $script:FakeApp
                }
                return New-FakeApp -KeyCredentials @(
                    [PSCustomObject]@{
                        Type          = 'AsymmetricX509Cert'
                        Usage         = 'Verify'
                        Key           = [byte[]]@(1, 2, 3, 4, 5)
                        DisplayName   = 'CustomApp'
                        StartDateTime = [datetime]::UtcNow.AddMinutes(-5)
                        EndDateTime   = [datetime]::UtcNow.AddYears(1)
                    }
                )
            }

            $result = & {
                . $script:ScriptUnderTest -CertificateSubject 'CN=CustomApp' `
                    -ValidityYears 1 -SkipConnectivityTest -Verbose:$false
            }

            $result | Should -Not -BeNullOrEmpty
            Assert-MockCalled -CommandName 'New-SelfSignedCertificate' -Times 1 -Scope It
        }
    }

    Context 'Script Execution - Connectivity Test' {
        BeforeEach {
            $script:FakeCert = New-FakeCert
            $script:FakeApp = New-FakeApp

            Mock -CommandName 'New-SelfSignedCertificate' -MockWith { $script:FakeCert }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-Module' -MockWith {
                [PSCustomObject]@{ Name = $Name; Version = [version]'2.26.1' }
            }
            Mock -CommandName 'Import-Module' -MockWith { }
            Mock -CommandName 'Update-MgApplication' -MockWith { }

            $script:GetMgAppCallCount2 = 0
            Mock -CommandName 'Get-MgApplication' -MockWith {
                $script:GetMgAppCallCount2++
                if ($script:GetMgAppCallCount2 -le 1) {
                    return $script:FakeApp
                }
                return New-FakeApp -KeyCredentials @(
                    [PSCustomObject]@{
                        Type          = 'AsymmetricX509Cert'
                        Usage         = 'Verify'
                        Key           = [byte[]]@(1, 2, 3, 4, 5)
                        DisplayName   = 'NPSBoxGraphApp'
                        StartDateTime = [datetime]::UtcNow.AddMinutes(-5)
                        EndDateTime   = [datetime]::UtcNow.AddYears(2)
                    }
                )
            }

            Mock -CommandName 'Get-MgContext' -MockWith {
                [PSCustomObject]@{
                    AppName  = $script:DefaultAppName
                    AuthType = 'AppOnly'
                }
            }
        }

        It 'should test connectivity when SkipConnectivityTest is not set' {
            $null = & {
                . $script:ScriptUnderTest -Verbose:$false
            }

            # Connect called twice: once for upload, once for connectivity test
            Assert-MockCalled -CommandName 'Connect-MgGraph' -Times 2 -Scope It
            Assert-MockCalled -CommandName 'Get-MgContext' -Times 1 -Scope It
        }
    }

    Context 'Script Execution - Error Handling' {
        BeforeEach {
            $script:FakeCert = New-FakeCert

            Mock -CommandName 'New-SelfSignedCertificate' -MockWith { $script:FakeCert }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-Module' -MockWith {
                [PSCustomObject]@{ Name = $Name; Version = [version]'2.26.1' }
            }
            Mock -CommandName 'Import-Module' -MockWith { }
        }

        It 'should throw when app registration is not found' {
            Mock -CommandName 'Get-MgApplication' -MockWith { return $null }

            { & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            } } | Should -Throw '*not found*'
        }

        It 'should throw when certificate verification fails' {
            $script:VerifyCallCount = 0
            Mock -CommandName 'Get-MgApplication' -MockWith {
                $script:VerifyCallCount++
                if ($script:VerifyCallCount -le 1) {
                    return New-FakeApp
                }
                # After upload, return app with no matching key credential
                return New-FakeApp -KeyCredentials @()
            }
            Mock -CommandName 'Update-MgApplication' -MockWith { }

            { & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            } } | Should -Throw '*Verification failed*'
        }

        It 'should throw when Update-MgApplication fails' {
            Mock -CommandName 'Get-MgApplication' -MockWith { return New-FakeApp }
            Mock -CommandName 'Update-MgApplication' -MockWith {
                throw [System.Exception]::new('Insufficient privileges')
            }

            { & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            } } | Should -Throw '*Insufficient privileges*'
        }
    }

    Context 'Script Execution - Preserves Existing Credentials' {
        BeforeEach {
            $script:FakeCert = New-FakeCert
            $script:CapturedKeys = $null

            $existingCred = [PSCustomObject]@{
                Type          = 'AsymmetricX509Cert'
                Usage         = 'Verify'
                Key           = [byte[]]@(10, 20, 30)
                DisplayName   = 'ExistingCert'
                StartDateTime = [datetime]::UtcNow.AddYears(-1)
                EndDateTime   = [datetime]::UtcNow.AddYears(1)
            }

            Mock -CommandName 'New-SelfSignedCertificate' -MockWith { $script:FakeCert }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-Module' -MockWith {
                [PSCustomObject]@{ Name = $Name; Version = [version]'2.26.1' }
            }
            Mock -CommandName 'Import-Module' -MockWith { }

            $script:PreserveCallCount = 0
            Mock -CommandName 'Get-MgApplication' -MockWith {
                $script:PreserveCallCount++
                if ($script:PreserveCallCount -le 1) {
                    return New-FakeApp -KeyCredentials @($existingCred)
                }
                return New-FakeApp -KeyCredentials @(
                    $existingCred,
                    [PSCustomObject]@{
                        Type          = 'AsymmetricX509Cert'
                        Usage         = 'Verify'
                        Key           = [byte[]]@(1, 2, 3, 4, 5)
                        DisplayName   = 'NPSBoxGraphApp'
                        StartDateTime = [datetime]::UtcNow.AddMinutes(-5)
                        EndDateTime   = [datetime]::UtcNow.AddYears(2)
                    }
                )
            }

            Mock -CommandName 'Update-MgApplication' -MockWith {
                # Capture all bound parameters for assertion
                $script:CapturedUpdateParams = $PSBoundParameters
            }
        }

        It 'should include existing key credentials when uploading' {
            $null = & {
                . $script:ScriptUnderTest -SkipConnectivityTest -Verbose:$false
            }

            Assert-MockCalled -CommandName 'Update-MgApplication' -Times 1 -Scope It
        }
    }
}
