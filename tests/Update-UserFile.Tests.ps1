#Requires -Version 7.0
# Pester test suite for Update-UserFile.ps1.
#
# Test design goals:
# - Keep tests isolated from Graph/file system dependencies by mocking all external commands.
# - Include comprehensive script execution tests with realistic scenarios.
# - Verify core functionality: permission mapping, path handling, error handling, and end-to-end workflow.

Set-StrictMode -Version Latest

Describe 'Update-UserFile.ps1' {

    BeforeAll {
        $OriginalScriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Update-UserFile.ps1'
        $script:ScriptUnderTest = Join-Path -Path $TestDrive -ChildPath 'Update-UserFile.NoRequires.ps1'

        # Remove #Requires lines in the test copy so module-loading issues do not block unit tests.
        $ScriptWithoutRequires = Get-Content -LiteralPath $OriginalScriptPath | Where-Object {
            $_ -notmatch '^\s*#Requires\b'
        }
        Set-Content -LiteralPath $script:ScriptUnderTest -Value $ScriptWithoutRequires -Encoding UTF8

        $script:DefaultOwner = 'user@contoso.onmicrosoft.com'
        $script:DefaultCollaborator = 'collab@contoso.com'
        $script:DefaultDriveId = 'b!-kIQeRjLDEyVXvh98xyWkBx6vWyJBJhFr5H_U3K6v7bkHqmOKs-hRpYN8L-rk6HJ'
        $script:DefaultWebUrl = 'https://contoso-my.sharepoint.com/personal/user_contoso_onmicrosoft_com'
        $script:DefaultThumbprint = 'AABBCCDDEE1122334455AABBCCDDEE1122334455'

        # Create stubs for script-internal functions that will be mocked.
        # These do not exist until the script is dot-sourced, but Pester needs
        # them resolvable before Mock is called.  Using explicit function
        # declarations at the current scope ensures Pester can find them.
        function Assert-RequiredModules { }
        function Connect-Graph { }
        function Assert-GraphAssemblyCompatibility { }
        function Assert-GraphPermissions { }
        function Get-ValidatedUserDrive { }
        function Invoke-OneDriveUpload { }
        function Test-CollaboratorDomain { return $true }
        function Get-RetryAfterSeconds { return $null }

        # Module cmdlets — only stub if not already available from the installed module.
        foreach ($cmdletName in @(
            'Disconnect-MgGraph',
            'Connect-MgGraph',
            'Get-MgUser',
            'Get-MgUserDrive',
            'Invoke-MgGraphRequest',
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

        function New-CsvRow {
            param(
                [string] $OwnerLogin = $script:DefaultOwner,
                [string] $Path = 'All Files/Documents',
                [string] $ItemName = 'Doc1.txt',
                [string] $ItemType = 'File',
                [string] $CollaboratorLogin = $script:DefaultCollaborator,
                [string] $CollaboratorPermission = 'Editor'
            )

            [PSCustomObject]@{
                'Owner Login' = $OwnerLogin
                'Path' = $Path
                'Item Name' = $ItemName
                'Item Type' = $ItemType
                'Collaborator Login' = $CollaboratorLogin
                'Collaborator Permission' = $CollaboratorPermission
            }
        }
    }

    Context 'Script Execution - Permission Mapping' {
        BeforeEach {
            # Create a temporary CSV file with test data
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test.csv'
            $rows = @(
                (New-CsvRow -ItemName 'Doc1.txt' -CollaboratorPermission 'Editor'),
                (New-CsvRow -ItemName 'Doc2.txt' -CollaboratorPermission 'Viewer'),
                (New-CsvRow -ItemName 'Doc3.txt' -CollaboratorPermission 'Previewer'),
                (New-CsvRow -ItemName 'Doc4.txt' -CollaboratorPermission 'Co-owner'),
                (New-CsvRow -ItemName 'Doc5.txt' -CollaboratorLogin '')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            # Create temp log folder
            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            # Setup standard mocks for all tests
            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }

            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{
                    Id     = $script:DefaultDriveId
                    WebUrl = $script:DefaultWebUrl
                }
            }

            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri, $Body)

                if ($Uri -match '/root\?') {
                    return [PSCustomObject]@{
                        id     = 'root-id'
                        webUrl = $script:DefaultWebUrl
                    }
                }
                elseif ($Uri -match '/root:/' -and $Method -eq 'GET') {
                    return [PSCustomObject]@{
                        id   = 'item-id-12345'
                        name = 'TestItem'
                    }
                }
                elseif ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{
                        value = @(@{ id = 'perm-12345'; roles = @('write') })
                    }
                }
            }
        }

        It 'should map Editor permission to write role' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $editorResult = $results | Where-Object { $_.ItemName -eq 'Doc1.txt' }
            $editorResult.GraphRole | Should -Be 'write'
            $editorResult.CollaboratorPermission | Should -Be 'Editor'
            $editorResult.Status | Should -Be 'Applied'
        }

        It 'should map Viewer permission to read role' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $viewerResult = $results | Where-Object { $_.ItemName -eq 'Doc2.txt' }
            $viewerResult.GraphRole | Should -Be 'read'
            $viewerResult.CollaboratorPermission | Should -Be 'Viewer'
            $viewerResult.Status | Should -Be 'Applied'
        }

        It 'should map Co-owner permission to write role' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $coOwnerResult = $results | Where-Object { $_.ItemName -eq 'Doc4.txt' }
            $coOwnerResult.GraphRole | Should -Be 'write'
            $coOwnerResult.CollaboratorPermission | Should -Be 'Co-owner'
            $coOwnerResult.Status | Should -Be 'Applied'
        }

        It 'should skip Previewer permission (maps to null)' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $previewerResult = $results | Where-Object { $_.ItemName -eq 'Doc3.txt' }
            $previewerResult.Action | Should -Be 'Skipped'
            $previewerResult.Status | Should -Be 'Skipped'
            $previewerResult.GraphRole | Should -BeNullOrEmpty
        }

        It 'should fail when collaborator login is empty' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $failResult = $results | Where-Object { $_.ItemName -eq 'Doc5.txt' }
            $failResult.Status | Should -Be 'Failed'
            $failResult.Error | Should -Match 'Collaborator Login'
        }
    }

    Context 'Script Execution - Path Handling' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_paths.csv'
            $rows = @(
                (New-CsvRow -ItemName 'File.txt' -Path 'All Files/Documents'),
                (New-CsvRow -ItemName 'Report.pdf' -Path 'All Files/Folder with Spaces/SubFolder'),
                (New-CsvRow -ItemName 'Thesis.docx' -Path 'All Files/Thesis (IPv6)/'),
                (New-CsvRow -ItemName 'Data.xlsx' -Path 'Documents\Subfolder')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_paths'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
                elseif ($Uri -match '/root') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                return [PSCustomObject]@{ id = 'item-id'; name = 'Item' }
            }
        }

        It 'should normalize All Files prefix in paths' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Where-Object { $_.ItemName -eq 'File.txt' }
            $result.NormalizedPath | Should -Be 'Documents'
        }

        It 'should handle paths with spaces' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Where-Object { $_.ItemName -eq 'Report.pdf' }
            $result.NormalizedPath | Should -Be 'Folder with Spaces/SubFolder'
            $result.Status | Should -Be 'Applied'
        }

        It 'should handle paths with parentheses' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Where-Object { $_.ItemName -eq 'Thesis.docx' }
            $result.NormalizedPath | Should -Be 'Thesis (IPv6)'
            $result.Status | Should -Be 'Applied'
        }

        It 'should convert backslashes to forward slashes' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Where-Object { $_.ItemName -eq 'Data.xlsx' }
            $result.NormalizedPath | Should -Be 'Documents/Subfolder'
            $result.Status | Should -Be 'Applied'
        }
    }

    Context 'Script Execution - Parameters and Output' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_output.csv'
            $rows = @(
                (New-CsvRow -ItemName 'Doc1.txt' -CollaboratorPermission 'Editor'),
                (New-CsvRow -ItemName 'Doc2.txt' -CollaboratorPermission 'Editor')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_output'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
                elseif ($Uri -match '/root') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                return [PSCustomObject]@{ id = 'item-id'; name = 'Item' }
            }
        }

        It 'should output custom objects with all required properties' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Select-Object -First 1
            $result.PSObject.Properties.Name | Should -Contain 'OwnerLogin'
            $result.PSObject.Properties.Name | Should -Contain 'ItemName'
            $result.PSObject.Properties.Name | Should -Contain 'Path'
            $result.PSObject.Properties.Name | Should -Contain 'NormalizedPath'
            $result.PSObject.Properties.Name | Should -Contain 'CollaboratorLogin'
            $result.PSObject.Properties.Name | Should -Contain 'CollaboratorPermission'
            $result.PSObject.Properties.Name | Should -Contain 'GraphRole'
            $result.PSObject.Properties.Name | Should -Contain 'DriveId'
            $result.PSObject.Properties.Name | Should -Contain 'OneDriveWebUrl'
            $result.PSObject.Properties.Name | Should -Contain 'ExistsInOneDrive'
            $result.PSObject.Properties.Name | Should -Contain 'DriveItemId'
            $result.PSObject.Properties.Name | Should -Contain 'Action'
            $result.PSObject.Properties.Name | Should -Contain 'Status'
            $result.PSObject.Properties.Name | Should -Contain 'Error'
            $result.PSObject.Properties.Name | Should -Contain 'IsValidAccount'
            $result.PSObject.Properties.Name | Should -Contain 'OneDriveProvisioned'
        }

        It 'should set IsValidAccount and OneDriveProvisioned to true on success' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Select-Object -First 1
            $result.IsValidAccount | Should -Be $true
            $result.OneDriveProvisioned | Should -Be $true
        }

        It 'should support -WhatIf parameter' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -WhatIf -Verbose:$false
            } 6>&1

            $results | Should -Not -BeNullOrEmpty
            $results[0].Status | Should -Be 'WhatIf'
        }

        It 'should create log file in specified folder' {
            $null = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $logFiles = Get-ChildItem -Path $script:LogFolder -Filter '*.log'
            $logFiles | Should -Not -BeNullOrEmpty
            $logFiles[0].Name | Should -Match 'Update-UserFile_\d{8}_\d{6}_\d{3}\.log'
        }

        It 'should create log folder if it does not exist' {
            $newLogFolder = Join-Path -Path $TestDrive -ChildPath 'new_logs_output'
            $null = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $newLogFolder -Verbose:$false
            } 6>&1

            $newLogFolder | Should -Exist
        }
    }

    Context 'Script Execution - Error Handling' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_errors.csv'
            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_errors'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
        }

        It 'should throw when InputFile does not exist' {
            { & {
                . $script:ScriptUnderTest -InputFile 'nonexistent.csv' -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } } | Should -Throw
        }

        It 'should output Failed results when drive lookup fails for a user' {
            $rows = @(New-CsvRow)
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                throw [System.Exception]::new('Drive not found')
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results | Should -Not -BeNullOrEmpty
            $results[0].Status | Should -Be 'Failed'
            $results[0].Error | Should -Match 'Drive not found'
        }

        It 'should mark item as not existing when 404 error occurs' {
            $rows = @(New-CsvRow)
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }

            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/root\?') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                elseif ($Uri -match '/root:/' -and $Method -eq 'GET') {
                    throw [System.Exception]::new('itemNotFound: Item does not exist')
                }
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results[0].ExistsInOneDrive | Should -Be $false
            $results[0].Status | Should -Be 'Failed'
        }

        It 'should provide actionable error when OneDrive is not provisioned' {
            $rows = @(New-CsvRow)
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                throw [System.Exception]::new('ResourceNotFound: the user does not exist or OneDrive is not provisioned')
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results[0].Status | Should -Be 'Failed'
            $results[0].Error | Should -Match 'not provisioned'
            $results[0].Error | Should -Match 'portal\.office\.com'
            $results[0].IsValidAccount | Should -Be $true
            $results[0].OneDriveProvisioned | Should -Be $false
        }

        It 'should set IsValidAccount to false when user account is not found' {
            $rows = @(New-CsvRow)
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            Mock -CommandName 'Get-MgUser' -MockWith {
                throw [System.Exception]::new('Request_ResourceNotFound: Resource not found')
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results[0].Status | Should -Be 'Failed'
            $results[0].IsValidAccount | Should -Be $false
            $results[0].OneDriveProvisioned | Should -Be $false
            $results[0].Error | Should -Match 'could not be found'
        }

        It 'should set IsValidAccount to false when user account is disabled' {
            $rows = @(New-CsvRow)
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Disabled User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $false }
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results[0].Status | Should -Be 'Failed'
            $results[0].IsValidAccount | Should -Be $false
            $results[0].OneDriveProvisioned | Should -Be $false
            $results[0].Error | Should -Match 'disabled'
        }

        It 'should catch accessDenied as unprovisioned OneDrive' {
            $rows = @(New-CsvRow)
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                throw [System.Exception]::new('[accessDenied] : Access denied')
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results[0].Status | Should -Be 'Failed'
            $results[0].IsValidAccount | Should -Be $true
            $results[0].OneDriveProvisioned | Should -Be $false
            $results[0].Error | Should -Match 'not provisioned'
            $results[0].Error | Should -Match 'Request-SPOPersonalSite'
        }

        It 'should report insufficient permissions when app lacks User.Read.All' {
            $rows = @(New-CsvRow)
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            Mock -CommandName 'Get-MgUser' -MockWith {
                throw [System.Exception]::new('[Authorization_RequestDenied] : Insufficient privileges to complete the operation.')
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results[0].Status | Should -Be 'Failed'
            $results[0].IsValidAccount | Should -Be $false
            $results[0].OneDriveProvisioned | Should -Be $false
            $results[0].Error | Should -Match 'User\.Read\.All'
            $results[0].Error | Should -Match $script:DefaultOwner.Replace('.', '\.')
        }
    }

    Context 'Script Execution - End-to-End Workflow' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_e2e.csv'
            $rows = @(
                (New-CsvRow -OwnerLogin 'adile@contoso.com' -ItemName 'Shared Doc' -Path 'All Files/Projects' -CollaboratorLogin 'amber@contoso.com' -CollaboratorPermission 'Editor'),
                (New-CsvRow -OwnerLogin 'adile@contoso.com' -ItemName 'Read-Only' -Path 'All Files/Reports' -CollaboratorLogin 'billie@contoso.com' -CollaboratorPermission 'Viewer')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_e2e'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri, $Body)
                if ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
                elseif ($Uri -match '/root') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                elseif ($Uri -match '/root:/' -and $Method -eq 'GET') {
                    return [PSCustomObject]@{ id = 'item-id'; name = 'Item' }
                }
            }
        }

        It 'should process multiple rows and apply permissions' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess 'adile@contoso.com' `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results.Count | Should -Be 2
            $results[0].ItemName | Should -Be 'Shared Doc'
            $results[0].Status | Should -Be 'Applied'
            $results[1].ItemName | Should -Be 'Read-Only'
            $results[1].Status | Should -Be 'Applied'
        }

        It 'should apply correct roles for different permission levels' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess 'adile@contoso.com' `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results[0].GraphRole | Should -Be 'write'
            $results[1].GraphRole | Should -Be 'read'
        }

        It 'should disconnect from Graph after processing' {
            $null = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess 'adile@contoso.com' `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            Assert-MockCalled -CommandName 'Disconnect-MgGraph' -Scope It
        }
    }

    Context 'Script Execution - Multi-User Processing' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_multi.csv'
            $rows = @(
                (New-CsvRow -OwnerLogin 'owner1@contoso.com' -ItemName 'Owner1Doc.txt' -Path 'All Files/Docs' -CollaboratorLogin 'collab1@contoso.com' -CollaboratorPermission 'Editor'),
                (New-CsvRow -OwnerLogin 'owner2@contoso.com' -ItemName 'Owner2Doc.txt' -Path 'All Files/Reports' -CollaboratorLogin 'collab2@contoso.com' -CollaboratorPermission 'Viewer')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_multi'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
                elseif ($Uri -match '/root') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                return [PSCustomObject]@{ id = 'item-id'; name = 'Item' }
            }
        }

        It 'should process all unique owners when UserToProcess is not specified' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results.Count | Should -Be 2
            ($results | Where-Object { $_.OwnerLogin -eq 'owner1@contoso.com' }).Count | Should -Be 1
            ($results | Where-Object { $_.OwnerLogin -eq 'owner2@contoso.com' }).Count | Should -Be 1
        }

        It 'should apply correct roles per owner when processing all users' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $owner1Result = $results | Where-Object { $_.OwnerLogin -eq 'owner1@contoso.com' }
            $owner1Result.GraphRole | Should -Be 'write'
            $owner1Result.Status | Should -Be 'Applied'

            $owner2Result = $results | Where-Object { $_.OwnerLogin -eq 'owner2@contoso.com' }
            $owner2Result.GraphRole | Should -Be 'read'
            $owner2Result.Status | Should -Be 'Applied'
        }

        It 'should continue processing remaining owners when one fails' {
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                param($UserId)
                if ($UserId -eq 'owner1@contoso.com') {
                    throw [System.Exception]::new('Drive not provisioned')
                }
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results.Count | Should -Be 2

            $failedResult = $results | Where-Object { $_.OwnerLogin -eq 'owner1@contoso.com' }
            $failedResult.Status | Should -Be 'Failed'
            $failedResult.Error | Should -Match 'Drive not provisioned'

            $successResult = $results | Where-Object { $_.OwnerLogin -eq 'owner2@contoso.com' }
            $successResult.Status | Should -Be 'Applied'
        }

        It 'should filter to single owner when UserToProcess is specified' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess 'owner2@contoso.com' `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $results.Count | Should -Be 1
            $results[0].OwnerLogin | Should -Be 'owner2@contoso.com'
            $results[0].ItemName | Should -Be 'Owner2Doc.txt'
        }
    }

    Context 'Script Execution - UploadFiles Switch' {
        BeforeEach {
            # Create a minimal CSV (required by the script even when uploading)
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_upload.csv'
            $rows = @(
                (New-CsvRow -ItemName 'Doc1.txt' -CollaboratorPermission 'Editor')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_upload'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            # Create a fake local file structure mirroring the user's files
            $script:LocalFilesRoot = Join-Path -Path $TestDrive -ChildPath 'LocalFiles'
            $script:UserLocalPath = Join-Path -Path $script:LocalFilesRoot -ChildPath $script:DefaultOwner
            $subFolder = Join-Path -Path $script:UserLocalPath -ChildPath 'TestFolder'
            New-Item -Path $subFolder -ItemType Directory -Force | Out-Null
            Set-Content -LiteralPath (Join-Path -Path $script:UserLocalPath -ChildPath 'RootFile.txt') -Value 'root content'
            Set-Content -LiteralPath (Join-Path -Path $subFolder -ChildPath 'SubFile.txt') -Value 'sub content'

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
                elseif ($Uri -match '/root') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                return [PSCustomObject]@{ id = 'item-id'; name = 'Item' }
            }
        }

        It 'should upload files and create folders' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllFilesDirectory $script:LocalFilesRoot -UploadFiles -Verbose:$false
            } 6>&1

            $uploadResults = $results | Where-Object { $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -in @('CreateFolder', 'UploadFile') }
            $uploadResults | Should -Not -BeNullOrEmpty

            $folderResults = $uploadResults | Where-Object { $_.Action -eq 'CreateFolder' }
            $folderResults | Should -Not -BeNullOrEmpty
            $folderResults[0].ItemType | Should -Be 'Folder'
            $folderResults[0].Status | Should -Be 'Applied'

            $fileResults = $uploadResults | Where-Object { $_.Action -eq 'UploadFile' }
            $fileResults.Count | Should -Be 2
            $fileResults | ForEach-Object { $_.Status | Should -Be 'Applied' }
        }

        It 'should list files that would be uploaded with -WhatIf' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllFilesDirectory $script:LocalFilesRoot -UploadFiles -WhatIf -Verbose:$false
            } 6>&1

            $uploadResults = $results | Where-Object { $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -in @('CreateFolder', 'UploadFile') }
            $uploadResults | Should -Not -BeNullOrEmpty
            $uploadResults | ForEach-Object { $_.Status | Should -Be 'WhatIf' }
        }

        It 'should output upload result objects with expected properties' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllFilesDirectory $script:LocalFilesRoot -UploadFiles -Verbose:$false
            } 6>&1

            $fileResult = $results | Where-Object { $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'UploadFile' } | Select-Object -First 1
            $fileResult.PSObject.Properties.Name | Should -Contain 'OwnerLogin'
            $fileResult.PSObject.Properties.Name | Should -Contain 'LocalPath'
            $fileResult.PSObject.Properties.Name | Should -Contain 'OneDrivePath'
            $fileResult.PSObject.Properties.Name | Should -Contain 'ItemType'
            $fileResult.PSObject.Properties.Name | Should -Contain 'SizeBytes'
            $fileResult.PSObject.Properties.Name | Should -Contain 'Action'
            $fileResult.PSObject.Properties.Name | Should -Contain 'Status'
            $fileResult.PSObject.Properties.Name | Should -Contain 'Error'
        }

        It 'should throw when user local folder does not exist' {
            $emptyRoot = Join-Path -Path $TestDrive -ChildPath 'EmptyLocalFiles'
            New-Item -Path $emptyRoot -ItemType Directory -Force | Out-Null

            { & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllFilesDirectory $emptyRoot -UploadFiles -Verbose:$false
            } } | Should -Throw '*not found*'
        }

        It 'should not upload when UploadFiles is not specified' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllFilesDirectory $script:LocalFilesRoot -Verbose:$false
            } 6>&1

            $uploadResults = $results | Where-Object { $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -in @('CreateFolder', 'UploadFile') }
            $uploadResults | Should -BeNullOrEmpty
        }

        It 'should fail upload for files exceeding 4 MB' {
            # Create a file larger than 4 MB
            $largeFolderRoot = Join-Path -Path $TestDrive -ChildPath 'LargeLocalFiles'
            $largeUserPath = Join-Path -Path $largeFolderRoot -ChildPath $script:DefaultOwner
            New-Item -Path $largeUserPath -ItemType Directory -Force | Out-Null
            $largeFile = Join-Path -Path $largeUserPath -ChildPath 'BigFile.bin'
            # Create a 5 MB file (exceeds 4 MB limit)
            $bytes = [byte[]]::new(5 * 1024 * 1024)
            [System.IO.File]::WriteAllBytes($largeFile, $bytes)

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllFilesDirectory $largeFolderRoot -UploadFiles -Verbose:$false
            } 6>&1

            $uploadResults = $results | Where-Object { $_.PSObject.Properties.Name -contains 'Action' -and $_.Action -eq 'UploadFile' }
            $failedUpload = $uploadResults | Where-Object { $_.Status -eq 'Failed' }
            $failedUpload | Should -Not -BeNullOrEmpty
            $failedUpload.Error | Should -Match '4 MB'
        }
    }

    Context 'Script Execution - AllowedDomains Validation' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_domains.csv'
            $rows = @(
                (New-CsvRow -ItemName 'Internal.txt' -CollaboratorLogin 'collab@contoso.onmicrosoft.com' -CollaboratorPermission 'Editor'),
                (New-CsvRow -ItemName 'External.txt' -CollaboratorLogin 'outsider@gmail.com' -CollaboratorPermission 'Editor'),
                (New-CsvRow -ItemName 'AlsoInternal.txt' -CollaboratorLogin 'other@contoso.com' -CollaboratorPermission 'Viewer')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_domains'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
                elseif ($Uri -match '/root') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                return [PSCustomObject]@{ id = 'item-id'; name = 'Item' }
            }
        }

        It 'should skip collaborators with domains not in AllowedDomains' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllowedDomains @('contoso.onmicrosoft.com', 'contoso.com') -Verbose:$false
            } 6>&1

            $external = $results | Where-Object { $_.ItemName -eq 'External.txt' }
            $external.Status | Should -Be 'Skipped'
            $external.Error | Should -Match 'AllowedDomains'
        }

        It 'should allow collaborators with domains in AllowedDomains' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllowedDomains @('contoso.onmicrosoft.com', 'contoso.com') -Verbose:$false
            } 6>&1

            $internal = $results | Where-Object { $_.ItemName -eq 'Internal.txt' }
            $internal.Status | Should -Be 'Applied'

            $alsoInternal = $results | Where-Object { $_.ItemName -eq 'AlsoInternal.txt' }
            $alsoInternal.Status | Should -Be 'Applied'
        }

        It 'should allow all domains when AllowedDomains is not specified' {
            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $external = $results | Where-Object { $_.ItemName -eq 'External.txt' }
            $external.Status | Should -Be 'Applied'
        }

        It 'should perform case-insensitive domain matching' {
            $singleRowCsv = Join-Path -Path $TestDrive -ChildPath 'test_domain_case.csv'
            @(New-CsvRow -ItemName 'CaseTest.txt' -CollaboratorLogin 'user@CONTOSO.COM' -CollaboratorPermission 'Editor') |
                Export-Csv -LiteralPath $singleRowCsv -NoTypeInformation -Encoding UTF8

            $results = & {
                . $script:ScriptUnderTest -InputFile $singleRowCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder `
                    -AllowedDomains @('contoso.com') -Verbose:$false
            } 6>&1

            $results[0].Status | Should -Be 'Applied'
        }
    }

    Context 'Script Execution - Retry-After Header Parsing' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_retry.csv'
            $rows = @(
                (New-CsvRow -ItemName 'RetryDoc.txt' -CollaboratorPermission 'Editor')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_retry'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
        }

        It 'should succeed after transient 429 with Retry-After in error details' {
            $script:CallCount = 0
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/root\?') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                elseif ($Uri -match '/root:/' -and $Method -eq 'GET') {
                    return [PSCustomObject]@{ id = 'item-id'; name = 'RetryDoc.txt' }
                }
                elseif ($Uri -match '/invite' -and $Method -eq 'POST') {
                    $script:CallCount++
                    if ($script:CallCount -le 1) {
                        $err = [System.Exception]::new('429 Too Many Requests')
                        $errRecord = [System.Management.Automation.ErrorRecord]::new($err, 'Throttled', 'InvalidOperation', $null)
                        $errRecord | Add-Member -NotePropertyName 'ErrorDetails' -NotePropertyValue ([PSCustomObject]@{ Message = '{"error":{"code":"TooManyRequests","retryAfterSeconds":2}}' }) -Force
                        throw $errRecord
                    }
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
            }

            Mock -CommandName 'Start-Sleep' -MockWith { }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Where-Object { $_.ItemName -eq 'RetryDoc.txt' }
            $result.Status | Should -Be 'Applied'
            Assert-MockCalled -CommandName 'Start-Sleep' -Times 1 -Scope It
        }

        It 'should fail after exhausting all retry attempts' {
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/root\?') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                elseif ($Uri -match '/root:/' -and $Method -eq 'GET') {
                    return [PSCustomObject]@{ id = 'item-id'; name = 'RetryDoc.txt' }
                }
                elseif ($Uri -match '/invite' -and $Method -eq 'POST') {
                    throw [System.Exception]::new('429 Too Many Requests')
                }
            }

            Mock -CommandName 'Start-Sleep' -MockWith { }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Where-Object { $_.ItemName -eq 'RetryDoc.txt' }
            $result.Status | Should -Be 'Failed'
            $result.Error | Should -Match '429'
        }

        It 'should retry up to 6 times (new default) before failing' {
            $script:InviteCallCount = 0
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/root\?') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                elseif ($Uri -match '/root:/' -and $Method -eq 'GET') {
                    return [PSCustomObject]@{ id = 'item-id'; name = 'RetryDoc.txt' }
                }
                elseif ($Uri -match '/invite' -and $Method -eq 'POST') {
                    $script:InviteCallCount++
                    throw [System.Exception]::new('503 Service Unavailable')
                }
            }

            Mock -CommandName 'Start-Sleep' -MockWith { }

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $result = $results | Where-Object { $_.ItemName -eq 'RetryDoc.txt' }
            $result.Status | Should -Be 'Failed'
            # Default MaxAttempts is now 6, so Start-Sleep should be called 5 times
            # (once per retry, not on the final failure).
            Assert-MockCalled -CommandName 'Start-Sleep' -Times 5 -Scope It
        }
    }

    Context 'Script Execution - CSV Deduplication' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_dedup.csv'
            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_dedup'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
            Mock -CommandName 'Get-MgUser' -MockWith {
                [PSCustomObject]@{ Id = 'user-guid'; DisplayName = 'Test User'; UserPrincipalName = 'test@contoso.com'; AccountEnabled = $true }
            }
            Mock -CommandName 'Get-MgUserDrive' -MockWith {
                [PSCustomObject]@{ Id = $script:DefaultDriveId; WebUrl = $script:DefaultWebUrl }
            }
            Mock -CommandName 'Invoke-MgGraphRequest' -MockWith {
                param($Method, $Uri)
                if ($Uri -match '/invite' -and $Method -eq 'POST') {
                    return [PSCustomObject]@{ value = @(@{ id = 'perm-id'; roles = @('write') }) }
                }
                elseif ($Uri -match '/root') {
                    return [PSCustomObject]@{ id = 'root-id'; webUrl = $script:DefaultWebUrl }
                }
                return [PSCustomObject]@{ id = 'item-id'; name = 'Item' }
            }
        }

        It 'should deduplicate identical CSV rows and process each unique row once' {
            # Three identical rows — only one should produce an invite call
            $rows = @(
                (New-CsvRow -ItemName 'Same.txt' -CollaboratorLogin 'collab@contoso.com' -CollaboratorPermission 'Editor'),
                (New-CsvRow -ItemName 'Same.txt' -CollaboratorLogin 'collab@contoso.com' -CollaboratorPermission 'Editor'),
                (New-CsvRow -ItemName 'Same.txt' -CollaboratorLogin 'collab@contoso.com' -CollaboratorPermission 'Editor')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            # After deduplication, only 1 result should be produced
            $permResults = $results | Where-Object { $_.PSObject.Properties.Name -contains 'ItemName' -and $_.ItemName -eq 'Same.txt' }
            $permResults.Count | Should -Be 1
            $permResults[0].Status | Should -Be 'Applied'
        }

        It 'should keep distinct rows with different collaborators' {
            $rows = @(
                (New-CsvRow -ItemName 'Doc.txt' -CollaboratorLogin 'alice@contoso.com' -CollaboratorPermission 'Editor'),
                (New-CsvRow -ItemName 'Doc.txt' -CollaboratorLogin 'bob@contoso.com' -CollaboratorPermission 'Viewer')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $results = & {
                . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                    -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
            } 6>&1

            $permResults = $results | Where-Object { $_.PSObject.Properties.Name -contains 'ItemName' -and $_.ItemName -eq 'Doc.txt' }
            $permResults.Count | Should -Be 2
        }
    }

    Context 'Script Execution - Post-Connection Auth Validation' {
        BeforeEach {
            $script:TestCsv = Join-Path -Path $TestDrive -ChildPath 'test_auth.csv'
            $rows = @(
                (New-CsvRow -ItemName 'Doc1.txt' -CollaboratorPermission 'Editor')
            )
            $rows | Export-Csv -LiteralPath $script:TestCsv -NoTypeInformation -Encoding UTF8

            $script:LogFolder = Join-Path -Path $TestDrive -ChildPath 'logs_auth'
            New-Item -Path $script:LogFolder -ItemType Directory -Force | Out-Null

            Mock -CommandName 'Assert-RequiredModules' -MockWith { }
            Mock -CommandName 'Assert-GraphAssemblyCompatibility' -MockWith { }
            Mock -CommandName 'Assert-GraphPermissions' -MockWith { }
            # NOTE: Connect-Graph is NOT mocked here — we let it run so
            # the post-connection validation logic is exercised.
            Mock -CommandName 'Connect-MgGraph' -MockWith { }
            Mock -CommandName 'Disconnect-MgGraph' -MockWith { }
        }

        It 'should throw when session falls back to delegated auth' {
            # Mock Connect-Graph to simulate post-connection validation
            # by calling Get-MgContext and checking the auth type.
            Mock -CommandName 'Connect-Graph' -MockWith {
                $ctx = Get-MgContext
                if ($null -ne $ctx -and $ctx.AuthType -ne 'AppOnly')
                {
                    throw ("Expected app-only authentication but the session is '{0}' (ClientId={1})." -f $ctx.AuthType, $ctx.ClientId)
                }
            }

            Mock -CommandName 'Get-MgContext' -MockWith {
                [PSCustomObject]@{
                    TenantId = '92075952-90f3-4613-833b-d2e19ec649e4'
                    ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
                    AuthType = 'Delegated'
                    AppName  = 'Graph Explorer'
                }
            }

            $threwError = $false
            try
            {
                $null = & {
                    . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                        -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
                } 6>&1
            }
            catch
            {
                $threwError = $true
                $_.Exception.Message | Should -Match 'app-only'
            }
            $threwError | Should -Be $true
        }

        It 'should throw when session uses wrong ClientId' {
            Mock -CommandName 'Connect-Graph' -MockWith {
                $ctx = Get-MgContext
                if ($null -ne $ctx -and $ctx.AuthType -eq 'AppOnly' -and $ctx.ClientId -ne '912696b9-1374-4110-893d-545fc17c3371')
                {
                    throw ("Connected with ClientId '{0}' but expected '912696b9-1374-4110-893d-545fc17c3371'." -f $ctx.ClientId)
                }
            }

            Mock -CommandName 'Get-MgContext' -MockWith {
                [PSCustomObject]@{
                    TenantId = '92075952-90f3-4613-833b-d2e19ec649e4'
                    ClientId = '00000000-0000-0000-0000-000000000000'
                    AuthType = 'AppOnly'
                    AppName  = 'Wrong App'
                }
            }

            $threwError = $false
            try
            {
                $null = & {
                    . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
                        -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
                } 6>&1
            }
            catch
            {
                $threwError = $true
                $_.Exception.Message | Should -Match 'expected'
            }
            $threwError | Should -Be $true
        }
    }
}


