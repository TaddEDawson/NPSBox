# The #Requires statement prevents the script from running unless the specified
# version of PowerShell is available.  PowerShell 7+ is required for features
# like ternary operators and improved module handling.
# https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_requires
#Requires -Version 7.0

<#
.SYNOPSIS
    Applies OneDrive item sharing permissions based on a CSV file using Microsoft Graph.

    Version: 1.2.2.13
    Date:    2026-05-13

.DESCRIPTION
    This script migrates Box collaboration data into OneDrive for Business.
    It reads a CSV file that describes who has access to what, then applies
    equivalent sharing permissions on the corresponding OneDrive items.

    HOW IT WORKS (step by step):
      1. Authenticates to Microsoft Graph using certificate-based app-only auth.
      2. Reads the CSV and identifies unique owners to process.
         If -UserToProcess is specified, only that user is processed;
         otherwise all unique owners in the CSV are processed.
      3. Looks up the user's OneDrive drive via the Graph API.
      4. Optionally uploads local files/folders to the user's OneDrive
         (when -UploadFiles is specified).
      5. For each CSV row, resolves the item by path in OneDrive, then
         grants the collaborator the appropriate permission (read or write)
         using the driveItem: invite API.
      6. No email notifications are sent (sendInvitation = false).
      7. Outputs a structured result object for each row so you can
         inspect what happened in the pipeline.

    WHAT IS MICROSOFT GRAPH?
      Microsoft Graph is a REST API that lets you interact with Microsoft 365
      services (OneDrive, SharePoint, Teams, Outlook, etc.) programmatically.
      This script uses the Microsoft Graph PowerShell SDK to call Graph.
      https://learn.microsoft.com/graph/overview
      https://learn.microsoft.com/powershell/microsoftgraph/overview

    WHAT IS A UPN (USER PRINCIPAL NAME)?
      A UPN looks like an email address (e.g. user@contoso.com) and uniquely
      identifies a user in Microsoft Entra ID (Azure AD).
      https://learn.microsoft.com/entra/identity/hybrid/connect/plan-connect-userprincipalname

    PREREQUISITES:
      - PowerShell 7.0 or later
        https://learn.microsoft.com/powershell/scripting/install/installing-powershell
      - Microsoft Graph PowerShell SDK modules (install once):
          Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
          Install-Module Microsoft.Graph.Users          -Scope CurrentUser
          Install-Module Microsoft.Graph.Files           -Scope CurrentUser
        https://learn.microsoft.com/powershell/microsoftgraph/installation
      - An Azure AD App Registration with the following APPLICATION permissions
        granted with admin consent:
          Files.ReadWrite.All   - Read/write all users' OneDrive files, upload
                                   content, and grant sharing permissions via the
                                   driveItem: invite API.
          User.Read.All         - Look up user accounts (Get-MgUser) to validate
                                   that owner and collaborator UPNs exist and are
                                   enabled before attempting drive operations.
        https://learn.microsoft.com/graph/permissions-reference
        https://learn.microsoft.com/entra/identity-platform/quickstart-register-app
      - A certificate uploaded to the app registration
        https://learn.microsoft.com/entra/identity-platform/certificate-credentials

    SAFETY:
      -WhatIf   : Shows what would happen without making changes.
      -Verbose   : Shows detailed progress messages.
      -Confirm   : Prompts for confirmation before each change.
      https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_commonparameters

.PARAMETER ConfigFile
    Path to a JSON configuration file used to supply default values for
    InputFile, TenantId, ClientId, CertificateThumbprint, LogFolder, and
    AllFilesDirectory.  Defaults to `config.json` in the script directory.
    Precedence (highest first):
      1. Explicit parameter values supplied on the command line.
      2. Values from this JSON file (when the file exists and the key is set).
      3. Environment variables (NPSBOX_TENANT_ID, NPSBOX_CLIENT_ID,
         NPSBOX_CERT_THUMBPRINT) or hard-coded defaults.
    Missing file is not an error - the script falls back to env vars/defaults.
    https://learn.microsoft.com/powershell/module/microsoft.powershell.utility/convertfrom-json

.PARAMETER InputFile
    Path to the CSV file containing collaboration data.
    The CSV must have these columns:
      - Owner Login             (UPN of the file owner)
      - Path                    (Box path, e.g. "All Files/Documents")
      - Item Name               (file or folder name)
      - Collaborator Login      (UPN of the person to share with)
      - Collaborator Permission (Box role: Editor, Viewer, Co-owner, etc.)

.PARAMETER UserToProcess
    The owner's UPN (User Principal Name) to process.
    Only CSV rows matching this owner will be processed.
    When omitted or empty, all unique owners in the CSV are processed.
    Accepts pipeline input so you can pipe a list of users.
    https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_pipelines

.PARAMETER TenantId
    Your Microsoft 365 tenant ID (a GUID).
    Find it in Azure Portal > Microsoft Entra ID > Overview > Tenant ID.
    Required for certificate auth to target the correct tenant.
    https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant

.PARAMETER ClientId
    The Application (client) ID of your Azure AD app registration.
    Find it in Azure Portal > App registrations > your app > Overview.
    Required for Certificate auth.
    https://learn.microsoft.com/entra/identity-platform/quickstart-register-app

.PARAMETER CertificateThumbprint
    The SHA-1 thumbprint of a certificate installed in Cert:\CurrentUser\My.
    Required for authentication.
    To find your thumbprint:  Get-ChildItem Cert:\CurrentUser\My
    https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph#example-2-using-a-certificate-thumbprint

.PARAMETER LogFolder
    Folder where timestamped log files are written.
    Created automatically if it does not exist.

.PARAMETER AllFilesDirectory
    Root directory containing per-user subfolders of local files to upload.
    Each subfolder must be named by the user's UPN
    (e.g. C:\Repos\NPSBox\LocalFiles\user@contoso.com\).
    Used together with the -UploadFiles switch.

.PARAMETER UploadFiles
    Switch parameter (no value needed - just include it or omit it).
    When present, uploads files and folders from AllFilesDirectory\<UserToProcess>
    to the user's OneDrive root before applying permissions.
    Combine with -WhatIf to preview what would be uploaded.
    https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_switch

.EXAMPLE
    # Preview what would happen (no changes made):
    .\Update-UserFile.ps1 -InputFile .\Box.csv -UserToProcess user@contoso.com -Verbose -WhatIf

.EXAMPLE
    # Apply permissions for a specific user:
    .\Update-UserFile.ps1 -InputFile .\Box.csv -UserToProcess user@contoso.com -Verbose

.EXAMPLE
    # Upload local files and apply permissions for all users in the CSV:
    .\Update-UserFile.ps1 -UploadFiles -Verbose

.NOTES
    DOCUMENTATION LINKS:
      PowerShell Basics:
        - Getting Started:          https://learn.microsoft.com/powershell/scripting/overview
        - About Parameters:         https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_parameters
        - About Pipelines:          https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_pipelines
        - About Try/Catch/Finally:  https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_try_catch_finally

      Microsoft Graph:
        - What is Graph:            https://learn.microsoft.com/graph/overview
        - Graph PowerShell SDK:     https://learn.microsoft.com/powershell/microsoftgraph/overview
        - driveItem: invite API:    https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0
        - Get item by path:         https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path
        - Upload small files:       https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0
        - Connect-MgGraph:          https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
        - Invoke-MgGraphRequest:    https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/invoke-mggraphrequest
        - Permission roles:         https://learn.microsoft.com/graph/api/resources/permission?view=graph-rest-1.0#roles-property-values

      Authentication:
        - App Registration:         https://learn.microsoft.com/entra/identity-platform/quickstart-register-app
        - Certificate credentials:  https://learn.microsoft.com/entra/identity-platform/certificate-credentials
        - Graph auth overview:      https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands
#>

# CmdletBinding enables -Verbose, -WhatIf, -Confirm, and other common parameters.
# SupportsShouldProcess = $true  lets us use $PSCmdlet.ShouldProcess() to guard
#   destructive operations so -WhatIf shows what WOULD happen without doing it.
# ConfirmImpact = 'Medium' means -Confirm prompts only when $ConfirmPreference
#   is Medium or lower (the default).
# https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute
# https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods#shouldprocess
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Help')]
param
(
    # Path to the input CSV file.  [System.IO.FileInfo] automatically resolves
    # the string to a file object with .Exists, .FullName, etc.
    [Parameter(ParameterSetName = 'Run')]
    [Parameter(ParameterSetName = 'Test')]
    [System.IO.FileInfo] $InputFile = "C:\Repos\NPSBox\UserInfo.csv"
    ,
    # Path to a JSON config file providing defaults for InputFile, TenantId,
    # ClientId, CertificateThumbprint, LogFolder, and AllFilesDirectory.
    # Defaults to config.json beside this script.  Missing file is OK.
    [Parameter(ParameterSetName = 'Run')]
    [Parameter(ParameterSetName = 'Test')]
    [string] $ConfigFile = (Join-Path -Path $PSScriptRoot -ChildPath 'config.json')
    ,
    # The owner's UPN to filter on in the CSV.
    # ValueFromPipeline lets you pipe UPNs:  'user1@contoso.com','user2@contoso.com' | .\Update-UserFile.ps1
    # Alias allows matching CSV column names directly for pipeline binding.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters#alias-attribute
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'Run')]
    [Alias('Owner Login', 'User', 'UPN', 'Account')]
    [string] $UserToProcess
    ,
    # Your tenant ID (GUID).  Find it: Azure Portal > Entra ID > Overview.
    # Defaults to $env:NPSBOX_TENANT_ID if set, otherwise from -ConfigFile.
    # Validated as non-empty in the begin block AFTER config merge.
    [Parameter(Mandatory = $false, ParameterSetName = 'Run')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Test')]
    [string] $TenantId = $env:NPSBOX_TENANT_ID
    ,
    # The app registration's client ID (GUID).
    # Defaults to $env:NPSBOX_CLIENT_ID if set, otherwise from -ConfigFile.
    # Validated as non-empty in the begin block AFTER config merge.
    [Parameter(Mandatory = $false, ParameterSetName = 'Run')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Test')]
    [string] $ClientId = $env:NPSBOX_CLIENT_ID
    ,
    # Certificate thumbprint for app-only auth.
    # Defaults to $env:NPSBOX_CERT_THUMBPRINT if set, otherwise from -ConfigFile.
    # Validated as non-empty in the begin block AFTER config merge.
    [Parameter(Mandatory = $false, ParameterSetName = 'Run')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Test')]
    [string] $CertificateThumbprint = $env:NPSBOX_CERT_THUMBPRINT
    ,
    # Where to write timestamped log files.  Created if it doesn't exist.
    [Parameter(ParameterSetName = 'Run')]
    [Parameter(ParameterSetName = 'Test')]
    [string] $LogFolder = "C:\Repos\NPSBox\Logs"
    ,
    # Root folder with per-user subfolders of files to upload.
    # Subfolder names must match the user's UPN exactly.
    [Parameter(ParameterSetName = 'Run')]
    [string] $AllFilesDirectory = "C:\Repos\NPSBox\LocalFiles"
    ,
    # Include this switch to upload local files to OneDrive before applying permissions.
    # A switch parameter is either present ($true) or absent ($false) - no value needed.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_switch
    [Parameter(ParameterSetName = 'Run')]
    [switch] $UploadFiles
    ,
    # Optional list of allowed email domains for collaborators.
    # When specified, collaborators whose domain is not in this list are skipped
    # with a warning instead of being granted access.  Prevents accidental
    # external sharing.  Example: @('contoso.com', 'contoso.onmicrosoft.com')
    [Parameter(ParameterSetName = 'Run')]
    [string[]] $AllowedDomains
    ,
    # When present, verifies authentication and access requirements only.
    # The script authenticates to Microsoft Graph, asserts required modules,
    # checks assembly compatibility, validates app permissions, and then exits
    # without processing any CSV rows.  Useful for pre-flight validation.
    [Parameter(Mandatory = $true, ParameterSetName = 'Test')]
    [switch] $Test
)

# #===============================================================================#
# #  BEGIN BLOCK                                                                 #
# #  Runs once before any pipeline input is processed.                           #
# #  Used here to define helper functions, import modules, and authenticate.     #
# #  https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods #
# #===============================================================================#
begin
{
    # -- Help mode: show help and exit when no parameters are provided --------
    # DefaultParameterSetName = 'Help' activates when no params are supplied.
    # Display the about_help topic and return immediately.
    $script:HelpModeActive = $false
    if ($PSCmdlet.ParameterSetName -eq 'Help')
    {
        $aboutFile = Join-Path -Path $PSScriptRoot -ChildPath 'en-US' |
            Join-Path -ChildPath 'about_Update-UserFile.help.txt'
        if (Test-Path -LiteralPath $aboutFile)
        {
            Get-Content -LiteralPath $aboutFile -Raw | Write-Output
        } # if
        else
        {
            Get-Help -Name $PSCommandPath -Detailed
        } # else
        $script:HelpModeActive = $true
        return
    } # if - Help mode

    # -- Load defaults from ConfigFile --------------------------------------------
    # Apply values from the JSON config to any parameter the caller did NOT
    # explicitly bind.  This lets users keep tenant/cert/path settings in a
    # config.json beside the script instead of passing them every invocation.
    # Precedence: explicit parameter > config.json > env var / hard-coded default.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.utility/convertfrom-json
    if (-not [string]::IsNullOrWhiteSpace($ConfigFile) -and (Test-Path -LiteralPath $ConfigFile -PathType Leaf))
    {
        try
        {
            $configData = Get-Content -LiteralPath $ConfigFile -Raw -ErrorAction Stop |
                ConvertFrom-Json -ErrorAction Stop
            Write-Verbose ("Loaded configuration from {0}" -f $ConfigFile)

            $configMap = @{
                InputFile             = 'InputFile'
                TenantId              = 'TenantId'
                ClientId              = 'ClientId'
                CertificateThumbprint = 'CertificateThumbprint'
                LogFolder             = 'LogFolder'
                AllFilesDirectory     = 'AllFilesDirectory'
            }
            foreach ($key in $configMap.Keys)
            {
                # Only override when the caller did not explicitly supply the parameter.
                if ($PSBoundParameters.ContainsKey($key)) { continue }
                if (-not ($configData.PSObject.Properties.Name -contains $key)) { continue }
                $value = $configData.$key
                if ($null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value))) { continue }
                Set-Variable -Name $configMap[$key] -Value $value -Scope 0
                Write-Verbose ("Config override: {0} = {1}" -f $key, $value)
            } # foreach - config key
        } # try
        catch
        {
            Write-Warning ("Failed to read ConfigFile '{0}': {1}" -f $ConfigFile, $_.Exception.Message)
        } # catch
    } # if - ConfigFile exists
    else
    {
        Write-Verbose ("ConfigFile not found or not specified: {0}" -f $ConfigFile)
    } # else

    # -- Validate required credentials (post-config) ------------------------------------
    # Validation runs here - after config merge - so config.json can supply
    # credentials when neither the caller nor env vars provide them.
    $missingCreds = @()
    if ([string]::IsNullOrWhiteSpace($TenantId))              { $missingCreds += 'TenantId' }
    if ([string]::IsNullOrWhiteSpace($ClientId))              { $missingCreds += 'ClientId' }
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $missingCreds += 'CertificateThumbprint' }
    if ($missingCreds.Count -gt 0)
    {
        throw ("Missing required credential value(s): {0}. Supply via parameter, -ConfigFile (config.json), or env vars NPSBOX_TENANT_ID/NPSBOX_CLIENT_ID/NPSBOX_CERT_THUMBPRINT." -f ($missingCreds -join ', '))
    } # if

    # -- Write-LogLine ------------------------------------------------------------
    # Writes a timestamped message to both the Verbose stream and a log file.
    # Write-Verbose sends output to the verbose stream (visible only with -Verbose).
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.utility/write-verbose
    #
    # Note: We temporarily disable $WhatIfPreference when writing to the log file
    # so that Add-Content actually writes even when the script is run with -WhatIf.
    function Write-LogLine
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $Message
            ,
            [Parameter()]
            [ValidateSet('INFO', 'WARN', 'ERROR')]
            [string] $Level = 'INFO'
        )

        # -f is the format operator:  "{0} {1}" -f 'Hello','World'  =>  "Hello World"
        # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_operators#format-operator--f
        #
        # Get-PSCallStack returns the call stack; index [1] is the immediate caller.
        # ScriptLineNumber gives us the line in the .ps1 file that invoked Write-LogLine.
        # https://learn.microsoft.com/powershell/module/microsoft.powershell.utility/get-pscallstack
        $callerFrame = (Get-PSCallStack)[1]
        $callerLine  = if ($null -ne $callerFrame) { $callerFrame.ScriptLineNumber } else { 0 }
        $line = "{0} [{1}] (line {2}) {3}" -f (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffK'), $Level, $callerLine, $Message
        Write-Verbose $line

        if (-not [string]::IsNullOrWhiteSpace($script:LogFilePath))
        {
            try
            {
                $previousWhatIfPreference = $WhatIfPreference
                try
                {
                    $WhatIfPreference = $false
                    Add-Content -LiteralPath $script:LogFilePath -Value $line -ErrorAction Stop
                } # try
                finally
                {
                    $WhatIfPreference = $previousWhatIfPreference
                } # finally - always restores the original $WhatIfPreference
            } # try
            catch
            {
                # Write-Warning outputs a non-terminating warning that appears in yellow.
                Write-Warning "Failed to write log line: $($_.Exception.Message)"
            } # catch
        } # if
    } # function Write-LogLine

    # -- Assert-RequiredModules ---------------------------------------------------
    # Ensures the Microsoft Graph PowerShell SDK modules are installed and imports them.
    # Modules are reusable packages of PowerShell commands.  The Graph SDK is split
    # into sub-modules (Authentication, Users, Files, etc.) to keep imports small.
    #
    # Install the required modules once (you only need to do this one time):
    #   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    #   Install-Module Microsoft.Graph.Users          -Scope CurrentUser
    #   Install-Module Microsoft.Graph.Files           -Scope CurrentUser
    # https://learn.microsoft.com/powershell/microsoftgraph/installation
    function Assert-RequiredModules
    {
        [CmdletBinding()]
        param()

        $requiredModules = @(
            'Microsoft.Graph.Authentication',   # Provides Connect-MgGraph, Invoke-MgGraphRequest
            'Microsoft.Graph.Users',            # Provides Get-MgUser and user-related cmdlets
            'Microsoft.Graph.Files',            # Provides Get-MgUserDrive and drive-related cmdlets
            'Microsoft.Graph.Applications'      # Provides Get-MgServicePrincipal for permission checks
        )

        foreach ($moduleName in $requiredModules)
        {
            # Get-Module -ListAvailable checks what is installed (not yet loaded).
            # We pick the newest version if multiple are installed.
            $availableModule = Get-Module -ListAvailable -Name $moduleName |
                Sort-Object -Property Version -Descending |
                Select-Object -First 1

            if ($null -eq $availableModule)
            {
                # 'throw' stops the script with an error.  It is a "terminating error".
                # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_throw
                throw (
                    "Required module not found: $moduleName. Install it with: Install-Module $moduleName -Scope CurrentUser"
                )
            } # if

            Write-Verbose ("Importing module {0} ({1})" -f $moduleName, $availableModule.Version)
            # Import-Module loads the module into the current session so its commands are available.
            # -RequiredVersion ensures we load the exact version we checked.
            # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/import-module
            Import-Module -Name $moduleName -RequiredVersion $availableModule.Version -ErrorAction Stop -Verbose:$false | Out-Null
        } # foreach
    } # function Assert-RequiredModules

    # -- ConvertTo-GraphRole ------------------------------------------------------
    # Maps a Box permission name to a Microsoft Graph sharing role.
    # Graph supports two sharing roles for the invite API:
    #   'read'   - view-only access
    #   'write'  - view + edit access
    # Box has more granular roles; some (Previewer, Uploader) have no equivalent
    # in Graph so they return $null and the row is skipped.
    # https://learn.microsoft.com/graph/api/resources/permission?view=graph-rest-1.0#roles-property-values
    function ConvertTo-GraphRole
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $BoxPermission
        )

        # The 'switch' statement is PowerShell's equivalent of if/else-if chains.
        # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_switch
        switch ($BoxPermission)
        {
            'Co-owner'           { return 'write' }   # Full edit access
            'Editor'             { return 'write' }   # Edit access
            'Viewer Uploader'    { return 'read'  }   # Read-only (upload aspect not supported)
            'Viewer'             { return 'read'  }   # Read-only
            'Previewer Uploader' { return $null   }   # No Graph equivalent - skip
            'Previewer'          { return $null   }   # No Graph equivalent - skip
            'Uploader'           { return $null   }   # No Graph equivalent - skip
            default              { return $null   }   # Unknown - skip
        } # switch
    } # function ConvertTo-GraphRole

    # -- Test-CollaboratorDomain ----------------------------------------------------
    # Validates that a collaborator's email domain is in the AllowedDomains list.
    # Returns $true if the domain is allowed (or if AllowedDomains is not set).
    # Returns $false if the domain is blocked.
    # This prevents accidental external sharing when the CSV contains addresses
    # outside the organisation.
    function Test-CollaboratorDomain
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $Email
            ,
            [Parameter()]
            [string[]] $Domains
        )

        # If no domain allowlist was provided, all domains are permitted.
        if ($null -eq $Domains -or $Domains.Count -eq 0)
        {
            return $true
        } # if

        $atIndex = $Email.LastIndexOf('@')
        if ($atIndex -lt 0)
        {
            return $false
        } # if

        $emailDomain = $Email.Substring($atIndex + 1).ToLowerInvariant()
        $lowerDomains = $Domains | ForEach-Object { $_.ToLowerInvariant() }
        return ($emailDomain -in $lowerDomains)
    } # function Test-CollaboratorDomain

    # -- Test-EmailFormat --------------------------------------------------------
    # Validates that a string is a plausible email address.
    # Uses a basic regex: local-part @ domain with at least one dot.
    # Not a full RFC 5322 implementation, but catches obvious non-email values
    # (e.g. "notanemail", "@domain", "user@").
    # Returns $true if the format is valid, $false otherwise.
    function Test-EmailFormat
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $Email
        )

        return ($Email -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
    } # function Test-EmailFormat

    # -- Assert-CsvColumns ---------------------------------------------------------
    # Validates that the CSV contains all required column headers.
    # Throws with a clear message listing any missing columns.
    # This prevents null-reference errors deep in the processing loop.
    function Assert-CsvColumns
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [object[]] $CsvRows
        )

        $requiredColumns = @('Owner Login', 'Path', 'Item Name', 'Collaborator Login', 'Collaborator Permission')
        $firstRow = $CsvRows | Select-Object -First 1
        if ($null -eq $firstRow)
        {
            throw 'CSV file is empty - no rows to process.'
        } # if

        $actualColumns = @($firstRow.PSObject.Properties.Name)
        $missingColumns = @($requiredColumns | Where-Object { $_ -notin $actualColumns })

        if ($missingColumns.Count -gt 0)
        {
            throw ("CSV is missing required column(s): {0}. Expected columns: {1}. Found columns: {2}" -f
                ($missingColumns -join ', '),
                ($requiredColumns -join ', '),
                ($actualColumns -join ', '))
        } # if
    } # function Assert-CsvColumns

    # -- ConvertTo-OneDriveRelativePath --------------------------------------------
    # Cleans up the Box export path so it can be used with the Graph API.
    # Box exports include a root label "All Files/" which does not exist in OneDrive.
    # This function strips that prefix, normalizes backslashes to forward slashes,
    # and trims extra slashes.
    #
    # Example: "All Files/Documents/Report.pdf" -> "Documents/Report.pdf"
    function ConvertTo-OneDriveRelativePath
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $Path
        )

        $normalized = $Path.Trim()
        if ([string]::IsNullOrWhiteSpace($normalized))
        {
            throw "Row Path is empty."
        } # if

        # Replace Windows-style backslashes with forward slashes for the Graph API.
        $normalized = $normalized -replace '\\', '/'
        $normalized = $normalized.Trim('/')

        # The -match operator tests a string against a regex pattern.
        # (?i) makes it case-insensitive.  (?:/|$) matches a slash or end-of-string.
        # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_regular_expressions
        if ($normalized -match '^(?i)all files(?:/|$)')
        {
            # The -replace operator substitutes matches with the replacement string.
            $normalized = $normalized -replace '^(?i)all files(?:/|$)', ''
            $normalized = $normalized.Trim('/')
        } # if

        if ([string]::IsNullOrWhiteSpace($normalized))
        {
            throw ("Row Path '{0}' resolves to empty OneDrive-relative path." -f $Path)
        } # if

        return $normalized
    } # function ConvertTo-OneDriveRelativePath

    # -- ConvertTo-GraphEncodedPath ------------------------------------------------
    # URL-encodes each segment of a relative path so special characters (spaces,
    # parentheses, etc.) are safe to use in Graph API URLs.
    #
    # Example: "Thesis (IPv6)/Report.pdf" -> "Thesis%20%28IPv6%29/Report.pdf"
    #
    # Graph uses the pattern /drives/{id}/root:/{encoded-path} to access items.
    # https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path
    function ConvertTo-GraphEncodedPath
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $RelativePath
        )

        # -split '/' breaks the path into individual folder/file names.
        # We encode each one separately so the '/' separators stay intact.
        $encodedSegments = foreach ($segment in ($RelativePath -split '/'))
        {
            if ([string]::IsNullOrWhiteSpace($segment))
            {
                continue
            } # if

            # EscapeDataString percent-encodes characters like spaces and parentheses.
            # https://learn.microsoft.com/dotnet/api/system.uri.escapedatastring
            [System.Uri]::EscapeDataString($segment)
        } # ?:$encodedSegments = foreach ($segment in ($RelativePath -split '/'))

        if ($null -eq $encodedSegments -or $encodedSegments.Count -eq 0)
        {
            throw ("Could not encode OneDrive-relative path: '{0}'" -f $RelativePath)
        } # if

        # -join '/' reassembles the encoded segments back into a path string.
        return ($encodedSegments -join '/')
    } # function ConvertTo-GraphEncodedPath

    # -- Test-IsRetryableGraphError ------------------------------------------------
    # Determines whether a Graph API error is transient and worth retrying.
    # Transient errors include:
    #   - HTTP 429 (Too Many Requests / throttling)
    #   - HTTP 500, 502, 503, 504 (server errors)
    #   - Timeouts, canceled requests, and temporary failures
    # Non-transient errors (401, 403, 404) are NOT retried.
    # https://learn.microsoft.com/graph/errors
    # https://learn.microsoft.com/graph/throttling
    function Test-IsRetryableGraphError
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.ErrorRecord] $ErrorRecord
        )

        $message = [string] $ErrorRecord.Exception.Message
        $details = [string] $ErrorRecord.ErrorDetails.Message
        $combined = ($message + " " + $details).ToLowerInvariant()

        # The -match operator tests against a regex pattern.  The | means "or".
        # \b is a word boundary so "429" doesn't accidentally match inside other numbers.
        return (
            $combined -match 'timeout|timed out|httpclient\.timeout|request was canceled|temporar|try again|throttl|too many requests|\b429\b|\b500\b|\b502\b|\b503\b|\b504\b|serviceunavailable|gatewaytimeout'
        )
    } # function Test-IsRetryableGraphError

    # -- Get-RetryAfterSeconds ---------------------------------------------------
    # Parses the Retry-After value from a Graph API error response.
    # Graph 429 responses include a Retry-After header (in seconds) that tells
    # the client exactly how long to wait before retrying.
    # Returns the parsed value in seconds, or $null if not found.
    # https://learn.microsoft.com/graph/throttling
    function Get-RetryAfterSeconds
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.ErrorRecord] $ErrorRecord
        )

        # Graph SDK errors embed the Retry-After value in the error details or
        # exception message.  Look for common patterns.
        $details = [string] $ErrorRecord.ErrorDetails.Message
        $message = [string] $ErrorRecord.Exception.Message
        $combined = $details + ' ' + $message

        # Pattern 1: JSON body with "retryAfterSeconds" or "retry-after" key.
        if ($combined -match '"retry[\-_]?after(?:Seconds)?"\s*:\s*(\d+)')
        {
            return [int] $Matches[1]
        } # if

        # Pattern 2: Header-style "Retry-After: <seconds>" in error text.
        if ($combined -match 'Retry-After\s*:\s*(\d+)')
        {
            return [int] $Matches[1]
        } # if

        # Pattern 3: Check the exception's Response.Headers if available
        # (HttpRequestException or similar with a Response property).
        $response = $ErrorRecord.Exception.PSObject.Properties['Response']
        if ($null -ne $response -and $null -ne $response.Value)
        {
            $headers = $response.Value.PSObject.Properties['Headers']
            if ($null -ne $headers -and $null -ne $headers.Value)
            {
                $retryHeader = $null
                if ($headers.Value -is [System.Collections.IDictionary])
                {
                    $retryHeader = $headers.Value['Retry-After']
                } # if

                if ($null -ne $retryHeader)
                {
                    $parsed = 0
                    if ([int]::TryParse([string] $retryHeader, [ref] $parsed))
                    {
                        return $parsed
                    } # if
                } # if
            } # if
        } # if

        return $null
    } # function Get-RetryAfterSeconds

    # -- Invoke-WithGraphRetry ----------------------------------------------------
    # Wraps a Graph API call with automatic retry and exponential backoff.
    # If the call fails with a transient error (timeout, 429, 5xx), it waits and
    # retries up to MaxAttempts times.  The wait doubles each time (exponential
    # backoff) to avoid hammering the server.
    #
    # A [scriptblock] is a block of PowerShell code you pass as a parameter.
    # The & operator executes it:  & { Get-Date }
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_script_blocks
    function Invoke-WithGraphRetry
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [scriptblock] $Operation
            ,
            [Parameter(Mandatory = $true)]
            [string] $OperationName
            ,
            [Parameter()]
            [ValidateRange(1, 10)]
            [int] $MaxAttempts = 6
            ,
            [Parameter()]
            [ValidateRange(1, 60)]
            [int] $InitialDelaySeconds = 2
            ,
            [Parameter()]
            [ValidateRange(1, 300)]
            [int] $MaxDelaySeconds = 60
        )

        $attempt = 1
        $delaySeconds = $InitialDelaySeconds

        while ($true)
        {
            try
            {
                return (& $Operation)
            } # try
            catch
            {
                $isRetryable = Test-IsRetryableGraphError -ErrorRecord $_
                if ((-not $isRetryable) -or $attempt -ge $MaxAttempts)
                {
                    throw
                } # if

                Write-LogLine -Level 'WARN' -Message (
                    "Transient Graph failure during '{0}' (attempt {1}/{2}): {3}. Retrying in {4}s." -f
                    $OperationName, $attempt, $MaxAttempts, $_.Exception.Message, $delaySeconds
                )

                # Honour the Retry-After header from Graph 429 responses when available.
                # This tells us exactly how long to wait, which is more accurate than
                # our exponential backoff guess.
                $retryAfter = Get-RetryAfterSeconds -ErrorRecord $_
                if ($null -ne $retryAfter -and $retryAfter -gt 0)
                {
                    # Honor the Retry-After value from the server, even if it
                    # exceeds MaxDelaySeconds.  Capping caused premature retry
                    # exhaustion under heavy throttling (finding T4).
                    $delaySeconds = $retryAfter
                    Write-LogLine -Level 'WARN' -Message ("Using Retry-After value: {0}s." -f $retryAfter)
                } # if

                Start-Sleep -Seconds $delaySeconds
                $attempt += 1
                $delaySeconds = [Math]::Min($delaySeconds * 2, $MaxDelaySeconds)
            } # catch
        } # while
    } # function Invoke-WithGraphRetry

    # -- Connect-GraphCertAuth -----------------------------------------------------
    # Authenticates to Microsoft Graph using certificate-based app-only auth.
    #
    # Certificate mode uses a certificate for "app-only" auth - no user sign-in
    # is required.  This is how you run the script unattended (e.g. scheduled).
    # The app registration must have Application permissions granted with admin consent.
    #
    # Connect-MgGraph reference:
    #   https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
    # Auth overview:
    #   https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands
    function Connect-GraphCertAuth
    {
        [CmdletBinding()]
        param()

        $previousWhatIfPreference = $WhatIfPreference
        try
        {
            # We disable $WhatIfPreference during authentication so that
            # Connect-MgGraph actually runs even when the script is invoked
            # with -WhatIf.  Authentication is a read-only operation.
            $WhatIfPreference = $false

            # Get-MgContext returns the current Graph session (or $null).
            # If we already have a session for the correct tenant, skip re-auth.
            $existingContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -ne $existingContext -and $existingContext.TenantId -eq $TenantId -and $existingContext.AuthType -eq 'AppOnly')
            {
                Write-LogLine -Message ("Reusing existing Microsoft Graph context (app-only). TenantId={0}, AppName={1}, AuthType={2}" -f
                    $existingContext.TenantId, $existingContext.AppName, $existingContext.AuthType)
                return
            } # if

            # Disconnect any existing delegated or mismatched session before
            # establishing certificate-based app-only auth.
            if ($null -ne $existingContext)
            {
                Write-LogLine -Message ("Disconnecting existing Graph session (AuthType={0}) to establish app-only auth." -f $existingContext.AuthType)
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            } # if

            if ([string]::IsNullOrWhiteSpace($TenantId))
            {
                throw "Certificate auth requires -TenantId."
            } # if

            if ([string]::IsNullOrWhiteSpace($ClientId))
            {
                throw "Certificate auth requires -ClientId."
            } # if

            Write-LogLine -Message ("Connecting to Microsoft Graph using Certificate thumbprint auth. TenantId={0}, ClientId={1}" -f $TenantId, $ClientId)
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop -NoWelcome | Out-Null

            # -- Post-connection validation ------------------------------------
            # Verify the resulting session is actually app-only and uses the
            # expected ClientId.  If Connect-MgGraph silently fell back to
            # delegated auth (e.g. certificate not found), the session will
            # lack the required Application permissions.
            $newContext = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -eq $newContext)
            {
                throw "Connect-MgGraph completed but Get-MgContext returned null. Authentication may have failed."
            } # if

            if ($newContext.AuthType -ne 'AppOnly')
            {
                throw (
                    ("Expected app-only authentication but the session is '{0}' (ClientId={1}). " +
                    "Verify that the certificate with thumbprint '{2}' is installed in Cert:\CurrentUser\My " +
                    "and is associated with the app registration (ClientId={3}).") -f
                    $newContext.AuthType, $newContext.ClientId, $CertificateThumbprint, $ClientId
                )
            } # if

            if ($newContext.ClientId -ne $ClientId)
            {
                throw (
                    ("Connected with ClientId '{0}' but expected '{1}'. " +
                    "Disconnect any existing Graph sessions and retry.") -f
                    $newContext.ClientId, $ClientId
                )
            } # if

            Write-LogLine -Message ("Connected to Microsoft Graph (AppOnly). TenantId={0}, ClientId={1}" -f $newContext.TenantId, $newContext.ClientId)
        } # try
        finally
        {
            $WhatIfPreference = $previousWhatIfPreference
        } # finally
    } # function Connect-GraphCertAuth

    # -- Invoke-OneDriveResumableUpload -------------------------------------------
    # Uploads a single local file to OneDrive using a resumable upload session.
    # Required for files larger than the 4 MB simple-upload limit.
    #
    # Flow:
    #   1. POST .../createUploadSession         (returns a short-lived uploadUrl)
    #   2. PUT  <uploadUrl>   for each chunk    (with Content-Range header)
    #
    # The uploadUrl is pre-authenticated, so chunk PUTs go through Invoke-WebRequest
    # WITHOUT adding Graph auth headers (using Invoke-MgGraphRequest here would
    # incorrectly add a bearer token).
    # https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
    function Invoke-OneDriveResumableUpload
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $DriveId
            ,
            [Parameter(Mandatory = $true)]
            [string] $EncodedRelPath
            ,
            [Parameter(Mandatory = $true)]
            [string] $LocalFilePath
            ,
            # Chunk size in bytes.  Microsoft Graph requires every chunk except the
            # last to be a multiple of 320 KiB (327,680 bytes).  10 MiB (32 x 320 KiB)
            # is the commonly recommended value.
            [Parameter()]
            [ValidateScript({ ($_ % 327680) -eq 0 })]
            [int] $ChunkSizeBytes = (10 * 1024 * 1024)
        )

        $fileInfo  = Get-Item -LiteralPath $LocalFilePath -ErrorAction Stop
        $totalBytes = [long]$fileInfo.Length
        $fileName  = [System.IO.Path]::GetFileName($LocalFilePath)

        # Step 1: create the upload session.
        $createSessionUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${EncodedRelPath}:/createUploadSession"
        $sessionBody = @{
            item = @{
                '@microsoft.graph.conflictBehavior' = 'replace'
                name                                = $fileName
            }
        } | ConvertTo-Json -Depth 4

        $session = Invoke-WithGraphRetry -OperationName ("Create upload session for '{0}'" -f $EncodedRelPath) -Operation {
            Invoke-MgGraphRequest -Method POST -Uri $createSessionUri -Body $sessionBody -ContentType 'application/json' -ErrorAction Stop
        } # inline:Invoke-WithGraphRetry - createUploadSession

        $uploadUrl = if ($session -is [System.Collections.IDictionary]) { $session['uploadUrl'] } else { $session.uploadUrl }
        if ([string]::IsNullOrWhiteSpace($uploadUrl))
        {
            throw 'createUploadSession did not return an uploadUrl.'
        } # if

        # Step 2: stream the file in chunks to the pre-authenticated uploadUrl.
        $stream = [System.IO.File]::OpenRead($LocalFilePath)
        try
        {
            $buffer       = [byte[]]::new($ChunkSizeBytes)
            $offset       = [long]0
            $lastResponse = $null

            while ($offset -lt $totalBytes)
            {
                $bytesToRead = [int][Math]::Min([long]$ChunkSizeBytes, $totalBytes - $offset)
                $bytesRead   = $stream.Read($buffer, 0, $bytesToRead)
                if ($bytesRead -le 0) { break }

                if ($bytesRead -ne $buffer.Length)
                {
                    $chunk = [byte[]]::new($bytesRead)
                    [Array]::Copy($buffer, 0, $chunk, 0, $bytesRead)
                } # if
                else
                {
                    $chunk = $buffer
                } # else

                $rangeEnd     = $offset + $bytesRead - 1
                $contentRange = 'bytes {0}-{1}/{2}' -f $offset, $rangeEnd, $totalBytes
                $headers      = @{ 'Content-Range' = $contentRange }

                $lastResponse = Invoke-WithGraphRetry -OperationName ('Upload chunk {0}-{1}/{2}' -f $offset, $rangeEnd, $totalBytes) -Operation {
                    Invoke-WebRequest -Method PUT -Uri $uploadUrl -Body $chunk -Headers $headers -ContentType 'application/octet-stream' -ErrorAction Stop
                } # inline:Invoke-WithGraphRetry - chunk PUT

                $offset += $bytesRead
            } # while

            return $lastResponse
        } # try
        finally
        {
            $stream.Dispose()
        } # finally
    } # function Invoke-OneDriveResumableUpload

    # -- Invoke-OneDriveUpload -----------------------------------------------------
    # Uploads local files and folders to a user's OneDrive.
    # Folders are created first (parents before children) via PATCH with a folder
    # body, and files are uploaded via PUT /content.
    #
    # Files up to 4 MB can use the simple upload endpoint:
    #   PUT /drives/{driveId}/root:/{path}:/content
    #   https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0
    #
    # For files larger than 4 MB, you would need a resumable upload session:
    #   https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
    #   (not implemented in this script)
    #
    # Supports -WhatIf:  when set, lists what WOULD be created/uploaded without
    # making any changes.
    function Invoke-OneDriveUpload
    {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $DriveId
            ,
            [Parameter(Mandatory = $true)]
            [string] $LocalSourcePath
            ,
            [Parameter(Mandatory = $true)]
            [string] $OwnerUpn
        )

        if (-not (Test-Path -LiteralPath $LocalSourcePath))
        {
            throw ("Local source path not found: '{0}'" -f $LocalSourcePath)
        } # if

        # Get-ChildItem -Recurse lists all files and folders under the path.
        # -Force includes hidden files.
        # https://learn.microsoft.com/powershell/module/microsoft.powershell.management/get-childitem
        $allItems = Get-ChildItem -LiteralPath $LocalSourcePath -Recurse -Force
        $baseLength = $LocalSourcePath.TrimEnd('\', '/').Length + 1

        # Process folders first (sorted by path length = depth) so parent folders
        # are created before their children.
        $folders = $allItems | Where-Object { $_.PSIsContainer } | Sort-Object { $_.FullName.Length }
        foreach ($folder in $folders)
        {
            $relativePath = $folder.FullName.Substring($baseLength) -replace '\\', '/'
            $encodedRelPath = ConvertTo-GraphEncodedPath -RelativePath $relativePath
            # Build the parent path and folder name for the POST /children API.
            # https://learn.microsoft.com/graph/api/driveitem-post-children?view=graph-rest-1.0
            $segments = $relativePath -split '/'
            $folderName = $segments[-1]
            if ($segments.Count -gt 1)
            {
                $parentRelPath = ($segments[0..($segments.Count - 2)]) -join '/'
                $encodedParentPath = ConvertTo-GraphEncodedPath -RelativePath $parentRelPath
                $childrenUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${encodedParentPath}:/children"
            } # if
            else
            {
                $childrenUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
            } # else

            $result = [pscustomobject]@{
                OwnerLogin   = $OwnerUpn
                LocalPath    = $folder.FullName
                OneDrivePath = $relativePath
                ItemType     = 'Folder'
                Action       = 'CreateFolder'
                Status       = 'Unknown'
                Error        = $null
            } # inline:$result = [pscustomobject]@{

            try
            {
                if ($PSCmdlet.ShouldProcess("OneDrive:/$relativePath", "Create folder"))
                {
                    $body = @{
                        name                              = $folderName
                        folder                            = @{}
                        '@microsoft.graph.conflictBehavior' = 'fail'
                    } | ConvertTo-Json -Depth 4
                    try
                    {
                        Invoke-WithGraphRetry -OperationName ("Create folder '{0}'" -f $relativePath) -Operation {
                            Invoke-MgGraphRequest -Method POST -Uri $childrenUri -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                        } # inline:Invoke-WithGraphRetry
                    } # try
                    catch
                    {
                        # 409 Conflict means the folder already exists - safe to continue.
                        if ($_.Exception.Message -match '409|nameAlreadyExists|conflict')
                        {
                            Write-LogLine -Message ("Folder already exists (409): OneDrive:/{0}" -f $relativePath)
                        } # if
                        else
                        {
                            throw
                        } # else
                    } # catch - 409 handling
                    $result.Status = 'Applied'
                    Write-LogLine -Message ("Created folder: OneDrive:/{0}" -f $relativePath)
                } # if
                else
                {
                    $result.Status = 'WhatIf'
                } # else
            } # try
            catch
            {
                $result.Status = 'Failed'
                $result.Error  = $_.Exception.Message
                Write-LogLine -Level 'ERROR' -Message ("Failed to create folder '{0}': {1}" -f $relativePath, $result.Error)
            } # catch

            $result
        } # foreach

        # Process files.
        $files = $allItems | Where-Object { -not $_.PSIsContainer }
        foreach ($file in $files)
        {
            $relativePath = $file.FullName.Substring($baseLength) -replace '\\', '/'
            $encodedRelPath = ConvertTo-GraphEncodedPath -RelativePath $relativePath
            $uploadUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${encodedRelPath}:/content"

            $result = [pscustomobject]@{
                OwnerLogin   = $OwnerUpn
                LocalPath    = $file.FullName
                OneDrivePath = $relativePath
                ItemType     = 'File'
                SizeBytes    = $file.Length
                Action       = 'UploadFile'
                Status       = 'Unknown'
                DurationSec  = 0.0
                RateKBps     = 0.0
                Error        = $null
            } # inline:$result = [pscustomobject]@{

            try
            {
                # Files <= 4 MB use the simple upload endpoint; larger files use a
                # resumable upload session (Invoke-OneDriveResumableUpload).
                # https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0
                # https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0
                $maxSimpleUploadBytes = 4 * 1024 * 1024  # 4 MB
                $useResumable = ($file.Length -gt $maxSimpleUploadBytes)

                $shouldProcessTarget = "OneDrive:/$relativePath ($($file.Length) bytes)"
                $shouldProcessAction = if ($useResumable) { 'Upload file (resumable)' } else { 'Upload file' }
                if ($PSCmdlet.ShouldProcess($shouldProcessTarget, $shouldProcessAction))
                {
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    try
                    {
                        if ($useResumable)
                        {
                            Invoke-OneDriveResumableUpload -DriveId $DriveId -EncodedRelPath $encodedRelPath -LocalFilePath $file.FullName | Out-Null
                        } # if
                        else
                        {
                            $fileBytes = [System.IO.File]::ReadAllBytes($file.FullName)
                            Invoke-WithGraphRetry -OperationName ("Upload file '{0}'" -f $relativePath) -Operation {
                                Invoke-MgGraphRequest -Method PUT -Uri $uploadUri -Body $fileBytes -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
                            } # inline:Invoke-WithGraphRetry - simple upload PUT
                        } # else
                    } # try
                    finally
                    {
                        $sw.Stop()
                    } # finally
                    $elapsedSec = $sw.Elapsed.TotalSeconds
                    $result.DurationSec = [Math]::Round($elapsedSec, 1)
                    $rateKBps = if ($elapsedSec -gt 0) { ($file.Length / 1KB) / $elapsedSec } else { 0.0 }
                    $result.RateKBps = [Math]::Round($rateKBps, 1)
                    $result.Status = 'Completed'
                    Write-LogLine -Message ("Uploaded file ({0}): OneDrive:/{1} ({2} bytes) in {3:N1}s @ {4:N1} kB/sec" -f ($useResumable ? 'resumable' : 'simple'), $relativePath, $file.Length, $result.DurationSec, $result.RateKBps)
                } # if
                else
                {
                    $result.Status = 'WhatIf'
                } # else
            } # try
            catch
            {
                $result.Status = 'Incomplete'
                $result.Error  = $_.Exception.Message
                Write-LogLine -Level 'ERROR' -Message ("Failed to upload file '{0}': {1}" -f $relativePath, $result.Error)
            } # catch

            $result
        } # foreach
    } # function Invoke-OneDriveUpload

    # -- Assert-GraphAssemblyCompatibility ------------------------------------------
    # Checks for a known conflict:  PnP.PowerShell loads an older version of
    # Microsoft.Graph.Core (1.x) which is incompatible with the Graph SDK v2 (3.x).
    # If both are loaded in the same session, Graph calls will fail with cryptic errors.
    # Solution: start a fresh pwsh session without PnP.PowerShell loaded.
    function Assert-GraphAssemblyCompatibility
    {
        [CmdletBinding()]
        param()

        $loadedPnp = Get-Module -Name 'PnP.PowerShell' -ErrorAction SilentlyContinue
        if ($null -ne $loadedPnp)
        {
            throw (
                "PnP.PowerShell is loaded in this session and can load Microsoft.Graph.Core 1.x, which conflicts with Microsoft Graph PowerShell SDK v2. " +
                "Start a new pwsh session (recommended) or run: Remove-Module PnP.PowerShell -Force, then re-run this script."
            )
        } # if

        $graphCoreAssembly = [AppDomain]::CurrentDomain.GetAssemblies() |
            Where-Object { $_.GetName().Name -eq 'Microsoft.Graph.Core' } |
            Select-Object -First 1

        if ($null -ne $graphCoreAssembly)
        {
            $loadedVersion = $graphCoreAssembly.GetName().Version
            if ($loadedVersion.Major -lt 3)
            {
                throw (
                    "Incompatible Microsoft.Graph.Core assembly already loaded in this session: $loadedVersion. " +
                    "This usually happens after importing PnP.PowerShell. Start a new pwsh session and run this script before importing PnP modules."
                )
            } # if
        } # if
    } # function Assert-GraphAssemblyCompatibility

    # -- Assert-GraphPermissions -------------------------------------------------
    # Verifies that the app registration has the required Microsoft Graph
    # application permissions (admin-consented).  Throws if any are missing.
    # Reuses the same logic as Test-AzureAppRegistration.ps1.
    #
    # https://learn.microsoft.com/powershell/module/microsoft.graph.applications/get-mgserviceprincipal
    # https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignments
    function Assert-GraphPermissions
    {
        [CmdletBinding()]
        param()

        $requiredPermissions = @('Files.ReadWrite.All', 'User.Read.All')

        Write-LogLine -Message "Validating app registration permissions..."

        # Resolve the app's service principal.
        $appSp = $null
        try
        {
            $appSp = Get-MgServicePrincipal -Filter "appId eq '$ClientId'" -ErrorAction Stop
        } # try
        catch
        {
            Write-LogLine -Level 'WARN' -Message ("Could not look up service principal for ClientId '{0}': {1}. Skipping permission check." -f $ClientId, $_.Exception.Message)
            return 'Skipped'
        } # catch

        if ($null -eq $appSp)
        {
            Write-LogLine -Level 'WARN' -Message ("Service principal not found for ClientId '{0}'. Skipping permission check." -f $ClientId)
            return 'Skipped'
        } # if

        # Resolve the Microsoft Graph service principal.
        $graphSp = $null
        try
        {
            $graphSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'" -ErrorAction Stop
        } # try
        catch
        {
            Write-LogLine -Level 'WARN' -Message ("Could not look up Microsoft Graph service principal: {0}. Skipping permission check." -f $_.Exception.Message)
            return 'Skipped'
        } # catch

        if ($null -eq $graphSp)
        {
            Write-LogLine -Level 'WARN' -Message "Microsoft Graph service principal not found. Skipping permission check."
            return 'Skipped'
        } # if

        # Get granted app role assignments.
        $appRoleAssignments = @()
        try
        {
            $appRoleAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.Id -ErrorAction Stop)
        } # try
        catch
        {
            Write-LogLine -Level 'WARN' -Message ("Could not retrieve app role assignments: {0}. Skipping permission check." -f $_.Exception.Message)
            return 'Skipped'
        } # catch

        $grantedRoleIds = @{}
        foreach ($assignment in $appRoleAssignments)
        {
            $grantedRoleIds[$assignment.AppRoleId] = $true
        } # foreach

        # Check each required permission.
        $missingPermissions = @()
        foreach ($permName in $requiredPermissions)
        {
            $roleDef = $graphSp.AppRoles | Where-Object { $_.Value -eq $permName } | Select-Object -First 1
            if ($null -eq $roleDef -or -not $grantedRoleIds.ContainsKey($roleDef.Id))
            {
                $missingPermissions += $permName
            } # if
        } # foreach

        if ($missingPermissions.Count -gt 0)
        {
            throw (
                ("The app registration '{0}' (ClientId={1}) is missing required Graph application permissions: {2}. " +
                "Grant these permissions with admin consent in Azure Portal > App registrations > API permissions.") -f
                $appSp.DisplayName, $ClientId, ($missingPermissions -join ', ')
            )
        } # if

        Write-LogLine -Message ("All required Graph permissions verified for app '{0}'." -f $appSp.DisplayName)
        return 'Verified'
    } # function Assert-GraphPermissions

    # -- Get-AppPermissionDetail --------------------------------------------------
    # Resolves the app's service principal in the tenant and checks each required
    # Graph application permission individually.  Returns an array of result objects
    # with per-permission grant status - the same detail as Test-AzureAppRegistration.ps1.
    #
    # Used by -Test mode to provide granular permission reporting.
    # https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignments
    function Get-AppPermissionDetail
    {
        [CmdletBinding()]
        param()

        $requiredPermissions = @('Files.ReadWrite.All', 'User.Read.All')
        $results = [System.Collections.Generic.List[pscustomobject]]::new()

        # Resolve the app's service principal.
        Write-LogLine -Message ("Looking up service principal for ClientId={0}" -f $ClientId)
        $appSp = Get-MgServicePrincipal -Filter "appId eq '$ClientId'" -ErrorAction Stop
        if ($null -eq $appSp)
        {
            throw ("Service principal not found for ClientId '{0}'. Ensure the app registration exists in tenant '{1}'." -f $ClientId, $TenantId)
        } # if

        # Resolve the Microsoft Graph service principal.
        Write-LogLine -Message "Looking up Microsoft Graph service principal."
        $graphSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'" -ErrorAction Stop
        if ($null -eq $graphSp)
        {
            throw "Could not find the 'Microsoft Graph' service principal in the tenant."
        } # if

        # Get granted app role assignments.
        Write-LogLine -Message "Retrieving app role assignments for the service principal."
        $appRoleAssignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.Id -ErrorAction Stop)

        $grantedRoleIds = @{}
        foreach ($assignment in $appRoleAssignments)
        {
            $grantedRoleIds[$assignment.AppRoleId] = $assignment
        } # foreach

        # Check each required permission.
        foreach ($permName in $requiredPermissions)
        {
            $roleDef    = $graphSp.AppRoles | Where-Object { $_.Value -eq $permName } | Select-Object -First 1
            $assignment = $null
            $isGranted  = $false

            if ($null -ne $roleDef -and $grantedRoleIds.ContainsKey($roleDef.Id))
            {
                $isGranted  = $true
                $assignment = $grantedRoleIds[$roleDef.Id]
            } # if

            $results.Add([pscustomobject]@{
                Permission  = $permName
                Type        = 'Application'
                IsGranted   = $isGranted
                RoleId      = if ($null -ne $roleDef) { $roleDef.Id } else { $null }
                GrantedOn   = if ($null -ne $assignment) { $assignment.CreatedDateTime } else { $null }
                AppId       = $ClientId
                TenantId    = $TenantId
                DisplayName = $appSp.DisplayName
            })
        } # foreach

        return $results
    } # function Get-AppPermissionDetail

    # -- Get-ValidatedUserDrive ----------------------------------------------------
    # Looks up a user's OneDrive drive via Microsoft Graph, validates the response,
    # and confirms the drive root is accessible.  Returns a custom object with the
    # drive info plus account/provisioning status flags.
    #
    # Steps:
    #   1. Verify the user account exists and is enabled via Get-MgUser.
    #   2. Resolve the user's OneDrive drive via Get-MgUserDrive.
    #   3. Validate the drive root is accessible.
    #
    # Uses Get-MgUser from the Microsoft.Graph.Users module:
    #   https://learn.microsoft.com/powershell/module/microsoft.graph.users/get-mguser
    # Uses Get-MgUserDrive from the Microsoft.Graph.Files module:
    #   https://learn.microsoft.com/powershell/module/microsoft.graph.files/get-mguserdrive
    #
    # If the user's OneDrive has not been provisioned yet (first-time user), this
    # will throw an error.  Provision it by visiting https://portal.office.com.
    function Get-ValidatedUserDrive
    {
        [CmdletBinding()]
        param
        (
            [Parameter(Mandatory = $true)]
            [string] $UserPrincipalName
        )

        # -- Step 1: Validate the user account --------------------------------
        # https://learn.microsoft.com/powershell/module/microsoft.graph.users/get-mguser
        Write-LogLine -Message ("Validating user account: {0}" -f $UserPrincipalName)
        $userAccount = $null
        try
        {
            $userAccount = Invoke-WithGraphRetry -OperationName ("Get-MgUser for '{0}'" -f $UserPrincipalName) -Operation {
                Get-MgUser -UserId $UserPrincipalName -Property Id, DisplayName, UserPrincipalName, AccountEnabled -ErrorAction Stop
            } # inline:$userAccount = Invoke-WithGraphRetry
        } # try
        catch
        {
            $errMsg = $_.Exception.Message
            # Distinguish permission errors from user-not-found errors.
            if ($errMsg -match 'Authorization_RequestDenied|Insufficient privileges|Forbidden')
            {
                throw (
                    ("The app registration lacks permission to read user '{0}'. " +
                    "Grant 'User.Read.All' (Application) with admin consent in Azure Portal > App registrations. " +
                    "Original error: {1}") -f $UserPrincipalName, $errMsg
                )
            } # if

            throw (
                ("User account '{0}' could not be found in the tenant. " +
                "Verify the UPN is correct and the account exists in Microsoft Entra ID. " +
                "Original error: {1}") -f $UserPrincipalName, $errMsg
            )
        } # catch

        if ($null -eq $userAccount)
        {
            throw ("Get-MgUser returned null for '{0}'." -f $UserPrincipalName)
        } # if

        if ($userAccount.AccountEnabled -eq $false)
        {
            throw (
                ("User account '{0}' (DisplayName='{1}') is disabled in Microsoft Entra ID. " +
                "Enable the account before running the migration.") -f $UserPrincipalName, $userAccount.DisplayName
            )
        } # if

        Write-LogLine -Message ("User account validated: {0} (DisplayName='{1}', Enabled={2})" -f
            $UserPrincipalName, $userAccount.DisplayName, $userAccount.AccountEnabled)

        # -- Step 2: Resolve the user's OneDrive drive ------------------------
        # Get-MgUserDrive can return multiple drives (e.g. OneDrive and
        # PersonalCacheLibrary).  We fetch all drives and filter to the one
        # named 'OneDrive' so we always target the correct drive.
        # https://learn.microsoft.com/powershell/module/microsoft.graph.files/get-mguserdrive
        Write-LogLine -Message ("Resolving OneDrive drive for owner: {0}" -f $UserPrincipalName)
        $userDrive = $null
        try
        {
            $allDrives = Invoke-WithGraphRetry -OperationName ("Get-MgUserDrive for '{0}'" -f $UserPrincipalName) -Operation {
                Get-MgUserDrive -UserId $UserPrincipalName -All -ErrorAction Stop
            } # inline:$allDrives = Invoke-WithGraphRetry -Oper

            # Filter to the OneDrive drive by Name.  Other drives such as
            # PersonalCacheLibrary share the same DriveType but are not the
            # user's primary document store.
            $allDrives = @($allDrives)
            if ($allDrives.Count -gt 1)
            {
                Write-LogLine -Message ("Multiple drives returned ({0}) for '{1}': {2}. Filtering to Name='OneDrive'." -f
                    $allDrives.Count, $UserPrincipalName, (($allDrives | ForEach-Object { $_.Name }) -join ', '))
            } # if

            $userDrive = $allDrives | Where-Object { $_.Name -eq 'OneDrive' } | Select-Object -First 1

            # Fallback: if no drive is named 'OneDrive' (older tenants or
            # single-drive result), use the first drive returned.
            if ($null -eq $userDrive -and $allDrives.Count -eq 1)
            {
                $userDrive = $allDrives[0]
                Write-LogLine -Level 'WARN' -Message ("No drive named 'OneDrive' found for '{0}'. Using the only drive returned: Name='{1}', Id='{2}'." -f
                    $UserPrincipalName, $userDrive.Name, $userDrive.Id)
            } # if
        } # try
        catch
        {
            $errMsg = $_.Exception.Message
            # Detect provisioning-related errors and provide actionable guidance.
            # Graph may return 404/ResourceNotFound when the drive doesn't exist,
            # or accessDenied when the OneDrive site collection is not provisioned.
            if ($errMsg -match '404|ResourceNotFound|not found|does not exist|no OneDrive|accessDenied|access denied')
            {
                throw (
                    ("OneDrive is not provisioned for user '{0}'. " +
                    "The user account exists but their OneDrive has not been created. " +
                    "Provision via: Request-SPOPersonalSite -UserEmails '{0}' " +
                    "or have the user sign in at https://portal.office.com. " +
                    "Original error: {1}") -f $UserPrincipalName, $errMsg
                )
            } # if
            throw
        } # catch

        if ($null -eq $userDrive -or [string]::IsNullOrWhiteSpace([string] $userDrive.Id))
        {
            throw ("No OneDrive drive was returned for user '{0}'." -f $UserPrincipalName)
        } # if

        if ([string]::IsNullOrWhiteSpace([string] $userDrive.WebUrl))
        {
            throw (
                "OneDrive WebUrl is empty for user '{0}'. The OneDrive site may not be provisioned yet." -f $UserPrincipalName
            )
        } # if

        $parsedOneDriveUrl = $null
        $isValidWebUrl = [System.Uri]::TryCreate(
            [string] $userDrive.WebUrl,
            [System.UriKind]::Absolute,
            [ref] $parsedOneDriveUrl
        )

        if (-not $isValidWebUrl)
        {
            throw (
                "OneDrive WebUrl is not a valid absolute URL for user '{0}': {1}" -f $UserPrincipalName, $userDrive.WebUrl
            )
        } # if

        # -- Step 3: Validate the drive root is accessible --------------------
        $rootCheckUri = "https://graph.microsoft.com/v1.0/drives/$($userDrive.Id)/root?`$select=id,webUrl"
        $driveRoot = Invoke-WithGraphRetry -OperationName ("Resolve drive root for '{0}'" -f $UserPrincipalName) -Operation {
            Invoke-MgGraphRequest -Method GET -Uri $rootCheckUri -ErrorAction Stop
        } # inline:$driveRoot = Invoke-WithGraphRetry -Oper
        if ($null -eq $driveRoot -or [string]::IsNullOrWhiteSpace([string] $driveRoot.id))
        {
            throw (
                "Could not resolve OneDrive root item for user '{0}' (DriveId={1})." -f $UserPrincipalName, $userDrive.Id
            )
        } # if

        Write-LogLine -Message ("Verified OneDrive WebUrl for '{0}': {1}" -f $UserPrincipalName, $userDrive.WebUrl)
        return $userDrive
    } # function Get-ValidatedUserDrive

    # -- Initialization (runs once at script start) -------------------------------
    # Set up logging, check for assembly conflicts, import modules, and authenticate.
    # $script: scope means the variable is visible across begin/process/end blocks.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_scopes
    $script:LogFilePath = $null
    try
    {
        # Test-Path checks whether a file or folder exists.
        if (-not (Test-Path -LiteralPath $LogFolder))
        {
            # New-Item -ItemType Directory creates the folder (like mkdir).
            New-Item -Path $LogFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } # if

        # Generate a unique log filename with a timestamp.
        $token = (Get-Date).ToString('yyyyMMdd_HHmmss_fff')
        $script:LogFilePath = Join-Path -Path $LogFolder -ChildPath ("Update-UserFile_{0}.log" -f $token)
    } # try
    catch
    {
        Write-Warning "Logging setup failed: $($_.Exception.Message)"
    } # catch

    # -- Run validation steps, collecting per-step results for -Test mode ----
    # Each step is tracked with a status (Passed / Failed / Skipped) so that
    # -Test can emit granular output showing exactly what was checked.
    # In normal (non-Test) mode, failures still throw as before.
    $script:TestStepResults = [System.Collections.Generic.List[pscustomobject]]::new()

    # Step 1: Assembly compatibility
    try
    {
        Assert-GraphAssemblyCompatibility
        $script:TestStepResults.Add([pscustomobject]@{ Step = 'Assembly compatibility'; Status = 'Passed'; Detail = $null })
    } # try
    catch
    {
        $script:TestStepResults.Add([pscustomobject]@{ Step = 'Assembly compatibility'; Status = 'Failed'; Detail = $_.Exception.Message })
        throw
    } # catch

    # Step 2: Required modules
    try
    {
        Assert-RequiredModules
        $script:TestStepResults.Add([pscustomobject]@{ Step = 'Required modules'; Status = 'Passed'; Detail = $null })
    } # try
    catch
    {
        $script:TestStepResults.Add([pscustomobject]@{ Step = 'Required modules'; Status = 'Failed'; Detail = $_.Exception.Message })
        throw
    } # catch

    # Step 3: Certificate authentication
    try
    {
        Connect-GraphCertAuth
        $script:TestStepResults.Add([pscustomobject]@{ Step = 'Certificate authentication'; Status = 'Passed'; Detail = $null })
    } # try
    catch
    {
        $script:TestStepResults.Add([pscustomobject]@{ Step = 'Certificate authentication'; Status = 'Failed'; Detail = $_.Exception.Message })
        throw
    } # catch

    # Step 4: Permission validation
    # Assert-GraphPermissions returns 'Verified' on success, 'Skipped' when it
    # cannot query the service principal (e.g. missing Application.Read.All),
    # and throws when permissions are confirmed missing.
    try
    {
        $permResult = Assert-GraphPermissions
        if ($permResult -eq 'Skipped')
        {
            $script:TestStepResults.Add([pscustomobject]@{ Step = 'Permission validation'; Status = 'Skipped'; Detail = 'Could not verify permissions. Grant Application.Read.All to enable this check.' })
        } # if
        else
        {
            $script:TestStepResults.Add([pscustomobject]@{ Step = 'Permission validation'; Status = 'Passed'; Detail = $null })
        } # else
    } # try
    catch
    {
        $script:TestStepResults.Add([pscustomobject]@{ Step = 'Permission validation'; Status = 'Failed'; Detail = $_.Exception.Message })
        throw
    } # catch

    # Step 5: Per-permission detail (same as Test-AzureAppRegistration.ps1)
    # Only runs in -Test mode and only when we can query the service principal.
    $script:PermissionDetails = $null
    if ($Test)
    {
        try
        {
            $script:PermissionDetails = Get-AppPermissionDetail
            $allGranted = ($script:PermissionDetails | Where-Object { -not $_.IsGranted }).Count -eq 0
            if ($allGranted)
            {
                $script:TestStepResults.Add([pscustomobject]@{ Step = 'Permission detail'; Status = 'Passed'; Detail = 'All required permissions are granted.' })
            } # if
            else
            {
                $missing = ($script:PermissionDetails | Where-Object { -not $_.IsGranted }).Permission -join ', '
                $script:TestStepResults.Add([pscustomobject]@{ Step = 'Permission detail'; Status = 'Failed'; Detail = "Missing permissions: $missing" })
            } # else
        } # try
        catch
        {
            $script:TestStepResults.Add([pscustomobject]@{ Step = 'Permission detail'; Status = 'Skipped'; Detail = $_.Exception.Message })
        } # catch
    } # if

    # -- Test mode: verify auth and access, then exit -------------------------
    # When -Test is specified, the script validates that authentication and
    # permissions are in order but does not process any CSV data.
    if ($Test)
    {
        Write-LogLine -Message "Test mode: authentication and access verification completed."
        $script:TestModeActive = $true
    } # if
    else
    {
        $script:TestModeActive = $false
    } # else

    # Cache the CSV data once in the begin block so piping multiple UPNs does
    # not re-read and re-parse the file for each pipeline input.
    $script:CachedCsvRows = $null
    if (-not $script:TestModeActive -and $InputFile.Exists)
    {
        $script:CachedCsvRows = Import-Csv -LiteralPath $InputFile.FullName

        # Validate that all required columns exist in the CSV before processing.
        Assert-CsvColumns -CsvRows $script:CachedCsvRows
    } # if
} # begin

# #===============================================================================#
# #  PROCESS BLOCK                                                               #
# #  Runs once for each pipeline input object ($UserToProcess).                  #
# #  If not piped and $UserToProcess is empty, processes all unique owners       #
# #  found in the CSV.                                                           #
# #  This is where the main work happens: read CSV, upload files, grant perms.   #
# #  https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods #
# #===============================================================================#
process
{
    # -- Help mode: skip all processing ----------------------------------------
    if ($script:HelpModeActive) { return }

    # -- Test mode: skip all CSV processing ------------------------------------
    if ($script:TestModeActive)
    {
        # Determine the overall status from the per-step results.
        # If any step was skipped, the overall status reflects that.
        $hasSkipped = $script:TestStepResults | Where-Object { $_.Status -eq 'Skipped' }
        $overallStatus = if ($hasSkipped) { 'Passed with warnings' } else { 'Passed' }

        foreach ($stepResult in $script:TestStepResults)
        {
            [pscustomobject]@{
                Test   = $true
                Step   = $stepResult.Step
                Status = $stepResult.Status
                Detail = $stepResult.Detail
            }
        } # foreach - step result

        # Emit per-permission detail (same data as Test-AzureAppRegistration.ps1).
        if ($null -ne $script:PermissionDetails)
        {
            foreach ($perm in $script:PermissionDetails)
            {
                $perm
            } # foreach - permission detail
        } # if

        [pscustomobject]@{
            Test   = $true
            Step   = 'Overall'
            Status = $overallStatus
            Detail = '{0} step(s) checked.' -f $script:TestStepResults.Count
        }
        return
    } # if

    if (-not $InputFile.Exists)
    {
        throw "InputFile not found: $($InputFile.FullName)"
    } # if

    # Use the CSV data cached in the begin block to avoid re-reading per pipeline input.
    $allRows = $script:CachedCsvRows

    # Filter to only the rows belonging to this user when specified.
    # Where-Object filters objects in the pipeline based on a condition.
    # $_ represents the current object in the pipeline.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/where-object
    if (-not [string]::IsNullOrWhiteSpace($UserToProcess))
    {
        $allRows = $allRows | Where-Object { $_.'Owner Login' -eq $UserToProcess }
    } # if

    if (-not $allRows)
    {
        Write-LogLine -Level 'WARN' -Message "No CSV rows found to process."
        if (-not $UploadFiles)
        {
            return
        } # if
    } # if

    # Get unique owner UPNs from the CSV rows.
    # Group-Object builds a hashtable in a single pass - O(n) instead of O(n^2)
    # from repeated Where-Object calls per owner.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.utility/group-object
    $ownerGroups = $allRows |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.'Owner Login') } |
        Group-Object -Property 'Owner Login' -AsHashTable -AsString

    if ($null -eq $ownerGroups) { $ownerGroups = @{} }

    # Start with owners that appear in the CSV.
    $uniqueOwners = [System.Collections.Generic.List[string]]::new()
    foreach ($k in $ownerGroups.Keys) { [void] $uniqueOwners.Add($k) }

    # When -UploadFiles is set, also include any UPN-named subdirectory under
    # AllFilesDirectory so files/folders are uploaded even for users that have
    # no permission rows in the CSV.
    if ($UploadFiles -and -not [string]::IsNullOrWhiteSpace($AllFilesDirectory) -and (Test-Path -LiteralPath $AllFilesDirectory))
    {
        $localDirs = Get-ChildItem -LiteralPath $AllFilesDirectory -Directory -ErrorAction SilentlyContinue
        foreach ($d in $localDirs)
        {
            if (-not [string]::IsNullOrWhiteSpace($UserToProcess) -and $d.Name -ne $UserToProcess) { continue }
            if (-not ($uniqueOwners -contains $d.Name))
            {
                [void] $uniqueOwners.Add($d.Name)
                Write-LogLine -Message ("Including upload-only owner from local directory: {0}" -f $d.Name)
            } # if
        } # foreach
    } # if

    if ($uniqueOwners.Count -eq 0)
    {
        throw "No owners to process: CSV has no 'Owner Login' values and no local upload folders were found."
    } # if

    Write-LogLine -Message ("Processing {0} unique owner(s): {1}" -f $uniqueOwners.Count, ($uniqueOwners -join ', '))

    # -- Iterate over each unique owner ---------------------------------------
    foreach ($ownerUpn in $uniqueOwners)
    {
        Write-LogLine -Message ("-- Begin processing owner: {0} --" -f $ownerUpn)

        # Retrieve pre-grouped rows for this owner and deduplicate.
        # Duplicate rows (same Owner + Path + Item Name + Collaborator Login)
        # waste API calls; the invite API is idempotent but we skip duplicates
        # to reduce noise and throttling risk.
        # $ownerRows is empty when this owner has no CSV rows (upload-only mode).
        $ownerRows = if ($ownerGroups.ContainsKey($ownerUpn)) { @($ownerGroups[$ownerUpn]) } else { @() }
        $deduplicatedRows = @($ownerRows |
            Sort-Object -Property 'Path', 'Item Name', 'Collaborator Login', 'Collaborator Permission' -Unique)
        $duplicateCount = $ownerRows.Count - $deduplicatedRows.Count
        if ($duplicateCount -gt 0)
        {
            Write-LogLine -Level 'WARN' -Message ("Removed {0} duplicate CSV row(s) for owner '{1}'." -f $duplicateCount, $ownerUpn)
        } # if
        $rows = $deduplicatedRows

        # Look up and validate the user's OneDrive drive.
        # Wrapped in try/catch so one failing user does not stop others.
        try
        {
            $drive = Get-ValidatedUserDrive -UserPrincipalName $ownerUpn
        } # try
        catch
        {
            Write-LogLine -Level 'ERROR' -Message ("Failed to resolve OneDrive for '{0}': {1}. Skipping this owner." -f $ownerUpn, $_.Exception.Message)

            # Determine which validation step failed for the status flags.
            $errText = $_.Exception.Message
            $isValid  = -not ($errText -match 'could not be found in the tenant|Get-MgUser returned null|is disabled|lacks permission to read user')
            $isDriveOk = $false

            foreach ($row in $rows)
            {
                [pscustomobject]@{
                    OwnerLogin             = $ownerUpn
                    ItemName               = $row.'Item Name'
                    Path                   = $row.Path
                    NormalizedPath         = $null
                    CollaboratorLogin      = $row.'Collaborator Login'
                    CollaboratorPermission = $row.'Collaborator Permission'
                    GraphRole              = $null
                    DriveId                = $null
                    OneDriveWebUrl         = $null
                    IsValidAccount         = $isValid
                    OneDriveProvisioned    = $isDriveOk
                    ExistsInOneDrive       = $null
                    DriveItemId            = $null
                    Action                 = $null
                    Status                 = 'Failed'
                    Error                  = $errText
                } # inline:[pscustomobject]@{
            } # foreach
            continue
        } # catch

        # -- Upload local files if -UploadFiles is specified ----------------------
        # The local folder must be named by the user's UPN under AllFilesDirectory.
        # Example: C:\Repos\NPSBox\LocalFiles\user@contoso.com\
        # Uploads run BEFORE permissions so that files/folders referenced by the
        # CSV exist in OneDrive when the invite calls are made. Owners that have
        # only a local folder (no CSV rows) are still uploaded for.
        if ($UploadFiles)
        {
            $userLocalPath = Join-Path -Path $AllFilesDirectory -ChildPath $ownerUpn
            if (Test-Path -LiteralPath $userLocalPath)
            {
                Write-LogLine -Message ("Uploading local files from '{0}' to OneDrive for '{1}'." -f $userLocalPath, $ownerUpn)
                Invoke-OneDriveUpload -DriveId $drive.Id -LocalSourcePath $userLocalPath -OwnerUpn $ownerUpn
            } # if
            else
            {
                Write-LogLine -Level 'WARN' -Message ("No local upload folder found for '{0}' at '{1}'. Skipping upload." -f $ownerUpn, $userLocalPath)
            } # else
        } # if

        # -- Process each CSV row for this owner: resolve item, grant permission --
        if ($rows.Count -eq 0)
        {
            Write-LogLine -Message ("No CSV permission rows for owner '{0}'; upload-only." -f $ownerUpn)
            continue
        } # if
        foreach ($row in $rows)
        {
            $itemPath = [string] $row.Path
            $itemName = [string] $row.'Item Name'
            $collab   = [string] $row.'Collaborator Login'
            $boxPerm  = [string] $row.'Collaborator Permission'

            # -- Unified empty-cell validation (D6) ---------------------------
            # Check all required fields at once before processing.  This gives
            # a consistent error for any empty/whitespace cell rather than
            # failing at different points with different messages.
            $emptyFields = @()
            if ([string]::IsNullOrWhiteSpace($itemPath))  { $emptyFields += 'Path' }
            if ([string]::IsNullOrWhiteSpace($itemName))  { $emptyFields += 'Item Name' }
            if ([string]::IsNullOrWhiteSpace($collab))    { $emptyFields += 'Collaborator Login' }
            if ([string]::IsNullOrWhiteSpace($boxPerm))   { $emptyFields += 'Collaborator Permission' }

            # Defer ConvertTo-GraphRole until after empty-cell validation so
            # that an empty Collaborator Permission gets a unified error.
            $graphRole = if ([string]::IsNullOrWhiteSpace($boxPerm)) { $null } else { ConvertTo-GraphRole -BoxPermission $boxPerm }

            # Create a result object to track what happens with this row.
            # [pscustomobject] is a lightweight object with named properties.
            # This object is output to the pipeline so callers can inspect results.
            # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_pscustomobject
            $result = [pscustomobject]@{
                OwnerLogin             = $ownerUpn
                ItemName               = $itemName
                Path                   = $itemPath
                NormalizedPath         = $null
                CollaboratorLogin      = $collab
                CollaboratorPermission = $boxPerm
                GraphRole              = $graphRole
                DriveId                = $drive.Id
                OneDriveWebUrl         = $drive.WebUrl
                IsValidAccount         = $true
                OneDriveProvisioned    = $true
                ExistsInOneDrive       = $null
                DriveItemId            = $null
                Action                 = $null
                Status                 = 'Unknown'
                Error                  = $null
            } # inline:$result = [pscustomobject]@{
            try
            {
                if ($emptyFields.Count -gt 0)
                {
                    throw ("Required CSV field(s) empty: {0}." -f ($emptyFields -join ', '))
                } # if

                # Validate basic email format before domain check (S6).
                if (-not (Test-EmailFormat -Email $collab))
                {
                    $result.Action = 'Skipped'
                    $result.Status = 'Skipped'
                    $result.Error  = "Collaborator Login '{0}' is not a valid email format." -f $collab
                    Write-LogLine -Level 'WARN' -Message $result.Error
                    $result
                    continue
                } # if

                # Validate that the collaborator's domain is permitted.
                if (-not (Test-CollaboratorDomain -Email $collab -Domains $AllowedDomains))
                {
                    $result.Action = 'Skipped'
                    $result.Status = 'Skipped'
                    $result.Error  = "Domain not in AllowedDomains list."
                    Write-LogLine -Level 'WARN' -Message ("Skipping external collaborator '{0}' - domain not in AllowedDomains." -f $collab)
                    $result
                    continue
                } # if

                if ([string]::IsNullOrWhiteSpace($graphRole))
                {
                    $result.Action = 'Skipped'
                    $result.Status = 'Skipped'
                    Write-LogLine -Message ("Skipping (role maps to None): Item='{0}', Collaborator='{1}', BoxPerm='{2}'" -f $itemName, $collab, $boxPerm)
                    $result
                    continue
                } # if

                # Clean up the Box path for use with the Graph API.
                $normalizedPath = ConvertTo-OneDriveRelativePath -Path $itemPath
                $result.NormalizedPath = $normalizedPath

                if (-not [string]::IsNullOrWhiteSpace([string] $drive.WebUrl))
                {
                    Write-LogLine -Message ("Resolving drive item at: {0}/{1}" -f $drive.WebUrl.TrimEnd('/'), $normalizedPath)
                } # if

                # URL-encode the path and look up the item in OneDrive.
                # The /root:/{path} syntax accesses a drive item by its path:
                # https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path
                $encodedPath = ConvertTo-GraphEncodedPath -RelativePath $normalizedPath
                $getItemUri = "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/root:/$encodedPath"
                $driveItem = Invoke-WithGraphRetry -OperationName ("Resolve drive item '{0}'" -f $normalizedPath) -Operation {
                    # Invoke-MgGraphRequest is a generic Graph API caller.
                    # It handles auth headers automatically.
                    # https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/invoke-mggraphrequest
                    Invoke-MgGraphRequest -Method GET -Uri $getItemUri -ErrorAction Stop
                } # inline:$driveItem = Invoke-WithGraphRetry -Oper

                $result.DriveItemId = $driveItem.id
                $result.ExistsInOneDrive = $true

                # ShouldProcess enables -WhatIf and -Confirm support.
                # When -WhatIf is used, it prints what WOULD happen and returns $false.
                # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods#shouldprocess
                $target = "DriveItemId=$($driveItem.id) Path='$itemPath' -> grant '$collab' Role='$graphRole'"
                if ($PSCmdlet.ShouldProcess($target, "Invite collaborator via Microsoft Graph (silent grant)"))
                {
                    # -- Grant permission using the driveItem: invite API ---------
                    # POST /drives/{driveId}/items/{itemId}/invite
                    # This creates a sharing permission on the item.
                    #
                    # Key body properties:
                    #   recipients     : array of { email } objects - who to share with
                    #   roles          : 'read' or 'write'
                    #   requireSignIn  : recipient must sign in to access
                    #   sendInvitation : false = NO EMAIL is sent; permission is granted silently
                    #
                    # https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0
                    $inviteUri = "https://graph.microsoft.com/v1.0/drives/$($drive.Id)/items/$($driveItem.id)/invite"

                    $body = @{
                        recipients      = @(@{ email = $collab })
                        roles           = @($graphRole)     # 'read' or 'write'
                        requireSignIn   = $true              # recipient must authenticate
                        sendInvitation  = $false             # NO email notification sent
                    } | ConvertTo-Json -Depth 6

                    $inviteResponse = Invoke-WithGraphRetry -OperationName ("Invite '{0}' on '{1}'" -f $collab, $normalizedPath) -Operation {
                        Invoke-MgGraphRequest -Method POST -Uri $inviteUri -Body $body -ContentType 'application/json' -ErrorAction Stop
                    } # inline:$inviteResponse = Invoke-WithGraphRetry 

                    # -- Validate the invite response ----------------------------
                    # The API returns { value: [ { id, roles, ... } ] }.
                    # A 207 Multi-Status can include per-recipient errors.
                    $grantedPermissions = $inviteResponse.value
                    if ($null -eq $grantedPermissions -or $grantedPermissions.Count -eq 0)
                    {
                        throw ("Invite API returned no permissions for collaborator '{0}' on item '{1}'." -f $collab, $normalizedPath)
                    } # if

                    $grantedEntry = $grantedPermissions | Select-Object -First 1
                    $grantedRoles = $grantedEntry.roles -join ', '

                    # Check for per-recipient errors (207 partial success).
                    if ($null -ne $grantedEntry.error)
                    {
                        $errCode = $grantedEntry.error.code
                        $errMsg  = $grantedEntry.error.message
                        throw ("Invite failed for '{0}': [{1}] {2}" -f $collab, $errCode, $errMsg)
                    } # if

                    $result.Action = 'Invited'
                    $result.Status = 'Applied'
                    Write-LogLine -Message ("Granted '{0}' roles=[{1}] on '{2}' (PermissionId={3}, sendInvitation=false)" -f
                        $collab, $grantedRoles, $itemPath, $grantedEntry.id)
                } # if
                else
                {
                    $result.Action = 'Invited'
                    $result.Status = 'WhatIf'
                } # else
            } # try
            catch
            {
                $result.Status = 'Failed'
                $result.Error  = $_.Exception.Message

                if ([object]::ReferenceEquals($result.ExistsInOneDrive, $null) -and $result.Error -match '404|itemNotFound|not found')
                {
                    $result.ExistsInOneDrive = $false
                } # if

                Write-LogLine -Level 'ERROR' -Message ("Failed row: Item='{0}', Path='{1}', Collaborator='{2}'. Error={3}" -f $itemName, $itemPath, $collab, $result.Error)
            } # catch

            $result
        } # foreach - row

        Write-LogLine -Message ("-- Finished processing owner: {0} --" -f $ownerUpn)
    } # foreach - owner
} # process

# #===============================================================================#
# #  END BLOCK                                                                   #
# #  Runs once after all pipeline input has been processed.                      #
# #  Used here to disconnect from Microsoft Graph and clean up the session.      #
# #===============================================================================#
end
{
    if ($script:HelpModeActive) { return }

    try
    {
        # Disconnect-MgGraph signs out of Microsoft Graph.
        # https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/disconnect-mggraph
        Disconnect-MgGraph | Out-Null
    } # try
    catch
    {
        # Non-fatal - the session will be cleaned up when PowerShell exits anyway.
    } # catch
} # end
