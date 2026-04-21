# The #Requires statement prevents the script from running unless the specified
# version of PowerShell is available.  PowerShell 7+ is required for features
# like ternary operators and improved module handling.
# https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_requires
#Requires -Version 7.0

<#
.SYNOPSIS
    Applies OneDrive item sharing permissions based on a CSV file using Microsoft Graph.

    Version: 1.2.0.0
    Date:    2026-04-20

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
          Files.ReadWrite.All
        https://learn.microsoft.com/entra/identity-platform/quickstart-register-app
      - A certificate uploaded to the app registration
        https://learn.microsoft.com/entra/identity-platform/certificate-credentials

    SAFETY:
      -WhatIf   : Shows what would happen without making changes.
      -Verbose   : Shows detailed progress messages.
      -Confirm   : Prompts for confirmation before each change.
      https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_commonparameters

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
    Switch parameter (no value needed — just include it or omit it).
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
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param
(
    # Path to the input CSV file.  [System.IO.FileInfo] automatically resolves
    # the string to a file object with .Exists, .FullName, etc.
    [Parameter()]
    [System.IO.FileInfo] $InputFile = "C:\Repos\NPSBox\UserInfo.csv"
    ,
    # The owner's UPN to filter on in the CSV.
    # ValueFromPipeline lets you pipe UPNs:  'user1@contoso.com','user2@contoso.com' | .\Update-UserFile.ps1
    # Alias allows matching CSV column names directly for pipeline binding.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters#alias-attribute
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [Alias('Owner Login', 'User', 'UPN', 'Account')]
    [string] $UserToProcess
    ,
    # Your tenant ID (GUID).  Find it: Azure Portal > Entra ID > Overview.
    [Parameter()]
    [string] $TenantId = "92075952-90f3-4613-833b-d2e19ec649e4"
    ,
    # The app registration's client ID (GUID).
    [Parameter()]
    [string] $ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    ,
    # Certificate thumbprint for app-only auth.
    [Parameter(Mandatory = $true)]
    [string] $CertificateThumbprint
    ,
    # Where to write timestamped log files.  Created if it doesn't exist.
    [Parameter()]
    [string] $LogFolder = "C:\Repos\NPSBox\Logs"
    ,
    # Root folder with per-user subfolders of files to upload.
    # Subfolder names must match the user's UPN exactly.
    [Parameter()]
    [string] $AllFilesDirectory = "C:\Repos\NPSBox\LocalFiles"
    ,
    # Include this switch to upload local files to OneDrive before applying permissions.
    # A switch parameter is either present ($true) or absent ($false) — no value needed.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_switch
    [Parameter()]
    [switch] $UploadFiles
)

# ╔═══════════════════════════════════════════════════════════════════════════════╗
# ║  BEGIN BLOCK                                                                 ║
# ║  Runs once before any pipeline input is processed.                           ║
# ║  Used here to define helper functions, import modules, and authenticate.     ║
# ║  https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods ║
# ╚═══════════════════════════════════════════════════════════════════════════════╝
begin
{
    # ── Write-LogLine ────────────────────────────────────────────────────────────
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
        $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffK'), $Level, $Message
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
                } # finally — always restores the original $WhatIfPreference
            } # try
            catch
            {
                # Write-Warning outputs a non-terminating warning that appears in yellow.
                Write-Warning "Failed to write log line: $($_.Exception.Message)"
            } # catch
        } # if
    } # function Write-LogLine

    # ── Assert-RequiredModules ───────────────────────────────────────────────────
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
            'Microsoft.Graph.Files'             # Provides Get-MgUserDrive and drive-related cmdlets
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

    # ── ConvertTo-GraphRole ──────────────────────────────────────────────────────
    # Maps a Box permission name to a Microsoft Graph sharing role.
    # Graph supports two sharing roles for the invite API:
    #   'read'   — view-only access
    #   'write'  — view + edit access
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
            'Previewer Uploader' { return $null   }   # No Graph equivalent — skip
            'Previewer'          { return $null   }   # No Graph equivalent — skip
            'Uploader'           { return $null   }   # No Graph equivalent — skip
            default              { return $null   }   # Unknown — skip
        } # switch
    } # function ConvertTo-GraphRole

    # ── ConvertTo-OneDriveRelativePath ────────────────────────────────────────────
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

    # ── ConvertTo-GraphEncodedPath ────────────────────────────────────────────────
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

    # ── Test-IsRetryableGraphError ────────────────────────────────────────────────
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

    # ── Invoke-WithGraphRetry ────────────────────────────────────────────────────
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
            [int] $MaxAttempts = 4
            ,
            [Parameter()]
            [ValidateRange(1, 60)]
            [int] $InitialDelaySeconds = 2
            ,
            [Parameter()]
            [ValidateRange(1, 120)]
            [int] $MaxDelaySeconds = 20
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

                Start-Sleep -Seconds $delaySeconds
                $attempt += 1
                $delaySeconds = [Math]::Min($delaySeconds * 2, $MaxDelaySeconds)
            } # catch
        } # while
    } # function Invoke-WithGraphRetry

    # ── Connect-Graph ────────────────────────────────────────────────────────────
    # Authenticates to Microsoft Graph using certificate-based app-only auth.
    #
    # Certificate mode uses a certificate for "app-only" auth — no user sign-in
    # is required.  This is how you run the script unattended (e.g. scheduled).
    # The app registration must have Application permissions granted with admin consent.
    #
    # Connect-MgGraph reference:
    #   https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
    # Auth overview:
    #   https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands
    function Connect-Graph
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
            if ($null -ne $existingContext -and $existingContext.TenantId -eq $TenantId)
            {
                Write-LogLine -Message ("Reusing existing Microsoft Graph context (app-only). TenantId={0}, AppName={1}, AuthType={2}" -f
                    $existingContext.TenantId, $existingContext.AppName, $existingContext.AuthType)
                return
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
        } # try
        finally
        {
            $WhatIfPreference = $previousWhatIfPreference
        } # finally
    } # function Connect-Graph

    # ── Invoke-OneDriveUpload ─────────────────────────────────────────────────────
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
            $folderUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/$encodedRelPath"

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
                    $body = @{ folder = @{}; '@microsoft.graph.conflictBehavior' = 'replace' } | ConvertTo-Json -Depth 4
                    Invoke-WithGraphRetry -OperationName ("Create folder '{0}'" -f $relativePath) -Operation {
                        Invoke-MgGraphRequest -Method PATCH -Uri $folderUri -Body $body -ContentType 'application/json' -ErrorAction Stop | Out-Null
                    } # inline:Invoke-WithGraphRetry -OperationName ("C
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
                Error        = $null
            } # inline:$result = [pscustomobject]@{

            try
            {
                if ($PSCmdlet.ShouldProcess("OneDrive:/$relativePath ($($file.Length) bytes)", "Upload file"))
                {
                    $fileBytes = [System.IO.File]::ReadAllBytes($file.FullName)
                    Invoke-WithGraphRetry -OperationName ("Upload file '{0}'" -f $relativePath) -Operation {
                        Invoke-MgGraphRequest -Method PUT -Uri $uploadUri -Body $fileBytes -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
                    } # inline:Invoke-WithGraphRetry -OperationName ("U
                    $result.Status = 'Applied'
                    Write-LogLine -Message ("Uploaded file: OneDrive:/{0} ({1} bytes)" -f $relativePath, $file.Length)
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
                Write-LogLine -Level 'ERROR' -Message ("Failed to upload file '{0}': {1}" -f $relativePath, $result.Error)
            } # catch

            $result
        } # foreach
    } # function Invoke-OneDriveUpload

    # ── Assert-GraphAssemblyCompatibility ──────────────────────────────────────────
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

    # ── Get-ValidatedUserDrive ────────────────────────────────────────────────────
    # Looks up a user's OneDrive drive via Microsoft Graph, validates the response,
    # and confirms the drive root is accessible.  Returns the drive object with
    # .Id (the driveId used in all subsequent API calls) and .WebUrl.
    #
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

        Write-LogLine -Message ("Resolving OneDrive drive for owner: {0}" -f $UserPrincipalName)
        $userDrive = Invoke-WithGraphRetry -OperationName ("Get-MgUserDrive for '{0}'" -f $UserPrincipalName) -Operation {
            Get-MgUserDrive -UserId $UserPrincipalName -ErrorAction Stop
        } # inline:$userDrive = Invoke-WithGraphRetry -Oper

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

    # ── Initialization (runs once at script start) ───────────────────────────────
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

    Assert-GraphAssemblyCompatibility   # Check for PnP.PowerShell conflicts
    Assert-RequiredModules              # Import Graph SDK modules
    Connect-Graph                       # Authenticate to Microsoft Graph
} # begin

# ╔═══════════════════════════════════════════════════════════════════════════════╗
# ║  PROCESS BLOCK                                                               ║
# ║  Runs once for each pipeline input object ($UserToProcess).                  ║
# ║  If not piped and $UserToProcess is empty, processes all unique owners       ║
# ║  found in the CSV.                                                           ║
# ║  This is where the main work happens: read CSV, upload files, grant perms.   ║
# ║  https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods ║
# ╚═══════════════════════════════════════════════════════════════════════════════╝
process
{
    if (-not $InputFile.Exists)
    {
        throw "InputFile not found: $($InputFile.FullName)"
    } # if

    # Import-Csv reads a CSV file and converts each row into a PowerShell object.
    # Column headers become property names (e.g. $row.'Owner Login').
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.utility/import-csv
    $allRows = Import-Csv -LiteralPath $InputFile.FullName

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
        return
    } # if

    # Get unique owner UPNs from the CSV rows.
    # Select-Object -Unique returns distinct values.
    # https://learn.microsoft.com/powershell/module/microsoft.powershell.utility/select-object
    $uniqueOwners = @($allRows |
        ForEach-Object { $_.'Owner Login' } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique)

    if ($uniqueOwners.Count -eq 0)
    {
        throw "Owner Login is empty in all CSV rows."
    } # if

    Write-LogLine -Message ("Processing {0} unique owner(s): {1}" -f $uniqueOwners.Count, ($uniqueOwners -join ', '))

    # ── Iterate over each unique owner ───────────────────────────────────────
    foreach ($ownerUpn in $uniqueOwners)
    {
        Write-LogLine -Message ("── Begin processing owner: {0} ──" -f $ownerUpn)

        # Filter rows for this owner.
        $rows = $allRows | Where-Object { $_.'Owner Login' -eq $ownerUpn }

        # Look up and validate the user's OneDrive drive.
        # Wrapped in try/catch so one failing user does not stop others.
        try
        {
            $drive = Get-ValidatedUserDrive -UserPrincipalName $ownerUpn
        } # try
        catch
        {
            Write-LogLine -Level 'ERROR' -Message ("Failed to resolve OneDrive for '{0}': {1}. Skipping this owner." -f $ownerUpn, $_.Exception.Message)
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
                    ExistsInOneDrive       = $null
                    DriveItemId            = $null
                    Action                 = $null
                    Status                 = 'Failed'
                    Error                  = $_.Exception.Message
                } # inline:[pscustomobject]@{
            } # foreach
            continue
        } # catch

        # ── Upload local files if -UploadFiles is specified ──────────────────────
        # The local folder must be named by the user's UPN under AllFilesDirectory.
        # Example: C:\Repos\NPSBox\LocalFiles\user@contoso.com\
        if ($UploadFiles)
        {
            $userLocalPath = Join-Path -Path $AllFilesDirectory -ChildPath $ownerUpn
            Write-LogLine -Message ("Uploading local files from '{0}' to OneDrive for '{1}'." -f $userLocalPath, $ownerUpn)
            Invoke-OneDriveUpload -DriveId $drive.Id -LocalSourcePath $userLocalPath -OwnerUpn $ownerUpn
        } # if

        # ── Process each CSV row for this owner: resolve item, grant permission ──
        foreach ($row in $rows)
        {
            $itemPath = [string] $row.Path
            $itemName = [string] $row.'Item Name'
            $collab   = [string] $row.'Collaborator Login'
            $boxPerm  = [string] $row.'Collaborator Permission'

            # Map the Box permission to a Graph role (read/write/null).
            $graphRole = ConvertTo-GraphRole -BoxPermission $boxPerm

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
                ExistsInOneDrive       = $null
                DriveItemId            = $null
                Action                 = $null
                Status                 = 'Unknown'
                Error                  = $null
            } # inline:$result = [pscustomobject]@{
            try
            {
                if ([string]::IsNullOrWhiteSpace($collab))
                {
                    throw "Collaborator Login is empty."
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
                    # ── Grant permission using the driveItem: invite API ─────────
                    # POST /drives/{driveId}/items/{itemId}/invite
                    # This creates a sharing permission on the item.
                    #
                    # Key body properties:
                    #   recipients     : array of { email } objects — who to share with
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

                    # ── Validate the invite response ────────────────────────────
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
        } # foreach — row

        Write-LogLine -Message ("── Finished processing owner: {0} ──" -f $ownerUpn)
    } # foreach — owner
} # process

# ╔═══════════════════════════════════════════════════════════════════════════════╗
# ║  END BLOCK                                                                   ║
# ║  Runs once after all pipeline input has been processed.                      ║
# ║  Used here to disconnect from Microsoft Graph and clean up the session.      ║
# ╚═══════════════════════════════════════════════════════════════════════════════╝
end
{
    try
    {
        # Disconnect-MgGraph signs out of Microsoft Graph.
        # https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/disconnect-mggraph
        Disconnect-MgGraph | Out-Null
    } # try
    catch
    {
        # Non-fatal — the session will be cleaned up when PowerShell exits anyway.
    } # catch
} # end
