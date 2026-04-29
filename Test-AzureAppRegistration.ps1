#Requires -Version 7.0

<#
.SYNOPSIS
    Verifies that the Azure AD app registration has the required Microsoft Graph
    permissions for the NPSBox migration script.

    Version: 1.2.0.12
    Date:    2026-04-29

.DESCRIPTION
    Connects to Microsoft Graph using the same certificate-based app-only auth
    as Update-UserFile.ps1 and checks whether the required application
    permissions have been granted with admin consent.

    Required permissions:
      Files.ReadWrite.All   — Drive access, file upload, sharing via invite API.
      User.Read.All         — User account validation (Get-MgUser).

    The script outputs a custom object per required permission showing whether
    it has been granted.

    PREREQUISITES:
      - PowerShell 7.0 or later
      - Microsoft Graph PowerShell SDK modules:
          Microsoft.Graph.Authentication
          Microsoft.Graph.Applications
        Install with:
          Install-Module Microsoft.Graph.Authentication  -Scope CurrentUser
          Install-Module Microsoft.Graph.Applications    -Scope CurrentUser
      - The same certificate, ClientId, and TenantId used by Update-UserFile.ps1

.PARAMETER TenantId
    Your Microsoft 365 tenant ID (GUID).

.PARAMETER ClientId
    The Application (client) ID of the Azure AD app registration.

.PARAMETER CertificateThumbprint
    The SHA-1 thumbprint of the certificate installed in Cert:\CurrentUser\My.

.EXAMPLE
    .\Test-AzureAppRegistration.ps1 -Verbose

.EXAMPLE
    .\Test-AzureAppRegistration.ps1 -ClientId '98454154-...' -TenantId '92075952-...' -CertificateThumbprint '9D0F9B...'

.NOTES
    Documentation:
      - Graph permissions reference: https://learn.microsoft.com/graph/permissions-reference
      - Connect-MgGraph:             https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
      - Get-MgServicePrincipal:      https://learn.microsoft.com/powershell/module/microsoft.graph.applications/get-mgserviceprincipal
      - App role assignments:        https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignments
#>

[CmdletBinding()]
param
(
    # Your tenant ID (GUID).  Find it: Azure Portal > Entra ID > Overview.
    [Parameter()]
    [string] $TenantId = "92075952-90f3-4613-833b-d2e19ec649e4"
    ,
    # The app registration's client ID (GUID).
    [Parameter()]
    [string] $ClientId = "912696b9-1374-4110-893d-545fc17c3371"
    ,
    # Certificate thumbprint for app-only auth.
    [Parameter()]
    [string] $CertificateThumbprint = "9D0F9B62AC3B002E56C2A304E88AD429813E55E2"
)

begin
{
    # ── Required permissions for Update-UserFile.ps1 ─────────────────────────
    # Each entry is the Graph application permission Value (role name).
    $script:RequiredPermissions = @(
        'Files.ReadWrite.All',
        'User.Read.All'
    )

    # ── Assert-RequiredModules ───────────────────────────────────────────────
    function Assert-RequiredModules
    {
        [CmdletBinding()]
        param()

        $requiredModules = @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Applications'
        )

        foreach ($moduleName in $requiredModules)
        {
            $availableModule = Get-Module -ListAvailable -Name $moduleName |
                Sort-Object -Property Version -Descending |
                Select-Object -First 1

            if ($null -eq $availableModule)
            {
                throw (
                    "Required module not found: $moduleName. Install it with: Install-Module $moduleName -Scope CurrentUser"
                )
            } # if

            Write-Verbose ("Importing module {0} ({1})" -f $moduleName, $availableModule.Version)
            Import-Module -Name $moduleName -RequiredVersion $availableModule.Version -ErrorAction Stop -Verbose:$false | Out-Null
        } # foreach
    } # function Assert-RequiredModules

    # ── Connect-Graph ────────────────────────────────────────────────────────
    function Connect-Graph
    {
        [CmdletBinding()]
        param()

        $existingContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -ne $existingContext -and $existingContext.TenantId -eq $TenantId -and $existingContext.AuthType -eq 'AppOnly')
        {
            Write-Verbose ("Reusing existing Microsoft Graph context. TenantId={0}, AppName={1}" -f
                $existingContext.TenantId, $existingContext.AppName)
            return
        } # if

        if ($null -ne $existingContext)
        {
            Write-Verbose ("Disconnecting existing Graph session (AuthType={0})." -f $existingContext.AuthType)
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        } # if

        if ([string]::IsNullOrWhiteSpace($TenantId))   { throw "Certificate auth requires -TenantId." }
        if ([string]::IsNullOrWhiteSpace($ClientId))    { throw "Certificate auth requires -ClientId." }

        Write-Verbose ("Connecting to Microsoft Graph. TenantId={0}, ClientId={1}" -f $TenantId, $ClientId)
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop -NoWelcome | Out-Null
    } # function Connect-Graph

    Assert-RequiredModules
    Connect-Graph
} # begin

process
{
    # ── Resolve the app's service principal in the tenant ────────────────────
    # https://learn.microsoft.com/powershell/module/microsoft.graph.applications/get-mgserviceprincipal
    Write-Verbose ("Looking up service principal for ClientId={0}" -f $ClientId)
    $appSp = Get-MgServicePrincipal -Filter "appId eq '$ClientId'" -ErrorAction Stop
    if ($null -eq $appSp)
    {
        throw ("Service principal not found for ClientId '{0}'. Ensure the app registration exists in tenant '{1}'." -f $ClientId, $TenantId)
    } # if

    # ── Resolve the Microsoft Graph service principal ────────────────────────
    # All Graph app roles are defined on this service principal.
    Write-Verbose "Looking up Microsoft Graph service principal."
    $graphSp = Get-MgServicePrincipal -Filter "displayName eq 'Microsoft Graph'" -ErrorAction Stop
    if ($null -eq $graphSp)
    {
        throw "Could not find the 'Microsoft Graph' service principal in the tenant."
    } # if

    # ── Get the app role assignments (granted application permissions) ───────
    # https://learn.microsoft.com/graph/api/serviceprincipal-list-approleassignments
    Write-Verbose "Retrieving app role assignments for the service principal."
    $appRoleAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $appSp.Id -ErrorAction Stop

    # Build a lookup of granted role IDs for quick matching.
    $grantedRoleIds = @{}
    foreach ($assignment in $appRoleAssignments)
    {
        $grantedRoleIds[$assignment.AppRoleId] = $assignment
    } # foreach

    # ── Check each required permission ───────────────────────────────────────
    $allGranted = $true
    foreach ($permissionName in $script:RequiredPermissions)
    {
        # Find the app role definition on the Graph service principal.
        $roleDef = $graphSp.AppRoles | Where-Object { $_.Value -eq $permissionName } | Select-Object -First 1

        $assignment = $null
        $isGranted  = $false
        if ($null -ne $roleDef -and $grantedRoleIds.ContainsKey($roleDef.Id))
        {
            $isGranted  = $true
            $assignment = $grantedRoleIds[$roleDef.Id]
        } # if

        if (-not $isGranted)
        {
            $allGranted = $false
        } # if

        [pscustomobject]@{
            Permission  = $permissionName
            Type        = 'Application'
            IsGranted   = $isGranted
            RoleId      = if ($null -ne $roleDef) { $roleDef.Id } else { $null }
            GrantedOn   = if ($null -ne $assignment) { $assignment.CreatedDateTime } else { $null }
            AppId       = $ClientId
            TenantId    = $TenantId
            DisplayName = $appSp.DisplayName
        } # inline:[pscustomobject]@{
    } # foreach

    if ($allGranted)
    {
        Write-Verbose "All required permissions are granted."
    } # if
    else
    {
        Write-Warning ("One or more required permissions are NOT granted. " +
            "Go to Azure Portal > App registrations > '{0}' > API permissions and grant admin consent." -f $appSp.DisplayName)
    } # else
} # process

end
{
    try
    {
        Disconnect-MgGraph | Out-Null
    } # try
    catch
    {
        # Non-fatal — session cleaned up at PowerShell exit.
    } # catch
} # end
