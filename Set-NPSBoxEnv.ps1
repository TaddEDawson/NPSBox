<#
.SYNOPSIS
    Sets environment variables used as defaults by Update-UserFile.ps1.

    Version: 1.0.0.0
    Date:    2026-05-05

.DESCRIPTION
    Creates or updates the NPSBOX_TENANT_ID, NPSBOX_CLIENT_ID, and
    NPSBOX_CERT_THUMBPRINT environment variables at the User scope so they
    persist across PowerShell sessions.

    Update-UserFile.ps1 reads these variables as parameter defaults.  Once
    set, you can run the script without specifying -TenantId, -ClientId,
    or -CertificateThumbprint each time.

.PARAMETER TenantId
    Microsoft 365 tenant ID (GUID).

.PARAMETER ClientId
    Azure AD app registration client ID (GUID).

.PARAMETER CertificateThumbprint
    SHA-1 thumbprint of the certificate in Cert:\CurrentUser\My.

.EXAMPLE
    .\Set-NPSBoxEnv.ps1

    Sets all three variables using the built-in defaults.

.EXAMPLE
    .\Set-NPSBoxEnv.ps1 -TenantId '00000000-...' -ClientId '11111111-...' -CertificateThumbprint 'AABB...'

    Sets custom values for a different environment.

.NOTES
    To view current values:
      $env:NPSBOX_TENANT_ID
      $env:NPSBOX_CLIENT_ID
      $env:NPSBOX_CERT_THUMBPRINT

    To remove:
      [Environment]::SetEnvironmentVariable('NPSBOX_TENANT_ID', $null, 'User')
      [Environment]::SetEnvironmentVariable('NPSBOX_CLIENT_ID', $null, 'User')
      [Environment]::SetEnvironmentVariable('NPSBOX_CERT_THUMBPRINT', $null, 'User')
#>
#Requires -Version 7.0

[CmdletBinding()]
param
(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $TenantId = '92075952-90f3-4613-833b-d2e19ec649e4'
    ,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $ClientId = '912696b9-1374-4110-893d-545fc17c3371'
    ,
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $CertificateThumbprint = '9D0F9B62AC3B002E56C2A304E88AD429813E55E2'
)

$vars = @{
    NPSBOX_TENANT_ID      = $TenantId
    NPSBOX_CLIENT_ID      = $ClientId
    NPSBOX_CERT_THUMBPRINT = $CertificateThumbprint
}

foreach ($name in $vars.Keys)
{
    $value = $vars[$name]

    # Set at User scope (persists across sessions).
    [Environment]::SetEnvironmentVariable($name, $value, 'User')

    # Also set in the current session so it takes effect immediately.
    Set-Item -Path "Env:\$name" -Value $value

    Write-Verbose ("{0} = {1}" -f $name, $value)
} # foreach

Write-Output "Environment variables set (User scope). Restart terminals to pick up changes."
Write-Output ""
foreach ($name in $vars.Keys | Sort-Object)
{
    [pscustomobject]@{
        Variable = $name
        Value    = $vars[$name]
    }
} # foreach
