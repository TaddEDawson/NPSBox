#Requires -Version 7.0

<#
.SYNOPSIS
    Creates a self-signed certificate and uploads it to an Azure AD app registration.

.DESCRIPTION
    - Creates a self-signed certificate in the current user's certificate store.
    - Connects to Microsoft Graph to upload the public key to the specified app registration.
    - Verifies the certificate was uploaded successfully.
    - Optionally tests Graph connectivity using the new certificate.

.PARAMETER AppClientId
    The Application (client) ID of the Azure AD app registration.

.PARAMETER TenantId
    The Azure AD tenant ID.

.PARAMETER CertificateSubject
    The subject name for the self-signed certificate.

.PARAMETER ValidityYears
    Number of years the certificate remains valid.

.PARAMETER SkipConnectivityTest
    Skip the final step that tests Graph connectivity with the new certificate.

.EXAMPLE
    .\Add-AppCertificateCredential.ps1 -AppClientId '14d82eec-204b-4c2f-b7e8-296a70dab67e' `
        -TenantId '92075952-90f3-4613-833b-d2e19ec649e4' -Verbose

.EXAMPLE
    .\Add-AppCertificateCredential.ps1 -AppClientId '14d82eec-204b-4c2f-b7e8-296a70dab67e' `
        -TenantId '92075952-90f3-4613-833b-d2e19ec649e4' `
        -CertificateSubject 'CN=MyGraphApp' -ValidityYears 1 -Verbose

.NOTES
    Requires Microsoft.Graph.Authentication and Microsoft.Graph.Applications modules.
    The interactive sign-in account must have Application.ReadWrite.All permission.

    Docs:
      - New-SelfSignedCertificate: https://learn.microsoft.com/powershell/module/pki/new-selfsignedcertificate
      - Update-MgApplication: https://learn.microsoft.com/powershell/module/microsoft.graph.applications/update-mgapplication
      - Get-MgApplication: https://learn.microsoft.com/powershell/module/microsoft.graph.applications/get-mgapplication
      - Connect-MgGraph: https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph
#>

[CmdletBinding()]
param
(
    [Parameter()]
    [string] $AppClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
    ,
    [Parameter()]
    [string] $TenantId = '92075952-90f3-4613-833b-d2e19ec649e4'
    ,
    [Parameter()]
    [string] $CertificateSubject = 'CN=NPSBoxGraphApp'
    ,
    [Parameter()]
    [ValidateRange(1, 10)]
    [int] $ValidityYears = 2
    ,
    [Parameter()]
    [switch] $SkipConnectivityTest
)

# ── Step 1: Create self-signed certificate ──────────────────────────────────
Write-Verbose "Creating self-signed certificate with subject '$CertificateSubject' (valid $ValidityYears year(s))."

$certParams = @{
    Subject           = $CertificateSubject
    CertStoreLocation = 'Cert:\CurrentUser\My'
    KeyExportPolicy   = 'Exportable'
    KeySpec           = 'Signature'
    KeyLength         = 2048
    KeyAlgorithm      = 'RSA'
    HashAlgorithm     = 'SHA256'
    NotAfter          = (Get-Date).AddYears($ValidityYears)
}

$cert = New-SelfSignedCertificate @certParams
Write-Verbose ("Certificate created. Thumbprint: {0}" -f $cert.Thumbprint)

# ── Step 2: Import required modules ─────────────────────────────────────────
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')
foreach ($moduleName in $requiredModules)
{
    $available = Get-Module -ListAvailable -Name $moduleName |
        Sort-Object -Property Version -Descending |
        Select-Object -First 1

    if ($null -eq $available)
    {
        throw "Required module not found: $moduleName. Install with: Install-Module $moduleName -Scope CurrentUser"
    }

    Write-Verbose ("Importing module {0} ({1})" -f $moduleName, $available.Version)
    Import-Module -Name $moduleName -RequiredVersion $available.Version -ErrorAction Stop -Verbose:$false | Out-Null
}

# ── Step 3: Connect to Graph (interactive) to manage app registration ───────
Write-Verbose "Connecting to Microsoft Graph (interactive) to upload certificate."
Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All' -ErrorAction Stop -NoWelcome | Out-Null

try
{
    # ── Step 4: Look up the app registration ────────────────────────────────
    Write-Verbose ("Looking up app registration with appId '{0}'." -f $AppClientId)
    $app = Get-MgApplication -Filter "appId eq '$AppClientId'" -ErrorAction Stop

    if ($null -eq $app)
    {
        throw ("App registration not found for appId '{0}'." -f $AppClientId)
    }

    Write-Verbose ("Found app registration: DisplayName='{0}', ObjectId='{1}'" -f $app.DisplayName, $app.Id)

    # ── Step 5: Upload the certificate public key ───────────────────────────
    Write-Verbose "Uploading certificate public key to app registration."

    $keyCredential = @{
        type          = 'AsymmetricX509Cert'
        usage         = 'Verify'
        key           = $cert.GetRawCertData()
        displayName   = $CertificateSubject -replace '^CN=', ''
        startDateTime = $cert.NotBefore.ToUniversalTime().ToString('o')
        endDateTime   = $cert.NotAfter.ToUniversalTime().ToString('o')
    }

    # Preserve any existing key credentials
    $existingKeys = @()
    if ($null -ne $app.KeyCredentials)
    {
        $existingKeys = @($app.KeyCredentials | ForEach-Object {
            @{
                type          = $_.Type
                usage         = $_.Usage
                key           = $_.Key
                displayName   = $_.DisplayName
                startDateTime = $_.StartDateTime.ToString('o')
                endDateTime   = $_.EndDateTime.ToString('o')
            }
        })
    }

    $allKeys = $existingKeys + @($keyCredential)
    Update-MgApplication -ApplicationId $app.Id -KeyCredentials $allKeys -ErrorAction Stop
    Write-Verbose "Certificate uploaded successfully."

    # ── Step 6: Verify the upload ───────────────────────────────────────────
    Write-Verbose "Verifying certificate was added to app registration."
    $updatedApp = Get-MgApplication -ApplicationId $app.Id -ErrorAction Stop

    $uploadedCert = $updatedApp.KeyCredentials | Where-Object {
        $_.DisplayName -eq ($CertificateSubject -replace '^CN=', '')
    } | Select-Object -Last 1

    if ($null -eq $uploadedCert)
    {
        throw "Verification failed: certificate not found on app registration after upload."
    }

    Write-Verbose "Verification passed: certificate is present on app registration."

    # ── Step 7: Optionally test Graph connectivity with the certificate ─────
    if (-not $SkipConnectivityTest)
    {
        Write-Verbose "Testing Graph connectivity with certificate auth."
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

        Connect-MgGraph -TenantId $TenantId -ClientId $AppClientId `
            -CertificateThumbprint $cert.Thumbprint -ErrorAction Stop -NoWelcome | Out-Null

        $context = Get-MgContext
        if ($null -eq $context -or $context.AppName -ne $app.DisplayName)
        {
            Write-Warning "Connected to Graph but app context may not match. Review Get-MgContext output."
        }
        else
        {
            Write-Verbose ("Connectivity verified. Connected as '{0}' (AuthType={1})." -f $context.AppName, $context.AuthType)
        }

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
finally
{
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
} # finally

# ── Output ──────────────────────────────────────────────────────────────────
[PSCustomObject]@{
    Thumbprint   = $cert.Thumbprint
    Subject      = $cert.Subject
    NotBefore    = $cert.NotBefore
    NotAfter     = $cert.NotAfter
    AppClientId  = $AppClientId
    AppName      = $app.DisplayName
    TenantId     = $TenantId
    StorePath    = 'Cert:\CurrentUser\My\{0}' -f $cert.Thumbprint
    Status       = 'Ready'
}
