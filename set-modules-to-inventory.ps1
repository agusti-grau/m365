<#
.SYNOPSIS
    Exporta la lista de módulos de PowerShell instalados al sitio SharePoint indicado.

.DESCRIPTION
    1️⃣ Obtiene un token de acceso a Microsoft Graph usando Client Id + certificado.  
    2️⃣ Recupera la colección de módulos instalados (Get‑InstalledModule / Get‑Module).  
    3️⃣ Para cada módulo crea (o actualiza) un elemento en la lista de SharePoint con:
        • Nombre del módulo
        • Versión instalada
        • Ruta del módulo (opcional)
        • Fecha de exportación

.PARAMETER ClientId
    ID de la aplicación registrada en Azure AD.

.PARAMETER CertPath
    Ruta completa al archivo .pfx que contiene el certificado del cliente.

.PARAMETER CertPassword
    Contraseña del certificado (SecureString).

.PARAMETER SharePointUrl
    URL base del sitio de SharePoint (ej.: https://contoso.sharepoint.com/sites/mysite).

.PARAMETER ListName
    Nombre de la lista de SharePoint donde se guardarán los registros.
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    [Parameter(Mandatory=$true)]
    [string]$CertPath,

    [Parameter(Mandatory=$true)]
    [securestring]$CertPassword,

    [Parameter(Mandatory=$true)]
    [string]$SharePointUrl,

    [Parameter(Mandatory=$true)]
    [string]$ListName
)

#region ── Funciones auxiliares ────────────────────────────────────────────────────────

function Get-AccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertPath,
        [securestring]$CertPassword
    )
    # Cargar certificado
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 `
        -ArgumentList $CertPath, $CertPassword, 'Exportable,PersistKeySet'

    # Construir JWT client assertion (simplificado: usamos el certificado como client_assertion)
    $clientAssertion = [System.Convert]::ToBase64String($cert.GetRawCertData())

    $body = @{
        grant_type               = "client_credentials"
        client_id                = $ClientId
        scope                    = "https://graph.microsoft.com/.default"
        client_assertion_type    = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        client_assertion        = $clientAssertion
    }

    $tokenResponse = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body $body

    return $tokenResponse.access_token
}

function Add-ItemToSharePoint {
    param (
        [string]$SiteUrl,
        [string]$ListName,
        [hashtable]$Item,
        [string]$AccessToken
    )
    $endpoint = "$SiteUrl/_api/web/lists/getbytitle('$ListName')/items"
    $headers  = @{ Authorization = "Bearer $AccessToken"
                   Accept        = "application/json;odata=verbose"
                   "Content-Type"= "application/json;odata=verbose" }

    $payload = @{
        __metadata = @{ type = "SP.Data.$($ListName.Replace(' ','_'))ListItem" }
    } + $Item

    $json = $payload | ConvertTo-Json -Depth 5

    Invoke-RestMethod -Method Post -Uri $endpoint -Headers $headers -Body $json
}
#endregion ── Fin funciones ─────────────────────────────────────────────────────────────

# -------------------------------------------------------------------------
# 1️⃣ Determinar TenantId a partir de la URL de SharePoint
# -------------------------------------------------------------------------
if ($SharePointUrl -match "^https?://([^\.]+)\.sharepoint\.com") {
    $tenantDomain = $Matches[1] + ".onmicrosoft.com"
} else {
    Write-Error "No se pudo extraer el tenant de la URL proporcionada."
    exit 1
}

# -------------------------------------------------------------------------
# 2️⃣ Obtener token de acceso
# -------------------------------------------------------------------------
$accessToken = Get-AccessToken -TenantId $tenantDomain `
                               -ClientId $ClientId `
                               -CertPath $CertPath `
                               -CertPassword $CertPassword

# -------------------------------------------------------------------------
# 3️⃣ Recopilar módulos instalados
# -------------------------------------------------------------------------
Write-Host "`nRecopilando módulos instalados..." -ForegroundColor Cyan

# Get‑InstalledModule requiere PowerShellGet ≥ 2.0; si no está disponible usamos Get‑Module -ListAvailable
$installedModules = @()
try {
    $installedModules = Get-InstalledModule -AllVersions -ErrorAction Stop
} catch {
    # Fallback al método tradicional
    $installedModules = Get-Module -ListAvailable | Select-Object Name, Version, ModuleBase
}

if (-not $installedModules) {
    Write-Warning "No se encontraron módulos instalados."
    exit 0
}

# -------------------------------------------------------------------------
# 4️⃣ Enviar cada módulo a la lista de SharePoint
# -------------------------------------------------------------------------
$exportDate = (Get-Date).ToString("s")   # ISO‑8601

foreach ($mod in $installedModules) {
    $item = @{
        Title          = $mod.Name                     # Campo obligatorio en listas genéricas
        Version        = $mod.Version.ToString()
        ModulePath     = $mod.ModuleBase   # Sólo está presente cuando usamos Get‑Module
        ExportedOn     = $exportDate
    }

    try {
        Add-ItemToSharePoint -SiteUrl $SharePointUrl `
                             -ListName $ListName `
                             -Item $item `
                             -AccessToken $accessToken
        Write-Host "✔️  $($mod.Name) v$($mod.Version) → añadido a SharePoint" -ForegroundColor Green
    } catch {
        Write-Warning "Error al agregar $($mod.Name): $_"
    }
}

Write-Host "`nExportación completada." -ForegroundColor Cyan
