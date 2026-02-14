<#
.SYNOPSIS
    Descarga e instala módulos de SharePoint especificados en una lista.

.DESCRIPTION
    - Se conecta a SharePoint usando ClientId y certificado.
    - Lee una lista de módulos (nombre y versión) desde una lista de SharePoint.
    - Descarga cada módulo desde la galería de PowerShell (o una ubicación personalizada).
    - Muestra en pantalla el nombre del módulo y la versión encontrada.
    - Pide confirmación al administrador antes de instalar cada módulo en el equipo local.

.PARAMETER ClientId
    El ID de aplicación (Client Id) registrado en Azure AD.

.PARAMETER CertPath
    Ruta al archivo .pfx que contiene el certificado del cliente.

.PARAMETER CertPassword
    Contraseña del certificado (si tiene protección).

.PARAMETER SharePointUrl
    URL base del sitio de SharePoint (ej.: https://contoso.sharepoint.com/sites/mysite).

.PARAMETER ListName
    Nombre de la lista de SharePoint que contiene los módulos a gestionar.
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

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
        client_assertion = [System.Convert]::ToBase64String($cert.GetRawCertData())
    }

    $tokenResponse = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -ContentType "application/x-www-form-urlencoded" `
        -Body $body

    return $tokenResponse.access_token
}

function Get-SharePointListItems {
    param (
        [string]$SiteUrl,
        [string]$ListName,
        [string]$AccessToken
    )
    $endpoint = "$SiteUrl/_api/web/lists/getbytitle('$ListName')/items?\$select=Title,Version"
    $headers = @{ Authorization = "Bearer $AccessToken" }

    $response = Invoke-RestMethod -Method Get -Uri $endpoint -Headers $headers -Headers $headers
    return $response.value
}

# -------------------------------------------------------------------------
# 1️⃣  Obtener token de acceso (requiere TenantId; lo extraemos del dominio)
# -------------------------------------------------------------------------
if ($SharePointUrl -match "^https?://([^\.]+)\.sharepoint\.com") {
    $tenantDomain = $Matches[1] + ".onmicrosoft.com"
} else {
    Write-Error "No se pudo determinar el tenant a partir de la URL de SharePoint."
    exit 1
}
$accessToken = Get-AccessToken -TenantId $tenantDomain -ClientId $ClientId -CertPath $CertPath -CertPassword $CertPassword

# -------------------------------------------------------------------------
# 2️⃣  Leer la lista de módulos desde SharePoint
# -------------------------------------------------------------------------
Write-Host "`nObteniendo la lista de módulos desde SharePoint..." -ForegroundColor Cyan
$modules = Get-SharePointListItems -SiteUrl $SharePointUrl -ListName $ListName -AccessToken $accessToken

if (-not $modules) {
    Write-Warning "La lista está vacía o no se pudieron obtener los elementos."
    exit 0
}

# -------------------------------------------------------------------------
# 3️⃣  Procesar cada módulo
# -------------------------------------------------------------------------
foreach ($module in $modules) {
    $moduleName    = $module.Title      # Asumimos que la columna "Title" contiene el nombre del módulo
    $moduleVersion = $module.Version    # Asumimos que la columna "Version" contiene la versión deseada

    Write-Host "`nMódulo: $moduleName"
    Write-Host "Versión solicitada: $moduleVersion"

    # 3️⃣🔎 Buscar el módulo en la galería de PowerShell (puedes cambiar la fuente si usas otra)
    $found = Find-Module -Name $moduleName -AllVersions -ErrorAction SilentlyContinue |
             Where-Object { $_.Version -eq [Version]$moduleVersion }

    if (-not $found) {
        Write-Warning "No se encontró la versión $moduleVersion del módulo $moduleName en la galería configurada."
        continue
    }

    # 3️⃣📦 Mostrar información del paquete encontrado
    Write-Host "Encontrado: $($found.Name) v$($found.Version) - $($found.Description)" -ForegroundColor Green

    # 3️⃣❓ Pedir confirmación al administrador
    $confirm = Read-Host "¿Deseas instalar este módulo en el sistema? (S/N)"
    if ($confirm -match '^[Ss]') {
        try {
            Install-Module -Name $moduleName -RequiredVersion $moduleVersion -Scope AllUsers -Force -AllowClobber
            Write-Host "✅ Instalado correctamente." -ForegroundColor Green
        } catch {
            Write-Error "Error al instalar $moduleName v$moduleVersion: $_"
        }
    } else {
        Write-Host "⚠️ Instalación omitida por el usuario." -ForegroundColor Yellow
    }
}

Write-Host "`nProceso completado." -ForegroundColor Cyan
