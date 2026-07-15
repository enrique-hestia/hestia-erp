# ─────────────────────────────────────────────────────────────────────────
#  deploy-gas.ps1 — Despliega el backend (Apps Script) en UN comando.
#  Uso:  powershell -ExecutionPolicy Bypass -File .\deploy-gas.ps1
#
#  Copia SOLO los .gs activos (lista de abajo) a gas-deploy\ (staging limpio),
#  hace `clasp push` y actualiza el deployment del Web App a una versión nueva
#  (MISMA URL). No usa .claspignore (clasp 3.x lo ignora): el staging solo
#  contiene los archivos correctos, por eso NUNCA sube respaldos ni el HTML.
#
#  Al agregar un .gs NUEVO al proyecto: añádelo a $FILES.
# ─────────────────────────────────────────────────────────────────────────
$ErrorActionPreference = 'Stop'
$repo  = $PSScriptRoot
$stage = Join-Path $repo 'gas-deploy'
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')
$clasp = Join-Path $env:APPDATA 'npm\clasp.cmd'
$DEPLOYMENT_ID = 'AKfycbw3pPDLfDFVmb2VF1FCGCTZGK8Ie1du7zUFyJPQT7Fwxi_ngsEmEL-o-ukv2QzqZ6Ro'

# Archivos .gs que SÍ van al proyecto de Apps Script (= los desplegados hoy).
# OJO: 'config_dropdowns' se ELIMINO (2026-07-15). Definia CFG_DD_TAB/CFG_DD_DEFAULTS/
# setupConfigDropdowns/readDropdowns/saveDropdownValues DUPLICADOS con finance.gs, pero en
# version vieja (10 dropdowns, sin Productos ni Empleados, sin _ddAddMissingDefaults).
# Como Apps Script comparte scope global, ganaba el archivo que cargara al final -> bomba
# de tiempo. finance.gs es superconjunto estricto; config_dropdowns no aportaba NADA propio.
$FILES = @(
  '_diagnostico','analisis','board','capture','catalogo','chat','cobranza',
  'comprobantes','config','core','cxp_creditos','declaraciones',
  'devengado','facturacion','finance','gastosfijos','identificar','inventario','inventario_migracion',
  'lab','medical','nomina','origenes','poliza_concil','presupuesto','privacidad',
  'productos_ss_id','prov_defaults','providers','recordatorios','scheduler','semanal','summary'
)

if (-not (Test-Path $clasp)) { Write-Error 'clasp no encontrado. Corre: npm i -g @google/clasp'; exit 1 }

# 1) Sincroniza el staging con los .gs actuales del repo.
Get-ChildItem (Join-Path $stage '*.gs') -ErrorAction SilentlyContinue | Remove-Item -Force
$falta = @()
foreach ($f in $FILES) {
  $src = Join-Path $repo "$f.gs"
  if (Test-Path $src) { Copy-Item $src (Join-Path $stage "$f.gs") -Force } else { $falta += "$f.gs" }
}
if ($falta.Count) { Write-Error "Faltan en el repo: $($falta -join ', ')"; exit 1 }
Write-Host "Staging: $($FILES.Count) .gs sincronizados." -ForegroundColor Cyan

# 2) Push al editor de Apps Script.
Set-Location $stage
Write-Host '=== clasp push ===' -ForegroundColor Cyan
& $clasp push --force
if ($LASTEXITCODE -ne 0) { Set-Location $repo; Write-Error 'clasp push fallo.'; exit 1 }

# 3) Actualiza el deployment del Web App (misma URL) a version nueva.
Write-Host '=== clasp deploy (misma URL) ===' -ForegroundColor Cyan
& $clasp deploy --deploymentId $DEPLOYMENT_ID --description ("deploy-gas.ps1 " + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
Set-Location $repo
Write-Host 'LISTO. Backend desplegado (misma URL).' -ForegroundColor Green
