# setup_dependencies.ps1
# Script de instalación y verificación automática para el Proyecto PDF-Word (v2.1)

Write-Host "=== INICIANDO CONFIGURACIÓN AUTOMÁTICA DEL ENTORNO ===" -ForegroundColor Cyan

# 1. VERIFICAR PERMISOS DE ADMINISTRADOR (Necesario para escribir en Program Files)
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
  Write-Warning "Este script necesita permisos de Administrador."
  Write-Host "Reiniciando como Administrador..." -ForegroundColor Yellow
  Start-Process powershell -Verb RunAs -ArgumentList "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path)
  Break
}

# --- BLOQUE A: POPPLER (Versión 25.12.0 corregida) ---
# Necesario para pdf2image
$popplerUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v25.12.0-0/Release-25.12.0-0.zip"
$installPathPoppler = "C:\Program Files\Poppler"
$zipPath = "$env:TEMP\poppler.zip"
$binPathPoppler = "$installPathPoppler\poppler-25.12.0\Library\bin"

Write-Host "`n[1/3] Verificando Poppler..." -ForegroundColor White
# COMPROBACIÓN DE SEGURIDAD 1: Verifica si el ejecutable ya existe.
if (Test-Path "$binPathPoppler\pdftoppm.exe") {
    Write-Host "   -> Poppler ya está instalado." -ForegroundColor Green
} else {
    Write-Host "   -> Descargando Poppler..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $popplerUrl -OutFile $zipPath -UseBasicParsing
    
    Write-Host "   -> Descomprimiendo en $installPathPoppler..." -ForegroundColor Yellow
    Expand-Archive -Path $zipPath -DestinationPath $installPathPoppler -Force
    
    # Agregar al PATH
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ($currentPath -notlike "*$binPathPoppler*") {
        Write-Host "   -> Agregando Poppler al PATH del sistema..." -ForegroundColor Yellow
        [Environment]::SetEnvironmentVariable("Path", $currentPath + ";$binPathPoppler", "Machine")
    }
    Write-Host "   -> [LISTO] Poppler instalado." -ForegroundColor Green
}

# --- BLOQUE B: TESSERACT-OCR ---
# Necesario para pytesseract
Write-Host "`n[2/3] Verificando Tesseract-OCR..." -ForegroundColor White
$tesseractPath = "C:\Program Files\Tesseract-OCR"
$tesseractExe = "$tesseractPath\tesseract.exe"

# COMPROBACIÓN DE SEGURIDAD 2: Verifica si el ejecutable ya existe.
if (Test-Path $tesseractExe) {
    Write-Host "   -> Tesseract ya está instalado." -ForegroundColor Green
} else {
    Write-Host "   -> Tesseract NO encontrado. Instalando automáticamente..." -ForegroundColor Magenta
    winget install -e --id UB-Mannheim.TesseractOCR --silent --accept-source-agreements --accept-package-agreements
    
    # Agregar Tesseract al PATH 
    if (Test-Path $tesseractExe) {
        Write-Host "   -> [LISTO] Tesseract instalado correctamente." -ForegroundColor Green
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($currentPath -notlike "*$tesseractPath*") {
            [Environment]::SetEnvironmentVariable("Path", $currentPath + ";$tesseractPath", "Machine")
            Write-Host "   -> Tesseract agregado al PATH." -ForegroundColor Green
        }
    } else {
        Write-Host "   -> [ERROR] Winget no pudo instalar Tesseract. Inténtalo manualmente." -ForegroundColor Red
    }
}

# --- BLOQUE C: LIBRERÍAS DE PYTHON (Lista de requisitos actualizada) ---
Write-Host "`n[3/3] Instalando librerías de Python..." -ForegroundColor White
# Incluye tu lista exacta, además de docx2pdf (necesaria para conversor_word.py) y Pillow (dependencia común de pdf2image).
pip install pdf2docx pymupdf pytesseract pdf2image tqdm Pillow python-docx docx2pdf

# --- BLOQUE D: VERIFICACIÓN FINAL ---
Write-Host "`n=== INSTALACIÓN FINALIZADA CON VERIFICACIÓN DE DEPENDENCIAS ===" -ForegroundColor Cyan

Write-Host "--- VERIFICANDO TESSERACT ---" -ForegroundColor Yellow
# VERIFICACIÓN 1: Ejecuta el Tesseract.exe directamente, sin depender del PATH inmediato.
if (Test-Path "$tesseractExe") {
    & "$tesseractExe" -v
} else {
    Write-Host "Tesseract: No se pudo verificar la versión (Error de instalación)." -ForegroundColor Red
}

Write-Host "`n--- VERIFICANDO POPPLER ---" -ForegroundColor Yellow
# VERIFICACIÓN 2: Ejecuta el pdftoppm.exe directamente, sin depender del PATH inmediato.
if (Test-Path "$binPathPoppler\pdftoppm.exe") {
    & "$binPathPoppler\pdftoppm.exe" -v
} else {
    Write-Host "Poppler: No se pudo verificar la versión (Error de instalación)." -ForegroundColor Red
}

Write-Host "`n--- INSTRUCCIONES FINALES ---" -ForegroundColor Magenta
Write-Host "Si las dos verificaciones mostraron la versión, la instalación fue EXITOSA." -ForegroundColor Green
Write-Host "⚠️ IMPORTANTE: CIERRA y RE-ABRE tu terminal (o VS Code) para que el nuevo PATH funcione correctamente." -ForegroundColor Yellow
Pause