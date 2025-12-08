# setup_dependencies.ps1
# Script de instalación automática para el Proyecto PDF-Word
# Instala: Poppler (Manual) y Tesseract (Automático via Winget)

Write-Host "=== INICIANDO CONFIGURACIÓN AUTOMÁTICA DEL ENTORNO ===" -ForegroundColor Cyan

# 1. VERIFICAR PERMISOS DE ADMINISTRADOR (Necesario para escribir en Program Files)
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
  Write-Warning "Este script necesita permisos de Administrador."
  Write-Host "Reiniciando como Administrador..." -ForegroundColor Yellow
  Start-Process powershell -Verb RunAs -ArgumentList "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path)
  Break
}

# --- BLOQUE A: POPPLER (Necesario para pdf2image) ---
$popplerUrl = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip"
$installPathPoppler = "C:\Program Files\Poppler"
$zipPath = "$env:TEMP\poppler.zip"
$binPathPoppler = "$installPathPoppler\poppler-24.08.0\Library\bin"

Write-Host "`n[1/3] Verificando Poppler..." -ForegroundColor White
if (Test-Path "$binPathPoppler\pdftoppm.exe") {
    Write-Host "   -> Poppler ya está instalado." -ForegroundColor Green
} else {
    Write-Host "   -> Descargando Poppler..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $popplerUrl -OutFile $zipPath
    
    Write-Host "   -> Descomprimiendo en $installPathPoppler..." -ForegroundColor Yellow
    Expand-Archive -Path $zipPath -DestinationPath $installPathPoppler -Force
    
    # Agregar al PATH
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ($currentPath -notlike "*$binPathPoppler*") {
        Write-Host "   -> Agregando Poppler al PATH del sistema..." -ForegroundColor Yellow
        [Environment]::SetEnvironmentVariable("Path", $currentPath + ";$binPathPoppler", "Machine")
    }
    Write-Host "   -> [LISTO] Poppler instalado." -ForegroundColor Green
}

# --- BLOQUE B: TESSERACT-OCR (Necesario para el texto en imágenes) ---
Write-Host "`n[2/3] Verificando Tesseract-OCR..." -ForegroundColor White
$tesseractPath = "C:\Program Files\Tesseract-OCR"
$tesseractExe = "$tesseractPath\tesseract.exe"

if (Test-Path $tesseractExe) {
    Write-Host "   -> Tesseract ya está instalado." -ForegroundColor Green
} else {
    Write-Host "   -> Tesseract NO encontrado. Instalando automáticamente..." -ForegroundColor Magenta
    
    # Comando Winget SILENCIOSO (Sin ventanas, sin preguntas)
    # -e: Exacto
    # --silent: No muestra interfaz gráfica
    # --accept-source-agreements: Acepta términos legales solo
    winget install -e --id UB-Mannheim.TesseractOCR --silent --accept-source-agreements --accept-package-agreements
    
    # Verificación post-instalación
    if (Test-Path $tesseractExe) {
        Write-Host "   -> [LISTO] Tesseract instalado correctamente." -ForegroundColor Green
        
        # Agregar Tesseract al PATH también (para que Python lo vea fácil)
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($currentPath -notlike "*$tesseractPath*") {
            [Environment]::SetEnvironmentVariable("Path", $currentPath + ";$tesseractPath", "Machine")
            Write-Host "   -> Tesseract agregado al PATH." -ForegroundColor Green
        }
    } else {
        Write-Host "   -> [ERROR] Winget no pudo instalar Tesseract. Inténtalo manualmente." -ForegroundColor Red
    }
}

# --- BLOQUE C: LIBRERÍAS DE PYTHON ---
Write-Host "`n[3/3] Instalando librerías de Python..." -ForegroundColor White
pip install pdf2docx pymupdf pytesseract pdf2image tqdm Pillow python-docx

Write-Host "`n=== INSTALACIÓN FINALIZADA CON ÉXITO ===" -ForegroundColor Cyan
Write-Host "IMPORTANTE: Cierra VS Code y todas las terminales y ábrelas de nuevo para que los cambios funcionen." -ForegroundColor Yellow
Pause