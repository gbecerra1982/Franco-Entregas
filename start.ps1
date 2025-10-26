# Script de PowerShell para iniciar la aplicación de Análisis de Entregas Duplicadas
# Uso: .\start.ps1

Write-Host ""
Write-Host "🚀 Iniciando Analizador de Entregas Duplicadas" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""

# Nombre del entorno
$ENV_NAME = "entregas_duplicadas"

# Verificar que conda esté instalado
try {
    $null = conda --version 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "Conda no encontrado"
    }
} catch {
    Write-Host "❌ Error: Conda no está instalado o no está en el PATH" -ForegroundColor Red
    Write-Host ""
    Write-Host "Para instalar Miniconda en Windows:" -ForegroundColor Yellow
    Write-Host "  1. Descarga desde: https://docs.conda.io/en/latest/miniconda.html" -ForegroundColor Yellow
    Write-Host "  2. Ejecuta el instalador" -ForegroundColor Yellow
    Write-Host "  3. Reinicia PowerShell" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Verificar si el entorno existe
$envList = conda env list | Select-String -Pattern "^$ENV_NAME "

if (-not $envList) {
    Write-Host "❌ Error: El entorno conda '$ENV_NAME' no existe" -ForegroundColor Red
    Write-Host ""
    Write-Host "Para crear el entorno, ejecuta primero:" -ForegroundColor Yellow
    Write-Host "  .\setup_conda.ps1" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Obtener la ruta base de conda
$condaBase = conda info --base 2>$null
if (-not $condaBase) {
    Write-Host "❌ Error: No se pudo determinar la ruta base de conda" -ForegroundColor Red
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Activar el entorno usando conda run (más confiable que activate en scripts)
Write-Host "🔧 Activando entorno: $ENV_NAME..." -ForegroundColor Yellow

# Verificar que el archivo app_entregas.py existe
if (-not (Test-Path "app_entregas.py")) {
    Write-Host "❌ Error: No se encuentra el archivo app_entregas.py" -ForegroundColor Red
    Write-Host ""
    Write-Host "Asegúrate de ejecutar este script desde el directorio del proyecto:" -ForegroundColor Yellow
    Write-Host "  cd C:\ruta\al\proyecto\Franco-Entregas" -ForegroundColor Yellow
    Write-Host "  .\start.ps1" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Presiona Enter para salir"
    exit 1
}

Write-Host "✅ Entorno activado" -ForegroundColor Green
Write-Host ""
Write-Host "🌐 Iniciando aplicación web Streamlit..." -ForegroundColor Cyan
Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "La aplicación se abrirá en tu navegador en:" -ForegroundColor White
Write-Host "  http://localhost:8501" -ForegroundColor Yellow
Write-Host ""
Write-Host "Para detener la aplicación, presiona Ctrl+C en esta ventana" -ForegroundColor White
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""

# Ejecutar streamlit usando conda run
try {
    & conda run -n $ENV_NAME streamlit run app_entregas.py
} catch {
    Write-Host ""
    Write-Host "❌ Error al ejecutar la aplicación: $_" -ForegroundColor Red
    Write-Host ""
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Si el usuario presionó Ctrl+C o la app terminó
Write-Host ""
Write-Host "👋 Aplicación cerrada" -ForegroundColor Yellow
Write-Host ""
