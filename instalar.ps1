# Script de PowerShell para instalar dependencias con pip
# Uso: .\instalar.ps1

Write-Host "🚀 Instalando Analizador de Entregas Duplicadas..." -ForegroundColor Cyan
Write-Host ""

# Verificar si Python está instalado
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✅ Python encontrado: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Python no está instalado. Por favor, instala Python 3.8 o superior primero." -ForegroundColor Red
    Write-Host ""
    Write-Host "Descarga Python desde: https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host ""

# Verificar si pip está instalado
try {
    $pipVersion = pip --version 2>&1
    Write-Host "✅ pip encontrado" -ForegroundColor Green
} catch {
    Write-Host "❌ pip no está instalado. Por favor, instala pip primero." -ForegroundColor Red
    exit 1
}

Write-Host ""

# Instalar dependencias de Python
Write-Host "📦 Instalando dependencias de Python desde requirements.txt..." -ForegroundColor Cyan
pip install -r requirements.txt

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Error al instalar dependencias de Python" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "✅ Dependencias de Python instaladas" -ForegroundColor Green
Write-Host ""

Write-Host "✨ ¡Instalación completada!" -ForegroundColor Green
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "📖 CÓMO USAR:" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Aplicación Web (Interfaz gráfica):" -ForegroundColor White
Write-Host "   streamlit run app_entregas.py" -ForegroundColor Yellow
Write-Host ""
Write-Host "2. Línea de Comandos:" -ForegroundColor White
Write-Host "   python procesar_entregas.py tu_archivo.xlsx" -ForegroundColor Yellow
Write-Host ""
Write-Host "Para más información, lee el archivo README.md" -ForegroundColor White
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
