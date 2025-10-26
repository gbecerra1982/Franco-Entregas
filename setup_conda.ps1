# Script de PowerShell para crear entorno conda en Windows
# Uso: .\setup_conda.ps1

Write-Host "🐍 Configurando entorno conda para Analizador de Entregas Duplicadas" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""

# Nombre del entorno
$ENV_NAME = "entregas_duplicadas"

# Verificar que conda esté instalado
try {
    $condaVersion = conda --version 2>$null
    Write-Host "✅ Conda encontrado: $condaVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Conda no está instalado." -ForegroundColor Red
    Write-Host ""
    Write-Host "Para instalar Miniconda en Windows:" -ForegroundColor Yellow
    Write-Host "  Descarga desde: https://docs.conda.io/en/latest/miniconda.html" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host ""

# Verificar si el entorno ya existe
$envList = conda env list | Select-String -Pattern "^$ENV_NAME "

if ($envList) {
    Write-Host "⚠️  El entorno '$ENV_NAME' ya existe." -ForegroundColor Yellow
    $respuesta = Read-Host "¿Deseas eliminarlo y recrearlo? (s/n)"
    if ($respuesta -eq 's' -or $respuesta -eq 'S') {
        Write-Host "🗑️  Eliminando entorno existente..." -ForegroundColor Yellow
        conda env remove -n $ENV_NAME -y
    } else {
        Write-Host "ℹ️  Usando entorno existente." -ForegroundColor Blue
        conda activate $ENV_NAME
        exit 0
    }
}

Write-Host "📦 Creando entorno conda: $ENV_NAME" -ForegroundColor Cyan
Write-Host "   (Python 3.11)" -ForegroundColor Cyan
Write-Host ""

# Crear el entorno con Python 3.11
conda create -n $ENV_NAME python=3.11 -y

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Error al crear el entorno conda" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "✅ Entorno creado exitosamente" -ForegroundColor Green
Write-Host ""
Write-Host "🔧 Activando entorno..." -ForegroundColor Cyan

# Activar el entorno
conda activate $ENV_NAME

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Error al activar el entorno" -ForegroundColor Red
    exit 1
}

Write-Host "✅ Entorno activado: $ENV_NAME" -ForegroundColor Green
Write-Host ""
Write-Host "📦 Instalando dependencias..." -ForegroundColor Cyan
Write-Host ""

# Instalar paquetes con conda
Write-Host "→ Instalando pandas, numpy, openpyxl, xlrd, xlwt, pyxlsb..." -ForegroundColor Yellow
conda install -c conda-forge pandas numpy openpyxl xlrd xlwt pyxlsb python-dateutil -y

Write-Host ""
Write-Host "→ Instalando streamlit..." -ForegroundColor Yellow
conda install -c conda-forge streamlit -y

Write-Host ""
Write-Host "→ Instalando matplotlib (para gráficos)..." -ForegroundColor Yellow
conda install -c conda-forge matplotlib -y

Write-Host ""
Write-Host "→ Instalando pywin32 (conversión avanzada de archivos .xls en Windows)..." -ForegroundColor Yellow
pip install pywin32

# Verificar instalación
Write-Host ""
Write-Host "🔍 Verificando instalación..." -ForegroundColor Cyan
python -c "import pandas; print(f'✅ pandas {pandas.__version__}')"
python -c "import numpy; print(f'✅ numpy {numpy.__version__}')"
python -c "import openpyxl; print(f'✅ openpyxl {openpyxl.__version__}')"
python -c "import xlrd; print(f'✅ xlrd {xlrd.__version__}')"
python -c "import xlwt; print(f'✅ xlwt {xlwt.__version__}')"
python -c "import pyxlsb; print(f'✅ pyxlsb {pyxlsb.__version__}')"
python -c "import streamlit; print(f'✅ streamlit {streamlit.__version__}')"
python -c "import matplotlib; print(f'✅ matplotlib {matplotlib.__version__}')"
python -c "import win32com.client; print('✅ pywin32 (win32com disponible)')"

Write-Host ""
Write-Host "✨ ¡Instalación completada exitosamente!" -ForegroundColor Green
Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "📖 CÓMO USAR EL ENTORNO:" -ForegroundColor Cyan
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Activar el entorno:" -ForegroundColor White
Write-Host "   conda activate $ENV_NAME" -ForegroundColor Yellow
Write-Host ""
Write-Host "2. Ejecutar la aplicación web:" -ForegroundColor White
Write-Host "   streamlit run app_entregas.py" -ForegroundColor Yellow
Write-Host ""
Write-Host "3. O ejecutar el script de línea de comandos:" -ForegroundColor White
Write-Host "   python procesar_entregas.py tu_archivo.xlsx" -ForegroundColor Yellow
Write-Host ""
Write-Host "4. Desactivar el entorno cuando termines:" -ForegroundColor White
Write-Host "   conda deactivate" -ForegroundColor Yellow
Write-Host ""
Write-Host "5. Para eliminar el entorno en el futuro:" -ForegroundColor White
Write-Host "   conda env remove -n $ENV_NAME" -ForegroundColor Yellow
Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
