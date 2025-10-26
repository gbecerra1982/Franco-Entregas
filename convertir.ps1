# Script de PowerShell para ejecutar el procesador masivo de entregas duplicadas
# Uso: 
#   .\convertir.ps1                                    # Procesa todos los XLS/XLSX en el directorio actual
#   .\convertir.ps1 "archivo.xls"                      # Procesa un archivo específico
#   .\convertir.ps1 "C:\ruta\directorio"               # Procesa todos los archivos en un directorio
#   .\convertir.ps1 "C:\ruta\directorio" --recursivo   # Procesa recursivamente subdirectorios

param(
    [Parameter(Mandatory=$false)]
    [string]$Ruta = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Recursivo
)

Write-Host ""
Write-Host "🔄 Procesador Masivo de Entregas Duplicadas" -ForegroundColor Cyan
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

# Verificar que el archivo procesar_entregas.py existe
if (-not (Test-Path "procesar_entregas.py")) {
    Write-Host "❌ Error: No se encuentra el archivo procesar_entregas.py" -ForegroundColor Red
    Write-Host ""
    Write-Host "Asegúrate de ejecutar este script desde el directorio del proyecto:" -ForegroundColor Yellow
    Write-Host "  cd C:\ruta\al\proyecto\Franco-Entregas" -ForegroundColor Yellow
    Write-Host "  .\convertir.ps1" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Activar el entorno
Write-Host "🔧 Activando entorno: $ENV_NAME..." -ForegroundColor Yellow
Write-Host "✅ Entorno activado" -ForegroundColor Green
Write-Host ""

# Construir el comando
$comando = "python procesar_entregas.py"

if ($Ruta -ne "") {
    # Si la ruta tiene espacios, agregar comillas
    if ($Ruta -match "\s") {
        $comando += " `"$Ruta`""
    } else {
        $comando += " $Ruta"
    }
}

if ($Recursivo) {
    $comando += " --recursivo"
}

# Mostrar información de uso
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "📋 MODOS DE USO:" -ForegroundColor White
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Procesar todos los archivos en el directorio actual:" -ForegroundColor White
Write-Host "   .\convertir.ps1" -ForegroundColor Yellow
Write-Host ""
Write-Host "2. Procesar un archivo específico:" -ForegroundColor White
Write-Host "   .\convertir.ps1 `"archivo.xls`"" -ForegroundColor Yellow
Write-Host ""
Write-Host "3. Procesar todos los archivos en un directorio:" -ForegroundColor White
Write-Host "   .\convertir.ps1 `"C:\ruta\al\directorio`"" -ForegroundColor Yellow
Write-Host ""
Write-Host "4. Procesar recursivamente (incluye subdirectorios):" -ForegroundColor White
Write-Host "   .\convertir.ps1 -Recursivo" -ForegroundColor Yellow
Write-Host "   .\convertir.ps1 `"C:\ruta\directorio`" -Recursivo" -ForegroundColor Yellow
Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""

# Mostrar comando a ejecutar
Write-Host "🚀 Ejecutando comando:" -ForegroundColor Cyan
Write-Host "   $comando" -ForegroundColor Yellow
Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""

# Ejecutar procesar_entregas.py usando conda run
try {
    if ($Ruta -eq "" -and -not $Recursivo) {
        # Sin argumentos - procesar directorio actual
        & conda run -n $ENV_NAME python procesar_entregas.py
    }
    elseif ($Ruta -ne "" -and -not $Recursivo) {
        # Con ruta pero sin recursivo
        & conda run -n $ENV_NAME python procesar_entregas.py $Ruta
    }
    elseif ($Ruta -eq "" -and $Recursivo) {
        # Sin ruta pero con recursivo
        & conda run -n $ENV_NAME python procesar_entregas.py --recursivo
    }
    else {
        # Con ruta y recursivo
        & conda run -n $ENV_NAME python procesar_entregas.py $Ruta --recursivo
    }
} catch {
    Write-Host ""
    Write-Host "❌ Error al ejecutar el procesamiento: $_" -ForegroundColor Red
    Write-Host ""
    Read-Host "Presiona Enter para salir"
    exit 1
}

# Finalización
Write-Host ""
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host "✅ Procesamiento completado" -ForegroundColor Green
Write-Host "==================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📁 Los archivos generados se encuentran en el mismo directorio" -ForegroundColor White
Write-Host "   de los archivos originales con los sufijos:" -ForegroundColor White
Write-Host ""
Write-Host "   • *_analisis_*.xlsx    - Archivo completo con análisis" -ForegroundColor Yellow
Write-Host "   • *_duplicados_*.xlsx  - Solo entregas duplicadas" -ForegroundColor Yellow
Write-Host "   • *_resumen_*.xlsx     - Resumen por dirección" -ForegroundColor Yellow
Write-Host ""
Write-Host "👋 Presiona Enter para salir" -ForegroundColor Cyan
Read-Host
