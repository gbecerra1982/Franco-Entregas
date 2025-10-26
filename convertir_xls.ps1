# Script PowerShell para convertir archivos XLS corruptos a XLSX
# Uso: .\convertir_xls.ps1 "ruta\al\archivo.xls"

param(
    [Parameter(Mandatory=$false)]
    [string]$ArchivoXLS
)

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "   CONVERTIDOR DE ARCHIVOS XLS CORRUPTOS" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

# Si no se proporciona archivo, pedir al usuario
if ([string]::IsNullOrEmpty($ArchivoXLS)) {
    Write-Host "Por favor, arrastra el archivo .xls aquí o escribe la ruta:" -ForegroundColor Yellow
    $ArchivoXLS = Read-Host "Ruta del archivo"
    
    # Limpiar comillas si las hay
    $ArchivoXLS = $ArchivoXLS.Trim('"')
}

# Verificar que el archivo existe
if (-not (Test-Path $ArchivoXLS)) {
    Write-Host ""
    Write-Host "ERROR: El archivo no existe:" -ForegroundColor Red
    Write-Host "  $ArchivoXLS" -ForegroundColor Red
    Write-Host ""
    Write-Host "Presiona cualquier tecla para salir..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "Archivo encontrado:" -ForegroundColor Green
Write-Host "  $ArchivoXLS" -ForegroundColor White
Write-Host ""

# Método 1: Usar Python si está disponible
Write-Host "Método 1: Intentando conversión con Python..." -ForegroundColor Cyan

try {
    $pythonPath = Get-Command python -ErrorAction SilentlyContinue
    
    if ($pythonPath) {
        Write-Host "  Python encontrado: $($pythonPath.Source)" -ForegroundColor Green
        
        # Ejecutar script Python
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
        $pythonScript = Join-Path $scriptDir "convertir_xls_a_xlsx.py"
        
        if (Test-Path $pythonScript) {
            Write-Host "  Ejecutando script de conversión..." -ForegroundColor Yellow
            
            & python $pythonScript $ArchivoXLS
            
            $archivoSalida = $ArchivoXLS -replace '\.xls$', '_convertido.xlsx'
            
            if (Test-Path $archivoSalida) {
                Write-Host ""
                Write-Host "================================================" -ForegroundColor Green
                Write-Host "   CONVERSION EXITOSA!" -ForegroundColor Green
                Write-Host "================================================" -ForegroundColor Green
                Write-Host ""
                Write-Host "Archivo convertido guardado en:" -ForegroundColor Green
                Write-Host "  $archivoSalida" -ForegroundColor White
                Write-Host ""
                Write-Host "Tamaño: $((Get-Item $archivoSalida).Length) bytes" -ForegroundColor Gray
                Write-Host ""
                Write-Host "Ahora puedes usar este archivo en la aplicación Streamlit" -ForegroundColor Yellow
                Write-Host ""
                
                # Preguntar si abrir la carpeta
                $respuesta = Read-Host "¿Deseas abrir la carpeta del archivo? (S/N)"
                if ($respuesta -eq 'S' -or $respuesta -eq 's') {
                    explorer.exe /select,"$archivoSalida"
                }
                
                Write-Host ""
                Write-Host "Presiona cualquier tecla para salir..."
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                exit 0
            }
        }
    } else {
        Write-Host "  Python no encontrado" -ForegroundColor Yellow
    }
} catch {
    Write-Host "  Error con Python: $($_.Exception.Message)" -ForegroundColor Red
}

# Método 2: Usar Excel COM Object
Write-Host ""
Write-Host "Método 2: Intentando conversión con Excel COM..." -ForegroundColor Cyan

try {
    Write-Host "  Iniciando Excel..." -ForegroundColor Yellow
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    Write-Host "  Abriendo archivo..." -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Open($ArchivoXLS)
    
    $archivoSalida = $ArchivoXLS -replace '\.xls$', '_convertido.xlsx'
    
    Write-Host "  Guardando como XLSX..." -ForegroundColor Yellow
    # 51 = xlOpenXMLWorkbook (formato .xlsx)
    $workbook.SaveAs($archivoSalida, 51)
    
    $workbook.Close($false)
    $excel.Quit()
    
    # Limpiar objetos COM
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    if (Test-Path $archivoSalida) {
        Write-Host ""
        Write-Host "================================================" -ForegroundColor Green
        Write-Host "   CONVERSION EXITOSA CON EXCEL!" -ForegroundColor Green
        Write-Host "================================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "Archivo convertido guardado en:" -ForegroundColor Green
        Write-Host "  $archivoSalida" -ForegroundColor White
        Write-Host ""
        Write-Host "Tamaño: $((Get-Item $archivoSalida).Length) bytes" -ForegroundColor Gray
        Write-Host ""
        
        # Preguntar si abrir la carpeta
        $respuesta = Read-Host "¿Deseas abrir la carpeta del archivo? (S/N)"
        if ($respuesta -eq 'S' -or $respuesta -eq 's') {
            explorer.exe /select,"$archivoSalida"
        }
        
        Write-Host ""
        Write-Host "Presiona cualquier tecla para salir..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 0
    }
    
} catch {
    Write-Host "  Error con Excel: $($_.Exception.Message)" -ForegroundColor Red
}

# Si todos los métodos fallaron
Write-Host ""
Write-Host "===============================================" -ForegroundColor Red
Write-Host "   METODOS AUTOMATICOS FALLARON" -ForegroundColor Red
Write-Host "===============================================" -ForegroundColor Red
Write-Host ""
Write-Host "SOLUCION MANUAL:" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. Abre el archivo en Microsoft Excel:" -ForegroundColor White
Write-Host "   $ArchivoXLS" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Selecciona todo (Ctrl + A)" -ForegroundColor White
Write-Host ""
Write-Host "3. Copia (Ctrl + C)" -ForegroundColor White
Write-Host ""
Write-Host "4. Abre un NUEVO libro de Excel (Ctrl + N)" -ForegroundColor White
Write-Host ""
Write-Host "5. Pega (Ctrl + V)" -ForegroundColor White
Write-Host ""
Write-Host "6. Guarda como:" -ForegroundColor White
Write-Host "   - Archivo -> Guardar como" -ForegroundColor Gray
Write-Host "   - Tipo: 'Libro de Excel (*.xlsx)'" -ForegroundColor Gray
Write-Host ""

$archivoSugerido = $ArchivoXLS -replace '\.xls$', '_convertido.xlsx'
Write-Host "7. Nombre sugerido:" -ForegroundColor White
Write-Host "   $archivoSugerido" -ForegroundColor Gray
Write-Host ""
Write-Host "8. Cierra Excel y usa el nuevo archivo" -ForegroundColor White
Write-Host ""

Write-Host "Presiona cualquier tecla para salir..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
exit 1
