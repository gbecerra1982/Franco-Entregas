#!/usr/bin/env python3
"""
Script para detectar y corregir entregas duplicadas en archivos de liquidación

Uso:
    python procesar_entregas.py                          # Procesa todos los XLS/XLSX en el directorio actual
    python procesar_entregas.py <archivo.xls>            # Procesa un archivo específico
    python procesar_entregas.py <directorio>             # Procesa todos los XLS/XLSX en el directorio
    python procesar_entregas.py <directorio> --recursivo # Procesa recursivamente subdirectorios
"""

import pandas as pd
from datetime import datetime
import sys
import os
import re
import tempfile
import glob

def detectar_formato_excel(file_path):
    """Detecta el formato del archivo Excel por firma binaria."""
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
            if header[:4] == b'\xD0\xCF\x11\xE0':
                return "xls_antiguo"
            if header[:4] == b'\x50\x4B\x03\x04':
                return "xlsx_moderno"
            return "desconocido"
    except:
        return "desconocido"

def convertir_xls_a_xlsx_metodo1(input_path, output_path):
    """
    Método 1: Conversión usando pandas con diferentes motores.
    Returns: bool - True si la conversión fue exitosa
    """
    try:
        engines = ['xlrd', 'openpyxl', None]

        for engine in engines:
            try:
                if engine == 'xlrd':
                    # Intentar con xlrd deshabilitando validaciones estrictas
                    import xlrd
                    xlrd.Book.logfile = open(os.devnull, 'w')
                    df = pd.read_excel(input_path, engine=engine)
                elif engine:
                    df = pd.read_excel(input_path, engine=engine)
                else:
                    df = pd.read_excel(input_path)

                # Guardar como XLSX
                df.to_excel(output_path, index=False, engine='openpyxl')
                print(f"✅ Conversión exitosa con motor: {engine or 'default'}")
                return True
            except Exception as e:
                # Silenciosamente intentar siguiente método
                continue

        return False
    except Exception as e:
        return False

def convertir_xls_a_xlsx_metodo2(input_path, output_path):
    """
    Método 2: Conversión usando xlrd con validación relajada y manejo de errores mejorado.
    Returns: bool - True si la conversión fue exitosa
    """
    try:
        import xlrd
        import io

        # Deshabilitar logging de xlrd para evitar spam en consola
        old_stderr = sys.stderr
        sys.stderr = io.StringIO()

        try:
            # Intentar abrir con diferentes configuraciones
            configs = [
                {'formatting_info': False, 'on_demand': True, 'ignore_workbook_corruption': True},
                {'formatting_info': False, 'on_demand': False},
                {'formatting_info': False},
            ]

            book = None
            for config in configs:
                try:
                    book = xlrd.open_workbook(input_path, **config)
                    break
                except:
                    continue

            if book is None:
                raise Exception("No se pudo abrir el archivo con xlrd")

            # Leer todas las hojas
            all_sheets = []
            for sheet_name in book.sheet_names():
                try:
                    sheet = book.sheet_by_name(sheet_name)

                    # Convertir a lista de listas
                    data = []
                    for row_idx in range(sheet.nrows):
                        row = []
                        for col_idx in range(sheet.ncols):
                            try:
                                cell = sheet.cell(row_idx, col_idx)
                                row.append(cell.value)
                            except:
                                row.append('')
                        data.append(row)

                    df = pd.DataFrame(data)
                    all_sheets.append((sheet_name, df))
                except:
                    continue

            if not all_sheets:
                raise Exception("No se pudo leer ninguna hoja")

            # Guardar como XLSX
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in all_sheets:
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

            print("✅ Conversión exitosa con xlrd (método alternativo)")
            return True

        finally:
            # Restaurar stderr
            sys.stderr = old_stderr

    except Exception as e:
        return False

def convertir_xls_a_xlsx_metodo3(input_path, output_path):
    """
    Método 3: Conversión usando LibreOffice (si está instalado).
    Returns: bool - True si la conversión fue exitosa
    """
    try:
        import subprocess

        # Buscar LibreOffice en ubicaciones comunes
        libreoffice_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/usr/bin/libreoffice",
            "/usr/local/bin/libreoffice"
        ]

        libreoffice_path = None
        for path in libreoffice_paths:
            if os.path.exists(path):
                libreoffice_path = path
                break

        if not libreoffice_path:
            return False

        output_dir = os.path.dirname(output_path)

        cmd = [
            libreoffice_path,
            "--headless",
            "--convert-to", "xlsx",
            "--outdir", output_dir,
            input_path
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        # Verificar si se creó el archivo
        expected_output = os.path.join(output_dir, os.path.basename(input_path).replace('.xls', '.xlsx'))

        if os.path.exists(expected_output):
            if expected_output != output_path:
                import shutil
                shutil.move(expected_output, output_path)
            print("✅ Conversión exitosa con LibreOffice")
            return True

        return False

    except Exception as e:
        return False

def convertir_xls_a_xlsx_metodo4(input_path, output_path):
    """
    Método 4: Conversión usando win32com (Excel COM automation en Windows).
    El más confiable en Windows si Excel está instalado.
    Returns: bool - True si la conversión fue exitosa
    """
    try:
        # Solo disponible en Windows
        if os.name != 'nt':
            return False

        import win32com.client
        import pythoncom

        # Inicializar COM
        pythoncom.CoInitialize()

        try:
            # Crear instancia de Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            # Convertir a ruta absoluta
            abs_input = os.path.abspath(input_path)
            abs_output = os.path.abspath(output_path)

            # Abrir archivo XLS
            workbook = excel.Workbooks.Open(abs_input)

            # Guardar como XLSX (formato 51)
            workbook.SaveAs(abs_output, FileFormat=51)

            # Cerrar
            workbook.Close(SaveChanges=False)
            excel.Quit()

            print("✅ Conversión exitosa con Excel COM (win32com)")
            return True

        finally:
            pythoncom.CoUninitialize()

    except Exception as e:
        return False

def intentar_conversion_automatica(input_path):
    """
    Intenta convertir archivo XLS antiguo a XLSX usando múltiples métodos.

    Args:
        input_path: Ruta del archivo XLS original

    Returns:
        str: Ruta del archivo XLSX convertido, o None si todos los métodos fallaron
    """
    # Crear archivo temporal para output
    output_path = input_path.replace('.xls', '_temp_convertido.xlsx')

    print("🔄 Intentando conversión automática de formato antiguo...")

    # Método 4: Win32com (Excel COM) - El más confiable en Windows
    if os.name == 'nt':
        print("   Método 1: Usando Microsoft Excel (COM)...")
        if convertir_xls_a_xlsx_metodo4(input_path, output_path):
            return output_path

    # Método 1: Pandas con diferentes motores
    print("   Método 2: Usando pandas...")
    if convertir_xls_a_xlsx_metodo1(input_path, output_path):
        return output_path

    # Método 2: xlrd con validación relajada
    print("   Método 3: Usando xlrd con validación relajada...")
    if convertir_xls_a_xlsx_metodo2(input_path, output_path):
        return output_path

    # Método 3: LibreOffice
    print("   Método 4: Intentando con LibreOffice...")
    if convertir_xls_a_xlsx_metodo3(input_path, output_path):
        return output_path

    # Si todos los métodos fallaron
    print("❌ No se pudo convertir el archivo automáticamente")
    print("\n💡 SOLUCIÓN:")
    print("1. Abre el archivo en Microsoft Excel")
    print("2. Selecciona todo (Ctrl+A)")
    print("3. Copia (Ctrl+C)")
    print("4. Crea un nuevo libro de Excel")
    print("5. Pega (Ctrl+V)")
    print("6. Guarda como .xlsx")
    return None

def normalizar_direccion(direccion):
    """Normaliza una dirección para comparación"""
    if pd.isna(direccion):
        return ""

    direccion = str(direccion).lower().strip()
    # Remover múltiples espacios
    direccion = re.sub(r'\s+', ' ', direccion)
    # Remover puntos y comas
    direccion = direccion.replace(',', '').replace('.', '')

    return direccion

def convertir_fecha(fecha_str):
    """Convierte string de fecha a datetime"""
    try:
        if pd.isna(fecha_str):
            return None
        return pd.to_datetime(fecha_str, format='%d/%m/%Y %H:%M')
    except:
        try:
            return pd.to_datetime(fecha_str, format='%d/%m/%Y')
        except:
            return None

def procesar_archivo(file_path):
    """Lee y procesa el archivo Excel con conversión automática de XLS"""
    print(f"📖 Leyendo archivo: {os.path.basename(file_path)}")

    archivo_convertido = None

    # Verificar si necesita conversión
    formato = detectar_formato_excel(file_path)
    extension = os.path.splitext(file_path)[1].lower()

    if formato == "xls_antiguo" or (extension == '.xls' and formato != "xlsx_moderno"):
        # Usar la función de conversión con múltiples métodos
        archivo_convertido = intentar_conversion_automatica(file_path)
        if archivo_convertido:
            file_path = archivo_convertido
        else:
            print(f"❌ No se pudo convertir: {os.path.basename(file_path)}")
            return None

    try:
        df = pd.read_excel(file_path)

        # La primera fila contiene los encabezados reales
        headers = df.iloc[0].tolist()
        df = df[1:].reset_index(drop=True)
        df.columns = headers

        print(f"✅ Archivo cargado: {len(df)} entregas")
        return df
    except Exception as e:
        print(f"❌ Error al leer archivo: {str(e)[:100]}")
        return None
    finally:
        # Limpiar archivo temporal convertido
        if archivo_convertido and os.path.exists(archivo_convertido):
            try:
                os.remove(archivo_convertido)
            except:
                pass

def analizar_duplicados(df):
    """Analiza y marca entregas duplicadas"""
    print("🔍 Analizando duplicados...")
    
    # Crear columna de dirección normalizada
    df['Dirección_Normalizada'] = df['Dirección'].apply(normalizar_direccion)
    
    # Convertir fechas
    df['Fecha_DateTime'] = df['Fecha Movimiento'].apply(convertir_fecha)
    df['Fecha_Solo'] = df['Fecha_DateTime'].dt.date
    
    # Ordenar por dirección y fecha
    df = df.sort_values(['Dirección_Normalizada', 'Fecha_DateTime']).reset_index(drop=True)
    
    # Crear columnas de análisis
    df['Es_Duplicado'] = False
    df['Costo_Original'] = df['Costo envio']
    df['Costo_Ajustado'] = df['Costo envio']
    df['Motivo_Ajuste'] = ''
    df['Grupo_Duplicado'] = None
    
    # Identificar duplicados
    grupo_actual = 0
    
    for idx in range(len(df)):
        if idx == 0:
            continue
            
        direccion_actual = df.loc[idx, 'Dirección_Normalizada']
        direccion_anterior = df.loc[idx-1, 'Dirección_Normalizada']
        
        fecha_actual = df.loc[idx, 'Fecha_Solo']
        fecha_anterior = df.loc[idx-1, 'Fecha_Solo']
        
        # Si es la misma dirección
        if direccion_actual == direccion_anterior and direccion_actual != '':
            # Calcular diferencia de días
            if fecha_actual and fecha_anterior:
                diferencia_dias = (fecha_actual - fecha_anterior).days
                
                # Si es el mismo día o 1-2 días después
                if 0 <= diferencia_dias <= 2:
                    # Marcar como duplicado
                    if df.loc[idx-1, 'Grupo_Duplicado'] is None:
                        grupo_actual += 1
                        df.loc[idx-1, 'Grupo_Duplicado'] = grupo_actual
                    
                    df.loc[idx, 'Es_Duplicado'] = True
                    df.loc[idx, 'Grupo_Duplicado'] = grupo_actual
                    
                    # Ajustar costo a 0
                    costo_actual = df.loc[idx, 'Costo envio']
                    if costo_actual and costo_actual > 0:
                        df.loc[idx, 'Costo_Ajustado'] = 0
                        df.loc[idx, 'Motivo_Ajuste'] = f'Entrega duplicada (día {diferencia_dias + 1})'
    
    return df

def generar_reporte(df_analizado, archivo_original, output_dir=None):
    """Genera reporte y archivos de salida"""
    
    # Calcular estadísticas
    total_entregas = len(df_analizado)
    total_duplicados = df_analizado['Es_Duplicado'].sum()
    total_grupos = df_analizado['Grupo_Duplicado'].nunique() - (1 if None in df_analizado['Grupo_Duplicado'].values else 0)
    
    # Calcular ahorro
    df_analizado['Costo_Original_Num'] = pd.to_numeric(df_analizado['Costo_Original'], errors='coerce').fillna(0)
    df_analizado['Costo_Ajustado_Num'] = pd.to_numeric(df_analizado['Costo_Ajustado'], errors='coerce').fillna(0)
    ahorro_total = (df_analizado['Costo_Original_Num'] - df_analizado['Costo_Ajustado_Num']).sum()
    
    print("\n" + "="*60)
    print("📊 RESUMEN DEL ANÁLISIS")
    print("="*60)
    print(f"📦 Total de entregas analizadas: {total_entregas}")
    print(f"🔄 Entregas duplicadas detectadas: {total_duplicados}")
    print(f"👥 Grupos de duplicados: {total_grupos}")
    print(f"💰 Ahorro total: ${ahorro_total:,.2f}")
    print("="*60)
    
    # Determinar directorio de salida
    if output_dir is None:
        output_dir = os.path.dirname(archivo_original) or '.'
    
    # Guardar archivos
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = os.path.splitext(os.path.basename(archivo_original))[0]
    
    # Archivo completo con análisis
    archivo_completo = os.path.join(output_dir, f"{base_name}_analisis_{timestamp}.xlsx")
    df_analizado.to_excel(archivo_completo, index=False)
    print(f"\n✅ Archivo completo guardado: {os.path.basename(archivo_completo)}")
    
    # Archivo solo con duplicados
    if total_duplicados > 0:
        df_duplicados = df_analizado[df_analizado['Es_Duplicado'] == True]
        archivo_duplicados = os.path.join(output_dir, f"{base_name}_duplicados_{timestamp}.xlsx")
        df_duplicados.to_excel(archivo_duplicados, index=False)
        print(f"✅ Duplicados guardados: {os.path.basename(archivo_duplicados)}")
        
        # Mostrar algunos ejemplos
        print("\n📋 EJEMPLOS DE ENTREGAS DUPLICADAS:")
        print("-" * 60)
        for grupo in sorted(df_duplicados['Grupo_Duplicado'].unique())[:3]:
            if pd.isna(grupo):
                continue
            
            grupo_df = df_analizado[df_analizado['Grupo_Duplicado'] == grupo]
            direccion = grupo_df['Dirección'].iloc[0][:50]
            
            print(f"\n🏠 Grupo {int(grupo)}: {direccion}...")
            for _, row in grupo_df.iterrows():
                fecha = row['Fecha Movimiento']
                costo_orig = row['Costo_Original']
                costo_ajust = row['Costo_Ajustado']
                estado = row['Estado']
                print(f"   • {fecha} - Estado: {estado} - ${costo_orig} → ${costo_ajust}")
    
    # Resumen por dirección
    resumen = df_analizado.groupby('Dirección_Normalizada').agg({
        'Fecha Movimiento': 'count',
        'Es_Duplicado': 'sum',
        'Costo_Original_Num': 'sum',
        'Costo_Ajustado_Num': 'sum'
    }).reset_index()
    
    resumen.columns = ['Dirección', 'Total Entregas', 'Entregas Duplicadas', 'Costo Original', 'Costo Ajustado']
    resumen['Ahorro'] = resumen['Costo Original'] - resumen['Costo Ajustado']
    resumen_duplicados = resumen[resumen['Entregas Duplicadas'] > 0].sort_values('Ahorro', ascending=False)
    
    if len(resumen_duplicados) > 0:
        archivo_resumen = os.path.join(output_dir, f"{base_name}_resumen_{timestamp}.xlsx")
        resumen_duplicados.to_excel(archivo_resumen, index=False)
        print(f"✅ Resumen por dirección guardado: {os.path.basename(archivo_resumen)}")
        
        print("\n💰 TOP 5 DIRECCIONES CON MAYOR AHORRO:")
        print("-" * 60)
        for idx, row in resumen_duplicados.head(5).iterrows():
            print(f"{row['Dirección'][:50]}...")
            print(f"   Entregas: {int(row['Total Entregas'])} | Duplicados: {int(row['Entregas Duplicadas'])} | Ahorro: ${row['Ahorro']:,.2f}")
    
    print("\n✨ Archivo procesado exitosamente!\n")
    
    return {
        'total_entregas': total_entregas,
        'total_duplicados': total_duplicados,
        'ahorro_total': ahorro_total
    }

def buscar_archivos_excel(directorio, recursivo=False):
    """Busca archivos Excel (.xls y .xlsx) en el directorio especificado"""
    archivos = []
    
    if recursivo:
        # Búsqueda recursiva
        for ext in ['*.xls', '*.xlsx']:
            pattern = os.path.join(directorio, '**', ext)
            archivos.extend(glob.glob(pattern, recursive=True))
    else:
        # Solo en el directorio actual
        for ext in ['*.xls', '*.xlsx']:
            pattern = os.path.join(directorio, ext)
            archivos.extend(glob.glob(pattern))
    
    # Excluir archivos que son resultados de este script
    archivos_filtrados = []
    for archivo in archivos:
        nombre = os.path.basename(archivo)
        # Excluir archivos que contienen _analisis_, _duplicados_, _resumen_, _temp_convertido
        if not any(x in nombre for x in ['_analisis_', '_duplicados_', '_resumen_', '_temp_convertido']):
            archivos_filtrados.append(archivo)
    
    return sorted(archivos_filtrados)

def procesar_multiple(archivos):
    """Procesa múltiples archivos y genera resumen consolidado"""
    print(f"\n🚀 Procesando {len(archivos)} archivo(s)")
    print("="*70)
    
    resultados = []
    archivos_exitosos = 0
    archivos_fallidos = 0
    
    for i, archivo in enumerate(archivos, 1):
        print(f"\n📁 [{i}/{len(archivos)}] Procesando: {os.path.basename(archivo)}")
        print("-"*70)
        
        try:
            # Procesar archivo
            df_original = procesar_archivo(archivo)
            if df_original is None:
                archivos_fallidos += 1
                print(f"⚠️  Saltando archivo por error en lectura")
                continue
            
            # Analizar duplicados
            df_analizado = analizar_duplicados(df_original)
            
            # Generar reporte
            stats = generar_reporte(df_analizado, archivo)
            
            resultados.append({
                'archivo': os.path.basename(archivo),
                'entregas': stats['total_entregas'],
                'duplicados': stats['total_duplicados'],
                'ahorro': stats['ahorro_total']
            })
            
            archivos_exitosos += 1
            
        except Exception as e:
            print(f"❌ Error procesando archivo: {str(e)[:100]}")
            archivos_fallidos += 1
            continue
    
    # Resumen consolidado
    if len(resultados) > 1:
        print("\n" + "="*70)
        print("📊 RESUMEN CONSOLIDADO DE TODOS LOS ARCHIVOS")
        print("="*70)
        
        total_entregas = sum(r['entregas'] for r in resultados)
        total_duplicados = sum(r['duplicados'] for r in resultados)
        total_ahorro = sum(r['ahorro'] for r in resultados)
        
        print(f"\n✅ Archivos procesados exitosamente: {archivos_exitosos}")
        if archivos_fallidos > 0:
            print(f"❌ Archivos con errores: {archivos_fallidos}")
        
        print(f"\n📦 Total de entregas procesadas: {total_entregas:,}")
        print(f"🔄 Total de duplicados detectados: {total_duplicados:,}")
        print(f"💰 Ahorro total acumulado: ${total_ahorro:,.2f}")
        
        print("\n📋 Detalle por archivo:")
        print("-"*70)
        for r in resultados:
            print(f"\n📄 {r['archivo']}")
            print(f"   Entregas: {r['entregas']:,} | Duplicados: {r['duplicados']:,} | Ahorro: ${r['ahorro']:,.2f}")
        
        print("\n" + "="*70)

def main():
    print("🚀 Analizador de Entregas Duplicadas - Procesamiento Masivo")
    print("="*70)
    
    # Determinar qué procesar
    recursivo = '--recursivo' in sys.argv or '-r' in sys.argv
    
    if len(sys.argv) < 2 or sys.argv[1] in ['--recursivo', '-r']:
        # Sin argumentos o solo --recursivo: usar directorio actual
        directorio = os.getcwd()
        print(f"📂 Buscando archivos Excel en: {directorio}")
        if recursivo:
            print("   (Búsqueda recursiva en subdirectorios)")
        
        archivos = buscar_archivos_excel(directorio, recursivo)
        
        if not archivos:
            print("\n❌ No se encontraron archivos Excel (.xls/.xlsx) en el directorio")
            print("\n💡 Uso:")
            print("   python procesar_entregas.py                    # Directorio actual")
            print("   python procesar_entregas.py <archivo.xls>      # Archivo específico")
            print("   python procesar_entregas.py <directorio>       # Directorio específico")
            print("   python procesar_entregas.py --recursivo        # Recursivo en directorio actual")
            sys.exit(1)
        
        print(f"✅ Encontrados {len(archivos)} archivo(s) para procesar")
        
        # Mostrar archivos
        for i, archivo in enumerate(archivos, 1):
            print(f"   {i}. {os.path.basename(archivo)}")
        
        # Procesar múltiples archivos
        procesar_multiple(archivos)
        
    else:
        # Con argumentos: verificar si es archivo o directorio
        ruta = sys.argv[1]
        
        if not os.path.exists(ruta):
            print(f"❌ Error: '{ruta}' no existe")
            sys.exit(1)
        
        if os.path.isfile(ruta):
            # Procesar archivo individual
            print(f"📄 Procesando archivo: {os.path.basename(ruta)}")
            print("="*70)
            
            df_original = procesar_archivo(ruta)
            if df_original is None:
                sys.exit(1)
            
            df_analizado = analizar_duplicados(df_original)
            generar_reporte(df_analizado, ruta)
            
        elif os.path.isdir(ruta):
            # Procesar directorio
            print(f"📂 Buscando archivos Excel en: {ruta}")
            if recursivo:
                print("   (Búsqueda recursiva en subdirectorios)")
            
            archivos = buscar_archivos_excel(ruta, recursivo)
            
            if not archivos:
                print(f"\n❌ No se encontraron archivos Excel en el directorio: {ruta}")
                sys.exit(1)
            
            print(f"✅ Encontrados {len(archivos)} archivo(s) para procesar")
            
            # Mostrar archivos
            for i, archivo in enumerate(archivos, 1):
                print(f"   {i}. {os.path.basename(archivo)}")
            
            # Procesar múltiples archivos
            procesar_multiple(archivos)
        
        else:
            print(f"❌ Error: '{ruta}' no es un archivo ni directorio válido")
            sys.exit(1)

if __name__ == "__main__":
    main()
