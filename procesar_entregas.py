#!/usr/bin/env python3
"""
Script para detectar y corregir entregas duplicadas en archivos de liquidación
Uso: python procesar_entregas.py <archivo_excel>
"""

import pandas as pd
from datetime import datetime
import sys
import os
import re
import tempfile

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

def convertir_xls_a_xlsx(input_path):
    """
    Convierte archivo XLS a XLSX usando el método más confiable disponible.
    Returns: Ruta del archivo convertido o None si falla.
    """
    output_path = input_path.replace('.xls', '_temp_convertido.xlsx')

    print("🔄 Detectado archivo XLS antiguo, convirtiendo...")

    # Método 1: Win32com (Excel COM) - El más confiable en Windows
    if os.name == 'nt':
        try:
            import win32com.client
            import pythoncom

            pythoncom.CoInitialize()
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False

                abs_input = os.path.abspath(input_path)
                abs_output = os.path.abspath(output_path)

                workbook = excel.Workbooks.Open(abs_input)
                workbook.SaveAs(abs_output, FileFormat=51)
                workbook.Close(SaveChanges=False)
                excel.Quit()

                print("✅ Conversión exitosa con Excel COM")
                return output_path
            finally:
                pythoncom.CoUninitialize()
        except:
            pass

    # Método 2: Pandas con xlrd
    try:
        import xlrd
        xlrd.Book.logfile = open(os.devnull, 'w')
        df = pd.read_excel(input_path, engine='xlrd')
        df.to_excel(output_path, index=False, engine='openpyxl')
        print("✅ Conversión exitosa con pandas/xlrd")
        return output_path
    except:
        pass

    # Método 3: LibreOffice
    try:
        import subprocess
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

        if libreoffice_path:
            output_dir = os.path.dirname(output_path)
            cmd = [libreoffice_path, "--headless", "--convert-to", "xlsx",
                   "--outdir", output_dir, input_path]
            subprocess.run(cmd, capture_output=True, timeout=60)

            expected = os.path.join(output_dir, os.path.basename(input_path).replace('.xls', '.xlsx'))
            if os.path.exists(expected):
                if expected != output_path:
                    import shutil
                    shutil.move(expected, output_path)
                print("✅ Conversión exitosa con LibreOffice")
                return output_path
    except:
        pass

    print("❌ No se pudo convertir automáticamente")
    print("\n💡 SOLUCIÓN:")
    print("1. Abre el archivo en Excel")
    print("2. Ctrl+A → Ctrl+C")
    print("3. Nuevo libro → Ctrl+V")
    print("4. Guarda como .xlsx")
    print("\nO instala: pip install pywin32")
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
    print(f"📖 Leyendo archivo: {file_path}")

    archivo_convertido = None

    # Verificar si necesita conversión
    formato = detectar_formato_excel(file_path)
    extension = os.path.splitext(file_path)[1].lower()

    if formato == "xls_antiguo" or (extension == '.xls' and formato != "xlsx_moderno"):
        archivo_convertido = convertir_xls_a_xlsx(file_path)
        if archivo_convertido:
            file_path = archivo_convertido
        else:
            print("\n❌ No se puede continuar sin conversión exitosa")
            sys.exit(1)

    try:
        df = pd.read_excel(file_path)

        # La primera fila contiene los encabezados reales
        headers = df.iloc[0].tolist()
        df = df[1:].reset_index(drop=True)
        df.columns = headers

        print(f"✅ Archivo cargado: {len(df)} entregas")
        return df
    finally:
        # Limpiar archivo temporal convertido
        if archivo_convertido and os.path.exists(archivo_convertido):
            try:
                os.remove(archivo_convertido)
            except:
                pass

def analizar_duplicados(df):
    """Analiza y marca entregas duplicadas"""
    print("\n🔍 Analizando duplicados...")
    
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

def generar_reporte(df_analizado, archivo_original):
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
    
    # Guardar archivos
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    base_name = os.path.splitext(os.path.basename(archivo_original))[0]
    
    # Archivo completo con análisis
    archivo_completo = f"{base_name}_analisis_{timestamp}.xlsx"
    df_analizado.to_excel(archivo_completo, index=False)
    print(f"\n✅ Archivo completo guardado: {archivo_completo}")
    
    # Archivo solo con duplicados
    if total_duplicados > 0:
        df_duplicados = df_analizado[df_analizado['Es_Duplicado'] == True]
        archivo_duplicados = f"{base_name}_duplicados_{timestamp}.xlsx"
        df_duplicados.to_excel(archivo_duplicados, index=False)
        print(f"✅ Duplicados guardados: {archivo_duplicados}")
        
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
        archivo_resumen = f"{base_name}_resumen_{timestamp}.xlsx"
        resumen_duplicados.to_excel(archivo_resumen, index=False)
        print(f"✅ Resumen por dirección guardado: {archivo_resumen}")
        
        print("\n💰 TOP 5 DIRECCIONES CON MAYOR AHORRO:")
        print("-" * 60)
        for idx, row in resumen_duplicados.head(5).iterrows():
            print(f"{row['Dirección'][:50]}...")
            print(f"   Entregas: {int(row['Total Entregas'])} | Duplicados: {int(row['Entregas Duplicadas'])} | Ahorro: ${row['Ahorro']:,.2f}")
    
    print("\n✨ Proceso completado exitosamente!\n")

def main():
    if len(sys.argv) < 2:
        print("❌ Error: Debes proporcionar un archivo de entrada")
        print("Uso: python procesar_entregas.py <archivo_excel>")
        print("Ejemplo: python procesar_entregas.py Liquidacion_Cliente.xlsx")
        sys.exit(1)
    
    archivo = sys.argv[1]
    
    if not os.path.exists(archivo):
        print(f"❌ Error: El archivo '{archivo}' no existe")
        sys.exit(1)
    
    print("🚀 Iniciando análisis de entregas duplicadas")
    print("="*60)
    
    # Procesar archivo
    df_original = procesar_archivo(archivo)
    if df_original is None:
        sys.exit(1)
    
    # Analizar duplicados
    df_analizado = analizar_duplicados(df_original)
    
    # Generar reporte
    generar_reporte(df_analizado, archivo)

if __name__ == "__main__":
    main()
