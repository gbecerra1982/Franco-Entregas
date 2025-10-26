"""
Script para convertir archivos XLS corruptos a XLSX
Usa múltiples métodos de conversión

Uso:
    python convertir_xls_a_xlsx.py "ruta/al/archivo.xls"
"""

import sys
import os
import pandas as pd
import shutil
from pathlib import Path

def metodo_1_pandas_basico(input_file, output_file):
    """Método 1: Pandas con diferentes motores"""
    print("🔄 Método 1: Intentando conversión con Pandas...")
    
    engines = ['openpyxl', 'xlrd', None]
    
    for engine in engines:
        try:
            print(f"   Probando motor: {engine or 'default'}...")
            if engine:
                df = pd.read_excel(input_file, engine=engine)
            else:
                df = pd.read_excel(input_file)
            
            # Guardar como XLSX
            df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"✅ Éxito con motor {engine or 'default'}!")
            return True
        except Exception as e:
            print(f"   ❌ Falló: {str(e)[:80]}")
            continue
    
    return False

def metodo_2_xlrd_ignorar_errores(input_file, output_file):
    """Método 2: xlrd ignorando errores de corrupción"""
    print("🔄 Método 2: xlrd con validación relajada...")
    
    try:
        import xlrd
        
        # Abrir libro ignorando algunos errores
        book = xlrd.open_workbook(input_file, formatting_info=False, on_demand=True)
        
        # Leer todas las hojas
        all_sheets = []
        for sheet_name in book.sheet_names():
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
        
        # Guardar como XLSX
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in all_sheets:
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        print("✅ Conversión exitosa con xlrd!")
        return True
        
    except Exception as e:
        print(f"❌ Método 2 falló: {str(e)[:80]}")
        return False

def metodo_3_copiar_y_pegar(input_file):
    """Método 3: Instrucciones para copia manual"""
    print("\n📋 Método 3: Conversión Manual")
    print("=" * 60)
    print("\nYa que los métodos automáticos fallaron, sigue estos pasos:")
    print("\n1. Abre el archivo en Microsoft Excel:")
    print(f"   {input_file}")
    print("\n2. Selecciona todo el contenido (Ctrl + A)")
    print("\n3. Copia (Ctrl + C)")
    print("\n4. Abre un NUEVO libro de Excel (Ctrl + N)")
    print("\n5. Pega en la primera celda (Ctrl + V)")
    print("\n6. Guarda como:")
    print("   - Archivo → Guardar como")
    print("   - Tipo: 'Libro de Excel (*.xlsx)'")
    print(f"   - Nombre sugerido: {input_file.replace('.xls', '_convertido.xlsx')}")
    print("\n7. Cierra Excel y vuelve a intentar con el nuevo archivo")
    print("\n" + "=" * 60)

def metodo_4_libreoffice(input_file, output_file):
    """Método 4: Usar LibreOffice si está instalado"""
    print("🔄 Método 4: Intentando con LibreOffice...")
    
    import subprocess
    
    # Buscar LibreOffice
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
        print("❌ LibreOffice no encontrado")
        print("💡 Descarga gratis desde: https://www.libreoffice.org/")
        return False
    
    try:
        output_dir = os.path.dirname(output_file)
        
        cmd = [
            libreoffice_path,
            "--headless",
            "--convert-to", "xlsx",
            "--outdir", output_dir,
            input_file
        ]
        
        print(f"   Ejecutando: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        # Verificar si se creó el archivo
        expected_output = os.path.join(output_dir, os.path.basename(input_file).replace('.xls', '.xlsx'))
        
        if os.path.exists(expected_output):
            if expected_output != output_file:
                shutil.move(expected_output, output_file)
            print("✅ Conversión exitosa con LibreOffice!")
            return True
        else:
            print(f"❌ No se generó el archivo esperado: {expected_output}")
            return False
            
    except Exception as e:
        print(f"❌ Error con LibreOffice: {str(e)}")
        return False

def convertir_archivo(input_file):
    """Función principal de conversión"""
    
    # Verificar que el archivo existe
    if not os.path.exists(input_file):
        print(f"❌ Error: El archivo no existe: {input_file}")
        return False
    
    # Preparar nombre de salida
    output_file = input_file.replace('.xls', '_convertido.xlsx')
    
    print("=" * 60)
    print("📁 CONVERTIDOR DE ARCHIVOS XLS CORRUPTOS")
    print("=" * 60)
    print(f"\n📄 Archivo entrada: {input_file}")
    print(f"📄 Archivo salida:  {output_file}")
    print(f"📊 Tamaño: {os.path.getsize(input_file):,} bytes")
    print("\n" + "=" * 60 + "\n")
    
    # Intentar diferentes métodos
    metodos = [
        metodo_1_pandas_basico,
        metodo_2_xlrd_ignorar_errores,
        metodo_4_libreoffice,
    ]
    
    for i, metodo in enumerate(metodos, 1):
        print(f"\n🔧 Probando Método {i}/{len(metodos)}...")
        try:
            if metodo(input_file, output_file):
                print("\n" + "=" * 60)
                print("✅ ¡CONVERSIÓN EXITOSA!")
                print("=" * 60)
                print(f"\n📁 Archivo convertido guardado en:")
                print(f"   {output_file}")
                print(f"\n📊 Tamaño del archivo convertido: {os.path.getsize(output_file):,} bytes")
                print("\n💡 Ahora puedes usar este archivo .xlsx en la aplicación")
                return True
        except Exception as e:
            print(f"❌ Error inesperado en método {i}: {str(e)}")
            continue
    
    # Si todos los métodos fallaron
    print("\n" + "=" * 60)
    print("❌ TODOS LOS MÉTODOS AUTOMÁTICOS FALLARON")
    print("=" * 60)
    metodo_3_copiar_y_pegar(input_file)
    
    return False

if __name__ == "__main__":
    print("\n")
    
    # Verificar argumentos
    if len(sys.argv) < 2:
        print("❌ Error: Debes proporcionar la ruta del archivo")
        print("\nUso:")
        print('   python convertir_xls_a_xlsx.py "ruta/al/archivo.xls"')
        print("\nEjemplo:")
        print('   python convertir_xls_a_xlsx.py "Liquidacion_Cliente_ANDROIDE AZUL (J&J)-202510232314.xls"')
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Convertir
    exito = convertir_archivo(input_file)
    
    if exito:
        print("\n✅ Proceso completado exitosamente\n")
        sys.exit(0)
    else:
        print("\n❌ No se pudo convertir el archivo automáticamente")
        print("💡 Sigue las instrucciones manuales mostradas arriba\n")
        sys.exit(1)
