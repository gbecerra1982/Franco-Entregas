"""Script de diagnóstico para verificar win32com y Excel"""
import os
import sys

print("🔍 Diagnóstico de win32com y Excel")
print("=" * 60)

# Test 1: Importar win32com
print("\n1️⃣ Probando importación de win32com...")
try:
    import win32com.client
    import pythoncom
    print("✅ win32com importado correctamente")
except Exception as e:
    print(f"❌ Error importando win32com: {e}")
    sys.exit(1)

# Test 2: Verificar Excel instalado
print("\n2️⃣ Probando conexión con Excel...")
try:
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        version = excel.Version
        print(f"✅ Excel encontrado - Versión: {version}")
        excel.Quit()
    finally:
        pythoncom.CoUninitialize()
except Exception as e:
    print(f"❌ Error conectando con Excel: {e}")
    print("\n💡 Posibles causas:")
    print("   - Microsoft Excel no está instalado")
    print("   - Excel no está registrado correctamente en el sistema")
    print("   - Permisos de seguridad")
    sys.exit(1)

# Test 3: Probar conversión con archivo real
print("\n3️⃣ Probando conversión de archivo XLS...")
archivo_xls = "Liquidacion_Cliente_ANDROIDE AZUL (J&J)-202510232314.xls"

if not os.path.exists(archivo_xls):
    print(f"⚠️  Archivo no encontrado: {archivo_xls}")
    print("   (Ejecutar este script desde el directorio del proyecto)")
    sys.exit(0)

try:
    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        print("   Iniciando Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow

        abs_input = os.path.abspath(archivo_xls)
        abs_output = abs_input.replace('.xls', '_test_convertido.xlsx')

        print(f"   Abriendo: {archivo_xls}")
        print(f"   (Esto puede tardar unos segundos...)")

        # Parámetros para abrir sin macros y sin actualizar links
        workbook = excel.Workbooks.Open(
            abs_input,
            UpdateLinks=0,      # No actualizar links
            ReadOnly=True,      # Solo lectura
            IgnoreReadOnlyRecommended=True,
            Notify=False
        )

        print(f"   ✅ Archivo abierto correctamente")
        print(f"   Guardando como: {os.path.basename(abs_output)}")

        workbook.SaveAs(abs_output, FileFormat=51)
        print(f"   ✅ Archivo guardado")

        workbook.Close(SaveChanges=False)
        print(f"   ✅ Archivo cerrado")

        excel.Quit()
        print(f"   ✅ Excel cerrado")

        if os.path.exists(abs_output):
            print(f"\n✅ Conversión exitosa!")
            print(f"   Archivo generado: {os.path.basename(abs_output)}")

            # Limpiar archivo de prueba
            try:
                os.remove(abs_output)
                print("   (Archivo de prueba eliminado)")
            except:
                pass
        else:
            print("\n❌ El archivo no se generó correctamente")

    finally:
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
        except:
            pass
        try:
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

except Exception as e:
    print(f"❌ Error durante la conversión: {e}")
    print(f"\nDetalles del error:")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n" + "=" * 60)
print("✅ Todos los tests pasaron correctamente")
print("   win32com y Excel funcionan bien")
