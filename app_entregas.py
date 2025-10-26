import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import tempfile
import os

st.set_page_config(page_title="Analizador de Entregas Duplicadas", layout="wide")

st.title("📦 Analizador de Entregas Duplicadas")
st.markdown("### Detecta entregas duplicadas a la misma dirección y ajusta costos")

# Función para intentar reparar archivo Excel corrupto usando pyxlsb
def intentar_reparar_xls(file_path):
    """Intenta leer archivo XLS usando métodos alternativos"""
    try:
        import subprocess
        import sys
        
        # Intentar con pandas usando python-xls
        st.info("🔧 Intentando método de reparación alternativo...")
        
        # Método 1: Intentar leer sin validación estricta
        try:
            import xlrd
            # Deshabilitar validación estricta
            xlrd.Book.logfile = open(os.devnull, 'w')
            df = pd.read_excel(file_path, engine='xlrd')
            st.success("✅ Archivo reparado exitosamente!")
            return df
        except:
            pass
        
        return None
    except Exception as e:
        st.warning(f"⚠️ Método de reparación falló: {str(e)[:100]}")
        return None

# Función para limpiar y procesar el archivo
def procesar_archivo(file_path):
    """Lee y procesa el archivo Excel con manejo robusto de diferentes formatos"""
    df = None
    error_msg = None
    
    # Intentar con diferentes motores en orden de preferencia
    engines = [
        ('openpyxl', 'Intentando con motor OpenPyXL (xlsx)'),
        ('xlrd', 'Intentando con motor xlrd (xls)'),
        (None, 'Intentando con motor por defecto')
    ]
    
    for engine, descripcion in engines:
        try:
            st.info(f"🔄 {descripcion}...")
            if engine:
                df = pd.read_excel(file_path, engine=engine)
            else:
                df = pd.read_excel(file_path)
            st.success(f"✅ Archivo leído exitosamente con motor: {engine or 'default'}")
            break  # Si tuvo éxito, salir del loop
        except Exception as e:
            error_msg = str(e)
            st.warning(f"⚠️ Fallo con motor {engine or 'default'}: {error_msg[:100]}")
            continue
    
    # Si todos los motores fallaron, intentar método de reparación
    if df is None:
        st.warning("⚠️ Todos los motores estándar fallaron. Intentando métodos alternativos...")
        df = intentar_reparar_xls(file_path)
    
    # Si sigue sin funcionar, sugerir conversión
    if df is None:
        raise Exception(f"No se pudo leer el archivo con ningún método. Último error: {error_msg}\n\n"
                       f"💡 SOLUCIÓN RECOMENDADA:\n"
                       f"1. Abre el archivo en Excel\n"
                       f"2. Selecciona todo (Ctrl+A)\n"
                       f"3. Copia (Ctrl+C)\n"
                       f"4. Abre un nuevo libro de Excel\n"
                       f"5. Pega (Ctrl+V)\n"
                       f"6. Guarda como .xlsx\n\n"
                       f"ALTERNATIVA RÁPIDA: Usa la herramienta de conversión XLS → XLSX incluida abajo")
    
    # La primera fila contiene los encabezados reales
    headers = df.iloc[0].tolist()
    df = df[1:].reset_index(drop=True)
    df.columns = headers
    
    return df

# Función para normalizar direcciones
def normalizar_direccion(direccion):
    """Normaliza una dirección para comparación"""
    if pd.isna(direccion):
        return ""
    
    direccion = str(direccion).lower().strip()
    # Remover múltiples espacios
    import re
    direccion = re.sub(r'\s+', ' ', direccion)
    # Remover puntos y comas
    direccion = direccion.replace(',', '').replace('.', '')
    
    return direccion

# Función para convertir fechas
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

# Función principal de análisis
def analizar_duplicados(df):
    """Analiza y marca entregas duplicadas"""
    
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
    grupos_duplicados = []
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

# Interfaz de la aplicación
st.sidebar.header("⚙️ Configuración")

# ========== HERRAMIENTA DE CONVERSIÓN XLS → XLSX ==========
st.sidebar.markdown("---")
st.sidebar.subheader("🔧 Herramienta de Conversión")
st.sidebar.markdown("**Si tu archivo .xls está corrupto:**")

with st.sidebar.expander("🔄 Convertir XLS → XLSX", expanded=False):
    st.markdown("""
    **Método Manual (Recomendado):**
    1. Abre el archivo en Excel
    2. Archivo → Guardar como
    3. Formato: **Excel Workbook (.xlsx)**
    4. Guarda y vuelve a cargar aquí
    
    ---
    
    **O usa LibreOffice (Gratis):**
    """)
    
    conversion_file = st.file_uploader(
        "Cargar archivo .xls corrupto para convertir", 
        type=['xls'],
        key="converter",
        help="Este archivo será convertido automáticamente"
    )
    
    if conversion_file:
        st.info("📝 Conversión usando LibreOffice...")
        
        # Guardar archivo temporal
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xls') as tmp_input:
            tmp_input.write(conversion_file.getbuffer())
            input_path = tmp_input.name
        
        output_path = input_path.replace('.xls', '.xlsx')
        
        try:
            # Intentar conversión con LibreOffice
            import subprocess
            import os
            
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
            
            if libreoffice_path:
                # Comando de conversión
                cmd = [
                    libreoffice_path,
                    "--headless",
                    "--convert-to", "xlsx",
                    "--outdir", os.path.dirname(output_path),
                    input_path
                ]
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                
                if os.path.exists(output_path):
                    with open(output_path, 'rb') as f:
                        xlsx_data = f.read()
                    
                    st.success("✅ Conversión exitosa!")
                    st.download_button(
                        label="📥 Descargar archivo convertido (.xlsx)",
                        data=xlsx_data,
                        file_name=conversion_file.name.replace('.xls', '.xlsx'),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Limpiar archivos temporales
                    os.remove(input_path)
                    os.remove(output_path)
                else:
                    raise Exception("No se pudo generar el archivo convertido")
            else:
                st.error("❌ LibreOffice no encontrado. Por favor instálalo desde https://www.libreoffice.org/")
                st.info("💡 O usa el método manual descrito arriba")
                
        except Exception as e:
            st.error(f"Error en conversión: {str(e)[:100]}")
            st.info("💡 **Usa el método manual:** Abre en Excel → Guardar como .xlsx")
            
            # Limpiar
            if os.path.exists(input_path):
                os.remove(input_path)

st.sidebar.markdown("---")

# Cargar archivo
uploaded_file = st.sidebar.file_uploader(
    "Cargar archivo Excel", 
    type=['xls', 'xlsx'],
    help="Sube tu archivo de liquidación de entregas"
)

# Agregar información de formato
st.sidebar.info("💡 **Importante:** Si el archivo .xls da error, ábrelo en Excel y guárdalo como .xlsx")

# Si hay archivo cargado, procesarlo automáticamente
if uploaded_file is None:
    st.info("👆 Por favor, carga un archivo Excel para comenzar el análisis")
    usar_archivo = False
else:
    # Guardar archivo subido temporalmente
    import tempfile
    import os
    
    # Determinar la extensión correcta
    file_extension = os.path.splitext(uploaded_file.name)[1]
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=file_extension) as tmp_file:
        tmp_file.write(uploaded_file.getbuffer())
        temp_path = tmp_file.name
    
    st.info(f"📁 Procesando archivo: {uploaded_file.name}")
    
    try:
        df_original = procesar_archivo(temp_path)
        usar_archivo = True
        st.success(f"✅ Archivo cargado exitosamente: {uploaded_file.name}")
        st.success(f"📊 Total de registros: {len(df_original)}")
    except Exception as e:
        st.error(f"❌ Error al procesar archivo: {e}")
        usar_archivo = False
        
        # Mostrar más información de debug
        with st.expander("🔍 Información de Debug"):
            st.code(f"""
Nombre del archivo: {uploaded_file.name}
Tamaño: {uploaded_file.size} bytes
Tipo MIME: {uploaded_file.type}

Error completo:
{str(e)}

Soluciones posibles:
1. Abre el archivo en Excel
2. Guarda como nuevo archivo con formato .xlsx
3. Vuelve a intentar cargar el archivo
            """)
    finally:
        # Limpiar archivo temporal
        if os.path.exists(temp_path):
            os.remove(temp_path)

# Si tenemos un archivo para procesar
if usar_archivo:
    # Opciones de filtrado
    st.sidebar.subheader("🔍 Filtros")
    
    # Mostrar estadísticas generales
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("📊 Total de Entregas", len(df_original))
    
    # Analizar duplicados
    with st.spinner('Analizando entregas duplicadas...'):
        df_analizado = analizar_duplicados(df_original.copy())
    
    # Calcular estadísticas
    total_duplicados = df_analizado['Es_Duplicado'].sum()
    total_grupos = df_analizado['Grupo_Duplicado'].nunique() - (1 if None in df_analizado['Grupo_Duplicado'].values else 0)
    
    with col2:
        st.metric("🔄 Entregas Duplicadas", total_duplicados)
    
    with col3:
        st.metric("👥 Grupos de Duplicados", total_grupos)
    
    # Calcular ahorro
    df_analizado['Costo_Original_Num'] = pd.to_numeric(df_analizado['Costo_Original'], errors='coerce').fillna(0)
    df_analizado['Costo_Ajustado_Num'] = pd.to_numeric(df_analizado['Costo_Ajustado'], errors='coerce').fillna(0)
    
    ahorro_total = (df_analizado['Costo_Original_Num'] - df_analizado['Costo_Ajustado_Num']).sum()
    
    if ahorro_total > 0:
        st.success(f"💰 **Ahorro Total: ${ahorro_total:,.2f}**")
    
    # Tabs para diferentes vistas
    tab1, tab2, tab3, tab4 = st.tabs(["📋 Todas las Entregas", "🔄 Solo Duplicados", "📊 Resumen por Dirección", "💾 Exportar"])
    
    with tab1:
        st.subheader("Todas las Entregas")
        
        # Preparar DataFrame para mostrar
        df_mostrar = df_analizado[[
            'Fecha Movimiento', 'Dirección', 'CP', 'Estado', 
            'Costo_Original', 'Costo_Ajustado', 'Es_Duplicado', 'Motivo_Ajuste'
        ]].copy()
        
        # Colorear duplicados
        def resaltar_duplicados(row):
            if row['Es_Duplicado']:
                return ['background-color: #ffebee'] * len(row)
            return [''] * len(row)
        
        st.dataframe(
            df_mostrar.style.apply(resaltar_duplicados, axis=1),
            use_container_width=True,
            height=400
        )
    
    with tab2:
        st.subheader("Entregas Duplicadas Detectadas")
        
        df_duplicados = df_analizado[df_analizado['Es_Duplicado'] == True].copy()
        
        if len(df_duplicados) > 0:
            # Agrupar por grupo de duplicados
            for grupo in sorted(df_duplicados['Grupo_Duplicado'].unique()):
                if pd.isna(grupo):
                    continue
                    
                grupo_df = df_analizado[df_analizado['Grupo_Duplicado'] == grupo]
                
                with st.expander(f"📦 Grupo {int(grupo)} - {grupo_df['Dirección'].iloc[0][:50]}... ({len(grupo_df)} entregas)"):
                    st.dataframe(
                        grupo_df[[
                            'Fecha Movimiento', 'Número Tracking', 'Dirección', 
                            'Estado', 'Costo_Original', 'Costo_Ajustado', 'Motivo_Ajuste'
                        ]],
                        use_container_width=True
                    )
                    
                    ahorro_grupo = (grupo_df['Costo_Original_Num'] - grupo_df['Costo_Ajustado_Num']).sum()
                    if ahorro_grupo > 0:
                        st.info(f"💰 Ahorro en este grupo: ${ahorro_grupo:,.2f}")
        else:
            st.info("No se detectaron entregas duplicadas en el archivo.")
    
    with tab3:
        st.subheader("Resumen por Dirección")
        
        # Crear resumen por dirección
        resumen = df_analizado.groupby('Dirección_Normalizada').agg({
            'Fecha Movimiento': 'count',
            'Es_Duplicado': 'sum',
            'Costo_Original_Num': 'sum',
            'Costo_Ajustado_Num': 'sum'
        }).reset_index()
        
        resumen.columns = ['Dirección', 'Total Entregas', 'Entregas Duplicadas', 'Costo Original', 'Costo Ajustado']
        resumen['Ahorro'] = resumen['Costo Original'] - resumen['Costo Ajustado']
        
        # Filtrar solo direcciones con duplicados
        resumen_duplicados = resumen[resumen['Entregas Duplicadas'] > 0].sort_values('Ahorro', ascending=False)
        
        if len(resumen_duplicados) > 0:
            st.dataframe(
                resumen_duplicados.style.format({
                    'Costo Original': '${:,.2f}',
                    'Costo Ajustado': '${:,.2f}',
                    'Ahorro': '${:,.2f}'
                }),
                use_container_width=True
            )
        else:
            st.info("No hay direcciones con entregas duplicadas.")
    
    with tab4:
        st.subheader("Exportar Resultados")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Archivo Original con Análisis**")
            st.write("Incluye todas las columnas originales más las columnas de análisis")
            
            # Preparar archivo para descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_analizado.to_excel(writer, index=False, sheet_name='Análisis Completo')
            
            st.download_button(
                label="📥 Descargar Excel Completo",
                data=output.getvalue(),
                file_name=f"analisis_entregas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            st.write("**Solo Entregas Duplicadas**")
            st.write("Archivo con únicamente las entregas marcadas como duplicadas")
            
            if len(df_duplicados) > 0:
                output2 = io.BytesIO()
                with pd.ExcelWriter(output2, engine='openpyxl') as writer:
                    df_duplicados.to_excel(writer, index=False, sheet_name='Duplicados')
                
                st.download_button(
                    label="📥 Descargar Solo Duplicados",
                    data=output2.getvalue(),
                    file_name=f"entregas_duplicadas_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("No hay duplicados para exportar")

# Footer
st.markdown("---")
st.markdown("### 📖 Cómo funciona")
st.markdown("""
Esta aplicación detecta entregas duplicadas basándose en:
1. **Misma dirección normalizada** (sin considerar mayúsculas, puntos, comas)
2. **Entregas en el mismo día o días consecutivos (0-2 días de diferencia)**
3. Cuando detecta un duplicado, **ajusta el costo de envío a $0** para el segundo día

**Columnas agregadas en el análisis:**
- `Es_Duplicado`: Indica si la entrega es un duplicado
- `Costo_Original`: Costo de envío original
- `Costo_Ajustado`: Costo ajustado (0 para duplicados)
- `Motivo_Ajuste`: Explicación del ajuste
- `Grupo_Duplicado`: Agrupa entregas a la misma dirección

### 🔧 Solución de Problemas

**Si recibes error "Workbook corruption":**
1. Abre el archivo en Microsoft Excel
2. Ve a Archivo → Guardar como
3. Selecciona formato "Excel Workbook (.xlsx)"
4. Guarda con un nuevo nombre
5. Vuelve a cargar el archivo nuevo en esta aplicación
""")