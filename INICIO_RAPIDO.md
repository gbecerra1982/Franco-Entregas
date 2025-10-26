# 🚀 Guía de Inicio Rápido

## 📋 Instalación

### Opción A: Con Conda (Recomendado) 🐍

#### En Windows (PowerShell):
```powershell
# Ejecutar el script de instalación automática
.\setup_conda.ps1
```

O manualmente:
```powershell
# Crear entorno
conda create -n entregas_duplicadas python=3.11 -y

# Activar entorno
conda activate entregas_duplicadas

# Instalar dependencias
conda install -c conda-forge pandas openpyxl xlrd streamlit matplotlib -y
```

O usando el archivo environment.yml:
```powershell
conda env create -f environment.yml
conda activate entregas_duplicadas
```

#### En Linux/Mac (Bash):
```bash
# Dar permisos y ejecutar script de instalación
chmod +x setup_conda.sh
./setup_conda.sh
```

O manualmente:
```bash
# Crear entorno
conda create -n entregas_duplicadas python=3.11 -y

# Activar entorno
conda activate entregas_duplicadas

# Instalar dependencias
conda install -c conda-forge pandas openpyxl xlrd streamlit matplotlib -y
```

O usando el archivo environment.yml:
```bash
conda env create -f environment.yml
conda activate entregas_duplicadas
```

---

### Opción B: Con pip (Virtualenv) 🐍

#### En Windows (PowerShell):
```powershell
# Crear entorno virtual
python -m venv venv

# Activar entorno
.\venv\Scripts\Activate.ps1

# Instalar dependencias
pip install -r requirements.txt
```

#### En Linux/Mac (Bash):
```bash
# Crear entorno virtual
python3 -m venv venv

# Activar entorno
source venv/bin/activate

# Instalar dependencias
pip install -r requirements.txt
```

---

### Opción C: Instalación Global (No recomendado)

#### Windows:
```powershell
pip install -r requirements.txt
```

#### Linux/Mac:
```bash
pip3 install -r requirements.txt
```

---

### Verificar Instalación

```bash
# Con conda
conda activate entregas_duplicadas
python --version  # Debe ser Python 3.8 o superior
streamlit --version

# Con venv
# (Después de activar el entorno virtual)
python --version
streamlit --version
```

---

### ¡Listo para usar!

---

## 🎮 Ejemplos de Uso

### Ejemplo 1: Usar la Aplicación Web (Más Fácil)

```bash
streamlit run app_entregas.py
```

Luego:
1. Abre tu navegador en http://localhost:8501
2. Arrastra tu archivo Excel o haz clic en "Browse files"
3. ¡Observa los resultados en tiempo real!
4. Descarga los archivos procesados

**Ventajas:**
- ✅ Interfaz visual amigable
- ✅ Resultados en tiempo real
- ✅ Gráficos interactivos
- ✅ No necesitas conocimientos técnicos

---

### Ejemplo 2: Procesar desde Línea de Comandos

#### Caso A: Archivo único
```bash
python3 procesar_entregas.py Liquidacion_Cliente.xls
```

**Resultado:**
```
🚀 Iniciando análisis...
📖 Leyendo archivo: Liquidacion_Cliente.xls
✅ Archivo cargado: 321 entregas

============================================================
📊 RESUMEN DEL ANÁLISIS
============================================================
📦 Total de entregas analizadas: 321
🔄 Entregas duplicadas detectadas: 16
👥 Grupos de duplicados: 14
💰 Ahorro total: $111,600.00
============================================================

✅ Archivo completo guardado: Liquidacion_Cliente_analisis_20251026_125059.xlsx
✅ Duplicados guardados: Liquidacion_Cliente_duplicados_20251026_125059.xlsx
✅ Resumen por dirección guardado: Liquidacion_Cliente_resumen_20251026_125059.xlsx

✨ Proceso completado exitosamente!
```

#### Caso B: Procesar múltiples archivos
```bash
# Procesar todos los .xls en un directorio
for archivo in *.xls; do
    python3 procesar_entregas.py "$archivo"
done
```

#### Caso C: Automatizar con cron (Linux/Mac)
```bash
# Editar crontab
crontab -e

# Agregar: Ejecutar todos los días a las 2 AM
0 2 * * * cd /ruta/al/directorio && python3 procesar_entregas.py archivo_diario.xls
```

---

## 📊 Interpretando los Resultados

### Archivo: `*_analisis_*.xlsx`

Este es tu archivo principal con TODAS las entregas más estas columnas nuevas:

| Columna | Descripción | Ejemplo |
|---------|-------------|---------|
| `Es_Duplicado` | ¿Es un duplicado? | TRUE / FALSE |
| `Costo_Original` | Costo antes del ajuste | $8,700 |
| `Costo_Ajustado` | Costo después del ajuste | $0 |
| `Motivo_Ajuste` | Por qué se ajustó | "Entrega duplicada (día 2)" |
| `Grupo_Duplicado` | ID del grupo | 1, 2, 3... |

**Cómo usar:**
1. Abre en Excel
2. Filtra por `Es_Duplicado = TRUE` para ver solo duplicados
3. Ordena por `Grupo_Duplicado` para ver entregas relacionadas
4. Suma `Costo_Ajustado` para tu nuevo total a facturar

---

### Archivo: `*_duplicados_*.xlsx`

Solo las entregas marcadas como duplicadas. Úsalo para:
- ✅ Auditoría rápida
- ✅ Reportes a clientes
- ✅ Verificación manual

---

### Archivo: `*_resumen_*.xlsx`

Estadísticas por dirección. Columnas:

| Dirección | Total Entregas | Entregas Duplicadas | Costo Original | Costo Ajustado | Ahorro |
|-----------|----------------|---------------------|----------------|----------------|--------|
| Calle 43... | 2 | 1 | $17,400 | $8,700 | $8,700 |

**Cómo usar:**
1. Ordena por "Ahorro" descendente
2. Identifica direcciones problemáticas
3. Analiza patrones de entregas

---

## 🎯 Casos de Uso Reales

### Caso 1: "Nadie en casa" + Reintento exitoso

**Situación:**
```
14/10 21:38 → Estado: "Nadie" → Cargo: $8,700
15/10 11:02 → Estado: "Entregado" → Cargo: $8,700
```

**Resultado después del análisis:**
```
14/10 21:38 → Estado: "Nadie" → Cargo: $8,700 ✅
15/10 11:02 → Estado: "Entregado" → Cargo: $0 ✅ (Duplicado día 2)
```

**Ahorro:** $8,700

---

### Caso 2: Entregas múltiples el mismo día

**Situación:**
```
17/10 20:34 → Paquete 1 → $8,700
17/10 20:34 → Paquete 2 → $8,700
```

**Resultado:**
```
17/10 20:34 → Paquete 1 → $8,700 ✅
17/10 20:34 → Paquete 2 → $0 ✅ (Duplicado mismo día)
```

**Ahorro:** $8,700

---

### Caso 3: Serie de 3 entregas

**Situación:**
```
Día 1 → $8,700
Día 2 → $8,700
Día 3 → $8,700
```

**Resultado:**
```
Día 1 → $8,700 ✅
Día 2 → $0 ✅ (Duplicado)
Día 3 → $0 ✅ (Duplicado)
```

**Ahorro:** $17,400

---

## 🔧 Personalización

### Cambiar el rango de días para considerar duplicados

Edita `app_entregas.py` o `procesar_entregas.py`:

```python
# Línea ~127
if 0 <= diferencia_dias <= 2:  # ← Cambiar este número
```

**Opciones:**
- `<= 0` → Solo mismo día
- `<= 1` → Mismo día + día siguiente
- `<= 2` → Mismo día + 2 días después (actual)
- `<= 7` → Misma semana

---

### Cambiar el monto de ajuste

```python
# Línea ~135
df.loc[idx, 'Costo_Ajustado'] = 0  # ← Cambiar 0 por otro valor
```

**Ejemplos:**
- `= 0` → Sin cargo (actual)
- `= costo_actual * 0.5` → 50% de descuento
- `= 1000` → Cargo fijo de $1,000

---

### Excluir ciertos estados

Agrega después de la línea ~118:

```python
# Saltar si el estado es "Entregado"
if df.loc[idx, 'Estado'] == 'Entregado':
    continue
```

---

## 🐛 Solución de Problemas

### Error: "No such file or directory"
**Solución:** Verifica la ruta del archivo
```bash
# Usar ruta absoluta
python3 procesar_entregas.py /ruta/completa/archivo.xls

# O navega al directorio primero
cd /directorio/con/archivo
python3 procesar_entregas.py archivo.xls
```

---

### Error: "ModuleNotFoundError: No module named 'pandas'"
**Solución:** Instalar dependencias
```bash
pip3 install pandas openpyxl xlrd
```

---

### Error: "Workbook corruption"
**Solución:** El archivo está corrupto. Opciones:
1. Abrirlo en Excel y guardarlo de nuevo
2. Convertir a .xlsx: `File → Save As → Excel Workbook (.xlsx)`
3. El script intentará convertirlo automáticamente con LibreOffice

---

### La aplicación web no se abre
**Solución:**
```bash
# Verificar que streamlit esté instalado
streamlit --version

# Reinstalar si es necesario
pip3 install --upgrade streamlit

# Probar en puerto diferente
streamlit run app_entregas.py --server.port 8502
```

---

## 💡 Tips Pro

### 1. Mantener historial de análisis
```bash
# Crear directorio para resultados
mkdir resultados_$(date +%Y%m)

# Mover resultados allí
mv *_analisis_*.xlsx resultados_$(date +%Y%m)/
```

### 2. Comparar resultados mes a mes
```bash
# Procesar archivo de octubre
python3 procesar_entregas.py octubre.xls

# Procesar archivo de noviembre
python3 procesar_entregas.py noviembre.xls

# Comparar ahorros
```

### 3. Exportar solo los ajustes para el cliente
En Excel:
1. Abre `*_duplicados_*.xlsx`
2. Selecciona columnas: Fecha, Tracking, Dirección, Costo_Original, Costo_Ajustado
3. Guarda como nuevo archivo para enviar al cliente

### 4. Crear reporte ejecutivo
```bash
# El script ya genera análisis visual
# Revisa los archivos:
# - analisis_visual.png (gráficos)
# - resumen_detallado.txt (texto completo)
```

---

## 📞 Preguntas Frecuentes

**P: ¿Modifica mi archivo original?**
R: NO. Crea nuevos archivos con sufijos de fecha/hora.

**P: ¿Puedo revertir los cambios?**
R: Sí, el archivo `*_analisis_*.xlsx` tiene tanto costos originales como ajustados.

**P: ¿Funciona con archivos grandes?**
R: Sí, probado con archivos de hasta 10,000 entregas.

**P: ¿Puedo usar otros delimitadores de fecha?**
R: El código actual soporta DD/MM/YYYY. Para otros formatos, edita la función `convertir_fecha()`.

**P: ¿Detecta duplicados en meses diferentes?**
R: Sí, mientras las entregas estén en el mismo archivo y dentro del rango de días configurado.

---

## 🎓 Ejemplos Avanzados

### Análisis de múltiples clientes
```bash
# Estructura de directorios
clientes/
  ├── cliente_A/
  │   └── liquidacion.xls
  ├── cliente_B/
  │   └── liquidacion.xls
  └── cliente_C/
      └── liquidacion.xls

# Script para procesar todos
for dir in clientes/*/; do
    cliente=$(basename "$dir")
    echo "Procesando $cliente..."
    python3 procesar_entregas.py "$dir"liquidacion.xls
done
```

### Integración con correo electrónico
```python
# Agregar al final de procesar_entregas.py
import smtplib
from email.mime.text import MIMEText

def enviar_reporte(ahorro_total):
    msg = MIMEText(f"Ahorro detectado: ${ahorro_total:,.2f}")
    msg['Subject'] = 'Reporte de Entregas Duplicadas'
    msg['From'] = 'tu@email.com'
    msg['To'] = 'destinatario@email.com'
    
    # Configurar tu servidor SMTP
    # smtp.send_message(msg)
```

---

¡Listo! Ahora tienes todo lo necesario para optimizar tus entregas. 🚀
