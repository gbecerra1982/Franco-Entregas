# ✨ Analizador de Entregas Duplicadas - INSTALADO

## 📦 ¡Todo Listo para Usar!

Este directorio contiene el sistema completo para detectar y ajustar entregas duplicadas.

---

## 🚀 Inicio Rápido (Windows)

### Opción 1: Con Conda (Recomendado)
```powershell
# 1. Abrir PowerShell en este directorio
# 2. Ejecutar:
.\setup_conda.ps1

# 3. Activar entorno:
conda activate entregas_duplicadas

# 4. Ejecutar la app:
streamlit run app_entregas.py
```

### Opción 2: Con pip
```powershell
# 1. Abrir PowerShell en este directorio  
# 2. Ejecutar:
.\instalar.ps1

# 3. Ejecutar la app:
streamlit run app_entregas.py
```

### Opción 3: Línea de Comandos
```powershell
# Activar entorno (si usas conda o venv)
conda activate entregas_duplicadas

# Procesar un archivo
python procesar_entregas.py tu_archivo.xlsx
```

---

## 📁 Archivos en este Directorio

### 📱 Aplicaciones
- **app_entregas.py** - Aplicación web interactiva (Streamlit)
- **procesar_entregas.py** - Script de línea de comandos

### ⚙️ Configuración e Instalación
- **requirements.txt** - Lista de dependencias Python
- **environment.yml** - Configuración de entorno Conda
- **setup_conda.ps1** - Instalador automático Conda (Windows)
- **instalar.ps1** - Instalador automático pip (Windows)
- **setup_conda.sh** - Instalador automático Conda (Linux/Mac)
- **instalar.sh** - Instalador automático pip (Linux/Mac)

### 📖 Documentación
- **README.md** - Documentación principal del proyecto
- **INICIO_RAPIDO.md** - Guía de inicio rápido con ejemplos
- **INDICE.md** - Índice completo de todos los archivos
- **INSTALACION_WINDOWS.md** - Guía detallada para Windows
- **CHEAT_SHEET.md** - Comandos rápidos y atajos

### 📊 Archivos de Ejemplo (del análisis demo)
- **Liquidacion_Cliente_..._analisis_....xlsx** - Ejemplo de análisis completo
- **Liquidacion_Cliente_..._duplicados_....xlsx** - Ejemplo de duplicados
- **Liquidacion_Cliente_..._resumen_....xlsx** - Ejemplo de resumen
- **analisis_visual.png** - Gráficos del análisis
- **resumen_detallado.txt** - Reporte en texto

---

## 🎯 Resultados del Ejemplo Procesado

Del archivo de demostración se detectó:
- ✅ **321 entregas** analizadas
- 🔄 **16 entregas duplicadas** (5%)
- 👥 **14 grupos** de direcciones con duplicados
- 💰 **Ahorro: $111,600**

---

## 🔧 Comandos Más Usados

### Iniciar la aplicación web:
```powershell
streamlit run app_entregas.py
```
Luego abre: http://localhost:8501

### Procesar un archivo:
```powershell
python procesar_entregas.py Liquidacion.xlsx
```

### Activar entorno conda:
```powershell
conda activate entregas_duplicadas
```

### Desactivar entorno:
```powershell
conda deactivate
```

---

## 📚 Documentación Completa

Lee estos archivos en orden:

1. **README.md** - Empieza aquí para entender qué hace el sistema
2. **INICIO_RAPIDO.md** - Ejemplos prácticos paso a paso
3. **INDICE.md** - Referencia completa de todos los archivos

---

## 💡 Consejos

- Si es tu primera vez, usa la **aplicación web** (más fácil)
- Los archivos .xlsx procesados tendrán columnas adicionales de análisis
- El sistema NO modifica tu archivo original
- Todos los archivos generados tienen timestamp para evitar sobrescritura

---

## 🐛 Problemas Comunes

### "No se reconoce conda"
→ Instala Miniconda o Anaconda desde: https://docs.conda.io/

### "No module named 'pandas'"
→ Ejecuta: `pip install -r requirements.txt`

### "Streamlit no abre"
→ Verifica que estés en el directorio correcto y que el puerto 8501 esté libre

---

## 📞 Necesitas Ayuda?

1. Revisa **INICIO_RAPIDO.md** - Sección "Solución de Problemas"
2. Verifica que todas las dependencias estén instaladas
3. Asegúrate de que tu archivo Excel tenga las columnas correctas

---

## 🎉 ¡A Optimizar Entregas!

**Tu siguiente paso:**
```powershell
streamlit run app_entregas.py
```

Luego arrastra tu archivo y observa la magia ✨

---

**Fecha de instalación:** 26 de Octubre de 2025
**Ubicación:** C:\workspaces\Franco\Franco-Entregas
