# 🪟 Guía de Instalación en Windows

## 🎯 Instalación Rápida (Recomendada)

### Prerequisitos
- Windows 10 o superior
- PowerShell 5.1 o superior (viene preinstalado)

### Opción 1: Con Conda (Más Estable) ⭐

#### Paso 1: Instalar Miniconda
1. Descarga Miniconda desde: https://docs.conda.io/en/latest/miniconda.html
2. Ejecuta el instalador `Miniconda3-latest-Windows-x86_64.exe`
3. Acepta los valores predeterminados
4. **Importante:** Marca la opción "Add Miniconda to PATH"

#### Paso 2: Abrir PowerShell
- Presiona `Win + X`
- Selecciona "Windows PowerShell" o "Terminal"

#### Paso 3: Ejecutar el script de instalación
```powershell
# Navegar a la carpeta del proyecto
cd C:\ruta\donde\descargaste\los\archivos

# Ejecutar el script
.\setup_conda.ps1
```

Si te sale un error de permisos, ejecuta primero:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Paso 4: ¡Listo!
```powershell
# Activar el entorno
conda activate entregas_duplicadas

# Ejecutar la aplicación
streamlit run app_entregas.py
```

---

### Opción 2: Con pip (Virtualenv)

#### Paso 1: Verificar Python
```powershell
python --version
```

Si no tienes Python instalado:
1. Descarga desde: https://www.python.org/downloads/
2. Ejecuta el instalador
3. **Importante:** Marca "Add Python to PATH"

#### Paso 2: Crear entorno virtual
```powershell
# Navegar a la carpeta del proyecto
cd C:\ruta\donde\descargaste\los\archivos

# Crear entorno virtual
python -m venv venv

# Activar entorno
.\venv\Scripts\Activate.ps1
```

Si te sale un error de permisos:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Paso 3: Instalar dependencias
```powershell
pip install -r requirements.txt
```

#### Paso 4: ¡Listo!
```powershell
# Ejecutar la aplicación
streamlit run app_entregas.py
```

---

## 🎮 Cómo Usar

### Aplicación Web (Interfaz Gráfica)

```powershell
# 1. Activar el entorno (si usaste conda)
conda activate entregas_duplicadas
# O si usaste venv:
# .\venv\Scripts\Activate.ps1

# 2. Ejecutar la aplicación
streamlit run app_entregas.py

# 3. Se abrirá automáticamente en tu navegador
# Si no, abre: http://localhost:8501
```

### Script de Línea de Comandos

```powershell
# Activar el entorno
conda activate entregas_duplicadas

# Procesar un archivo
python procesar_entregas.py "C:\ruta\a\tu\archivo.xls"
```

---

## 📁 Ubicación de Archivos Generados

Los archivos procesados se guardan en la misma carpeta donde está el script:

```
C:\Users\TuUsuario\Downloads\
├── procesar_entregas.py
├── tu_archivo.xls
├── tu_archivo_analisis_20251026_125059.xlsx  ← NUEVO
├── tu_archivo_duplicados_20251026_125059.xlsx ← NUEVO
└── tu_archivo_resumen_20251026_125059.xlsx   ← NUEVO
```

---

## 🔧 Solución de Problemas en Windows

### Error: "No se puede ejecutar scripts en este sistema"

**Solución:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### Error: "conda no se reconoce como un comando"

**Solución:**
1. Abre una nueva ventana de PowerShell
2. O reinicia tu computadora
3. O agrega Conda al PATH manualmente:
   - Busca "Variables de entorno" en el menú de Windows
   - Edita "Path" en "Variables del sistema"
   - Agrega: `C:\Users\TuUsuario\miniconda3\Scripts`

---

### Error: "python no se reconoce como un comando"

**Solución:**
1. Reinstala Python desde https://www.python.org/
2. Marca la opción "Add Python to PATH" durante la instalación
3. Reinicia PowerShell

---

### Error: "ModuleNotFoundError: No module named 'pandas'"

**Solución:**
```powershell
# Si usas conda:
conda activate entregas_duplicadas
conda install -c conda-forge pandas openpyxl xlrd streamlit matplotlib -y

# Si usas pip:
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

### Error: "streamlit no se encuentra"

**Solución:**
```powershell
# Asegúrate de que el entorno esté activado
conda activate entregas_duplicadas

# Reinstala streamlit
conda install -c conda-forge streamlit -y

# O con pip:
pip install streamlit
```

---

### El navegador no se abre automáticamente

**Solución:**
Abre manualmente: http://localhost:8501

O si ese puerto está ocupado:
```powershell
streamlit run app_entregas.py --server.port 8502
```
Luego abre: http://localhost:8502

---

## 💡 Tips para Windows

### 1. Crear Acceso Directo en el Escritorio

Crea un archivo `.bat`:

```batch
@echo off
cd C:\ruta\a\tus\archivos
call conda activate entregas_duplicadas
streamlit run app_entregas.py
pause
```

Guárdalo como `Analizador_Entregas.bat` en tu escritorio.

---

### 2. Usar Windows Terminal (Recomendado)

Windows Terminal es mejor que PowerShell tradicional:
1. Descarga desde Microsoft Store: "Windows Terminal"
2. Es más moderno y fácil de usar

---

### 3. Procesar Múltiples Archivos

Crea un script `.bat` para procesar varios archivos:

```batch
@echo off
cd C:\ruta\a\tus\archivos
call conda activate entregas_duplicadas

for %%f in (*.xls) do (
    echo Procesando: %%f
    python procesar_entregas.py "%%f"
)

pause
```

---

### 4. Programar Ejecución Automática

Usa el "Programador de Tareas" de Windows:
1. Abre "Programador de Tareas"
2. Crear Tarea Básica
3. Nombre: "Análisis Entregas Diario"
4. Desencadenador: Diariamente a las 2:00 AM
5. Acción: Iniciar programa
6. Programa: `C:\Windows\System32\cmd.exe`
7. Argumentos: `/c conda activate entregas_duplicadas && python C:\ruta\procesar_entregas.py C:\ruta\archivo.xls`

---

## 📞 ¿Necesitas Ayuda?

Si algo no funciona:
1. Asegúrate de estar en la carpeta correcta: `cd C:\ruta\correcta`
2. Verifica que el entorno esté activado: `conda activate entregas_duplicadas`
3. Revisa que los archivos estén ahí: `dir`
4. Lee el mensaje de error completo

---

## 🎓 Comandos Útiles en PowerShell

```powershell
# Ver dónde estás
pwd

# Listar archivos
dir

# Cambiar de carpeta
cd C:\ruta\a\carpeta

# Volver a la carpeta anterior
cd ..

# Limpiar pantalla
cls

# Ver versión de Python
python --version

# Ver paquetes instalados
conda list
# o
pip list

# Desactivar entorno
conda deactivate
```

---

¡Listo! Ahora ya puedes usar el sistema en Windows sin problemas. 🚀
