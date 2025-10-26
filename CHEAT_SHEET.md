# 📝 Hoja de Referencia Rápida (Cheat Sheet)

## 🚀 Instalación Rápida

### Windows (PowerShell)
```powershell
# Con Conda
.\setup_conda.ps1

# Con pip
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### Linux/Mac (Bash)
```bash
# Con Conda
chmod +x setup_conda.sh && ./setup_conda.sh

# Con pip
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## 🎯 Uso Diario

### Activar Entorno

**Windows + Conda:**
```powershell
conda activate entregas_duplicadas
```

**Windows + venv:**
```powershell
.\venv\Scripts\Activate.ps1
```

**Linux/Mac + Conda:**
```bash
conda activate entregas_duplicadas
```

**Linux/Mac + venv:**
```bash
source venv/bin/activate
```

---

### Ejecutar Aplicación Web

```bash
streamlit run app_entregas.py
```

Luego abre: http://localhost:8501

---

### Ejecutar Script de Línea de Comandos

**Windows:**
```powershell
python procesar_entregas.py "C:\ruta\a\archivo.xls"
```

**Linux/Mac:**
```bash
python3 procesar_entregas.py /ruta/a/archivo.xls
```

---

### Procesar Múltiples Archivos

**Windows (PowerShell):**
```powershell
Get-ChildItem -Filter *.xls | ForEach-Object {
    python procesar_entregas.py $_.FullName
}
```

**Linux/Mac (Bash):**
```bash
for archivo in *.xls; do
    python3 procesar_entregas.py "$archivo"
done
```

---

## 🔧 Comandos de Mantenimiento

### Ver Paquetes Instalados

**Conda:**
```bash
conda list
```

**pip:**
```bash
pip list
```

---

### Actualizar Paquetes

**Conda:**
```bash
conda update --all
```

**pip:**
```bash
pip install --upgrade -r requirements.txt
```

---

### Desactivar Entorno

**Conda:**
```bash
conda deactivate
```

**venv:**
```bash
deactivate
```

---

### Eliminar Entorno

**Conda:**
```bash
conda env remove -n entregas_duplicadas
```

**venv:**
```bash
# Solo elimina la carpeta
rm -rf venv  # Linux/Mac
rmdir /s venv  # Windows
```

---

### Recrear Entorno desde Cero

**Conda + environment.yml:**
```bash
conda env remove -n entregas_duplicadas
conda env create -f environment.yml
```

**pip + requirements.txt:**
```bash
# Eliminar venv viejo
rm -rf venv

# Crear nuevo venv
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

---

## 📊 Comandos de Análisis

### Ver Primeras Líneas de un Archivo

**Windows:**
```powershell
Get-Content archivo_analisis.xlsx | Select-Object -First 10
```

**Linux/Mac:**
```bash
head -10 archivo_analisis.txt
```

---

### Buscar en Archivos

**Windows (PowerShell):**
```powershell
Select-String -Path "*.txt" -Pattern "duplicado"
```

**Linux/Mac:**
```bash
grep -r "duplicado" *.txt
```

---

### Contar Líneas en Archivo

**Windows:**
```powershell
(Get-Content archivo.txt).Count
```

**Linux/Mac:**
```bash
wc -l archivo.txt
```

---

## 🐛 Solución Rápida de Problemas

### Error: Comando no encontrado

**Solución:**
1. Asegúrate de que el entorno esté activado
2. Reinicia la terminal
3. Verifica la instalación: `python --version`

---

### Error: Permisos en PowerShell

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### Error: Puerto 8501 ocupado

```bash
# Usar puerto diferente
streamlit run app_entregas.py --server.port 8502
```

---

### Error: Módulo no encontrado

```bash
# Reactivar entorno e instalar
conda activate entregas_duplicadas
pip install -r requirements.txt
```

---

## 🎨 Personalización Rápida

### Cambiar Días para Detectar Duplicados

Edita `procesar_entregas.py` línea ~127:
```python
if 0 <= diferencia_dias <= 2:  # Cambiar el 2
```

Valores comunes:
- `<= 0` = Solo mismo día
- `<= 1` = Hasta día siguiente
- `<= 2` = Hasta 2 días (actual)
- `<= 7` = Misma semana

---

### Cambiar Costo Ajustado

Edita `procesar_entregas.py` línea ~135:
```python
df.loc[idx, 'Costo_Ajustado'] = 0  # Cambiar 0
```

Ejemplos:
- `= 0` = Sin cargo (actual)
- `= costo_actual * 0.5` = 50% descuento
- `= 1000` = Cargo fijo $1,000

---

## 📁 Ubicaciones Importantes

### Archivos del Proyecto

**Windows:**
```
C:\Users\TuUsuario\Downloads\analizador-entregas\
├── app_entregas.py
├── procesar_entregas.py
├── requirements.txt
├── environment.yml
└── setup_conda.ps1
```

**Linux/Mac:**
```
~/Downloads/analizador-entregas/
├── app_entregas.py
├── procesar_entregas.py
├── requirements.txt
├── environment.yml
└── setup_conda.sh
```

---

### Archivos Generados

Los archivos procesados se guardan en la misma carpeta:
```
├── archivo_original.xls
├── archivo_analisis_YYYYMMDD_HHMMSS.xlsx ← Principal
├── archivo_duplicados_YYYYMMDD_HHMMSS.xlsx
└── archivo_resumen_YYYYMMDD_HHMMSS.xlsx
```

---

## ⚡ Atajos de Teclado

### En PowerShell/Terminal

| Atajo | Acción |
|-------|--------|
| `Ctrl + C` | Detener aplicación |
| `Ctrl + L` | Limpiar pantalla |
| `↑` / `↓` | Historial de comandos |
| `Tab` | Autocompletar |
| `Ctrl + R` | Buscar en historial |

---

### En Streamlit

| Atajo | Acción |
|-------|--------|
| `R` | Recargar aplicación |
| `Ctrl + C` | Cerrar aplicación |

---

## 📞 Comandos de Ayuda

```bash
# Ver ayuda de Python
python --help

# Ver ayuda de conda
conda --help

# Ver ayuda de streamlit
streamlit --help

# Ver versión de un paquete
python -c "import pandas; print(pandas.__version__)"
```

---

## 🎓 Comandos de Navegación

### Windows (PowerShell)

| Comando | Descripción |
|---------|-------------|
| `pwd` | Ver directorio actual |
| `cd ruta` | Cambiar directorio |
| `dir` | Listar archivos |
| `cls` | Limpiar pantalla |
| `cd ..` | Subir un nivel |

### Linux/Mac (Bash)

| Comando | Descripción |
|---------|-------------|
| `pwd` | Ver directorio actual |
| `cd ruta` | Cambiar directorio |
| `ls` | Listar archivos |
| `clear` | Limpiar pantalla |
| `cd ..` | Subir un nivel |

---

## 💾 Backup y Exportación

### Exportar Entorno Conda

```bash
# Exportar
conda env export > mi_entorno.yml

# Importar en otra máquina
conda env create -f mi_entorno.yml
```

---

### Exportar Paquetes pip

```bash
# Exportar
pip freeze > mis_requirements.txt

# Importar
pip install -r mis_requirements.txt
```

---

## 🔗 Enlaces Útiles

- **Documentación Python:** https://docs.python.org/3/
- **Documentación Conda:** https://docs.conda.io/
- **Documentación Pandas:** https://pandas.pydata.org/docs/
- **Documentación Streamlit:** https://docs.streamlit.io/
- **Stack Overflow:** https://stackoverflow.com/

---

¡Guarda esta hoja de referencia para consultas rápidas! 📌
