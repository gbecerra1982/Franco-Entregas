# 📁 Índice de Archivos Generados

## 📚 Documentación

### 1. **README.md** (5 KB)
Documentación principal del proyecto
- Explicación completa del sistema
- Qué hace y cómo funciona
- Casos de uso típicos
- Personalización

**Léelo primero para entender el proyecto**

---

### 2. **INICIO_RAPIDO.md** (8.6 KB)
Guía práctica con ejemplos paso a paso
- Instalación en 3 pasos
- Ejemplos de uso real
- Solución de problemas
- Tips avanzados

**Léelo segundo para empezar a usar el sistema**

---

### 3. **resumen_detallado.txt** (8.6 KB)
Reporte completo en texto del análisis realizado
- Estadísticas generales
- Análisis financiero
- Top 10 direcciones con duplicados
- Detalle de cada grupo de duplicados

**Léelo para ver el análisis completo en texto**

---

### 4. **INSTALACION_WINDOWS.md** (Nuevo)
Guía específica de instalación para Windows
- Instalación paso a paso en Windows
- Solución de problemas en PowerShell
- Tips para usuarios de Windows
- Programación de tareas automáticas

**Léelo si usas Windows**

---

## 🖼️ Visualizaciones

### 4. **analisis_visual.png** (510 KB)
Dashboard visual con 4 gráficos:
1. **Gráfico de pastel:** Entregas normales vs duplicadas
2. **Gráfico de barras:** Top 10 grupos con mayor ahorro
3. **Gráfico horizontal:** Estados de entrega
4. **Panel de métricas:** Resumen ejecutivo con todos los números

**Ideal para presentaciones y reportes ejecutivos**

---

## 💻 Código Ejecutable

### 5. **app_entregas.py** (12 KB)
Aplicación web interactiva con Streamlit
- Interfaz visual completa
- Carga de archivos drag & drop
- Visualización en tiempo real
- Descarga de resultados

**Ejecutar con:** `streamlit run app_entregas.py`

---

### 6. **procesar_entregas.py** (8.9 KB)
Script de línea de comandos
- Procesamiento automático
- Generación de 3 archivos Excel
- Reporte en consola
- Automatizable con cron

**Ejecutar con:** `python3 procesar_entregas.py archivo.xls`

---

### 7. **instalar.sh** (1.8 KB)
Script de instalación automática para Linux/Mac
- Verifica Python y pip
- Instala dependencias
- Verifica LibreOffice
- Configura permisos

**Ejecutar con:** `chmod +x instalar.sh && ./instalar.sh`

---

### 8. **setup_conda.sh** (Nuevo)
Script de instalación con Conda para Linux/Mac
- Crea entorno conda automáticamente
- Instala todas las dependencias
- Verifica instalación

**Ejecutar con:** `chmod +x setup_conda.sh && ./setup_conda.sh`

---

### 9. **setup_conda.ps1** (Nuevo)
Script de instalación con Conda para Windows (PowerShell)
- Crea entorno conda automáticamente
- Instala todas las dependencias
- Verifica instalación

**Ejecutar con:** `.\setup_conda.ps1`

---

### 10. **requirements.txt** (Nuevo)
Lista de dependencias para pip
- Especifica versiones exactas
- Compatible con virtualenv

**Usar con:** `pip install -r requirements.txt`

---

### 11. **environment.yml** (Nuevo)
Archivo de entorno para conda
- Reproducible en cualquier máquina
- Incluye canal conda-forge

**Usar con:** `conda env create -f environment.yml`

---

## 📊 Resultados del Análisis

### 8. **Liquidacion_Cliente_ANDROIDE_AZUL__J_J_-202510232314.xlsx** (37 KB)
Archivo original convertido a formato .xlsx

---

### 9. **..._analisis_20251026_125059.xlsx** (50 KB)
**⭐ ARCHIVO PRINCIPAL**

Contiene TODAS las 321 entregas con columnas adicionales:
- `Es_Duplicado`: TRUE/FALSE
- `Costo_Original`: Costo antes del ajuste
- `Costo_Ajustado`: Costo después del ajuste ($0 para duplicados)
- `Motivo_Ajuste`: Explicación del ajuste
- `Grupo_Duplicado`: Agrupa entregas a la misma dirección
- `Dirección_Normalizada`: Dirección limpia para comparación
- `Fecha_DateTime` y `Fecha_Solo`: Fechas procesadas

**Este es el archivo que debes usar para facturación**

Estadísticas:
- 321 entregas totales
- 16 entregas duplicadas detectadas (5%)
- 14 grupos de duplicados
- **Ahorro total: $111,600**

---

### 10. **..._duplicados_20251026_125059.xlsx** (8.7 KB)
Solo las 16 entregas marcadas como duplicadas

Útil para:
- Auditoría rápida
- Verificación manual
- Reportes al cliente
- Análisis de patrones

---

### 11. **..._resumen_20251026_125059.xlsx** (6.1 KB)
Resumen estadístico por dirección

Columnas:
- Dirección
- Total de Entregas
- Entregas Duplicadas
- Costo Original
- Costo Ajustado
- Ahorro

**Top 5 direcciones con mayor ahorro:**
1. Calle Chubut 415, Villa Rosa: $17,400
2. A Porto y Dellepiane, Campana: $8,700
3. 601 entre 121 bis y 122, La Plata: $8,700
4. Avenida San Martín 1222, Marcos Paz: $8,700
5. Avenida Montevideo 95, Berisso: $8,700

---

## 🚀 Inicio Rápido

### Para Usuarios de Windows:
```powershell
# Opción 1: Con Conda (Recomendado)
.\setup_conda.ps1
conda activate entregas_duplicadas
streamlit run app_entregas.py

# Opción 2: Con pip
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app_entregas.py
```

### Para Usuarios de Linux/Mac:
```bash
# Opción 1: Con Conda (Recomendado)
chmod +x setup_conda.sh
./setup_conda.sh
conda activate entregas_duplicadas
streamlit run app_entregas.py

# Opción 2: Con pip
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
streamlit run app_entregas.py
```

### Para Usuarios No Técnicos (Cualquier Sistema):
1. Descarga Miniconda: https://docs.conda.io/en/latest/miniconda.html
2. Ejecuta el script correspondiente a tu sistema:
   - Windows: `.\setup_conda.ps1`
   - Linux/Mac: `./setup_conda.sh`
3. Ejecuta: `streamlit run app_entregas.py`
4. ¡Abre tu navegador en http://localhost:8501 y listo!

---

## 📈 Resultados Clave del Análisis

```
════════════════════════════════════════════════════════════════
                    RESUMEN EJECUTIVO
════════════════════════════════════════════════════════════════

📦 Total de Entregas:              321
🔄 Entregas Duplicadas:             16 (5.0%)
👥 Grupos de Duplicados:            14

💰 Costo Original Total:      $2,826,300.00
✅ Costo Ajustado Total:      $2,714,700.00

💵 AHORRO TOTAL:                $111,600.00
   (3.9% de reducción)

📊 Ahorro promedio por duplicado:  $6,975.00

════════════════════════════════════════════════════════════════
```

### Tipos de duplicados encontrados:

1. **Reintentos después de "Nadie"** (7 casos)
   - Primera visita: Estado "Nadie"
   - Segunda visita: Estado "Entregado"
   - Ahorro: $60,900

2. **Entregas múltiples mismo día** (5 casos)
   - Varias entregas a la misma dirección en el mismo horario
   - Ahorro: $43,500

3. **Entregas días consecutivos** (4 casos)
   - Segunda entrega 1-2 días después de la primera
   - Ahorro: $7,200

---

## 💼 Caso de Negocio

**Proyección anual:**
Si este patrón se mantiene, con 321 entregas/mes:
- Entregas anuales: 3,852
- Duplicados anuales (5%): ~193
- **Ahorro anual estimado: $1,339,200**

**ROI del sistema:**
- Tiempo de implementación: ~1 hora
- Ahorro primer mes: $111,600
- Ahorro anual: $1,339,200
- **ROI: Inmediato**

---

## 🎯 Próximos Pasos

1. **Revisar los duplicados**
   - Abre: `..._duplicados_20251026_125059.xlsx`
   - Verifica manualmente algunos casos
   - Confirma que el criterio es correcto

2. **Aplicar los ajustes**
   - Usa: `..._analisis_20251026_125059.xlsx`
   - Columna `Costo_Ajustado` tiene los nuevos costos
   - Factura usando estos valores

3. **Configurar para uso regular**
   - Si estás satisfecho, automatiza el proceso
   - Ejecuta el script cada vez que recibas un archivo
   - O usa la app web cuando lo necesites

4. **Personalizar si es necesario**
   - Ajusta el rango de días (actualmente 0-2)
   - Modifica el monto del ajuste (actualmente $0)
   - Agrega filtros por estado si lo deseas

---

## 📞 Soporte

Si tienes preguntas:
1. Lee **README.md** para conceptos
2. Lee **INICIO_RAPIDO.md** para ejemplos prácticos
3. Revisa la sección "Solución de Problemas" en INICIO_RAPIDO.md

---

## 📝 Notas Finales

- ✅ El sistema NO modifica tu archivo original
- ✅ Todos los archivos tienen timestamp para evitar sobrescritura
- ✅ Puedes ejecutar el análisis múltiples veces sin problemas
- ✅ Los archivos Excel son compatibles con todas las versiones
- ✅ El código es open source y personalizable

---

**Desarrollado con ❤️ para optimizar la gestión de entregas**

**Fecha de análisis:** 26 de Octubre de 2025
**Cliente:** ANDROIDE AZUL (J&J)
**Período analizado:** Octubre 2025
