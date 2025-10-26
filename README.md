# 📦 Analizador de Entregas Duplicadas

Sistema para detectar y corregir entregas duplicadas a la misma dirección, eliminando costos de envío del segundo día.

## 🎯 ¿Qué hace?

Analiza archivos de liquidación de entregas y detecta cuando:
- Se entrega **a la misma dirección**
- En **días consecutivos** (mismo día, día siguiente, o hasta 2 días después)
- Automáticamente **ajusta el costo a $0** para la segunda entrega y siguientes

## 📊 Resultados del Análisis

Del archivo procesado se detectó:
- **321 entregas totales**
- **16 entregas duplicadas** (5%)
- **14 grupos de duplicados**
- **💰 Ahorro total: $111,600**

### Ejemplos de duplicados encontrados:

1. **Calle Chubut 415, Villa Rosa**
   - 3 entregas (2 duplicadas)
   - Ahorro: $17,400

2. **A Porto y Dellepiane, Campana**
   - 2 entregas el mismo día
   - Ahorro: $8,700

3. **601 entre 121 bis y 122, La Plata**
   - Primera entrega: "Nadie" → Segunda entrega exitosa al día siguiente
   - Ahorro: $8,700

## 📁 Archivos Generados

### 1. `*_analisis_*.xlsx` - Análisis Completo
Contiene todas las entregas con columnas adicionales:
- `Es_Duplicado`: TRUE si es un duplicado
- `Costo_Original`: Costo original de envío
- `Costo_Ajustado`: Costo corregido ($0 para duplicados)
- `Motivo_Ajuste`: Explicación del ajuste
- `Grupo_Duplicado`: Agrupa entregas a la misma dirección

### 2. `*_duplicados_*.xlsx` - Solo Duplicados
Solo las 16 entregas identificadas como duplicadas

### 3. `*_resumen_*.xlsx` - Resumen por Dirección
Estadísticas agrupadas por dirección:
- Total de entregas por dirección
- Cantidad de duplicados
- Costo original vs ajustado
- Ahorro por dirección

## 🚀 Cómo Usar

### Opción 1: Aplicación Web Interactiva (Recomendado)

```bash
streamlit run app_entregas.py
```

Luego abre tu navegador en: http://localhost:8501

**Características:**
- ✅ Interfaz visual intuitiva
- ✅ Sube archivos directamente
- ✅ Visualiza duplicados en tiempo real
- ✅ Descarga resultados con un click
- ✅ Gráficos y estadísticas

### Opción 2: Script de Línea de Comandos

```bash
python3 procesar_entregas.py archivo.xls
```

o

```bash
python3 procesar_entregas.py archivo.xlsx
```

**Características:**
- ✅ Procesamiento rápido
- ✅ Automatizable (scripts, cron jobs)
- ✅ Genera 3 archivos Excel automáticamente
- ✅ Reporte en consola

## 🔧 Requisitos

```bash
pip install pandas openpyxl streamlit
```

Para archivos .xls antiguos, también necesitas LibreOffice:
```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# MacOS
brew install --cask libreoffice
```

## 🧮 Lógica de Detección

El sistema considera una entrega como duplicada cuando:

1. **Misma dirección normalizada**
   - Convierte a minúsculas
   - Elimina puntos, comas y espacios extra
   - Ejemplo: "Calle 43, 1105" = "calle 43 1105"

2. **Entregas cercanas en el tiempo**
   - Mismo día (0 días)
   - Día siguiente (1 día)
   - Hasta 2 días después

3. **Ajuste de costos**
   - Primera entrega: costo original
   - Segundas entregas y siguientes: $0

## 📈 Casos de Uso Típicos

### Caso 1: Intento fallido + Entrega exitosa
```
Día 1: 14/10 21:38 - Estado: "Nadie" - $8,700
Día 2: 15/10 11:02 - Estado: "Entregado" - $8,700 → $0 ✅
Ahorro: $8,700
```

### Caso 2: Múltiples entregas el mismo día
```
17/10 20:34 - Entrega 1 - $8,700
17/10 20:34 - Entrega 2 - $8,700 → $0 ✅
Ahorro: $8,700
```

### Caso 3: Serie de entregas
```
Día 1: Primera entrega - $8,700
Día 2: Segunda entrega - $8,700 → $0 ✅
Día 3: Tercera entrega - $8,700 → $0 ✅
Ahorro: $17,400
```

## 💡 Tips de Uso

1. **Revisar antes de aplicar**: Los archivos `*_analisis_*.xlsx` contienen tanto los costos originales como ajustados para que puedas revisar antes de procesar

2. **Filtrar por estado**: Si solo quieres ajustar entregas con estado "Nadie" o específicos, puedes filtrar el Excel después

3. **Automatización**: Usa el script de línea de comandos en un proceso automatizado para analizar lotes de archivos

4. **Validación manual**: El archivo `*_duplicados_*.xlsx` te permite revisar manualmente cada caso detectado

## 🎨 Personalización

### Cambiar el rango de días
En el código, busca:
```python
if 0 <= diferencia_dias <= 2:  # Cambiar el 2 por el número deseado
```

### Cambiar el monto de ajuste
En el código, busca:
```python
df.loc[idx, 'Costo_Ajustado'] = 0  # Cambiar 0 por otro valor
```

## 📞 Soporte

Si encuentras algún problema o tienes preguntas:
- Revisa que el archivo Excel tenga el formato correcto
- Asegúrate de que las columnas "Fecha Movimiento", "Dirección" y "Costo envio" existan
- Verifica que las fechas estén en formato DD/MM/YYYY

## 📝 Notas Importantes

- ⚠️ El sistema NO modifica el archivo original
- ✅ Genera nuevos archivos con sufijo de fecha/hora
- 🔒 Todas las operaciones son de solo lectura sobre el archivo original
- 💾 Los resultados se guardan en formato Excel (.xlsx) compatible con todas las versiones

---

**Desarrollado con ❤️ para optimizar la gestión de entregas**
