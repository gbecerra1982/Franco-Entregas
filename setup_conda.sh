#!/bin/bash
# Script para crear entorno conda para el Analizador de Entregas Duplicadas

echo "🐍 Configurando entorno conda para Analizador de Entregas Duplicadas"
echo "=================================================================="
echo ""

# Nombre del entorno
ENV_NAME="entregas_duplicadas"

# Verificar que conda esté instalado
if ! command -v conda &> /dev/null; then
    echo "❌ Conda no está instalado."
    echo ""
    echo "Para instalar Miniconda:"
    echo "  Linux/Mac: https://docs.conda.io/en/latest/miniconda.html"
    echo "  Windows: https://docs.conda.io/en/latest/miniconda.html"
    echo ""
    exit 1
fi

echo "✅ Conda encontrado: $(conda --version)"
echo ""

# Verificar si el entorno ya existe
if conda env list | grep -q "^${ENV_NAME} "; then
    echo "⚠️  El entorno '${ENV_NAME}' ya existe."
    read -p "¿Deseas eliminarlo y recrearlo? (s/n): " respuesta
    if [[ $respuesta =~ ^[Ss]$ ]]; then
        echo "🗑️  Eliminando entorno existente..."
        conda env remove -n $ENV_NAME -y
    else
        echo "ℹ️  Usando entorno existente."
        conda activate $ENV_NAME
        exit 0
    fi
fi

echo "📦 Creando entorno conda: $ENV_NAME"
echo "   (Python 3.11)"
echo ""

# Crear el entorno con Python 3.11
conda create -n $ENV_NAME python=3.11 -y

if [ $? -ne 0 ]; then
    echo "❌ Error al crear el entorno conda"
    exit 1
fi

echo ""
echo "✅ Entorno creado exitosamente"
echo ""
echo "🔧 Activando entorno..."

# Activar el entorno
source "$(conda info --base)/etc/profile.d/conda.sh"
conda activate $ENV_NAME

if [ $? -ne 0 ]; then
    echo "❌ Error al activar el entorno"
    exit 1
fi

echo "✅ Entorno activado: $ENV_NAME"
echo ""
echo "📦 Instalando dependencias..."
echo ""

# Instalar paquetes con conda (más rápido y estable)
echo "→ Instalando pandas, openpyxl, xlrd..."
conda install -c conda-forge pandas openpyxl xlrd -y

echo ""
echo "→ Instalando streamlit..."
conda install -c conda-forge streamlit -y

echo ""
echo "→ Instalando matplotlib (para gráficos)..."
conda install -c conda-forge matplotlib -y

# Verificar instalación
echo ""
echo "🔍 Verificando instalación..."
python -c "import pandas; print(f'✅ pandas {pandas.__version__}')"
python -c "import openpyxl; print(f'✅ openpyxl {openpyxl.__version__}')"
python -c "import xlrd; print(f'✅ xlrd {xlrd.__version__}')"
python -c "import streamlit; print(f'✅ streamlit {streamlit.__version__}')"
python -c "import matplotlib; print(f'✅ matplotlib {matplotlib.__version__}')"

echo ""
echo "✨ ¡Instalación completada exitosamente!"
echo ""
echo "=================================================================="
echo "📖 CÓMO USAR EL ENTORNO:"
echo "=================================================================="
echo ""
echo "1. Activar el entorno:"
echo "   conda activate $ENV_NAME"
echo ""
echo "2. Ejecutar la aplicación web:"
echo "   streamlit run app_entregas.py"
echo ""
echo "3. O ejecutar el script de línea de comandos:"
echo "   python procesar_entregas.py tu_archivo.xls"
echo ""
echo "4. Desactivar el entorno cuando termines:"
echo "   conda deactivate"
echo ""
echo "5. Para eliminar el entorno en el futuro:"
echo "   conda env remove -n $ENV_NAME"
echo ""
echo "=================================================================="
