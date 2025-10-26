#!/bin/bash
# Script de instalación rápida para el Analizador de Entregas Duplicadas

echo "🚀 Instalando Analizador de Entregas Duplicadas..."
echo ""

# Verificar si Python está instalado
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 no está instalado. Por favor, instala Python 3 primero."
    exit 1
fi

echo "✅ Python 3 encontrado: $(python3 --version)"

# Verificar si pip está instalado
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3 no está instalado. Por favor, instala pip3 primero."
    exit 1
fi

echo "✅ pip3 encontrado"
echo ""

# Instalar dependencias de Python
echo "📦 Instalando dependencias de Python..."
pip3 install pandas openpyxl streamlit xlrd

if [ $? -ne 0 ]; then
    echo "❌ Error al instalar dependencias de Python"
    exit 1
fi

echo "✅ Dependencias de Python instaladas"
echo ""

# Verificar si LibreOffice está instalado (opcional pero recomendado)
if command -v libreoffice &> /dev/null; then
    echo "✅ LibreOffice encontrado (para soporte de archivos .xls)"
else
    echo "⚠️  LibreOffice no encontrado"
    echo "   El script puede procesar .xlsx pero para .xls necesitas LibreOffice"
    echo ""
    echo "   Para instalar LibreOffice:"
    echo "   - Ubuntu/Debian: sudo apt-get install libreoffice"
    echo "   - MacOS: brew install --cask libreoffice"
    echo "   - Windows: Descarga desde https://www.libreoffice.org/"
    echo ""
fi

# Dar permisos de ejecución al script
chmod +x procesar_entregas.py 2>/dev/null

echo "✨ ¡Instalación completada!"
echo ""
echo "📖 Uso:"
echo ""
echo "1. Aplicación Web (Interfaz gráfica):"
echo "   streamlit run app_entregas.py"
echo ""
echo "2. Línea de Comandos:"
echo "   python3 procesar_entregas.py tu_archivo.xls"
echo ""
echo "Para más información, lee el archivo README.md"
echo ""
