@echo off
echo ============================================
echo    UPCS GENERATOR - Instalador
echo ============================================
echo.

python --version >/dev/null 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado.
    echo Descargalo de https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" al instalar.
    pause
    exit
)

echo Instalando dependencias...
pip install streamlit pandas openpyxl requests playwright -q

echo Instalando browser para automatizacion...
playwright install chromium

echo.
echo ============================================
echo  Instalacion completa!
echo  Ejecuta "iniciar.bat" para abrir la app.
echo ============================================
pause
