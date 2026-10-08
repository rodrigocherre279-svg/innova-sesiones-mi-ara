@echo off
title Innova Schools - Generador de Sesiones
echo ========================================================
echo   Innova Schools - Generador de Presentaciones de Clase
echo   Iniciando interfaz web para la profesora Araceli...
echo ========================================================
echo.
cd /d "%~dp0"
python -m streamlit run app.py --browser.gatherUsageStats=false
pause
