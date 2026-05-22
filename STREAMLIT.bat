@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   LOGISTIX - Socorro Cargo Express
echo   Archivo: %~dp0STREAMFINAL.py
echo ============================================
echo.

echo Cerrando Streamlit viejo en puerto 8501 (si existe)...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8501" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
)

timeout /t 2 /nobreak >nul

echo Iniciando dashboard en http://localhost:8501
echo NO uses otro .py ni otro puerto salvo que lo sepas.
echo.

start "" "http://localhost:8501"

python -m streamlit run "%~dp0STREAMFINAL.py" --server.port 8501 --browser.serverAddress localhost

pause
