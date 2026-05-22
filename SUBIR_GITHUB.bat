@echo off
chcp 65001 >nul
cd /d "%~dp0"

set REPO_URL=https://github.com/FaridRojas23/logistix-sce.git

echo ============================================
echo   Subir a FaridRojas23/logistix-sce
echo ============================================

if not exist .git (
    git init -b main
)

git add STREAMFINAL.py requirements.txt README.md render.yaml runtime.txt start_render.sh RENDER_PASOS.md .gitignore .streamlit/config.toml STREAMLIT.bat
git commit -m "Dashboard Streamlit LOGISTIX para Render" 2>nul

git remote remove origin 2>nul
git remote add origin "%REPO_URL%"

echo.
echo Remoto configurado:
git remote -v
echo.
echo Subiendo (autentica en el navegador si pide)...
git push -u origin main

pause
