@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

echo ============================================
echo  Extrator de Notas de Entrada - SPED p/ CSV
echo ============================================
echo.
echo Selecione as pastas nas janelas que vao abrir.
echo.

python -u "%~dp0sped_para_csv.py"

echo.
pause
