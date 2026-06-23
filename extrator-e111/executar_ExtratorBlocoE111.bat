@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

echo ============================================
echo  SPED-ExtratorBlocoE111 - Bloco E (E111/E112/E113)
echo ============================================
echo.
echo Verificando dependencia (openpyxl)...
python -c "import openpyxl" 2>nul || python -m pip install openpyxl
echo.
echo Selecione nas janelas que vao abrir:
echo   1) a pasta-raiz com os arquivos SPED (.txt)
echo   2) o arquivo de parametros PostosR7.csv
echo   3) a pasta onde salvar os XLSX
echo.

python -u "%~dp0SPED-ExtratorBlocoE111.py"

echo.
pause
