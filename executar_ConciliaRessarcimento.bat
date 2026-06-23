@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

echo ============================================
echo  SPED-GIA - Conciliacao do Ressarcimento ST
echo ============================================
echo.
echo Verificando dependencia (openpyxl)...
python -c "import openpyxl" 2>nul || python -m pip install openpyxl
echo.
echo Selecione nas janelas que vao abrir:
echo   1) a pasta (e subpastas) onde esta o SPED Fiscal (.txt)
echo   2) a pasta (e subpastas) onde esta a GIA (.prf)
echo   3) a pasta onde salvar os resultados
echo.

python -u "%~dp0SPED-GIA-ConciliaRessarcimento.py"

echo.
pause
