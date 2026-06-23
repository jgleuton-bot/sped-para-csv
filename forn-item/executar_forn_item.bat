@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

echo ============================================
echo  Notas e Itens de Entrada - SPED p/ CSV/XLSX
echo ============================================
echo.

echo Verificando bibliotecas necessarias...
python -c "import openpyxl, pandas, lxml, xlsxwriter" 2>nul
if errorlevel 1 (
    echo  - Instalando/atualizando: openpyxl pandas lxml xlsxwriter ...
    python -m pip install --quiet --disable-pip-version-check openpyxl pandas lxml xlsxwriter
    if errorlevel 1 (
        echo.
        echo [ERRO] Nao foi possivel instalar as bibliotecas automaticamente.
        echo Tente manualmente:  py -m pip install openpyxl pandas lxml xlsxwriter
        echo O programa ainda gera os CSVs, mas o XLSX pode nao ser gerado.
        echo.
        pause
    ) else (
        echo  - Bibliotecas instaladas com sucesso.
    )
) else (
    echo  - Bibliotecas OK ^(openpyxl, pandas, lxml, xlsxwriter^).
)
echo.
echo Selecione as pastas/arquivos nas janelas que vao abrir.
echo.

python -u "%~dp0sped_para_csv_forn_Item.py"

echo.
pause
