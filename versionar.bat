@echo off
cd /d "%~dp0"

echo ============================================
echo  Versionando sped-para-csv v1.0.0
echo ============================================
echo.

if not exist ".git" (
    git init
    git config user.name "jgleuton"
    git config user.email "jgleuton@gmail.com"
    git branch -M main
    git remote add origin https://github.com/jgleuton-bot/sped-para-csv.git
)

git add README.md .gitignore sped_para_csv.py executar_sped_csv.bat versionar.bat

git commit -m "v1.0.0 - extrator de notas de entrada SPED Fiscal para CSV (paralelo, com dedup e log)"

git tag v1.0.0 2>nul

git push -u origin main --tags

echo.
if %ERRORLEVEL% EQU 0 (
    echo [OK] Versao publicada com sucesso!
) else (
    echo [ERRO] Verifique autenticacao ou status do repositorio.
)
echo.
pause
