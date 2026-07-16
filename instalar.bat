@echo off
title Instalacao - Sistema LMC
color 2F

echo.
echo  ==========================================
echo   INSTALACAO - SISTEMA LMC
echo   Cleodon Contabilidade
echo  ==========================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado!
    echo.
    echo  Abrindo pagina de download do Python...
    start https://www.python.org/downloads/
    echo.
    echo  IMPORTANTE ao instalar o Python:
    echo  Marque a caixa "Add Python to PATH"
    echo  antes de clicar em Install Now!
    echo.
    echo  Apos instalar o Python, execute
    echo  este arquivo novamente.
    pause
    exit /b 1
)

echo  [OK] Python encontrado!
echo.
echo  Instalando dependencias do sistema...
echo  Aguarde, isso pode levar alguns minutos...
echo.

python -m venv venv
call venv\Scripts\activate.bat
pip install flask werkzeug openpyxl requests Pillow pypdf xlrd --quiet

echo.
echo  ==========================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo  ==========================================
echo.
echo  Para usar o sistema, clique duas vezes em:
echo  INICIAR.bat
echo.
pause
