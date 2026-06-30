@echo off
title Sistema LMC - Cleodon Contabilidade
color 1F

echo.
echo  ==========================================
echo   SISTEMA LMC - CLEODON CONTABILIDADE
echo  ==========================================
echo.

:: Verificar se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado!
    echo.
    echo  Por favor, instale o Python em:
    echo  https://www.python.org/downloads/
    echo.
    echo  IMPORTANTE: Marque a opcao "Add Python to PATH"
    echo  durante a instalacao!
    echo.
    pause
    exit /b 1
)

echo  [OK] Python encontrado!

:: Instalar dependencias se necessário
if not exist "venv\" (
    echo.
    echo  Configurando o sistema pela primeira vez...
    echo  Isso pode demorar alguns minutos...
    echo.
    python -m venv venv
    call venv\Scripts\activate.bat
    pip install flask werkzeug openpyxl requests Pillow pypdf xlrd --quiet
    echo.
    echo  [OK] Sistema configurado com sucesso!
) else (
    call venv\Scripts\activate.bat
)

:: Abrir navegador automaticamente após 3 segundos
echo.
echo  Iniciando o sistema...
echo  O navegador abrira automaticamente em instantes.
echo.
echo  Para encerrar o sistema, feche esta janela.
echo.

:: Aguardar 3s e abrir navegador
start /b cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:5000"

:: Iniciar o servidor
python -m flask --app app run --host=0.0.0.0 --port=5000

pause
