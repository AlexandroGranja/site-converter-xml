@echo off
echo Iniciando o Processador de XML...
echo.

REM Verifica se o Python está instalado
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: Python não encontrado. Por favor, instale o Python 3.8 ou superior.
    echo Visite: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Verifica se o ambiente virtual existe, se não, cria
if not exist venv (
    echo Criando ambiente virtual...
    python -m venv venv
)

REM Ativa o ambiente virtual
call venv\Scripts\activate

REM Instala as dependências se necessário
if not exist venv\Lib\site-packages\flask (
    echo Instalando dependências...
    pip install -r requirements.txt
)

REM Inicia a aplicação
echo Iniciando o servidor web local...
echo.
echo Quando o servidor iniciar, acesse: http://localhost:5000
echo Para encerrar o servidor, pressione CTRL+C
echo.
python -m src.main

pause
