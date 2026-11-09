@echo off
setlocal EnableDelayedExpansion
title Build — Aprovar Pedidos RPA
chcp 65001 >nul

echo ==============================================
echo   Build — Aprovar Pedidos RPA
echo ==============================================
echo.

:: Ativa o venv se existir
if exist "venv\Scripts\activate.bat" (
    echo Ativando ambiente virtual...
    call venv\Scripts\activate.bat
)

:: Instala/atualiza PyInstaller
echo Verificando PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Instalando PyInstaller...
    pip install pyinstaller
)

:: Remove builds anteriores
echo.
echo Limpando builds anteriores...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist
if exist aprovar_pedidos.spec del /q aprovar_pedidos.spec

echo.
echo Compilando executavel...
echo.

pyinstaller ^
  --onefile ^
  --windowed ^
  --icon=logo.ico ^
  --name=aprovar_pedidos ^
  --add-data="logo.ico;." ^
  --add-data="conveniar.py;." ^
  --add-data="version.py;." ^
  --add-data="auto_update.py;." ^
  --hidden-import=conveniar ^
  --hidden-import=auto_update ^
  --hidden-import=version ^
  --hidden-import=dotenv ^
  --hidden-import=dotenv.main ^
  --hidden-import=openpyxl ^
  --hidden-import=openpyxl.styles ^
  --hidden-import=openpyxl.utils ^
  --hidden-import=pandas ^
  --collect-all=selenium ^
  --clean ^
  --noconfirm ^
  aprovar_pedidos.py

if errorlevel 1 (
    echo.
    echo [ERRO] A compilacao falhou. Verifique as mensagens acima.
    pause
    exit /b 1
)

echo.
echo ==============================================
echo   Concluido!
echo   Arquivo: dist\aprovar_pedidos.exe
echo ==============================================
echo.
echo Antes de distribuir:
echo   1. Copie dist\aprovar_pedidos.exe para a pasta de destino
echo   2. Crie um arquivo .env ao lado do .exe com:
echo        USUARIO=seu.login
echo        SENHA=SuaSenha
echo   3. Coloque o Chrome portatil em  %%USERPROFILE%%\Chrome\
echo.
pause
