@echo off
cls
color 0B
title Instalacion Cameron Pipeline

echo.
echo ========================================================
echo.
echo        INSTALACION CAMERON INDIRECT PIPELINE
echo.
echo ========================================================
echo.
echo [%date% %time%] Iniciando instalacion...
echo.

REM Cambiar al directorio Cameron Tool
cd /d "%~dp0\Cameron Tool"

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] No se encuentra la carpeta Cameron Tool
    echo    Verifica que instalacion.bat este junto a Cameron Tool
    pause
    exit /b 1
)

echo [INFO] Directorio de trabajo: %CD%
echo.

REM ============================================
REM 1. Verificar Python
REM ============================================
echo [1/5] Verificando Python...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Python no encontrado
    echo.
    echo Instala Python desde: https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

python --version
echo    [OK] Python instalado correctamente
echo.

REM ============================================
REM 2. Actualizar pip
REM ============================================
echo [2/5] Actualizando pip...
python -m pip install --upgrade pip --quiet
if %ERRORLEVEL% EQU 0 (
    echo    [OK] pip actualizado
) else (
    echo    [WARN] No se pudo actualizar pip, continuando...
)
echo.

REM ============================================
REM 3. Instalar dependencias Python
REM ============================================
echo [3/5] Instalando dependencias Python...
echo    (Esto puede tardar 2-3 minutos)
echo.

if not exist "requirements.txt" (
    echo [ERROR] No se encuentra requirements.txt en cameron_tool
    pause
    exit /b 1
)

echo    Instalando paquetes...
pip install -r requirements.txt --disable-pip-version-check

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Error instalando dependencias
    echo.
    echo Intenta ejecutar manualmente:
    echo   cd cameron_tool
    echo   pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo.
echo    [OK] Dependencias Python instaladas
echo.

REM ============================================
REM 3.5 Verificar Playwright instalado
REM ============================================
echo [3.5/5] Verificando instalacion de Playwright...
python -c "import playwright" >nul 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Playwright no se instalo correctamente
    echo.
    echo Instalando Playwright manualmente...
    pip install playwright
    
    if %ERRORLEVEL% NEQ 0 (
        echo [ERROR] No se pudo instalar Playwright
        pause
        exit /b 1
    )
)

echo    [OK] Playwright verificado
echo.

REM ============================================
REM 4. Instalar navegadores Playwright
REM ============================================
echo [4/5] Instalando navegadores Playwright...
echo    (Chromium, Firefox, WebKit)
echo    (Esto puede tardar 5-10 minutos)
echo.

playwright install

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] No se pudieron instalar los navegadores
    echo.
    echo Verifica tu conexion a internet e intenta de nuevo
    echo.
    pause
    exit /b 1
) else (
    echo.
    echo    [OK] Navegadores Playwright instalados
    echo.
)


REM ============================================
REM 5. Verificar instalacion
REM ============================================
echo [5/5] Verificando instalacion...
echo.

python -c "import pandas; print('   [OK] Pandas instalado')" 2>nul
python -c "import openpyxl; print('   [OK] Openpyxl instalado')" 2>nul
python -c "import playwright; print('   [OK] Playwright instalado')" 2>nul

echo.
echo ========================================================
echo.
echo   INSTALACION COMPLETADA EXITOSAMENTE
echo.
echo ========================================================
echo.
echo   Navegadores instalados:
echo   - Chromium (para Chrome/Edge)
echo   - Firefox
echo   - WebKit (para Safari)
echo.
echo   Siguiente paso:
echo   1. Configura tus credenciales si es necesario
echo   2. Ejecuta: Ejecutar.bat
echo.
echo ========================================================
echo.
pause
