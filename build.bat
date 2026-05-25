@echo off
setlocal EnableDelayedExpansion

set "PROJECT_DIR=%~dp0"
set "PYTHON=C:\Users\Peter\AppData\Local\Programs\Python\Python313\python.exe"

echo.
echo =========================================================
echo   MetaCleaner - Script de compilacion
echo =========================================================
echo.

:: Verificar Python
if not exist "%PYTHON%" (
    echo [ERROR] Python no encontrado en:
    echo         %PYTHON%
    echo.
    echo Edita la variable PYTHON al inicio de build.bat
    pause
    exit /b 1
)
echo [OK] Python encontrado.

:: Instalar/actualizar dependencias
echo.
echo [1/4] Instalando dependencias...
"%PYTHON%" -m pip install -r "%PROJECT_DIR%requirements.txt" --quiet
if errorlevel 1 (
    echo [ERROR] Fallo al instalar dependencias.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.

:: Instalar PyInstaller si no esta disponible
"%PYTHON%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [INFO] Instalando PyInstaller...
    "%PYTHON%" -m pip install pyinstaller --quiet
    if errorlevel 1 (
        echo [ERROR] No se pudo instalar PyInstaller.
        pause
        exit /b 1
    )
)
echo [OK] PyInstaller disponible.

:: Generar icono si no existe
if not exist "%PROJECT_DIR%assets\icon.ico" (
    echo.
    echo [INFO] Generando icono...
    "%PYTHON%" "%PROJECT_DIR%assets\create_icon.py"
)
echo [OK] Icono listo.

:: Limpiar compilacion anterior
echo.
echo [2/4] Limpiando build anterior...
if exist "%PROJECT_DIR%dist\MetaCleaner" (
    rmdir /s /q "%PROJECT_DIR%dist\MetaCleaner"
)
if exist "%PROJECT_DIR%build\metacleaner" (
    rmdir /s /q "%PROJECT_DIR%build\metacleaner"
)
echo [OK] Limpieza completada.

:: Compilar
echo.
echo [3/4] Compilando (puede tardar 1-3 minutos)...
echo.
cd /d "%PROJECT_DIR%"
"%PYTHON%" -m PyInstaller metacleaner.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo [ERROR] La compilacion fallo.
    pause
    exit /b 1
)

:: Verificar resultado
echo.
echo [4/4] Verificando resultado...
if not exist "%PROJECT_DIR%dist\MetaCleaner\MetaCleaner.exe" (
    echo [ERROR] El ejecutable no fue generado.
    pause
    exit /b 1
)

echo.
echo =========================================================
echo   COMPILACION EXITOSA
echo =========================================================
echo.
echo   Ejecutable: dist\MetaCleaner\MetaCleaner.exe
echo   Carpeta:    dist\MetaCleaner\
echo.
echo   Copia TODA la carpeta dist\MetaCleaner\ para distribuir.
echo =========================================================
echo.

explorer "%PROJECT_DIR%dist\MetaCleaner"
pause
