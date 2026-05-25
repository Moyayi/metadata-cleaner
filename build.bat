@echo off
setlocal EnableDelayedExpansion

set "PROJECT_DIR=%~dp0"

echo.
echo =========================================================
echo   MetaCleaner - Script de compilacion
echo =========================================================
echo.

:: ── Detectar Python automaticamente ──────────────────────────────────
::
:: Orden de busqueda:
::   1. Launcher "py" (instalado con Python en Windows, el mas fiable)
::   2. "python" en el PATH del sistema
::   3. Ruta tipica de instalacion para el usuario actual
::   4. Rutas tipicas de instalacion global

set "PYTHON="

:: Intento 1: Python Launcher (py.exe)
where py >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%P in ('py -c "import sys; print(sys.executable)"') do set "PYTHON=%%P"
)

:: Intento 2: python en PATH
if not defined PYTHON (
    where python >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%P in ('python -c "import sys; print(sys.executable)"') do set "PYTHON=%%P"
    )
)

:: Intento 3: python3 en PATH
if not defined PYTHON (
    where python3 >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%P in ('python3 -c "import sys; print(sys.executable)"') do set "PYTHON=%%P"
    )
)

:: Intento 4: rutas tipicas de instalacion por usuario
if not defined PYTHON (
    for %%V in (313 312 311 310) do (
        if not defined PYTHON (
            if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
                set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
            )
        )
    )
)

:: Intento 5: rutas tipicas de instalacion global
if not defined PYTHON (
    for %%V in (313 312 311 310) do (
        if not defined PYTHON (
            if exist "C:\Python%%V\python.exe" (
                set "PYTHON=C:\Python%%V\python.exe"
            )
        )
    )
)

:: Verificar que se encontro Python
if not defined PYTHON (
    echo [ERROR] No se encontro Python en este sistema.
    echo.
    echo Soluciones:
    echo   1. Instala Python desde https://www.python.org/downloads/
    echo      Marca la opcion "Add Python to PATH" durante la instalacion.
    echo   2. O edita este archivo y escribe la ruta manualmente:
    echo      set "PYTHON=C:\ruta\a\tu\python.exe"
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado en: %PYTHON%

:: Verificar version minima (3.10+)
for /f "delims=" %%V in ('"%PYTHON%" -c "import sys; print(sys.version_info.major*10+sys.version_info.minor)"') do set "PY_VER=%%V"
if %PY_VER% LSS 310 (
    echo [ERROR] Se requiere Python 3.10 o superior. Version encontrada es anterior.
    pause
    exit /b 1
)

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
