@echo off
setlocal EnableDelayedExpansion

set "PROJECT_DIR=%~dp0"
set "VERSION=1.0.0"
set "ZIP_NAME=MetaCleaner_v%VERSION%_Windows.zip"
set "ZIP_PATH=%PROJECT_DIR%%ZIP_NAME%"
set "DIST_DIR=%PROJECT_DIR%dist\MetaCleaner"

echo.
echo =========================================================
echo   MetaCleaner - Empaquetar Release v%VERSION%
echo =========================================================
echo.

:: Verificar que el ejecutable existe
if not exist "%DIST_DIR%\MetaCleaner.exe" (
    echo [ERROR] No se encuentra el ejecutable.
    echo         Ejecuta primero build.bat para compilar la aplicacion.
    echo.
    pause
    exit /b 1
)

:: Eliminar zip anterior si existe
if exist "%ZIP_PATH%" (
    echo [INFO] Eliminando zip anterior...
    del /q "%ZIP_PATH%"
)

:: Crear el zip con PowerShell
echo [1/2] Creando %ZIP_NAME%...
powershell -NoProfile -Command "Compress-Archive -Path '%DIST_DIR%' -DestinationPath '%ZIP_PATH%' -Force"
if errorlevel 1 (
    echo [ERROR] No se pudo crear el zip.
    pause
    exit /b 1
)

:: Mostrar tamano del zip
for %%F in ("%ZIP_PATH%") do set "TAM=%%~zF"
set /a "TAM_MB=!TAM! / 1048576"
echo [OK] Zip creado: %ZIP_NAME% (!TAM_MB! MB)

echo.
echo =========================================================
echo   SIGUIENTE PASO: Subir a GitLab como Release
echo =========================================================
echo.
echo   Archivo a subir: %ZIP_NAME%
echo   Ubicacion:       %PROJECT_DIR%
echo.
echo   Pasos en GitLab:
echo   1. Abre tu proyecto en GitLab
echo   2. Menu izquierdo: Deploy ^> Releases
echo   3. Boton "New release"
echo   4. Tag: v%VERSION%  /  Title: MetaCleaner v%VERSION%
echo   5. En "Release assets" ^> "Add asset as a link"
echo      sube el archivo %ZIP_NAME%
echo   6. Publica la release
echo.
echo =========================================================
echo.

:: Abrir la carpeta del proyecto en el Explorador
explorer "%PROJECT_DIR%"
pause
