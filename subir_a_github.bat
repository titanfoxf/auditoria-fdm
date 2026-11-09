@echo off
setlocal
title Publicar en GitHub Pages
cd /d "%~dp0"

echo =======================================================
echo    PUBLICADOR EN GITHUB PAGES - APLICATIVO FDM
echo =======================================================
echo.

git add index.html projects_data.json README.md .gitignore
git commit -m "Actualizacion aplicativo web auditoria FDM" 2>nul
git branch -M main

echo.
echo =======================================================
echo PASO REQUERIDO:
echo 1. Crea tu repositorio en: https://github.com/new
echo    (Ejemplo: auditoria-fdm, publico, sin README)
echo 2. Copia la URL del repositorio (termina en .git)
echo =======================================================
echo.

set "REPO_URL="
set /p REPO_URL=Pega aqui la URL de tu repositorio de GitHub: 

if "%REPO_URL%"=="" (
    echo.
    echo [ERROR] No ingresaste ninguna URL. Operacion cancelada.
    pause
    exit /b
)

echo.
echo Conectando con: %REPO_URL%
git remote remove origin 2>nul
git remote add origin %REPO_URL%

echo Subiendo archivos a GitHub...
git push -u origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo =======================================================
    echo   SUBIDA EXITOSA A GITHUB!
    echo =======================================================
    echo.
    echo Activa la web en 2 clics:
    echo 1. En tu repositorio GitHub, ve a: Settings -^> Pages
    echo 2. En 'Branch', selecciona 'main' y haz clic en 'Save'.
    echo En 1 minuto estara visible en tu enlace de GitHub Pages.
) else (
    echo.
    echo [AVISO] Hubo un problema al subir. Revisa tu conexion o credenciales.
)

echo.
pause
