@echo off
setlocal EnableExtensions
title Lancement de SANAD

REM ============================================================
REM  Double-cliquez sur ce fichier pour demarrer SANAD :
REM   1. le serveur Flask demarre dans une fenetre WSL
REM   2. Microsoft Edge s'ouvre sur l'application
REM  Fermez la fenetre du serveur pour arreter SANAD.
REM ============================================================

set "APP_DIR=D:\Apps\sanad_app"
set "WSL_DIR=/mnt/d/Apps/sanad_app"
set "PORT=5050"
set "EDGE=C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

REM --- Port lu dans data\network.json (il change si vous changez le port dans l'app) ---
if exist "%APP_DIR%\data\network.json" (
    for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "$j = Get-Content -Raw '%APP_DIR%\data\network.json' | ConvertFrom-Json; $j.port"`) do set "PORT=%%P"
)
set "URL=http://127.0.0.1:%PORT%/"
echo SANAD : %URL%
echo.

REM --- Le serveur tourne deja ? ---
curl.exe -s -o nul --max-time 2 "%URL%" >nul 2>&1
if not errorlevel 1 goto ouvrir_edge

REM --- Environnement Python (venv) present dans WSL ? ---
wsl.exe --exec bash -lc "test -x %WSL_DIR%/.venv/bin/python" >nul 2>&1
if errorlevel 1 goto venv_manquant

echo Demarrage du serveur dans une fenetre WSL...
start "SANAD - Serveur" wsl.exe --exec bash -lc "cd %WSL_DIR% && exec .venv/bin/python app.py"

REM --- Attente du demarrage : 30 x 2 s max ---
set /a _essai=0
:attendre
curl.exe -s -o nul --max-time 2 "%URL%" >nul 2>&1
if not errorlevel 1 goto ouvrir_edge
set /a _essai+=1
if %_essai% geq 30 goto echec
ping -n 3 127.0.0.1 >nul 2>&1
goto attendre

:ouvrir_edge
echo Serveur accessible - ouverture de Microsoft Edge...
if exist "%EDGE%" (
    start "" "%EDGE%" "%URL%"
) else (
    start "" "msedge" "%URL%"
)
ping -n 3 127.0.0.1 >nul 2>&1
exit /b 0

:venv_manquant
echo.
echo [ERREUR] Environnement Python introuvable dans WSL.
echo.
echo Pour l'installer, ouvrez une invite WSL et executez :
echo     cd %WSL_DIR%
echo     python3 -m venv .venv
echo     .venv/bin/pip install -r requirements.txt
echo.
pause
exit /b 1

:echec
echo.
echo [ERREUR] Le serveur ne repond pas sur %URL%
echo Verifiez la fenetre "SANAD - Serveur" pour voir l'erreur.
echo.
pause
exit /b 1
