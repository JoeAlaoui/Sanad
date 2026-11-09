@echo off
setlocal EnableExtensions
title Sauvegarde de la base SANAD

REM ============================================================
REM  Double-cliquez pour sauvegarder la base SANAD (SQLite).
REM  Fonctionne serveur demarre. Garde les 30 dernieres copies.
REM  Pour une copie HORS du PC : mettez en parametre un dossier
REM  OneDrive / cle USB / partage reseau, par exemple :
REM     Sauvegarde-SANAD.bat "E:\Sauvegardes\SANAD"
REM ============================================================

set "DEST=%~1"
if "%DEST%"=="" set "DEST=D:\Apps\sanad_backups_db"

REM Chemin Windows -> chemin WSL (D:\x -> /mnt/d/x)
for /f "usebackq delims=" %%P in (`wsl.exe wslpath -a "%DEST%"`) do set "WSL_DEST=%%P"

wsl.exe --exec bash -lc "cd /mnt/d/Apps/sanad_app && .venv/bin/python tools/backup_sqlite.py '%WSL_DEST%' --db $HOME/sanad-data/sanad.db"
if errorlevel 1 (
    echo.
    echo [ERREUR] La sauvegarde a echoue.
    pause
    exit /b 1
)
echo.
echo Sauvegarde terminee dans %DEST%
pause
