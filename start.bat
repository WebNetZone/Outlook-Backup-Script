@echo off
chcp 65001 >nul
title Outlook Migration Tool - Launcher
color 0B

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║         Outlook Migration Tool - Launcher            ║
echo  ║                    Version 2.0.0                     ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: ── ADMIN RECHTE PRÜFEN ──────────────────────────────────
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo  [!] Admin-Rechte erforderlich. Starte neu als Administrator...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo  [✓] Admin-Rechte vorhanden.

:: ── WINDOWS VERSION PRÜFEN ───────────────────────────────
echo  [~] Prüfe Windows Version...
for /f "tokens=4-5 delims=. " %%i in ('ver') do set VERSION=%%i.%%j
if "%VERSION%" lss "10.0" (
    echo.
    echo  [✗] FEHLER: Windows Version zu alt!
    echo      Mindestanforderung: Windows 10
    echo      Gefunden: %VERSION%
    echo.
    pause
    exit /b 1
)
echo  [✓] Windows Version OK (%VERSION%)

:: ── INTERNETVERBINDUNG PRÜFEN ────────────────────────────
echo  [~] Prüfe Internetverbindung...
ping -n 1 8.8.8.8 >nul 2>&1
set INTERNET=%errorLevel%

:: ── PYTHON PRÜFEN ────────────────────────────────────────
echo  [~] Prüfe Python Installation...
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo  [!] Python nicht gefunden.
    if %INTERNET% neq 0 (
        echo.
        echo  [✗] FEHLER: Kein Internet und kein Python!
        echo      Bitte Python 3.10+ manuell installieren:
        echo      https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    )
    echo  [~] Lade Python herunter und installiere...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe' -OutFile '%TEMP%\python_installer.exe'}"
    %TEMP%\python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    if %errorLevel% neq 0 (
        echo  [✗] FEHLER: Python Installation fehlgeschlagen!
        pause
        exit /b 1
    )
    echo  [✓] Python erfolgreich installiert.
    :: PATH neu laden
    call refreshenv >nul 2>&1
) else (
    :: Python Version prüfen (min 3.10)
    for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
    for /f "tokens=1,2 delims=." %%a in ("%PYVER%") do (
        set PYMAJ=%%a
        set PYMIN=%%b
    )
    if %PYMAJ% lss 3 (
        echo  [!] Python Version zu alt: %PYVER%
        goto :update_python
    )
    if %PYMAJ% equ 3 if %PYMIN% lss 10 (
        echo  [!] Python Version zu alt: %PYVER%
        goto :update_python
    )
    echo  [✓] Python %PYVER% OK
    goto :install_libs
)

:update_python
if %INTERNET% neq 0 (
    echo  [✗] FEHLER: Kein Internet - Python kann nicht aktualisiert werden!
    echo      Bitte Python 3.10+ manuell installieren.
    pause
    exit /b 1
)
echo  [~] Aktualisiere Python...
powershell -Command "& {Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe' -OutFile '%TEMP%\python_installer.exe'}"
%TEMP%\python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
echo  [✓] Python aktualisiert.

:install_libs
:: ── LIBRARIES INSTALLIEREN ───────────────────────────────
echo  [~] Installiere benötigte Libraries...
if exist "%~dp0requirements.txt" (
    python -m pip install -r "%~dp0requirements.txt" --quiet
    if %errorLevel% neq 0 (
        echo  [✗] FEHLER: Libraries konnten nicht installiert werden!
        pause
        exit /b 1
    )
    echo  [✓] Libraries installiert.
) else (
    python -m pip install psutil --quiet
    echo  [✓] Libraries installiert.
)

:: ── SCRIPT STARTEN ───────────────────────────────────────
echo.
echo  ══════════════════════════════════════════════════════
echo  [✓] Alle Voraussetzungen erfüllt!
echo  ══════════════════════════════════════════════════════
echo.
set /p START="  Script jetzt starten? (J/N): "
if /i "%START%"=="J" (
    echo.
    echo  [~] Starte Outlook Migration Tool...
    python "%~dp0outlook_migration.py"
) else (
    echo  [~] Beendet.
)

pause
