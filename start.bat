@echo off
chcp 1252 >nul
title Outlook Migration Tool

:: Admin-Check
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Starte als Administrator...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: Alles weitere in PowerShell
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher.ps1"
pause
