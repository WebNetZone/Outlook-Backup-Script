$ErrorActionPreference = "Continue"

# Skriptverzeichnis robust ermitteln
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $ScriptDir) {
    $ScriptDir = (Get-Location).Path
}

Write-Host ""
Write-Host "  +======================================================+"
Write-Host "  |         Outlook Migration Tool - Launcher            |"
Write-Host "  |                    Version 2.0.0                     |"
Write-Host "  +======================================================+"
Write-Host ""

# Python suchen (Store-Stub ausschliessen)
function Find-Python {
    # 1. Bekannte Installationspfade direkt prüfen
    $versions = @("313","312","311","310")
    $bases = @(
        "$env:LOCALAPPDATA\Programs\Python",
        "C:\Program Files",
        "C:\Program Files (x86)",
        "C:\"
    )
    foreach ($v in $versions) {
        foreach ($b in $bases) {
            $path = "$b\Python$v\python.exe"
            if (Test-Path $path) { return $path }
        }
    }

    # 2. Python Launcher py.exe versuchen
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py -and $py.Source -and $py.Source -notlike "*WindowsApps*") {
        try {
            $ver = & $py.Source --version 2>&1
            if ($ver -match "Python 3\.(\d+)") {
                $minor = [int]$Matches[1]
                if ($minor -ge 10) { return $py.Source }
            }
        } catch {}
    }

    return $null
}

# Internet prüfen
Write-Host "  [..] Prüfe Internetverbindung..."
try {
    $internet = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet -ErrorAction SilentlyContinue
} catch {
    $internet = $false
}

# Python prüfen
Write-Host "  [..] Suche Python 3.10+..."
$pythonExe = Find-Python

if (-not $pythonExe) {
    Write-Host "  [!] Python nicht gefunden."
    if (-not $internet) {
        Write-Host "  [!!] FEHLER: Kein Internet und kein Python!"
        Write-Host "       Bitte Python 3.10+ manuell installieren:"
        Write-Host "       https://www.python.org/downloads/"
        Read-Host "  Druecken Sie Enter zum Beenden"
        exit 1
    }
    Write-Host "  [..] Lade Python 3.12 herunter (bitte warten)..."
    $installer = "$env:TEMP\python_installer.exe"
    try {
        Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe" -OutFile $installer
        Write-Host "  [..] Installiere Python..."
        Start-Process -FilePath $installer -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait
        $pythonExe = "C:\Program Files\Python312\python.exe"
    } catch {
        Write-Host "  [!!] FEHLER: Python-Download fehlgeschlagen: $_"
    }
    if (-not (Test-Path $pythonExe)) {
        Write-Host "  [!!] FEHLER: Python Installation fehlgeschlagen!"
        Read-Host "  Enter zum Beenden"
        exit 1
    }
}

Write-Host "  [OK] Python gefunden: $pythonExe"

# Libraries installieren
Write-Host "  [..] Installiere benoetigte Libraries..."
$req = Join-Path $ScriptDir "requirements.txt"
try {
    if (Test-Path $req) {
        & $pythonExe -m pip install -r $req --quiet 2>&1 | Out-Null
    } else {
        & $pythonExe -m pip install psutil --quiet 2>&1 | Out-Null
    }
} catch {
    Write-Host "  [!] Warnung: Libraries konnten nicht installiert werden: $_"
}
Write-Host "  [OK] Libraries installiert."

Write-Host ""
Write-Host "  +======================================================+"
Write-Host "  [OK] Alle Voraussetzungen erfüllt!"
Write-Host "  +======================================================+"
Write-Host ""

$start = Read-Host "  Script jetzt starten? (J/N)"
if ($start -eq "J" -or $start -eq "j") {
    Write-Host ""
    Write-Host "  [..] Starte Outlook Migration Tool..."
    & $pythonExe (Join-Path $ScriptDir "outlook_migration.py")
} else {
    Write-Host "  [..] Beendet."
}
