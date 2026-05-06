# 📦 Outlook Backup & Migration Tool
> **Status:** Planung abgeschlossen – bereit zur Entwicklung  
> **Sprache:** Python + BAT Launcher  
> **Plattform:** Windows 10 / 11  
> **Version:** 3.0.0  

---

## 1. Konzept & Ziel

Ein einzelnes Python-Script das auf **beiden PCs** läuft – alter PC und neuer PC.
Das Script erkennt automatisch in welchem Modus es läuft anhand der Konfig-Datei auf dem USB-Stick.
Ein BAT-Launcher sorgt dafür dass alle Voraussetzungen erfüllt sind bevor das Script startet.

---

## 2. Launcher (start.bat)

### 2.1 Ablauf

```
start.bat doppelklick
│
├── Internetverbindung prüfen
│   └── Keine Verbindung + Python fehlt → Fehlermeldung + Abbruch
│
├── Windows Version prüfen (min. Windows 10)
│   └── Zu alt → Fehlermeldung + Abbruch
│
├── Python Version prüfen (min. 3.10)
│   ├── Nicht installiert → Python herunterladen & installieren
│   └── Zu alt → Python aktualisieren
│
├── Libraries installieren (pip install -r requirements.txt)
│
└── Benutzer fragen: Script jetzt starten?
    ├── JA → Script starten
    └── NEIN → Beenden
```

### 2.2 Fehlerbehandlung Launcher

| Situation | Lösung |
|-----------|--------|
| Keine Internetverbindung + kein Python | Fehlermeldung + Abbruch |
| Windows zu alt (unter Win 10) | Fehlermeldung + Abbruch |
| Python unter 3.10 | Automatisch aktualisieren |
| Python nicht installiert | Automatisch herunterladen & installieren |

---

## 3. Script Hauptablauf

### 3.1 Start – USB Erkennung

```
Script startet
│
├── Admin-Rechte prüfen
│   └── Keine Rechte → Automatisch UAC Fenster → Neustart als Admin
│
├── USB-Stick suchen
│   ├── Konfig gefunden → Szenario 2 (Neuer PC – Import)
│   └── Keine Konfig gefunden →
│       ├── Hardware_Info.txt vorhanden?
│       │   ├── JA → Szenario 1 (Alter PC – Outlook Backup)
│       │   │       + Option "Erneut scannen" anbieten
│       │   └── NEIN → Abfrage: Hardware scannen?
│       │       ├── JA → Szenario 0 (Hardware Scan)
│       │       └── NEIN → Szenario 1 (Alter PC – Outlook Backup)
```

---


## 4. Szenario 0 – Hardware Scan (Neuer PC, brandneuer Stick)

### 4.1 Wann wird dieser Modus ausgeführt?

- Kein Konfig auf Stick
- Keine Hardware_Info.txt auf Stick
- Benutzer wählt "Hardware scannen"

### 4.2 Ablauf

```
Szenario 0 startet
│
├── Hardware Komponenten auslesen
│   ├── Mainboard Modell
│   ├── GPU Modell
│   ├── CPU Modell
│   ├── LAN Karte
│   ├── Audio Chip
│   ├── RAM Größe
│   └── Fehler beim Auslesen → Trotzdem speichern was gefunden wurde + Warnung
│
├── Treiber Ordner erstellen
│   ├── USB-Stick/Treiber/ erstellen
│   └── Fehler → Alternativen Speicherort vorschlagen
│
├── Hardware_Info.txt speichern
│   ├── Komponenten-Namen
│   └── Download-Links zu Treiber-Webseiten
│
└── Fertig → Stick raus ✅
    Nächster Schritt: Zum alten PC gehen
```

### 4.3 Hardware_Info.txt Beispiel

```
Hardware Info - 06.05.2026
==============================

Mainboard: ASUS ROG STRIX Z890-E
→ https://www.asus.com/motherboards/rog-strix-z890-e/helpdesk_download

GPU: GIGABYTE RTX 5070 Ti
→ https://www.nvidia.com/drivers

CPU: Intel Core Ultra 7 265KF
→ https://www.intel.com/content/www/us/en/download-center

LAN: Intel I226-V
→ https://www.intel.com/content/www/us/en/download-center

Audio: Realtek ALC4080
→ https://www.realtek.com/downloads

RAM: 32 GB DDR5
CPU Kerne: 24
```

### 4.4 Ordnerstruktur auf Stick nach Scan

```
USB-Stick/
│
├── start.bat
├── outlook_migration.py
├── requirements.txt
├── README.md
│
└── Treiber/
    └── Hardware_Info.txt    ← Automatisch erstellt
        (Treiber hier manuell ablegen)
```

### 4.5 Fehlerbehandlung Hardware Scan

| # | Situation | Lösung |
|---|-----------|--------|
| H1 | Hardware nicht auslesbar | Trotzdem speichern + Warnung |
| H2 | Treiber Ordner nicht erstellbar | Alternativen Speicherort vorschlagen |
| H3 | Benutzer will erneut scannen | Option "Erneut scannen" Button |

---

## 5. Szenario 1 – Alter PC (Keine Konfig gefunden)

### 4.1 Ablauf

```
Szenario 1 startet
│
├── Outlook Version automatisch erkennen (2016/2019/2021/365)
│   └── Nicht gefunden → Fehlermeldung + Abbruch
│
├── Windows Benutzerprofile erkennen
│   ├── Nur eines → Automatisch auswählen
│   └── Mehrere → Benutzer wählt aus
│
├── Zielordner auswählen (USB-Stick)
│
├── PST-Dateien erkennen & auswählen
│   ├── Aktive PSTs aus Registry → Automatisch ✅ markiert
│   ├── Gefundene inaktive PSTs → ☐ (manuell aktivierbar)
│   ├── PST in OneDrive → ⚠️ Gesperrt + Lösung anbieten
│   │   └── OneDrive pausieren + PST verschieben anbieten
│   ├── PST auf Netzlaufwerk nicht erreichbar → Benutzer kann verbinden
│   ├── PST beschädigt → ❌ Überspringen + Warnung
│   └── Manuell hinzufügen möglich
│
├── Speicherplatz auf USB prüfen
│   └── Nicht genug → Anderen Zielort auswählen
│   (Hinweis: PST Dateien werden NICHT auf Stick kopiert – nur als Netzwerkfreigabe)
│
├── Outlook schließen
│   ├── Normal schließen → 10 Sekunden warten
│   ├── Nicht geschlossen → Zwangsbeenden (taskkill)
│   └── Zwangsbeenden fehlgeschlagen → Fehlermeldung + Abbruch
│
├── Kleine Dateien auf Stick kopieren
│   ├── Signaturen
│   ├── Outlook Einstellungen (Roaming)
│   ├── Outlook Einstellungen (Local)
│   ├── Regeln (.rwz)
│   └── Kontodaten (automatisch oder manuell eingeben)
│
├── PST Dateien als Netzwerkfreigabe einrichten
│   ├── Freigabe automatisch erstellen
│   ├── Computername speichern
│   ├── IP-Adresse speichern
│   └── MAC-Adresse speichern (falls IP sich ändert)
│
├── Konfig-Datei auf Stick speichern
│   ├── Computername
│   ├── IP-Adresse
│   ├── MAC-Adresse
│   ├── Freigabename & Pfad
│   └── Liste aller PST-Dateien
│
└── Benachrichtigung: Stick kann sicher entfernt werden ✅
```

### 4.2 Fortschritt & Fortsetzung

- Fortschritt wird in temporärer Datei gespeichert
- Bei Unterbrechung → Weitermachen wo abgebrochen
- Bei Fehler einer Datei → Weitermachen + alle Fehler am Ende anzeigen

### 4.3 Was wird auf dem Stick gespeichert

```
Outlook_Backup_2026-05-06/
│
├── Signaturen/
├── Einstellungen/
│   ├── Roaming_Outlook/
│   └── Local_Outlook/
├── Regeln/
│   └── Regeln.rwz
├── Konten_Info.txt
├── config.json          ← Konfig-Datei (Netzwerkdaten alter PC)
└── start.bat            ← Launcher für neuen PC
```

---

## 6. Szenario 2 – Neuer PC (Konfig gefunden)

### 5.1 Start – Abfrage

```
Konfig gefunden → Abfrage:
│
├── A) Kopieren & Importieren → Weiter mit 5.2
└── B) Konfig löschen (Stick zurücksetzen)
    └── Alle Daten + Konfig löschen
        Script bleibt auf Stick ✅
        Stick bereit für neuen alten PC ✅
```

### 5.2 Ablauf Kopieren & Importieren

```
Kopieren & Importieren gewählt
│
├── Admin-Rechte prüfen → UAC falls nötig
│
├── System prüfen
│   ├── Windows Version (min. Win 10)
│   │   └── Zu alt → Fehlermeldung + Abbruch
│   ├── RAM & CPU Mindestanforderungen
│   │   └── Nicht erfüllt → Abfrage: Trotzdem weitermachen oder Abbruch
│   └── Outlook installiert?
│       └── Nicht installiert → Warnung + Abbruch
│
├── Speicherplatz prüfen
│   └── Nicht genug → Anderen Zielort auswählen
│
├── Verbindung prüfen – Automatische Entscheidung:
│   ├── Alter PC im Netzwerk erreichbar?
│   │   ├── JA → Über Netzwerk kopieren (Computername → IP → MAC)
│   │   └── NEIN → Vom Stick kopieren
│   │
│   └── Während Netzwerkkopieren:
│       ├── Verbindung abbricht → Warnung + Warten + Weiter
│       ├── Alter PC ausgeschaltet → Warnung + Warten bis erreichbar
│       └── Fortsetzungsfunktion (weiter wo abgebrochen)
│
├── Daten kopieren mit Fortschrittsanzeige
│   ├── PST Dateien (Netzwerk oder Stick)
│   ├── Signaturen
│   ├── Einstellungen
│   ├── Regeln
│   └── Kontodaten
│
├── Vollständigkeit prüfen
│   └── Dateigröße vergleichen (alter PC vs neuer PC)
│       └── Datei fehlerhaft → Abfrage:
│           ├── A) Nochmal versuchen
│           ├── B) Überspringen
│           └── C) Abbrechen
│
├── Outlook Import
│   ├── PST importieren
│   ├── Signaturen einspielen
│   ├── Einstellungen wiederherstellen
│   └── Regeln importieren
│
├── Netzwerkfreigabe schließen (Optional)
│   └── Abfrage: Freigabe auf altem PC schließen?
│       ├── JA → Freigabe entfernen
│       └── NEIN → Freigabe bleibt
│
└── Abschlussbericht
    ├── Anzeige im Script-Fenster
    └── Als Datei auf Stick speichern
```

---

## 7. Abschlussbericht

```
✅ Erfolgreich:
   → PST: Outlook.pst (4,2 GB)
   → Signaturen (3 Dateien)
   → Einstellungen
   → Regeln
   → Kontodaten

⚠️ Warnungen:
   → Archiv_2022.pst – Übersprungen (beschädigt)

❌ Fehler:
   → Regeln: Keine .rwz Datei gefunden

📊 Zusammenfassung:
   → Kopiert: 5,1 GB
   → Dauer: 4 Minuten 32 Sekunden
   → Methode: Netzwerk
```

---

## 8. Vollständige Fehlerbehandlung

### 7.1 Launcher

| # | Situation | Lösung |
|---|-----------|--------|
| L1 | Keine Internetverbindung + kein Python | Fehlermeldung + Abbruch |
| L2 | Windows unter Win 10 | Fehlermeldung + Abbruch |
| L3 | Python nicht installiert | Automatisch installieren |
| L4 | Python unter 3.10 | Automatisch aktualisieren |

### 7.2 Alter PC (Szenario 1)

| # | Situation | Lösung |
|---|-----------|--------|
| A1 | Keine Admin-Rechte | Automatisch UAC Neustart |
| A2 | Outlook nicht gefunden | Fehlermeldung + Abbruch |
| A3 | Mehrere Benutzerprofile | Benutzer wählt aus |
| A4 | PST in OneDrive | Überspringen + Lösung anbieten |
| A5 | PST auf Netzlaufwerk nicht erreichbar | Benutzer kann verbinden |
| A6 | PST beschädigt | Überspringen + Warnung |
| A7 | Nicht genug Speicherplatz USB | Anderen Zielort wählen |
| A8 | Outlook schließt nicht | Zwangsbeenden |
| A9 | Zwangsbeenden fehlgeschlagen | Fehlermeldung + Abbruch |
| A10 | Eine Datei schlägt fehl | Weitermachen + Fehler am Ende |
| A11 | Registry nicht lesbar | Manuelle Eingabe Kontodaten |
| A12 | Kopieren unterbrochen | Weitermachen wo abgebrochen |

### 7.3 Neuer PC (Szenario 2)

| # | Situation | Lösung |
|---|-----------|--------|
| N1 | Keine Admin-Rechte | Automatisch UAC Neustart |
| N2 | Windows zu alt | Fehlermeldung + Abbruch |
| N3 | RAM/CPU unzureichend | Abfrage: Weitermachen oder Abbruch |
| N4 | Outlook nicht installiert | Warnung + Abbruch |
| N5 | Nicht genug Speicherplatz | Anderen Zielort wählen |
| N6 | Alter PC nicht erreichbar | Automatisch vom Stick kopieren |
| N7 | Netzwerkverbindung abbricht | Warnung + Warten + Weitermachen |
| N8 | Alter PC ausgeschaltet | Warnung + Warten bis erreichbar |
| N9 | Datei fehlerhaft (Prüfung) | Abfrage: Retry / Überspringen / Abbruch |
| N10 | Kopieren unterbrochen | Weitermachen wo abgebrochen |

---

## 9. Technische Details

### 8.1 Konfig-Datei (config.json)

```json
{
  "computer_name": "PC-ALT",
  "ip_address": "192.168.1.100",
  "mac_address": "AA:BB:CC:DD:EE:FF",
  "share_name": "OutlookBackup",
  "share_path": "\\\\PC-ALT\\OutlookBackup",
  "pst_files": [
    "C:\\Users\\Name\\Documents\\Outlook.pst"
  ],
  "backup_date": "2026-05-06",
  "outlook_version": "16.0"
}
```

### 8.2 Netzwerkverbindung Reihenfolge

1. Zuerst über **Computername** verbinden
2. Falls nicht → über **IP-Adresse**
3. Falls nicht → über **MAC-Adresse** IP suchen

### 8.3 Python Libraries

| Library | Verwendung |
|---------|------------|
| `tkinter` | Grafische Benutzeroberfläche |
| `shutil` | Dateien kopieren |
| `winreg` | Windows Registry auslesen |
| `os` / `pathlib` | Dateipfade & Ordner |
| `psutil` | Outlook Prozess & System Info |
| `subprocess` | taskkill / Admin-Rechte / Netzwerkfreigabe |
| `json` | Konfig & Fortschritt speichern |
| `datetime` | Backup-Ordner Datum |
| `threading` | Kopieren im Hintergrund |
| `socket` | Netzwerk & IP-Adresse |
| `uuid` | MAC-Adresse auslesen |

---

## 10. Entwicklungsschritte

1. BAT Launcher erstellen
2. Script Grundgerüst & GUI
3. Admin-Rechte & USB Erkennung
4. Szenario 0 – Hardware Scan
5. Szenario 1 – Alter PC komplett
6. Netzwerkfreigabe einrichten
7. Konfig-Datei erstellen
8. Szenario 2 – Neuer PC komplett
9. Netzwerk & Stick Kopierfunktion
10. Vollständigkeitsprüfung
11. Outlook Import
12. Abschlussbericht
13. Fehlerbehandlung alle Fälle
14. Testen & Feinschliff

---

> **Nächster Schritt:** Go sagen → Entwicklung beginnt!
