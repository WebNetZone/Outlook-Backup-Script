# 🔄 Outlook Migration Tool v3.0.0

Vollautomatische Outlook Sicherung & Migration mit grafischer Benutzeroberfläche.
Ein einziges Script das auf dem **alten UND neuen PC** läuft – erkennt automatisch den Modus.

---

## 🚀 So benutzt man es

### Schritt 1 – Auf dem alten PC
1. Alle 4 Dateien auf den USB-Stick kopieren
2. `start.bat` doppelklicken
3. Script sichert alles & erstellt Netzwerkfreigabe für PST-Dateien
4. Stick raus wenn Benachrichtigung erscheint ✅

### Schritt 2 – Auf dem neuen PC
1. USB-Stick einstecken
2. `start.bat` doppelklicken
3. Script erkennt Konfig automatisch
4. "Kopieren & Importieren" wählen
5. Alles wird automatisch kopiert & in Outlook importiert ✅

---

## 📦 Was wird gesichert?

| Was | Methode |
|-----|---------|
| PST-Dateien | Netzwerkfreigabe (direkt) oder USB-Stick |
| Signaturen | USB-Stick |
| Outlook Einstellungen | USB-Stick |
| Regeln (.rwz) | USB-Stick |
| Kontodaten (Server, Port) | USB-Stick |

---

## ⚙️ Voraussetzungen

- Windows 10 / 11
- Python 3.10+ (wird automatisch installiert falls nicht vorhanden)
- Microsoft Outlook (2016 / 2019 / 2021 / 365)
- USB-Stick
- Beide PCs im gleichen Netzwerk (für Netzwerk-Modus)

---

## 🔍 Automatische Erkennung

| Situation | Modus |
|-----------|-------|
| Kein Konfig auf Stick | ➡️ Alter PC Modus |
| Konfig gefunden | ➡️ Neuer PC Modus |

---

## 🌐 Netzwerk-Kopier-Reihenfolge

1. **Computername**
2. **IP-Adresse**
3. **MAC-Adresse**
4. **USB-Stick** (Fallback)

---

## 🛡️ Fehlerbehandlung

| Situation | Lösung |
|-----------|--------|
| Outlook schließt nicht | Zwangsbeenden → Abbruch |
| Mehrere Windows Profile | Benutzer wählt aus |
| PST in OneDrive | Warnung + automatisch verschieben |
| PST beschädigt | Überspringen + Warnung |
| Kein Speicherplatz | Anderen Zielort wählen |
| Keine Admin-Rechte | Automatisch UAC Neustart |
| Netzwerk bricht ab | Warten + weitermachen |
| Datei fehlerhaft | Abfrage: Nochmal / Überspringen / Abbrechen |
| Kopieren unterbrochen | Weitermachen wo abgebrochen |
| Registry nicht lesbar | Manuelle Eingabe |

---

## 👤 Autor

**WebNetZone**
