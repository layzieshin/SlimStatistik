# LabStats App (PyQt6)

Eine kleine, modular aufgebaute Statistik-App für Laborbefunde (SQLite), entworfen nach MVC + Logic/Service Layer.

## Features (Mapping zu deinen Tasks)
1. **Anzahl Anforderungen pro Analyt** in Zeitraum (Task 1).  
2. **Anzahl Befunde** (offen/fertig/alle) in Zeitraum (Task 2).  
3. **Anzahl Befunde pro Wochentag** inkl. Durchschnitt (Task 3).  
4. **Filter auf Analyt** für Task 3 (Task 4).  
5. **Offene Anforderungen** für Auswahl an Analyten, älter als _x_ (Task 5).  
6. **„Blutentnahme vermutlich nicht erfolgt“**: Fälle älter als 24h, mehrere Anforderungen, **kein** Ergebnis — automatisch **überall ausgeschlossen**, separat auflistbar & **löschbar** (Task 6).  
7. **Monatliche Analyt-Statistik** seit Monatsanfang; am 1. des Monats Auto-Export nach Excel (wird erstellt, falls nicht vorhanden) (Task 7).  
8. **Config-Datei** für Arbeits­pfade + Pfadwahl im UI (Task 8).  
9. **Analyt-Liste** aus `config/analytes.txt` (zur Laufzeit änderbar), **verknüpft mit `TestKB`** (Task 9).  
10. **PyQt6-UI** (Task 10).  

## Architektur
- `models/` – Datamodelle + DB-Repository
- `logic/` – Business-Logik (Statistiken, Regeln, Exporte)
- `controller/` – Vermittler zwischen UI und Logik
- `ui/` – PyQt6-Fenster/Widgets
- `config/` – `settings.ini`, `analytes.txt`, `mapping.json` (Feld-Zuordnung)
- `resources/` – Styles
- `export/` – Excel-Dateien

## Datenbank-Annahmen (konfigurierbar)
Basierend auf deinen Tabellen (siehe `config/mapping.json`). Du kannst dort jederzeit Spalten umbenennen/neu mappen, ohne Code zu ändern.

- **Header**: Tabelle `Befund` (Primärschlüssel `ProbenNr`).
  - *Order-/Abnahmezeit*: `AbnahmeDatum` (Fallback: `TimeStamp` → `TransDatum`).
- **Zeilen/Anforderungen**: Tabelle `BefTag` (Primärschlüssel `ProbenNr, MatCode, APID, TestKB`).
  - *Analyt-Code (Suche/Filter)*: `TestKB`
  - *Analyt-Name*: `LDTName`
  - *Ergebnis-Wert*: `Ergebnis`

**Statuslogik (gemäß Vorgabe):**
- **Fertig** = **alle** Zeilen besitzen im Feld **`Ergebnis`** einen Eintrag (nicht NULL/leer).  
- **Offen** = mindestens **eine** Zeile besitzt im Feld **`Ergebnis`** **keinen** Eintrag.  
- Zeitliche Einordnung (Zeitraumfilter) erfolgt über `AbnahmeDatum` (Fallbacks via `mapping.json`).

**„Nicht entnommen?“ – Globale Exklusion + eigener Task**
- Probe mit **>1 Anforderungen**, in **keiner** Zeile ein **`Ergebnis`**, **älter als 24h** bezogen auf `AbnahmeDatum`.  
- Wird bei allen Statistiken ausgeschlossen, kann separat angezeigt & gelöscht werden (löscht `Befund` + `BefTag`).

## Installation
```bash
python -m venv .venv
# Windows:
.venv\Scriptsctivate
# Linux/macOS:
source .venv/bin/activate

pip install PyQt6 openpyxl
python app.py
```

## Konfiguration
- `config/settings.ini` – Pfade (DB, Export, Analytes, Excel-Datei)
- `config/analytes.txt` – **TestKB-Codes**, eine Zeile pro Analyt (`CODE;Optionaler Anzeigename`)
- `config/mapping.json` – Zuordnung „Fachfeld → DB-Spalte“
