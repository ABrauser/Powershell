# Invoke-FolderScan — Offene Punkte & Status

## Stand: 2026-02-13

### ✅ Erledigt

- **Pipeline-Status Tracking** — `-StagingPath` / `-ErgebnisPath` Parameter, Dashboard zeigt Offen/Kopiert/Veredelt
- **Pipeline Filter Buttons** — Offen / Kopiert / Veredelt Buttons im Dashboard
- **Workflow Guide (5 Schritte)** — Schritt-für-Schritt Anleitung im Dashboard:
  1. Dateien filtern (Live-Zähler)
  2. Arbeitsliste exportieren & verschieben (Pipeline CSV mit Timestamp, Move-Befehl)
  3. Öffnen & prüfen (Out-GridView Befehl)
  4. In Staging kopieren (Command Builder mit Pipeline-CSV Pfad)
  5. Rescan — Status prüfen (Rescan-Befehl)
- **Einheitliche Copy-Buttons** — Alle Schritte haben identischen 📋 Kopieren Button
- **Pipeline CSV Export** — Timestamp im Dateinamen (`Pipeline_2026-02-13_1647.csv`)
- **Move-Befehl** — Downloads → scan-results, mit `downloadsPath` JS-Konstante (Backslash-Fix)
- **Copy-ScannedFiles Safety Check** — Blockiert `FullScan.csv` als Input, erfordert `Pipeline_*.csv`
- **Copy-ScannedFiles Incremental** — Skip-if-exists, Write-Progress, -Force Flag
- **-PassThru Parameter** — Gibt FileList-Objekte nur auf Wunsch zurück
- **-NoGui Parameter** — Überspringt HTML/data.js Generierung
- **Parameter-Dokumentation** — .SYNOPSIS mit allen neuen Parametern + Beispielen

### 🔲 Offen

- **Docling API Integration** — Neues Skript `Invoke-DoclingConversion.ps1`:
  - Dateien aus Staging per REST API an Docling schicken
  - Markdown-Ergebnisse im Ergebnis-Ordner ablegen
  - Write-Progress Fortschritt
  - Voraussetzung: Docling API URL/Endpoint klären (Docker lokal? Externer Service?)
- **Copy-ScannedFiles.ps1 Examples** — `.EXAMPLE` Abschnitte: `FullScan.csv` durch `Pipeline_*.csv` ersetzen
- **Optionaler Step 6: Docling** — Im Workflow-Guide einbauen sobald API-Skript fertig

### 📝 Architektur-Hinweise

- **Ordner-Vertrag**: Source (Original) → Staging (Kopie) → Ergebnis (Docling Output)
- **CSV-Vertrag**: `FullScan.csv` = Inventar (nie direkt kopieren), `Pipeline_*.csv` = kuratierte Arbeitsliste
- **Dashboard**: Statische HTML-Datei, JavaScript im Browser, keine Server-Komponente
- **Pfade in JS**: Müssen mit `-replace '\\', '\\\\'` escaped werden (PowerShell Here-String → JS String)
