# DocPipeline — Offene Punkte & Status

## Stand: 2026-03-11

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
- **Docling API Integration** — `Invoke-DoclingConversion.ps1`:
  - POST `/v1/convert/file` Multipart-Upload an Docling Serve API
  - Bulk-Verarbeitung mit Write-Progress, ETA-Berechnung
  - Retry mit exponential Backoff (konfigurierbar, Default: 3 Retries)
  - API Health-Check vor Batch-Start
  - Skip-existing (Resume-safe bei Abbruch)
  - Konfigurierbare Ausgabeformate: Markdown (Default), HTML, Text, JSON, DocTags
  - Vollständige Docling-Optionen: Pipeline Type, OCR, PDF Backend, Table Mode, Image Export, Enrichment
  - Multi-Format Output mit Unterordnern pro Format
  - Per-File Log (`docling_log.csv`) + Run Summary (`docling_runs.json`)
  - Folder-basierter Input (`-InputPath` + `-Folders`) oder CSV-basiert (`-CsvPath`)
  - `-TimeoutSec`, `-MaxConcurrency`, `-Force`, `-AbortOnError`
- **Dashboard Docling-Panel** — Neuer Abschnitt im Dashboard:
  - Hierarchischer Ordner-Baum mit Tri-State-Checkboxen (rekursiv)
  - Lazy-Rendering für Performance bei vielen Ordnern
  - Ordner-Suche, Alle/Keine/Nur konvertierbare Buttons
  - Summary-Bar: Ordner, Dateien, %, Größe, geschätzte Dauer
  - Zeitschätzung aus `docling_runs.json` (Ø s/Datei)
  - Vollständiges Options-Panel (To Formats, Pipeline, OCR, PDF Backend, Table, Enrichment)
  - Command Builder generiert kopierbaren PowerShell-Befehl
  - localStorage-Persistenz für Docling URL, Input/Output Pfade
- **Pipeline Manifest** — `pipeline_manifest.json` als portabler Status-Tracker:
  - `Copy-ScannedFiles.ps1`: Schreibt Status `"kopiert"` nach Copy
  - `Invoke-DoclingConversion.ps1`: Schreibt Status `"veredelt"` nach Konvertierung
  - `Invoke-FolderScan.ps1`: Smart Path Detection (Offline-Fallback), Manifest-Lookup vor Test-Path
- **AnythingLLM Upload** — `Invoke-AnythingLLMUpload.ps1`:
  - Two-Phase: Upload (synchron) → Embedding (Batch-Polling auf `cached: true`)
  - Konfigurierbar: BatchSize, BatchPauseSec, EmbeddingTimeout, EmbeddingPollInterval
  - SecureString API Key Caching (Session-basiert)
  - Skip-existing (Resume-safe), Force, UploadOnly
  - Retry mit exponential Backoff
  - Logs: `upload_log.csv`, `upload_errors.csv`, `upload_runs.json`
  - Pipeline Manifest: Schreibt Status `"hochgeladen"` mit Workspace + DocumentLocation
  - Folder- und Extension-Filter
- **Dashboard AnythingLLM-Panel** — Neuer Abschnitt im Dashboard:
  - AnythingLLM URL, Workspace Slug, Document Folder Eingabefelder
  - Input-Ordner (auto-sync aus Docling Ergebnis-Ordner)
  - Extensions-Filter (.md als Default)
  - Options-Panel: Batch, Embedding, Retries, Timeout, Modus (Skip/UploadOnly)
  - Command Builder generiert kopierbaren PowerShell-Befehl
  - API Key Hinweis (SecureString, nicht im Dashboard gespeichert)
  - localStorage-Persistenz für alle Einstellungen

### 🔲 Offen

- **Copy-ScannedFiles.ps1 Examples** — `.EXAMPLE` Abschnitte: `FullScan.csv` durch `Pipeline_*.csv` ersetzen
- **README.md Update** — Neue Parameter, Docling-Workflow, AnythingLLM-Workflow, Manifest-Dokumentation
- **Praxistest Docling** — Invoke-DoclingConversion gegen echten Docling Serve testen, Response-Parsing validieren
- **Praxistest AnythingLLM** — Invoke-AnythingLLMUpload gegen laufende Instanz testen, Embedding-Polling validieren
- **MaxConcurrency** — Parallele API-Calls (Runspace/Job-basiert) implementieren (aktuell sequentiell)

### 📝 Architektur-Hinweise

- **Ordner-Vertrag**: Source (Original) → Staging (Kopie) → Ergebnis (Docling Output) → AnythingLLM (Upload)
- **CSV-Vertrag**: `FullScan.csv` = Inventar (nie direkt kopieren), `Pipeline_*.csv` = kuratierte Arbeitsliste
- **Dashboard**: Statische HTML-Datei, JavaScript im Browser, keine Server-Komponente
- **Pfade in JS**: Müssen mit `-replace '\\\\', '\\\\\\\\'` escaped werden (PowerShell Here-String → JS String)
- **Docling API**: Base URL z.B. `http://janus:8080`, Endpoint `POST /v1/convert/file` (Multipart Form)
- **Docling Logs**: `docling_log.csv` (per-file Detail) + `docling_runs.json` (Batch-Zusammenfassungen für Dashboard ETA)
- **AnythingLLM API**: Base URL z.B. `http://localhost:3001`, Upload `POST /api/v1/document/upload/{folder}`, Embedding `POST /api/v1/workspace/{slug}/update-embeddings`
- **AnythingLLM Logs**: `upload_log.csv` + `upload_errors.csv` + `upload_runs.json`
- **Pipeline Manifest**: `pipeline_manifest.json` in jedem Ordner (Staging, Ergebnis, InputPath) — portabler Status-Tracker für Offline-Szenarien
