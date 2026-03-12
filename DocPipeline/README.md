# 📂 Invoke-FolderScan

A PowerShell script that scans a directory and generates a beautiful, interactive HTML dashboard with file statistics, charts, and a treemap visualization.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![License](https://img.shields.io/badge/License-MIT-green)

## ✨ Features

- **Interactive Dashboard** — Single-file HTML report with dark/light theme
- **Treemap Visualization** — Folder structure analysis powered by Plotly.js, click to filter
- **Pie Charts** — File extension distribution and convertibility breakdown (SVG-based)
- **Top-10 Biggest Files** — Quick overview with size warnings (>50 MB yellow, >500 MB red)
- **DataTables** — Sortable, filterable, paginated file list with search
- **Folder Path Search** — Quickly filter files by directory path
- **Delta Detection** — Highlights new and modified files since the last scan
- **Export** — Copy, CSV, Excel, and Print export buttons
- **Responsive** — Works on desktop, tablet, and mobile
- **Offline Detection** — Shows a warning when CDN libraries are unavailable

## 🚀 Quick Start

```powershell
# 1. Load the script
. .\Invoke-FolderScan.ps1

# 2. Scan a folder
Invoke-FolderScan -Path "C:\MyData" -Recurse

# 3. The dashboard opens automatically in your browser
```

## 📋 Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `-Path` | Target directory to scan | *(required)* |
| `-Recurse` | Scan subdirectories recursively | `$false` |
| `-OutputDir` | Output directory for CSV, data.js, HTML | `./scan-results` |
| `-Theme` | Dashboard theme: `dark` or `light` | `dark` |
| `-LoadLatest` | Load the most recent scan instead of scanning | — |

## 📖 Examples

```powershell
# Recursive scan with dark theme (default)
Invoke-FolderScan -Path "D:\Projects" -Recurse

# Light theme
Invoke-FolderScan -Path "D:\Docs" -Recurse -Theme light

# Custom output directory
Invoke-FolderScan -Path "C:\Data" -Recurse -OutputDir "C:\Reports\scan"

# Reload the last scan without re-scanning
Invoke-FolderScan -LoadLatest -OutputDir "./scan-results"
```

## 📊 Dashboard Features

### Stat Cards
Summary statistics that update dynamically when filtering: file count, total size, folder count, max depth, convertible files, file types.

### Treemap (Plotly.js)
- Click a folder to **filter the file list** to that folder
- **"Ebene höher"** button to navigate up one level
- Collapsible section (lazy-loaded for performance)
- Toggle between file count and size view

### Pie Charts
- **Extension distribution** — click a segment to filter by file type
- **Convertibility** — shows how many files can be converted to Markdown/Duckling

### File Table (DataTables)
- Sort by any column
- Quick-filter buttons for top file extensions
- Folder path search field
- Size warnings: yellow (>50 MB), red (>500 MB)
- 📂 icon copies the folder path to clipboard
- Export to CSV, Excel, or Print

### Delta Detection
Run the scan twice on the same folder — the dashboard highlights:
- 🟢 **New** files added since the last scan
- 🟡 **Modified** files changed since the last scan

## 📁 Output Structure

```
scan-results/
├── Dashboard.html          # Interactive HTML dashboard
├── data.js                 # File data for the dashboard
├── FullScan.csv            # Latest scan (full CSV export)
└── FullScan_YYYY-MM-DD_HHMMSS.csv  # Timestamped scan history
```

## 🔧 Requirements

- **PowerShell 5.1+** (Windows PowerShell or PowerShell Core)
- **Internet connection** for CDN libraries (Plotly.js, jQuery, DataTables) — the dashboard shows a warning if offline

## 📄 License

[MIT](LICENSE) — Benjamin Rauser
