<#
.SYNOPSIS
    Scans a target directory and generates an interactive HTML dashboard with statistics.

.DESCRIPTION
    Invoke-FolderScan analyzes a target folder (optionally recursive), collects metadata for all files,
    exports a complete CSV, and generates an interactive HTML dashboard with:
    - Summary statistics (file count, total size, date range)
    - Pie charts by file extension and convertibility (SVG-based)
    - Interactive treemap for folder structure analysis (Plotly.js)
    - Click-to-filter: click treemap folders to filter the file list
    - Top-10 biggest files overview with size warnings (>50MB / >500MB)
    - DataTables-powered sortable, filterable, paginated file table
    - Folder path search field for quick directory filtering
    - Copy/CSV/Excel/Print export buttons
    - Delta detection (new/deleted/modified files since last scan)
    - Theme support (dark/light)
    - Responsive design for mobile/tablet
    - Offline detection with CDN fallback warning

.PARAMETER Path
    The target directory to scan. Required unless -LoadLatest is used.

.PARAMETER Recurse
    If specified, scans all subdirectories recursively.

.PARAMETER OutputDir
    Output directory for CSV, data.js, and HTML files. Defaults to './scan-results'.

.PARAMETER Theme
    Dashboard color theme. Valid values: 'dark', 'light'. Defaults to 'dark'.

.PARAMETER LoadLatest
    If specified, loads the most recent FullScan.csv from OutputDir instead of scanning.

.NOTES
    Name:    Invoke-FolderScan
    Author:  Benjamin Rauser
    Contact: Benjamin.Rauser@outlook.com
    Version History:
            2.0 - 2026-02-11 - Benjamin Rauser
            - Interactive treemap with click-to-filter and folder navigation
            - Dynamic stat cards, pie charts, and quick-filter buttons (update on folder filter)
            - Top-10 biggest files section with size warnings
            - Folder path search field
            - Collapsible treemap section (lazy-loaded for performance)
            - "Ebene hoeher" navigation button for treemap folder filter
            - Responsive design (768px / 480px breakpoints)
            - Offline detection banner
            - Column visibility dropdown styling (dark mode fix)
            - Folder icon copies path to clipboard
            - Export excludes folder icon column

            1.0 - 2026-02-11 - Benjamin Rauser
            - Initial version with file scanning, CSV export, HTML dashboard
            - Pie charts (extension distribution, convertibility)
            - DataTables with sorting, filtering, pagination, export
            - Delta detection (new/modified files since last scan)
            - Dark/light theme support

.EXAMPLE
    . .\Invoke-FolderScan.ps1
    Invoke-FolderScan -Path "C:\MyData" -Recurse

.EXAMPLE
    Invoke-FolderScan -Path "D:\Docs" -Recurse -Theme light

.EXAMPLE
    Invoke-FolderScan -LoadLatest -OutputDir "./scan-results"
#>

function Invoke-FolderScan {
  [CmdletBinding(DefaultParameterSetName = 'Scan')]
  param(
    [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'Scan')]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$Path,

    [Parameter(ParameterSetName = 'Scan')]
    [switch]$Recurse,

    [string]$OutputDir = ".\scan-results",

    [ValidateSet('dark', 'light')]
    [string]$Theme = 'dark',

    [Parameter(Mandatory = $true, ParameterSetName = 'Load')]
    [switch]$LoadLatest
  )

  # ═══════════════════════════════════════════════════════════════
  # CONFIG
  # ═══════════════════════════════════════════════════════════════
  $ConvertibleExtensions = @(
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt',
    '.txt', '.md', '.html', '.htm', '.csv', '.json', '.xml',
    '.rtf', '.odt', '.ods', '.odp', '.epub', '.eml', '.msg'
  )

  # ═══════════════════════════════════════════════════════════════
  # ENSURE OUTPUT DIR
  # ═══════════════════════════════════════════════════════════════
  if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
  }
  $OutputDir = (Resolve-Path $OutputDir).Path

  $csvPath = Join-Path $OutputDir "FullScan.csv"
  $FileList = $null
  $deltaInfo = $null
  $scanDuration = 0
  $ResolvedPath = ""

  # ═══════════════════════════════════════════════════════════════
  # MODE: LOAD LATEST
  # ═══════════════════════════════════════════════════════════════
  if ($LoadLatest) {
    if (-not (Test-Path $csvPath)) {
      Write-Warning "[Invoke-FolderScan] No previous scan found at '$csvPath'. Run a fresh scan first."
      return @()
    }
    Write-Host "`n[Invoke-FolderScan] Loading latest scan from: $csvPath" -ForegroundColor Cyan
    $FileList = Import-Csv -Path $csvPath -Encoding UTF8
    # Convert types
    $FileList = $FileList | ForEach-Object {
      $_.SizeBytes = [long]$_.SizeBytes
      $_.SizeKB = [double]$_.SizeKB
      $_.SizeMB = [double]$_.SizeMB
      $_.IsReadOnly = [bool]($_.IsReadOnly -eq 'True')
      $_.IsConvertible = [bool]($_.IsConvertible -eq 'True')
      $_
    }
    if ($FileList -isnot [System.Array]) { $FileList = @($FileList) }
    $ResolvedPath = if ($FileList.Count -gt 0) {
      # Try to find common root
      ($FileList[0].DirectoryName)
    }
    else { $OutputDir }
    Write-Host "[Invoke-FolderScan] Loaded $($FileList.Count) files from CSV." -ForegroundColor Green
  }

  # ═══════════════════════════════════════════════════════════════
  # MODE: FRESH SCAN
  # ═══════════════════════════════════════════════════════════════
  if (-not $LoadLatest) {
    $ResolvedPath = (Resolve-Path $Path).Path
    Write-Host "`n[Invoke-FolderScan] Scanning: $ResolvedPath" -ForegroundColor Cyan
    Write-Host "[Invoke-FolderScan] Recursive: $Recurse" -ForegroundColor Cyan

    $gciParams = @{ Path = $ResolvedPath; File = $true; ErrorAction = 'SilentlyContinue' }
    if ($Recurse) { $gciParams['Recurse'] = $true }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    $FileList = Get-ChildItem @gciParams | ForEach-Object {
      $ext = if ($_.Extension) { $_.Extension.ToLower() } else { '(none)' }
      [PSCustomObject]@{
        Name          = $_.Name
        Extension     = $ext
        FullPath      = $_.FullName
        DirectoryName = $_.DirectoryName
        SizeBytes     = $_.Length
        SizeKB        = [math]::Round($_.Length / 1KB, 2)
        SizeMB        = [math]::Round($_.Length / 1MB, 2)
        CreationTime  = $_.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
        LastWriteTime = $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
        IsReadOnly    = $_.IsReadOnly
        IsConvertible = ($ConvertibleExtensions -contains $ext)
      }
    }

    $stopwatch.Stop()
    $scanDuration = $stopwatch.Elapsed.TotalSeconds

    if (-not $FileList -or $FileList.Count -eq 0) {
      Write-Warning "[Invoke-FolderScan] No files found in '$ResolvedPath'."
      return @()
    }
    if ($FileList -isnot [System.Array]) { $FileList = @($FileList) }

    Write-Host "[Invoke-FolderScan] Found $($FileList.Count) files ($([math]::Round(($FileList | Measure-Object -Property SizeBytes -Sum).Sum / 1MB, 2)) MB) in $([math]::Round($scanDuration, 2))s" -ForegroundColor Green

    # ─── DELTA DETECTION ───
    if (Test-Path $csvPath) {
      Write-Host "[Invoke-FolderScan] Previous scan found. Calculating delta..." -ForegroundColor Yellow
      $previousFiles = Import-Csv -Path $csvPath -Encoding UTF8

      $prevPaths = @{}
      foreach ($pf in $previousFiles) {
        $prevPaths[$pf.FullPath] = $pf.LastWriteTime
      }

      $currPaths = @{}
      foreach ($cf in $FileList) {
        $currPaths[$cf.FullPath] = $cf.LastWriteTime
      }

      $newFiles = @($FileList | Where-Object { -not $prevPaths.ContainsKey($_.FullPath) })
      $deletedPaths = @($previousFiles | Where-Object { -not $currPaths.ContainsKey($_.FullPath) })
      $modifiedFiles = @($FileList | Where-Object {
          $prevPaths.ContainsKey($_.FullPath) -and $prevPaths[$_.FullPath] -ne $_.LastWriteTime
        })

      $deltaInfo = @{
        NewCount      = $newFiles.Count
        DeletedCount  = $deletedPaths.Count
        ModifiedCount = $modifiedFiles.Count
        NewPaths      = ($newFiles | Select-Object -ExpandProperty FullPath)
        DeletedPaths  = ($deletedPaths | Select-Object -ExpandProperty FullPath)
        ModifiedPaths = ($modifiedFiles | Select-Object -ExpandProperty FullPath)
        PreviousDate  = (Get-Item $csvPath).LastWriteTime.ToString('yyyy-MM-dd HH:mm')
      }

      Write-Host "[Invoke-FolderScan] Delta: $($deltaInfo.NewCount) new, $($deltaInfo.DeletedCount) deleted, $($deltaInfo.ModifiedCount) modified" -ForegroundColor Cyan

      # Rename old CSV
      $oldDate = (Get-Item $csvPath).LastWriteTime.ToString('yyyy-MM-dd_HHmmss')
      $archivePath = Join-Path $OutputDir "FullScan_$oldDate.csv"
      Move-Item -Path $csvPath -Destination $archivePath -Force
      Write-Host "[Invoke-FolderScan] Previous scan archived: $archivePath" -ForegroundColor DarkGray
    }

    # ─── EXPORT CSV ───
    $FileList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "[Invoke-FolderScan] CSV exported: $csvPath" -ForegroundColor Green
  }

  # ═══════════════════════════════════════════════════════════════
  # BUILD STATISTICS
  # ═══════════════════════════════════════════════════════════════
  $totalFiles = $FileList.Count
  $totalSizeBytes = ($FileList | Measure-Object -Property SizeBytes -Sum).Sum
  $convertibleCount = @($FileList | Where-Object { $_.IsConvertible }).Count
  $nonConvertibleCount = $totalFiles - $convertibleCount

  $extGroups = $FileList | Group-Object Extension | Sort-Object Count -Descending
  $top10 = $extGroups | Select-Object -First 10
  $othersCount = ($extGroups | Select-Object -Skip 10 | Measure-Object -Property Count -Sum).Sum
  $othersSize = 0
  if ($extGroups.Count -gt 10) {
    $othersSize = ($extGroups | Select-Object -Skip 10 | ForEach-Object {
        ($_.Group | Measure-Object -Property SizeBytes -Sum).Sum
      } | Measure-Object -Sum).Sum
  }

  # Chart JSON
  $extChartData = @()
  foreach ($g in $top10) {
    $gSize = ($g.Group | Measure-Object -Property SizeBytes -Sum).Sum
    $extChartData += @{
      label   = $g.Name
      count   = $g.Count
      sizeMB  = [math]::Round($gSize / 1MB, 2)
      percent = [math]::Round(($g.Count / $totalFiles) * 100, 1)
    }
  }
  if ($othersCount -gt 0) {
    $extChartData += @{
      label   = 'Other'
      count   = $othersCount
      sizeMB  = [math]::Round($othersSize / 1MB, 2)
      percent = [math]::Round(($othersCount / $totalFiles) * 100, 1)
    }
  }

  $extChartJson = ($extChartData | ForEach-Object {
      "{`"label`":`"$($_.label)`",`"count`":$($_.count),`"sizeMB`":$($_.sizeMB),`"percent`":$($_.percent)}"
    }) -join ','

  # Human-readable total size
  $totalSizeDisplay = if ($totalSizeBytes -ge 1GB) {
    "$([math]::Round($totalSizeBytes / 1GB, 2)) GB"
  }
  elseif ($totalSizeBytes -ge 1MB) {
    "$([math]::Round($totalSizeBytes / 1MB, 2)) MB"
  }
  else {
    "$([math]::Round($totalSizeBytes / 1KB, 2)) KB"
  }

  # Folders & depth
  $uniqueFolders = ($FileList | Select-Object -ExpandProperty DirectoryName -Unique).Count
  $maxDepth = 0
  if (-not $LoadLatest) {
    foreach ($f in $FileList) {
      if ($f.DirectoryName.Length -gt $ResolvedPath.Length) {
        $rel = $f.DirectoryName.Substring($ResolvedPath.Length)
        $depth = ($rel.Split('\', [System.StringSplitOptions]::RemoveEmptyEntries)).Count
        if ($depth -gt $maxDepth) { $maxDepth = $depth }
      }
    }
  }

  # ═══════════════════════════════════════════════════════════════
  # GENERATE data.js (separate file for performance)
  # ═══════════════════════════════════════════════════════════════
  Write-Host "[Invoke-FolderScan] Generating data.js..." -ForegroundColor Yellow
  $dataJsPath = Join-Path $OutputDir "data.js"

  # Build delta lookup sets for status column
  $deltaNewSet = @{}
  $deltaModSet = @{}
  if ($deltaInfo) {
    foreach ($p in $deltaInfo.NewPaths) { $deltaNewSet[$p] = $true }
    foreach ($p in $deltaInfo.ModifiedPaths) { $deltaModSet[$p] = $true }
  }

  $sw = [System.IO.StreamWriter]::new($dataJsPath, $false, [System.Text.Encoding]::UTF8)
  $sw.WriteLine("const allFiles = [")
  $i = 0
  foreach ($f in $FileList) {
    $n = ($f.Name -replace '\\', '\\\\' -replace '"', '\"').Replace("'", "\\'")
    $fp = ($f.FullPath -replace '\\', '\\\\' -replace '"', '\"').Replace("'", "\\'")
    $dn = ($f.DirectoryName -replace '\\', '\\\\' -replace '"', '\"').Replace("'", "\\'")

    # Delta status
    $status = 'unchanged'
    if ($deltaNewSet.ContainsKey($f.FullPath)) { $status = 'new' }
    elseif ($deltaModSet.ContainsKey($f.FullPath)) { $status = 'modified' }

    $comma = if ($i -lt $FileList.Count - 1) { ',' } else { '' }
    $sw.WriteLine("{`"name`":`"$n`",`"ext`":`"$($f.Extension)`",`"fullPath`":`"$fp`",`"dir`":`"$dn`",`"sizeBytes`":$($f.SizeBytes),`"sizeKB`":$($f.SizeKB),`"sizeMB`":$($f.SizeMB),`"created`":`"$($f.CreationTime)`",`"modified`":`"$($f.LastWriteTime)`",`"readOnly`":$($f.IsReadOnly.ToString().ToLower()),`"convertible`":$($f.IsConvertible.ToString().ToLower()),`"status`":`"$status`"}$comma")
    $i++
  }
  $sw.WriteLine("];")

  # Also write deleted files for delta display
  if ($deltaInfo -and $deltaInfo.DeletedCount -gt 0) {
    $sw.WriteLine("const deletedFiles = [")
    for ($d = 0; $d -lt $deltaInfo.DeletedPaths.Count; $d++) {
      $dp = ($deltaInfo.DeletedPaths[$d] -replace '\\', '\\\\' -replace '"', '\"').Replace("'", "\\'")
      $comma = if ($d -lt $deltaInfo.DeletedPaths.Count - 1) { ',' } else { '' }
      $sw.WriteLine("`"$dp`"$comma")
    }
    $sw.WriteLine("];")
  }
  else {
    $sw.WriteLine("const deletedFiles = [];")
  }

  $sw.Close()
  $dataJsSize = [math]::Round((Get-Item $dataJsPath).Length / 1MB, 2)
  Write-Host "[Invoke-FolderScan] data.js: $dataJsPath ($dataJsSize MB)" -ForegroundColor Green

  # ═══════════════════════════════════════════════════════════════
  # THEME CSS VARIABLES
  # ═══════════════════════════════════════════════════════════════
  $themeCss = if ($Theme -eq 'light') {
    @"
    --bg-primary: #F0F4F9;
    --bg-secondary: #E9EEF6;
    --bg-card: #FFFFFF;
    --bg-glass: #E9EEF6;
    --border-glass: transparent;
    --text-primary: #1F1F1F;
    --text-secondary: #444746;
    --text-muted: #747775;
    --accent: #0B57D0;
    --accent-glow: rgba(11, 87, 208, 0.15);
    --success: #146C2E;
    --warning: #EF6C00;
    --danger: #B3261E;
    --table-stripe: #F8F9FA;
    --table-hover: #E9EEF6;
    --shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    --radius-card: 24px;
    --radius-btn: 20px;
    --font-head: 'Google Sans', 'Inter', sans-serif;
"@
  }
  else {
    @"
    --bg-primary: #0E0E11; /* Deep Space */
    --bg-secondary: #131314;
    --bg-card: #1E1F20;
    --bg-glass: rgba(255, 255, 255, 0.05);
    --border-glass: rgba(255, 255, 255, 0.1);
    --text-primary: #E3E3E3;
    --text-secondary: #C4C7C5;
    --text-muted: #8E918F;
    --accent: #A8C7FA;
    --accent-glow: rgba(168, 199, 250, 0.1);
    --success: #6DD58C; /* Pastel Green */
    --warning: #FFB74D; /* Pastel Orange */
    --danger: #E25C5C;  /* Pastel Red */
    --table-stripe: rgba(255, 255, 255, 0.02);
    --table-hover: rgba(168, 199, 250, 0.08);
    --shadow: 0 4px 8px rgba(0,0,0,0.3);
    --radius-card: 24px;
    --radius-btn: 20px;
    --font-head: 'Google Sans', 'Inter', sans-serif;
"@
  }

  # ═══════════════════════════════════════════════════════════════
  # DELTA HTML SECTION
  # ═══════════════════════════════════════════════════════════════
  $deltaHtml = ""
  $deltaJs = "const HAS_DELTA = false; const DELTA_INFO = null;"
  if ($deltaInfo) {
    $deltaJs = "const HAS_DELTA = true; const DELTA_INFO = {newCount:$($deltaInfo.NewCount),deletedCount:$($deltaInfo.DeletedCount),modifiedCount:$($deltaInfo.ModifiedCount),previousDate:`"$($deltaInfo.PreviousDate)`"};"
    $deltaHtml = @"
  <div class="delta-section">
    <h2>🔄 Änderungen seit $($deltaInfo.PreviousDate)</h2>
    <div class="delta-grid">
      <div class="delta-card delta-new">
        <div class="delta-value">$($deltaInfo.NewCount)</div>
        <div class="delta-label">🟢 Neu</div>
      </div>
      <div class="delta-card delta-deleted">
        <div class="delta-value">$($deltaInfo.DeletedCount)</div>
        <div class="delta-label">🔴 Gelöscht</div>
      </div>
      <div class="delta-card delta-modified">
        <div class="delta-value">$($deltaInfo.ModifiedCount)</div>
        <div class="delta-label">🟡 Geändert</div>
      </div>
    </div>
  </div>
"@
  }

  # Scan info text
  $scanInfoText = if ($LoadLatest) {
    "Loaded from CSV · $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
  }
  else {
    "Scanned in $([math]::Round($scanDuration, 2))s · $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
  }

  $convertibleListHtml = ($ConvertibleExtensions | Sort-Object | ForEach-Object { "<span class='ext-tag'>$_</span>" }) -join ''

  # ═══════════════════════════════════════════════════════════════
  # GENERATE HTML DASHBOARD
  # ═══════════════════════════════════════════════════════════════
  $htmlContent = @"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Folder Scan Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<!-- DataTables CSS -->
<link rel="stylesheet" href="https://cdn.datatables.net/2.3.7/css/dataTables.dataTables.min.css">
<link rel="stylesheet" href="https://cdn.datatables.net/buttons/3.2.0/css/buttons.dataTables.min.css">
<style>
  :root {
$themeCss
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg-primary);
    color: var(--text-primary);
    line-height: 1.6;
    min-height: 100vh;
  }
  .container { max-width: 98%; margin: 0 auto; padding: 1.5rem 2rem; }
  .path-cell-inner { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 0; min-width: 100%; direction: rtl; text-align: left; }
  /* Active states for filters */
  .qf-btn.active-filter {
    background: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important;
  }
  .chart-segment { transition: transform 0.2s, filter 0.2s; cursor: pointer; }
  .chart-segment:hover, .chart-segment.active { opacity: 0.8; }
  .chart-segment.active { stroke: var(--bg-card); stroke-width: 2px; transform: scale(1.05); filter: brightness(1.1); }

  /* === HEADER === */
  .header { margin-bottom: 3rem; text-align: left; position: relative; }
  .header h1 {
    font-family: var(--font-head); font-size: 2.2rem; font-weight: 500;
    color: var(--text-primary); letter-spacing: -0.02em; margin-bottom: 0.5rem;
  }
  .header h1 span {
    background: linear-gradient(90deg, #4285F4, #9B72CB, #D96570);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;
    font-weight: 600;
  }
  .scan-path {
    font-family: 'JetBrains Mono', monospace; font-size: 0.85rem; color: var(--text-secondary);
    background: var(--bg-secondary); display: inline-block; padding: 0.4rem 1rem;
    border-radius: 99px; margin-top: 0.8rem;
  }
  .scan-meta { font-size: 0.85rem; color: var(--text-muted); margin-top: 0.5rem; }

  /* === STAT CARDS === */
  .stats-grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 1rem; margin-bottom: 2rem;
  }
  .stat-card {
    background: var(--bg-card); border-radius: var(--radius-card); padding: 1.5rem;
    display: flex; flex-direction: column; align-items: flex-start; justify-content: space-between;
    box-shadow: none; border: 1px solid var(--border-glass); transition: all 0.2s;
  }
  .stat-card:hover { background: var(--bg-secondary); transform: translateY(-2px); }
  .stat-value { font-family: var(--font-head); font-size: 2rem; font-weight: 400; color: var(--text-primary); line-height: 1.1; }
  .stat-label { font-size: 0.75rem; color: var(--text-secondary); margin-top: 0.5rem; font-weight: 500; }

  /* === CHARTS === */
  .chart-section {
    display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; margin-bottom: 2rem;
  }
  @media (max-width: 900px) { .chart-section { grid-template-columns: 1fr; } }
  .chart-card {
    background: var(--bg-card); border-radius: var(--radius-card); padding: 2rem;
    border: 1px solid var(--border-glass);
  }
  .chart-card h2 { font-family: var(--font-head); font-size: 1.1rem; font-weight: 500; margin-bottom: 1.5rem; color: var(--text-primary); }
  .chart-container { display: flex; align-items: center; justify-content: center; gap: 2rem; flex-wrap: wrap; }
  .legend { display: flex; flex-direction: column; gap: 0.6rem; font-size: 0.85rem; }
  .legend-item { display: flex; align-items: center; gap: 0.8rem; color: var(--text-secondary); }
  .legend-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
  .legend-count { color: var(--text-muted); font-size: 0.85rem; margin-left: auto; padding-left: 1rem; }
  .pie-tooltip {
    position: absolute; background: var(--bg-secondary); border-radius: 8px;
    padding: 0.6rem 1rem; font-size: 0.85rem; color: var(--text-primary);
    pointer-events: none; opacity: 0; transition: opacity 0.15s; z-index: 100;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15); border: 1px solid var(--border-glass);
  }

  /* === DELTA === */
  .delta-section {
    background: var(--bg-card); border-radius: var(--radius-card); padding: 1.5rem;
    margin-bottom: 2rem; border: 1px solid var(--border-glass); border-left: 4px solid var(--accent); box-shadow: var(--shadow);
  }
  .delta-section h2 { font-family: var(--font-head); font-size: 1.1rem; font-weight: 500; margin-bottom: 1rem; }
  .delta-grid { display: flex; gap: 1rem; flex-wrap: wrap; }
  .delta-card {
    flex: 1; min-width: 140px; padding: 1rem; border-radius: 16px; text-align: center;
    background: var(--bg-secondary); border: 1px solid var(--border-glass);
  }
  .delta-new .delta-value { color: var(--success); }
  .delta-deleted .delta-value { color: var(--danger); }
  .delta-modified .delta-value { color: var(--warning); }
  .delta-value { font-size: 1.8rem; font-weight: 500; }
  .delta-label { font-size: 0.8rem; color: var(--text-secondary); margin-top: 0.2rem; }

  /* === TABLE SECTION === */
  .table-section {
    background: var(--bg-card); border: 1px solid var(--border-glass); border-radius: var(--radius-card); padding: 2rem; margin-bottom: 2rem;
  }
  .table-section h2 { font-family: var(--font-head); font-size: 1.1rem; font-weight: 500; margin-bottom: 1.5rem; }

  /* DataTables Theme Override */
  table.dataTable { color: var(--text-secondary); border-collapse: separate !important; border-spacing: 0; width: 100% !important; }
  table.dataTable thead th {
    background: var(--bg-card) !important; color: var(--text-secondary) !important;
    border-bottom: 1px solid var(--border-glass) !important; font-size: 0.75rem;
    text-transform: uppercase; letter-spacing: 0.05em; padding: 1rem 0.8rem !important;
    font-weight: 600 !important;
  }
  table.dataTable tbody td {
    padding: 0.8rem !important; border-bottom: 1px solid var(--border-glass) !important;
    font-size: 0.85rem; vertical-align: middle;
  }
  table.dataTable tbody tr:hover { background: var(--table-hover) !important; }

  /* DataTables controls */
  .dt-container .dt-search input, .dt-container .dt-length select {
    background: var(--bg-secondary) !important; border: none !important;
    border-radius: 8px !important; color: var(--text-primary) !important; padding: 0.6rem 1rem !important;
  }
  .dt-paging button.current {
    background: var(--accent) !important; color: #fff !important; border-radius: 50% !important;
    width: 32px; height: 32px; padding: 0 !important; line-height: 32px; border: none !important;
  }
  .dt-paging button:hover:not(.current) {
    background: var(--bg-secondary) !important; border-radius: 50% !important; border: none !important;
  }

  /* Buttons */
  div.dt-buttons .dt-button {
    background: var(--bg-secondary) !important; border: none !important;
    color: var(--text-primary) !important; border-radius: 99px !important;
    padding: 0.5rem 1.2rem !important; font-size: 0.85rem !important; margin-right: 0.5rem !important;
    box-shadow: none !important;
  }
  div.dt-buttons .dt-button:hover { background: var(--border-glass) !important; }

  /* ColVis Dropdown */
  div.dt-button-collection { background: var(--bg-card) !important; border: 1px solid var(--border-glass) !important; border-radius: var(--radius-card) !important; padding: 0.5rem 0 !important; box-shadow: 0 8px 24px rgba(0,0,0,0.3) !important; }
  div.dt-button-collection .dt-button { background: transparent !important; color: var(--text-primary) !important; border-radius: 0 !important; padding: 0.5rem 1.2rem !important; margin: 0 !important; display: block !important; width: 100% !important; text-align: left !important; }
  div.dt-button-collection .dt-button:hover { background: var(--bg-secondary) !important; color: var(--accent) !important; }
  div.dt-button-collection .dt-button.dt-button-active { background: var(--bg-secondary) !important; color: var(--accent) !important; font-weight: 600 !important; }
  div.dt-button-collection .dt-button.dt-button-active::before { content: '\2713 '; color: var(--accent); }

  /* Cell styles */
  .cell-name { font-weight: 500; color: var(--text-primary); }
  .badge { padding: 0.25rem 0.8rem; border-radius: 99px; font-size: 0.75rem; font-weight: 500; }
  .badge-yes { background: #E6F4EA; color: #137333; }
  .badge-no { background: #F1F3F4; color: #5F6368; }
  /* Dark mode specific badges adjustment via opacity if needed, simplfied here */
  :root[style*="--bg-primary: #0E0E11"] .badge-yes { background: rgba(109, 213, 140, 0.2); color: #6DD58C; }
  :root[style*="--bg-primary: #0E0E11"] .badge-no { background: rgba(255, 255, 255, 0.1); color: #C4C7C5; }

  .folder-link {
    color: var(--text-muted); text-decoration: none; font-size: 1.1rem;
    width: 32px; height: 32px; display: inline-flex; align-items: center; justify-content: center;
    border-radius: 50%; transition: background 0.2s;
  }
  .folder-link:hover { background: var(--bg-secondary); color: var(--accent); }

  /* Quick filters */
  .quick-filters { display: flex; gap: 0.8rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
  .qf-btn {
    padding: 0.5rem 1.2rem; border-radius: 99px; border: 1px solid var(--border-glass);
    background: var(--bg-secondary); color: var(--text-primary); cursor: pointer; transition: all 0.2s;
    font-size: 0.85rem; font-weight: 500;
  }
  .qf-btn:hover, .qf-btn.active {
    background: var(--bg-secondary); color: var(--accent); border-color: var(--accent);
  }
  .qf-btn-accent { background: var(--bg-secondary); color: var(--text-primary); border: none; }
  .qf-btn-accent:hover { background: #E9EEF6; color: #0B57D0; }
  :root[style*="--bg-primary: #0E0E11"] .qf-btn-accent:hover { background: #2D2E30; color: #A8C7FA; }

  /* Info box */
  .info-box {
    background: var(--bg-card); border-radius: var(--radius-card); padding: 2rem; margin-top: 2rem;
  }
  .ext-tag {
    display: inline-block; padding: 0.2rem 0.6rem; margin: 0.2rem;
    background: var(--bg-secondary); border-radius: 8px; color: var(--text-secondary); font-size: 0.8rem; font-family: 'JetBrains Mono';
  }

  /* Animation overrides */
  @keyframes fadeInDown { from { opacity: 0; transform: translateY(-10px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes fadeInUp { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
  .ext-tag {
    display: inline-block; padding: 0.15rem 0.5rem; margin: 0.1rem;
    background: rgba(74, 222, 128, 0.1); border: 1px solid rgba(74, 222, 128, 0.2);
    border-radius: 6px; color: var(--success);
  }

  /* Animations */
  @keyframes fadeInDown { from { opacity: 0; transform: translateY(-15px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes fadeInUp { from { opacity: 0; transform: translateY(15px); } to { opacity: 1; transform: translateY(0); } }

  /* Structure Section */
  .structure-section { background: var(--bg-card); border-radius: var(--radius-card); padding: 2rem; margin-bottom: 2rem; border: 1px solid var(--border-glass); }
  .section-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; flex-wrap: wrap; gap: 1rem; }
  .section-header h2 { margin: 0; font-family: var(--font-head); font-size: 1.1rem; font-weight: 500; cursor: pointer; }
  .section-header h2 .collapse-arrow { display: inline-block; transition: transform 0.3s; margin-right: 0.3rem; font-size: 0.8rem; }
  .section-header h2 .collapse-arrow.collapsed { transform: rotate(-90deg); }
  .treemap-body { overflow: hidden; transition: max-height 0.4s ease, opacity 0.3s ease; max-height: 800px; opacity: 1; }
  .treemap-body.collapsed { max-height: 0; opacity: 0; margin: 0; padding: 0; }
  #treemap-container { width: 100%; height: 600px; background: var(--bg-secondary); border-radius: 12px; }
  .treemap-hint { margin-top: 0.5rem; font-size: 0.8rem; color: var(--text-muted); text-align: right; }
  .toggle-group { display: flex; gap: 0.5rem; background: var(--bg-secondary); padding: 4px; border-radius: 99px; border: 1px solid var(--border-glass); }
  .toggle-group .qf-btn { border: none; background: transparent; padding: 0.3rem 1rem; font-size: 0.8rem; margin: 0; }
  .toggle-group .qf-btn.active { background: var(--bg-card); color: var(--accent); box-shadow: 0 1px 3px rgba(0,0,0,0.1); }

  /* Compact GridView Style */
  table.dataTable.compact { border-spacing: 0 !important; }
  table.dataTable.compact thead th, table.dataTable.compact thead td { padding: 4px 6px; font-size: 0.75rem; border-bottom: 1px solid var(--border-glass); }
  table.dataTable.compact tbody th, table.dataTable.compact tbody td { padding: 0 4px !important; font-size: 0.75rem; line-height: 18px !important; height: 18px !important; border-bottom: 1px solid var(--border-glass); vertical-align: middle !important; }
  /* Compact internal elements */
  table.dataTable.compact .path-cell-inner { line-height: 18px !important; font-size: 0.75rem; max-width: 0; min-width: 100%; }
  table.dataTable.compact .badge { padding: 0 4px; font-size: 0.7rem; line-height: 14px; height: 16px; display: inline-flex; align-items: center; }
  table.dataTable.compact .folder-link { width: 16px; height: 16px; font-size: 0.8rem; line-height: 14px; }
  table.dataTable.compact .cell-size { font-family: 'JetBrains Mono'; }

  /* Size warnings */
  .size-warn { color: #F4B400 !important; font-weight: 600; }
  .size-danger { color: #D96570 !important; font-weight: 700; }

  /* Top-10 Section */
  .top10-section { background: var(--bg-card); border-radius: var(--radius-card); padding: 1.5rem 2rem; margin-bottom: 2rem; border: 1px solid var(--border-glass); }
  .top10-section h2 { font-family: var(--font-head); font-size: 1.1rem; font-weight: 500; margin-bottom: 1rem; cursor: pointer; }
  .top10-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 0.5rem; }
  .top10-item { display: flex; justify-content: space-between; align-items: center; padding: 0.5rem 0.8rem; background: var(--bg-secondary); border-radius: 8px; font-size: 0.8rem; gap: 0.5rem; }
  .top10-name { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--text-primary); font-weight: 500; }
  .top10-size { font-family: 'JetBrains Mono', monospace; color: var(--accent); white-space: nowrap; }
  .top10-bar { height: 3px; background: var(--accent); border-radius: 2px; margin-top: 3px; transition: width 0.3s; }

  /* Folder search */
  .folder-search { display: flex; gap: 0.5rem; align-items: center; margin-bottom: 1rem; }
  .folder-search input { flex: 1; padding: 0.5rem 1rem; border-radius: 99px; border: 1px solid var(--border-glass); background: var(--bg-secondary); color: var(--text-primary); font-size: 0.85rem; outline: none; }
  .folder-search input:focus { border-color: var(--accent); }
  .folder-search input::placeholder { color: var(--text-muted); }

  /* Offline warning */
  .offline-banner { background: #F4B400; color: #000; padding: 0.6rem 1.5rem; border-radius: var(--radius-card); margin-bottom: 1rem; font-size: 0.85rem; display: none; text-align: center; }

  /* Responsive */
  @media (max-width: 768px) {
    .container { padding: 1rem; }
    .header h1 { font-size: 1.5rem; }
    .stats-grid { grid-template-columns: repeat(2, 1fr); gap: 0.5rem; }
    .stat-card { padding: 1rem; }
    .stat-value { font-size: 1.4rem; }
    .chart-section { grid-template-columns: 1fr; }
    .chart-card { padding: 1rem; }
    .chart-container { flex-direction: column; }
    .section-header { flex-direction: column; align-items: flex-start; }
    #treemap-container { height: 400px; }
    .top10-list { grid-template-columns: 1fr; }
    table.dataTable.compact tbody td { font-size: 0.7rem; }
  }
  @media (max-width: 480px) {
    .stats-grid { grid-template-columns: 1fr 1fr; }
    .stat-value { font-size: 1.2rem; }
    .quick-filters { gap: 0.4rem; }
    .qf-btn { font-size: 0.75rem; padding: 0.4rem 0.8rem; }
  }

  /* Print */
  @media print {
    body { background: #fff; color: #000; }
    .chart-section, .delta-section, .info-box, .quick-filters, .dt-buttons, .dt-search, .dt-length, .structure-section, .top10-section, .offline-banner { display: none !important; }
    .stat-card { border: 1px solid #ddd; box-shadow: none; }
    .stat-value { color: #333; }
    table.dataTable thead th { background: #f0f0f0 !important; color: #333 !important; }
    table.dataTable tbody td { color: #333; }
  }
</style>
</head>
<body>
<div class="container">

  <!-- Offline Warning -->
  <div class="offline-banner" id="offline-banner">⚠️ Einige Funktionen (Treemap, Tabellen-Export) benötigen eine Internetverbindung für CDN-Bibliotheken.</div>

  <!-- Header -->
  <div class="header">
    <h1>Folder Scan Dashboard <span>&bull; AI Analysis</span></h1>
    <div class="scan-path">$ResolvedPath</div>
    <div class="scan-meta">$scanInfoText</div>
  </div>

  <!-- Stats -->
  <div class="stats-grid" id="stats-grid">
    <div class="stat-card"><div class="stat-value" id="stat-files">$totalFiles</div><div class="stat-label">Dateien</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-size">$totalSizeDisplay</div><div class="stat-label">Gesamtgröße</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-folders">$uniqueFolders</div><div class="stat-label">Ordner</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-depth">$maxDepth</div><div class="stat-label">Max. Tiefe</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-conv">$convertibleCount</div><div class="stat-label">Konvertierbar</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-types">$($extGroups.Count)</div><div class="stat-label">Dateitypen</div></div>
  </div>

  $deltaHtml

  <!-- Charts -->
  <div class="chart-section">
    <div class="chart-card">
      <h2>📊 Verteilung nach Dateiendung</h2>
      <div class="chart-container">
        <div style="position:relative">
          <svg id="pie-ext" width="240" height="240" viewBox="-120 -120 240 240"></svg>
          <div class="pie-tooltip" id="tip-ext"></div>
        </div>
        <div class="legend" id="leg-ext"></div>
      </div>
    </div>
    <div class="chart-card">
      <h2>✅ Konvertierbarkeit</h2>
      <div class="chart-container">
        <div style="position:relative">
          <svg id="pie-conv" width="200" height="200" viewBox="-100 -100 200 200"></svg>
          <div class="pie-tooltip" id="tip-conv"></div>
        </div>
        <div class="legend" id="leg-conv"></div>
      </div>
    </div>
  </div>

  <!-- Structure Analysis -->
  <div class="structure-section">
    <div class="section-header">
      <h2 onclick="toggleTreemap()"><span class="collapse-arrow collapsed" id="tm-arrow">▼</span> 🧱 Struktur-Analyse (Treemap)</h2>
      <div class="toggle-group" id="tm-toggle-group" style="display:none">
        <button class="qf-btn active" id="btn-tm-count" onclick="drawTreemap('count')">Anzahl Dateien</button>
        <button class="qf-btn" id="btn-tm-size" onclick="drawTreemap('size')">Größe</button>
      </div>
    </div>
    <div class="treemap-body collapsed" id="treemap-body">
      <div id="treemap-container"></div>
      <div class="treemap-hint">Klicke auf einen Ordner, um die Dateiliste unten auf diesen Ordner zu filtern. Doppelklick zum Zoomen.</div>
    </div>
  </div>

  <!-- Top 10 Biggest Files -->
  <div class="top10-section">
    <h2>🏋️ Top 10 größte Dateien</h2>
    <div class="top10-list" id="top10-list"></div>
  </div>

  <!-- Folder Filter Indicator -->
  <div id="folder-filter-indicator" style="display:none;background:var(--bg-card);border-radius:var(--radius-card);padding:1rem 1.5rem;margin-bottom:1rem;border:1px solid var(--accent);font-size:0.9rem;"></div>

  <!-- File Table -->
  <div class="table-section">
    <h2>📋 Dateiliste</h2>
    <div class="folder-search">
      <input type="text" id="folder-search-input" placeholder="🔍 Ordner filtern (z.B. Mockup-Data)..." oninput="filterByPath(this.value)">
    </div>
    <div class="quick-filters" id="qf-container">
      <button class="qf-btn qf-btn-accent" onclick="filterConvertible()">✅ Nur Konvertierbare</button>
      <button class="qf-btn qf-btn-accent" onclick="clearFilter()">✖ Filter zurücksetzen</button>
    </div>
    <table id="file-table" class="display compact" style="width:100%">
      <thead>
        <tr>
          <th>Name</th>
          <th>Typ</th>
          <th>Pfad</th>
          <th>Größe</th>
          <th>Geändert</th>
          <th>Konv.</th>
          <th>Status</th>
          <th>Ordner</th>
        </tr>
      </thead>
    </table>
  </div>

  <!-- Info Box -->
  <div class="info-box">
    <h2>ℹ️ Konvertierbare Dateitypen (Ducling / Markdown)</h2>
    <div class="ext-list">$convertibleListHtml</div>
  </div>

</div>

<!-- Scripts -->
<script src="data.js"></script>
<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script src="https://cdn.datatables.net/2.3.7/js/dataTables.min.js"></script>
<script src="https://cdn.datatables.net/buttons/3.2.0/js/dataTables.buttons.min.js"></script>
<script src="https://cdn.datatables.net/buttons/3.2.0/js/buttons.html5.min.js"></script>
<script src="https://cdn.datatables.net/buttons/3.2.0/js/buttons.print.min.js"></script>
<script src="https://cdn.datatables.net/buttons/3.2.0/js/buttons.colVis.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<script>

// === DATA SAFETY CHECK ===
if (typeof allFiles === 'undefined') {
  console.error("[FolderScan] 'data.js' did not load. Using empty dataset.");
  window.allFiles = [];
}

// === CHART COLORS (Gemini Palette) ===
const CC = ['#4285F4','#9B72CB','#D96570','#F4B400','#0F9D58','#8AB4F8','#C58AF9','#F0929C','#FDD663','#6DD58C','#E8EAED'];

// === DATA ===
const SCAN_ROOT = "$($ResolvedPath -replace '\\', '\\\\')";
const scanRoot = SCAN_ROOT.replace(/\\\\/g, '\\');
const extData = [$extChartJson];
const convData = [
  { label: 'Konvertierbar', count: $convertibleCount, percent: Math.round(($convertibleCount / ($totalFiles || 1)) * 1000) / 10 },
  { label: 'Nicht konvertierbar', count: $nonConvertibleCount, percent: Math.round(($nonConvertibleCount / ($totalFiles || 1)) * 1000) / 10 }
];

// === FILTER STATE ===
let currentTreemapMetric = 'count';
const activeFilters = {
  ext: new Set(),
  conv: null, // null, 'Ja', 'Nein'
  folder: null // absolute folder path for treemap click filter
};

function applyFilters() {
  // Extension Filter (OR logic)
  if (activeFilters.ext.size > 0) {
    // Escape special chars for regex (e.g. . ++ etc)
    const regex = Array.from(activeFilters.ext).map(e => '^' + e.replace(/[.*+?^\x24{}()|[\x5D\x5C]/g, '\\$&') + '$').join('|');
    table.column(1).search(regex, true, false);
  } else {
    table.column(1).search('');
  }

  // Convertible Filter
  if (activeFilters.conv) {
    table.column(5).search(activeFilters.conv);
  } else {
    table.column(5).search('');
  }

  // Folder Filter (from treemap click)
  if (activeFilters.folder) {
    table.column(2).search('^' + escRx(activeFilters.folder), true, false);
  } else {
    table.column(2).search('');
  }

  table.draw();
  updateVisuals();
  updateFolderIndicator();
}

function getFilteredFiles() {
  if (!activeFilters.folder) return allFiles;
  var folder = activeFilters.folder;
  return allFiles.filter(function(f) {
    var dir = f.dir ? f.dir.replace(/\\\\/g, '\\') : '';
    return dir === folder || dir.indexOf(folder + '\\') === 0;
  });
}

function updateVisuals() {
  var filtered = getFilteredFiles();

  // Update quick-filter button counts
  var extCounts = {};
  var convCount = 0;
  filtered.forEach(function(f) {
    var ext = f.ext || '(none)';
    extCounts[ext] = (extCounts[ext] || 0) + 1;
    if (f.convertible) convCount++;
  });

  document.querySelectorAll('.qf-btn[data-ext]').forEach(function(b) {
    var ext = b.dataset.ext;
    var cnt = extCounts[ext] || 0;
    b.textContent = ext + ' (' + cnt + ')';
    if (activeFilters.ext.has(ext)) b.classList.add('active-filter');
    else b.classList.remove('active-filter');
  });

  // Conv button
  var btnConv = document.getElementById('btn-conv-yes');
  if (btnConv) {
    btnConv.textContent = '\u2705 Nur Konvertierbare (' + convCount + ')';
    btnConv.classList.toggle('active-filter', activeFilters.conv === 'Ja');
  }

  // Update stat cards
  updateStatCards(filtered);

  // Update Top-10 biggest files
  try { buildTop10(); } catch(e) {}

  // Update pie charts with filtered data (redraws SVG)
  updatePieCharts(filtered);

  // Pie Segments highlight (AFTER redraw so classes are not wiped)
  document.querySelectorAll('.chart-segment[data-ext]').forEach(function(p) {
    if (activeFilters.ext.has(p.dataset.ext)) p.classList.add('active');
    else p.classList.remove('active');
  });
  document.querySelectorAll('.chart-segment[data-conv]').forEach(function(p) {
    if (activeFilters.conv === p.dataset.conv) p.classList.add('active');
    else p.classList.remove('active');
  });
}

function updateStatCards(filtered) {
  var totalBytes = 0;
  var dirs = new Set();
  var convCnt = 0;
  var exts = new Set();
  filtered.forEach(function(f) {
    totalBytes += f.sizeBytes;
    var dir = f.dir ? f.dir.replace(/\\\\/g, '\\') : '';
    if (dir) dirs.add(dir);
    if (f.convertible) convCnt++;
    if (f.ext) exts.add(f.ext);
  });

  var el;
  el = document.getElementById('stat-files');
  if (el) el.textContent = filtered.length;
  el = document.getElementById('stat-size');
  if (el) el.textContent = fmtSize(totalBytes);
  el = document.getElementById('stat-folders');
  if (el) el.textContent = dirs.size;
  el = document.getElementById('stat-conv');
  if (el) el.textContent = convCnt;
  el = document.getElementById('stat-types');
  if (el) el.textContent = exts.size;
}

function updatePieCharts(filtered) {
  // Build ext distribution from filtered files
  var extMap = {};
  filtered.forEach(function(f) {
    var ext = f.ext || '(none)';
    if (!extMap[ext]) extMap[ext] = { label: ext, count: 0, sizeBytes: 0 };
    extMap[ext].count++;
    extMap[ext].sizeBytes += f.sizeBytes;
  });
  var dynExtData = Object.values(extMap).sort(function(a, b) { return b.count - a.count; });
  var total = filtered.length || 1;
  dynExtData.forEach(function(d) {
    d.percent = Math.round((d.count / total) * 1000) / 10;
    d.sizeMB = (d.sizeBytes / 1048576).toFixed(2);
  });

  // Build conv distribution
  var cYes = 0;
  filtered.forEach(function(f) { if (f.convertible) cYes++; });
  var cNo = filtered.length - cYes;
  var dynConvData = [
    { label: 'Konvertierbar', count: cYes, percent: Math.round((cYes / total) * 1000) / 10 },
    { label: 'Nicht konvertierbar', count: cNo, percent: Math.round((cNo / total) * 1000) / 10 }
  ];

  drawPie('pie-ext','tip-ext','leg-ext', dynExtData, CC, 100, 'ext');
  drawPie('pie-conv','tip-conv','leg-conv', dynConvData, ['#4ade80','#475569'], 80, 'conv');
}

function toggleExt(ext) {
  if (activeFilters.ext.has(ext)) activeFilters.ext.delete(ext);
  else activeFilters.ext.add(ext);
  applyFilters();
}

function toggleConv(val) {
  if (activeFilters.conv === val) activeFilters.conv = null;
  else activeFilters.conv = val;
  applyFilters();
}

function clearAllFilters() {
  activeFilters.ext.clear();
  activeFilters.conv = null;
  var hadFolder = activeFilters.folder;
  activeFilters.folder = null;
  var searchInput = document.getElementById('folder-search-input');
  if (searchInput) searchInput.value = '';
  applyFilters();
  if (hadFolder) {
    try { drawTreemap(currentTreemapMetric); } catch(e) { console.error(e); }
  }
}

function filterByFolder(absPath) {
  // Clicking root = show all; clicking same folder again = clear filter
  if (!absPath || absPath === scanRoot || activeFilters.folder === absPath) {
    activeFilters.folder = null;
  } else {
    activeFilters.folder = absPath;
  }
  applyFilters();
  // Redraw treemap to update highlighting
  try { drawTreemap(currentTreemapMetric); } catch(e) { console.error(e); }
}

function updateFolderIndicator() {
  const el = document.getElementById('folder-filter-indicator');
  if (!el) return;
  if (activeFilters.folder) {
    const display = activeFilters.folder.replace(/\\\\/g, '\\');
    var html = '<span style="color:var(--accent)">\ud83d\udcc2 Ordnerfilter:</span> <strong>' + display + '</strong> ';
    html += '<button class="qf-btn" style="margin-left:0.5rem;padding:0.2rem 0.8rem;font-size:0.8rem" onclick="folderUp()">\u2b06 Ebene h\u00f6her</button>';
    html += '<button class="qf-btn" style="margin-left:0.5rem;padding:0.2rem 0.8rem;font-size:0.8rem" onclick="clearAllFilters()">\u2716 Aufheben</button>';
    el.innerHTML = html;
    el.style.display = 'block';
  } else {
    el.innerHTML = '';
    el.style.display = 'none';
  }
}

function folderUp() {
  if (!activeFilters.folder) return;
  var idx = activeFilters.folder.lastIndexOf('\\');
  if (idx <= 0) { clearAllFilters(); return; }
  var parent = activeFilters.folder.substring(0, idx);
  if (parent.length < scanRoot.length || parent === scanRoot) {
    clearAllFilters();
  } else {
    filterByFolder(parent);
  }
}

var treemapDrawn = false;
function toggleTreemap() {
  var body = document.getElementById('treemap-body');
  var arrow = document.getElementById('tm-arrow');
  var toggleGrp = document.getElementById('tm-toggle-group');
  if (!body) return;
  var isCollapsed = body.classList.toggle('collapsed');
  if (arrow) arrow.classList.toggle('collapsed', isCollapsed);
  if (toggleGrp) toggleGrp.style.display = isCollapsed ? 'none' : '';
  if (!isCollapsed && !treemapDrawn) {
    treemapDrawn = true;
    setTimeout(function() { try { drawTreemap('count'); } catch(e) { console.error(e); } }, 100);
  }
}

// === PIE CHART ===
function drawPie(svgId, tipId, legId, data, colors, r, type) {
  const svg = document.getElementById(svgId);
  const tip = document.getElementById(tipId);
  const leg = document.getElementById(legId);
  const total = data.reduce((s, d) => s + d.count, 0);
  const ri = r * 0.55;
  let a = -Math.PI / 2;
  svg.innerHTML = ''; leg.innerHTML = '';
  data.forEach((d, i) => {
    const ang = (d.count / total) * 2 * Math.PI;
    const x1 = Math.cos(a)*r, y1 = Math.sin(a)*r, x1i = Math.cos(a)*ri, y1i = Math.sin(a)*ri;
    a += ang;
    const x2 = Math.cos(a)*r, y2 = Math.sin(a)*r, x2i = Math.cos(a)*ri, y2i = Math.sin(a)*ri;
    const la = ang > Math.PI ? 1 : 0, c = colors[i % colors.length];
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    p.setAttribute('d', 'M'+x1i+' '+y1i+'L'+x1+' '+y1+'A'+r+' '+r+' 0 '+la+' 1 '+x2+' '+y2+'L'+x2i+' '+y2i+'A'+ri+' '+ri+' 0 '+la+' 0 '+x1i+' '+y1i+'Z');
    p.setAttribute('fill', c);
    p.setAttribute('class', 'chart-segment');
    if (type === 'ext') p.setAttribute('data-ext', d.label);
    if (type === 'conv') p.setAttribute('data-conv', d.label === 'Konvertierbar' ? 'Ja' : 'Nein');
    
    p.addEventListener('click', () => {
        if (type === 'ext') toggleExt(d.label);
        if (type === 'conv') toggleConv(d.label === 'Konvertierbar' ? 'Ja' : 'Nein');
    });

    p.addEventListener('mouseenter', () => {
      tip.style.opacity = '1';
      tip.innerHTML = '<strong>'+d.label+'</strong><br>'+d.count+' ('+d.percent+'%)'+(d.sizeMB!==undefined?'<br>'+d.sizeMB+' MB':'');
    });
    p.addEventListener('mousemove', e => {
      const rect = svg.closest('.chart-container').getBoundingClientRect();
      tip.style.left = (e.clientX-rect.left+12)+'px'; tip.style.top = (e.clientY-rect.top-8)+'px';
    });
    p.addEventListener('mouseleave', () => { tip.style.opacity='0'; });
    svg.appendChild(p);
    
    // Legend click
    const li = document.createElement('div');
    li.className = 'legend-item';
    li.style.cursor = 'pointer';
    li.innerHTML = '<span class="legend-dot" style="background:'+c+'"></span><span>'+d.label+'</span><span class="legend-count">'+d.count+'</span>';
    li.onclick = () => {
        if (type === 'ext') toggleExt(d.label);
        if (type === 'conv') toggleConv(d.label === 'Konvertierbar' ? 'Ja' : 'Nein');
    };
    leg.appendChild(li);
  });
}

drawPie('pie-ext','tip-ext','leg-ext', extData, CC, 100, 'ext');
drawPie('pie-conv','tip-conv','leg-conv', convData, ['#4ade80','#475569'], 80, 'conv');

// === QUICK FILTER BUTTONS ===
(function(){
  const c = document.getElementById('qf-container');
  c.innerHTML = ''; // Clear static obsolete buttons
  
  // Reset Button
  const btnReset = document.createElement('button');
  btnReset.className = 'qf-btn qf-btn-accent';
  btnReset.textContent = '✖ Filter zurücksetzen';
  btnReset.onclick = clearAllFilters;
  c.appendChild(btnReset);

  // Add Conv Button
  const btnConv = document.createElement('button');
  btnConv.className = 'qf-btn qf-btn-accent'; 
  btnConv.id = 'btn-conv-yes';
  btnConv.textContent = '✅ Nur Konvertierbare';
  btnConv.onclick = () => toggleConv('Ja');
  c.appendChild(btnConv);

  const topExts = extData.filter(d => d.label !== 'Other').slice(0, 10);
  topExts.forEach(d => {
    const b = document.createElement('button');
    b.className = 'qf-btn'; 
    b.dataset.ext = d.label;
    b.textContent = d.label + ' (' + d.count + ')';
    b.addEventListener('click', () => toggleExt(d.label));
    c.appendChild(b);
  });
})();

// === FORMAT HELPERS ===
function fmtSize(b) {
  if (b >= 1073741824) return (b/1073741824).toFixed(2)+' GB';
  if (b >= 1048576) return (b/1048576).toFixed(2)+' MB';
  if (b >= 1024) return (b/1024).toFixed(1)+' KB';
  return b+' B';
}

function escH(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

function escRx(s) { return s.replace(/[-\/\x5C^\x24*+?.()|[\x5D]/g, '\x5C\x24&'); }

function copyFolder(el) {
  var path = el.getAttribute('data-folder');
  if (!path) return;
  var orig = el.textContent;
  try {
    var ta = document.createElement('textarea');
    ta.value = path;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    el.textContent = '\u2705';
    setTimeout(function() { el.textContent = orig; }, 1200);
  } catch(e) {
    window.prompt('Pfad kopieren:', path);
  }
}

// === DATATABLES ===
$deltaJs

const table = new DataTable('#file-table', {
  data: allFiles,
  deferRender: true,
  pageLength: 25,
  lengthMenu: [[15, 25, 50, 100, -1], [15, 25, 50, 100, 'Alle']],
  order: [[0, 'asc']],
  layout: {
    topStart: 'buttons',
    topEnd: 'search',
    bottomStart: ['pageLength', 'info'],
    bottomEnd: 'paging'
  },
  buttons: [
    'copy',
    { extend: 'colvis', text: 'Spalten' },
    { extend: 'csvHtml5', title: 'FolderScan', exportOptions: { columns: ':visible:not(:last-child)' } },
    { extend: 'excelHtml5', title: 'FolderScan', exportOptions: { columns: ':visible:not(:last-child)' } },
    { extend: 'print', title: 'Folder Scan', exportOptions: { columns: ':visible:not(:last-child)' } }
  ],
  columns: [
    {
      data: 'name',
      title: 'Name',
      render: function(data) { return '<span class="cell-name" title="'+escH(data)+'">'+escH(data)+'</span>'; }
    },
    {
      data: 'ext',
      title: 'Typ',
      render: function(data) { return '<span class="cell-ext">'+escH(data)+'</span>'; }
    },
    {
      data: 'dir',
      title: 'Verzeichnis',
      render: function(data) {
        // Fix double backslashes for display
        const cleanPath = data.replace(/\\\\/g, '\\');
        return '<div class="path-cell-inner" title="'+escH(cleanPath)+'">'+escH(cleanPath)+'</div>';
      }
    },
    {
      data: 'fullPath',
      title: 'Vollst. Pfad',
      visible: false,
      render: function(data) {
        const cleanPath = data.replace(/\\\\/g, '\\');
        return '<div class="path-cell-inner" title="'+escH(cleanPath)+'">'+escH(cleanPath)+'</div>';
      }
    },
    {
      data: 'sizeBytes',
      title: 'Größe',
      render: function(data) {
        var cls = 'cell-size';
        if (data >= 524288000) cls += ' size-danger';
        else if (data >= 52428800) cls += ' size-warn';
        return '<span class="'+cls+'">'+fmtSize(data)+'</span>';
      },
      type: 'num'
    },
    { data: 'sizeKB', title: 'KB', visible: false, render: DataTable.render.number('.', ',', 2, '', ' KB') },
    { data: 'sizeMB', title: 'MB', visible: false, render: DataTable.render.number('.', ',', 2, '', ' MB') },
    
    { data: 'created', title: 'Erstellt', visible: false },
    { data: 'modified', title: 'Geändert' },
    
    { 
      data: 'readOnly', title: '🔒', visible: false,
      render: function(d) { return d ? '🔒' : ''; }
    },
    {
      data: 'convertible',
      title: 'Konv.',
      render: function(data) {
        return data ? '<span class="badge badge-yes">Ja</span>' : '<span class="badge badge-no">Nein</span>';
      }
    },
    {
      data: 'status',
      title: 'Status',
      render: function(data) {
        if (data === 'new') return '<span class="badge badge-new">🟢 Neu</span>';
        if (data === 'modified') return '<span class="badge badge-modified">🟡 Geändert</span>';
        return '';
      },
      visible: HAS_DELTA
    },
    {
      data: 'dir',
      title: '\ud83d\udcc2',
      orderable: false,
      searchable: false,
      className: 'no-export',
      render: function(data, type, row, meta) {
        const clean = data.replace(/\\\\/g, '\\');
        return '<span class="folder-link" title="'+escH(clean)+'" data-folder="'+escH(clean)+'" onclick="copyFolder(this)">\ud83d\udcc2</span>';
      }
    }
  ],
  language: {
    search: '🔍 Suchen:',
    lengthMenu: '_MENU_ pro Seite',
    info: '_START_ – _END_ von _TOTAL_ Dateien',
    infoFiltered: '(gefiltert von _MAX_)',
    infoEmpty: 'Keine Dateien',
    zeroRecords: 'Keine Treffer',
    paginate: { first: '⏮', previous: '◀', next: '▶', last: '⏭' },
    buttons: { copy: '📋 Copy', csv: '📄 CSV', excel: '📊 Excel', print: '🖨️ Print' }
  }
});



// === TREEMAP ===

function buildTreemapData(metric) {
  const folders = new Map();
  const rootLabel = scanRoot.split('\\').pop() || scanRoot;

  // Root node = scanned folder
  folders.set(scanRoot, { id: scanRoot, label: rootLabel, parent: '', size: 0, count: 0 });

  // Collect all unique directories from files
  allFiles.forEach(function(f) {
    if (!f.dir) return;
    const dir = f.dir.replace(/\\\\/g, '\\');
    if (!folders.has(dir)) {
      // Build chain from dir up to scanRoot
      let current = dir;
      const chain = [];
      while (current && current !== scanRoot && current.length > scanRoot.length) {
        chain.push(current);
        const idx = current.lastIndexOf('\\');
        if (idx <= 0) break;
        current = current.substring(0, idx);
      }
      // Add folders from closest-to-root to deepest
      chain.reverse().forEach(function(p) {
        if (!folders.has(p)) {
          const parentPath = p.substring(0, p.lastIndexOf('\\')) || scanRoot;
          const lbl = p.substring(p.lastIndexOf('\\') + 1);
          folders.set(p, { id: p, label: lbl || p, parent: parentPath, size: 0, count: 0 });
        }
      });
    }
  });

  // Accumulate file sizes/counts into their direct folder
  allFiles.forEach(function(f) {
    const dir = f.dir ? f.dir.replace(/\\\\/g, '\\') : scanRoot;
    const entry = folders.get(dir);
    if (entry) {
      entry.size += f.sizeBytes;
      entry.count++;
    }
  });

  // Build arrays for Plotly
  const ids = [], labels = [], parents = [], values = [], text = [];
  folders.forEach(function(f) {
    ids.push(f.id);
    labels.push(f.label);
    parents.push(f.parent);
    values.push(metric === 'size' ? f.size : f.count);
    text.push(f.count + ' Dateien<br>' + fmtSize(f.size));
  });

  return { ids: ids, labels: labels, parents: parents, values: values, text: text };
}

function drawTreemap(metric) {
  const container = document.getElementById('treemap-container');
  if (!container) return;

  if (typeof Plotly === 'undefined') {
    container.innerHTML = '<div style="padding:2rem;text-align:center;opacity:0.6">Treemap needs internet (Plotly CDN).</div>';
    return;
  }

  currentTreemapMetric = metric;
  document.getElementById('btn-tm-count').classList.toggle('active', metric === 'count');
  document.getElementById('btn-tm-size').classList.toggle('active', metric === 'size');

  const data = buildTreemapData(metric);
  const bodyStyle = getComputedStyle(document.body);
  const textColor = bodyStyle.getPropertyValue('--text-primary').trim();
  const fontFamily = bodyStyle.getPropertyValue('font-family');
  const bgCard = bodyStyle.getPropertyValue('--bg-card').trim();

  // Build marker colors: highlight active folder + children
  var markerColors = data.values.slice();
  var lineWidths = data.ids.map(function() { return 1; });
  var lineColors = data.ids.map(function() { return bgCard; });
  if (activeFilters.folder) {
    var af = activeFilters.folder;
    markerColors = data.ids.map(function(id, i) {
      if (id === af || id.indexOf(af + '\\') === 0) return data.values[i];
      return 0;
    });
    lineWidths = data.ids.map(function(id) { return id === af ? 3 : 1; });
    lineColors = data.ids.map(function(id) { return id === af ? '#4285F4' : bgCard; });
  }

  const trace = {
    type: 'treemap',
    ids: data.ids,
    labels: data.labels,
    parents: data.parents,
    values: data.values,
    text: data.text,
    textinfo: 'label+value',
    hoverinfo: 'label+text+value',
    branchvalues: 'remainder',
    marker: {
      colors: markerColors,
      colorscale: metric === 'size' ? 'Blues' : 'Greens',
      showscale: false,
      line: { width: lineWidths, color: lineColors }
    },
    pathbar: { visible: true, thickness: 36, textfont: { size: 14 } },
    tiling: { packing: 'squarify' }
  };

  const layout = {
    margin: { t: 44, l: 0, r: 0, b: 0 },
    paper_bgcolor: 'transparent',
    font: { family: fontFamily, color: textColor }
  };

  Plotly.newPlot('treemap-container', [trace], layout).then(function() {
    // Click handler: filter DataTable by clicked folder
    container.on('plotly_click', function(eventData) {
      if (eventData && eventData.points && eventData.points.length > 0) {
        var clickedId = eventData.points[0].id;
        if (clickedId) {
          filterByFolder(clickedId);
        }
      }
    });
  });
}

// === TOP 10 BIGGEST FILES ===
function buildTop10() {
  var el = document.getElementById('top10-list');
  if (!el) return;
  var filtered = getFilteredFiles();
  var sorted = filtered.slice().sort(function(a, b) { return b.sizeBytes - a.sizeBytes; });
  var top = sorted.slice(0, 10);
  if (top.length === 0) { el.innerHTML = '<div style="color:var(--text-muted);font-size:0.85rem">Keine Dateien.</div>'; return; }
  var maxSize = top[0].sizeBytes || 1;
  var html = '';
  top.forEach(function(f) {
    var pct = Math.round((f.sizeBytes / maxSize) * 100);
    var cls = f.sizeBytes >= 524288000 ? 'size-danger' : (f.sizeBytes >= 52428800 ? 'size-warn' : '');
    html += '<div class="top10-item"><span class="top10-name" title="' + escH(f.name) + '">' + escH(f.name) + '</span><span class="top10-size ' + cls + '">' + fmtSize(f.sizeBytes) + '</span></div>';
  });
  el.innerHTML = html;
}

// === FOLDER PATH SEARCH ===
function filterByPath(query) {
  if (!query || query.trim() === '') {
    table.column(2).search('');
  } else {
    table.column(2).search(query);
  }
  table.draw();
}

// === OFFLINE DETECTION ===
function checkOffline() {
  var banner = document.getElementById('offline-banner');
  if (!banner) return;
  if (!navigator.onLine) {
    banner.style.display = 'block';
  } else {
    // Check if jQuery loaded (CDN dependency)
    if (typeof jQuery === 'undefined' || typeof Plotly === 'undefined') {
      banner.style.display = 'block';
    }
  }
}

// Initial draw
jQuery(document).ready(function() {
  try {
    if (typeof updateVisuals === 'function') updateVisuals();
  } catch (e) { console.error('UpdateVisuals Error:', e); }

  try { buildTop10(); } catch(e) { console.error('Top10 Error:', e); }
  try { checkOffline(); } catch(e) {}
});

</script>
</body>
</html>
"@

  # ═══════════════════════════════════════════════════════════════
  # WRITE HTML & OPEN
  # ═══════════════════════════════════════════════════════════════
  $htmlPath = Join-Path $OutputDir "Dashboard.html"
  [System.IO.File]::WriteAllText($htmlPath, $htmlContent, [System.Text.UTF8Encoding]::new($false))
  Write-Host "[Invoke-FolderScan] Dashboard: $htmlPath" -ForegroundColor Green

  Start-Process $htmlPath
  Write-Host "[Invoke-FolderScan] Dashboard opened in browser." -ForegroundColor Cyan

  Write-Host "`n[Invoke-FolderScan] Done! $totalFiles files. Results: $OutputDir" -ForegroundColor Green
  Write-Host "[Invoke-FolderScan] Tip: `$result | Where-Object { `$_.IsConvertible } | Out-GridView" -ForegroundColor DarkGray

  return $FileList
}
