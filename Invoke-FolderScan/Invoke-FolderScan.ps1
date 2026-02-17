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

.PARAMETER AllowCrossPathDelta
    If specified, delta comparison is allowed even when the previous scan was created for a different root path.
    By default, delta is skipped across different scan roots to avoid false comparisons.

.PARAMETER StagingPath
    Path to the staging folder. Used for pipeline status detection (kopiert/veredelt).
    Files found in this folder are marked as 'kopiert' in the dashboard.

.PARAMETER ErgebnisPath
    Path to the result folder (Docling output). Files found here are marked as 'veredelt'.

.PARAMETER PassThru
    If specified, returns the file list as PowerShell objects for further processing.
    Without this switch, no objects are written to the pipeline.

.PARAMETER NoGui
    If specified, skips HTML dashboard generation and browser launch.
    Useful for scripted/automated runs where only the CSV is needed.

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

.EXAMPLE
    Invoke-FolderScan -Path "D:\Docs" -Recurse -StagingPath "D:\Staging"
    # Scan with pipeline status detection

.EXAMPLE
    $files = Invoke-FolderScan -Path "D:\Docs" -Recurse -PassThru -NoGui
    $files | Where-Object { $_.IsConvertible } | Out-GridView
    # Scan without dashboard, capture results for scripting
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
    [switch]$LoadLatest,

    [Parameter(ParameterSetName = 'Scan')]
    [switch]$AllowCrossPathDelta,

    [string]$StagingPath,

    [string]$ErgebnisPath,

    [switch]$PassThru,

    [switch]$NoGui
  )

  # ═══════════════════════════════════════════════════════════════
  # CONFIG
  # ═══════════════════════════════════════════════════════════════
  $ConvertibleExtensions = @(
    '.pdf', '.docx', '.xlsx', '.pptx',
    '.html', '.htm', '.md', '.csv', '.xml', '.json',
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.webp',
    '.asciidoc', '.adoc', '.vtt',
    '.mp3', '.wav', '.flac', '.m4a', '.ogg', '.wma'
  )

  # ═══════════════════════════════════════════════════════════════
  # ENSURE OUTPUT DIR
  # ═══════════════════════════════════════════════════════════════
  if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
  }
  $OutputDir = (Resolve-Path $OutputDir).Path

  $csvPath = Join-Path $OutputDir "FullScan.csv"
  $metaPath = Join-Path $OutputDir "FullScan.meta.json"
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
    $resolvedFromMeta = $null
    if (Test-Path $metaPath) {
      try {
        $scanMeta = Get-Content -Path $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($scanMeta.ScanRoot) {
          $resolvedFromMeta = [string]$scanMeta.ScanRoot
        }
      }
      catch {}
    }

    $ResolvedPath = if ($resolvedFromMeta) {
      $resolvedFromMeta
    }
    elseif ($FileList.Count -gt 0) {
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

      $previousScanRoot = $null
      $previousRootSource = $null
      if (Test-Path $metaPath) {
        try {
          $previousMeta = Get-Content -Path $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
          if ($previousMeta.ScanRoot) {
            $previousScanRoot = [string]$previousMeta.ScanRoot
            $previousRootSource = 'meta'
          }
        }
        catch {}
      }
      if (-not $previousScanRoot -and $previousFiles.Count -gt 0) {
        $previousScanRoot = [string]$previousFiles[0].DirectoryName
        $previousRootSource = 'csv'
      }

      $shouldComputeDelta = $true
      if (-not $AllowCrossPathDelta -and $previousScanRoot) {
        $currNorm = $ResolvedPath.TrimEnd('\').ToLowerInvariant()
        $prevNorm = $previousScanRoot.TrimEnd('\').ToLowerInvariant()
        $rootsCompatible = $false

        if ($previousRootSource -eq 'meta') {
          $rootsCompatible = ($currNorm -eq $prevNorm)
        }
        else {
          # Fallback (legacy CSV without metadata): treat parent/child as compatible to reduce false negatives.
          $rootsCompatible = (
            $currNorm -eq $prevNorm -or
            $prevNorm.StartsWith($currNorm + '\') -or
            $currNorm.StartsWith($prevNorm + '\')
          )
        }

        if (-not $rootsCompatible) {
          $shouldComputeDelta = $false
          Write-Warning "[Invoke-FolderScan] Delta skipped: previous scan root '$previousScanRoot' differs from current '$ResolvedPath'. Use -AllowCrossPathDelta to force comparison."
        }
      }

      if ($shouldComputeDelta) {
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
          NewPaths      = @($newFiles | Select-Object -ExpandProperty FullPath)
          DeletedPaths  = @($deletedPaths | Select-Object -ExpandProperty FullPath)
          ModifiedPaths = @($modifiedFiles | Select-Object -ExpandProperty FullPath)
          PreviousDate  = (Get-Item $csvPath).LastWriteTime.ToString('yyyy-MM-dd HH:mm')
        }

        Write-Host "[Invoke-FolderScan] Delta: $($deltaInfo.NewCount) new, $($deltaInfo.DeletedCount) deleted, $($deltaInfo.ModifiedCount) modified" -ForegroundColor Cyan
      }

      # Rename old CSV
      $oldDate = (Get-Item $csvPath).LastWriteTime.ToString('yyyy-MM-dd_HHmmss')
      $archivePath = Join-Path $OutputDir "FullScan_$oldDate.csv"
      Move-Item -Path $csvPath -Destination $archivePath -Force
      Write-Host "[Invoke-FolderScan] Previous scan archived: $archivePath" -ForegroundColor DarkGray

      if (Test-Path $metaPath) {
        $archiveMetaPath = Join-Path $OutputDir "FullScan_$oldDate.meta.json"
        Move-Item -Path $metaPath -Destination $archiveMetaPath -Force
        Write-Host "[Invoke-FolderScan] Previous metadata archived: $archiveMetaPath" -ForegroundColor DarkGray
      }
    }

    # ─── EXPORT CSV ───
    $FileList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "[Invoke-FolderScan] CSV exported: $csvPath" -ForegroundColor Green

    $scanMeta = [PSCustomObject]@{
      ScanRoot     = $ResolvedPath
      ScannedAt    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
      Recursive    = [bool]$Recurse
      FileCount    = $FileList.Count
      StagingPath  = if ($StagingPath) { $StagingPath } else { $null }
      ErgebnisPath = if ($ErgebnisPath) { $ErgebnisPath } else { $null }
    }
    $scanMeta | ConvertTo-Json | Set-Content -Path $metaPath -Encoding UTF8
    Write-Host "[Invoke-FolderScan] Metadata exported: $metaPath" -ForegroundColor Green
  }

  # ═══════════════════════════════════════════════════════════════
  # PIPELINE STATUS (Staging / Ergebnis check)
  # ═══════════════════════════════════════════════════════════════
  # If LoadLatest, try to restore StagingPath/ErgebnisPath from metadata
  if ($LoadLatest -and -not $StagingPath -and -not $ErgebnisPath) {
    if (Test-Path $metaPath) {
      try {
        $savedMeta = Get-Content -Path $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($savedMeta.StagingPath) { $StagingPath = $savedMeta.StagingPath }
        if ($savedMeta.ErgebnisPath) { $ErgebnisPath = $savedMeta.ErgebnisPath }
      }
      catch {}
    }
  }

  $hasPipeline = $false
  $pipelineStats = @{ offen = 0; kopiert = 0; veredelt = 0 }

  if ($StagingPath -or $ErgebnisPath) {
    $hasPipeline = $true
    $scanRootNorm = $ResolvedPath.TrimEnd('\')

    if ($StagingPath) { Write-Host "[Invoke-FolderScan] Staging:  $StagingPath" -ForegroundColor Cyan }
    if ($ErgebnisPath) { Write-Host "[Invoke-FolderScan] Ergebnis: $ErgebnisPath" -ForegroundColor Cyan }

    $pipelineStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($f in $FileList) {
      $relativePath = ''
      if ($f.FullPath.StartsWith($scanRootNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
        $relativePath = $f.FullPath.Substring($scanRootNorm.Length).TrimStart('\')
      }
      else {
        $relativePath = $f.Name
      }

      $status = 'offen'

      # Check Ergebnis first (higher priority: veredelt > kopiert > offen)
      if ($ErgebnisPath -and $relativePath) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($relativePath)
        $relDir = [System.IO.Path]::GetDirectoryName($relativePath)
        $mdPath = Join-Path $ErgebnisPath (Join-Path $relDir "$baseName.md")
        if (Test-Path $mdPath) {
          $status = 'veredelt'
        }
      }

      # Check Staging (only if not already veredelt)
      if ($status -eq 'offen' -and $StagingPath -and $relativePath) {
        $stagingFile = Join-Path $StagingPath $relativePath
        if (Test-Path $stagingFile) {
          $status = 'kopiert'
        }
      }

      $f | Add-Member -NotePropertyName 'PipelineStatus' -NotePropertyValue $status -Force
      $pipelineStats[$status]++
    }

    $pipelineStopwatch.Stop()
    Write-Host "[Invoke-FolderScan] Pipeline status: $($pipelineStats.offen) offen, $($pipelineStats.kopiert) kopiert, $($pipelineStats.veredelt) veredelt ($([math]::Round($pipelineStopwatch.Elapsed.TotalSeconds, 2))s)" -ForegroundColor Green
  }
  else {
    # No pipeline paths: set empty status
    foreach ($f in $FileList) {
      $f | Add-Member -NotePropertyName 'PipelineStatus' -NotePropertyValue '' -Force
    }
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
  $totalFolders = if (-not $LoadLatest) {
    (Get-ChildItem -Path $ResolvedPath -Directory -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
  }
  else { $uniqueFolders }
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
  # GENERATE DASHBOARD (skip with -NoGui)
  # ═══════════════════════════════════════════════════════════════
  if ($NoGui) {
    Write-Host "[Invoke-FolderScan] -NoGui: Dashboard generation skipped." -ForegroundColor DarkGray
    Write-Host "[Invoke-FolderScan] CSV: $csvPath" -ForegroundColor Green
    Write-Host "`n[Invoke-FolderScan] Done! $($FileList.Count) files. Results: $OutputDir" -ForegroundColor Green
    if ($PassThru) { return $FileList } else { return }
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
    $n = ($f.Name -replace '\\', '\\\\' -replace '"', '\"')
    $fp = ($f.FullPath -replace '\\', '\\\\' -replace '"', '\"')
    $dn = ($f.DirectoryName -replace '\\', '\\\\' -replace '"', '\"')

    # Delta status
    $status = 'unchanged'
    if ($deltaNewSet.ContainsKey($f.FullPath)) { $status = 'new' }
    elseif ($deltaModSet.ContainsKey($f.FullPath)) { $status = 'modified' }

    $comma = if ($i -lt $FileList.Count - 1) { ',' } else { '' }
    $pStatus = if ($f.PipelineStatus) { $f.PipelineStatus } else { '' }
    $sw.WriteLine("{`"name`":`"$n`",`"ext`":`"$($f.Extension)`",`"fullPath`":`"$fp`",`"dir`":`"$dn`",`"sizeBytes`":$($f.SizeBytes),`"sizeKB`":$($f.SizeKB),`"sizeMB`":$($f.SizeMB),`"created`":`"$($f.CreationTime)`",`"modified`":`"$($f.LastWriteTime)`",`"readOnly`":$($f.IsReadOnly.ToString().ToLower()),`"convertible`":$($f.IsConvertible.ToString().ToLower()),`"status`":`"$status`",`"pipeline`":`"$pStatus`"}$comma")
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
    --font-head: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
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
    --font-head: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
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

  # Docling time estimation (read from previous runs if available)
  $doclingAvgPerFile = 0
  if ($ErgebnisPath) {
    $doclingRunsPath = Join-Path $ErgebnisPath "docling_runs.json"
    if (Test-Path $doclingRunsPath) {
      try {
        $runsData = Get-Content -Path $doclingRunsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($runsData.runs -and $runsData.runs.Count -gt 0) {
          $totalAvg = ($runsData.runs | Where-Object { $_.avgPerFile -gt 0 } | Measure-Object -Property avgPerFile -Average).Average
          if ($totalAvg) { $doclingAvgPerFile = [math]::Round($totalAvg, 2) }
        }
      }
      catch {}
    }
  }

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
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet" onerror="this.remove()">
<!-- DataTables CSS -->
<link rel="stylesheet" href="https://cdn.datatables.net/2.3.7/css/dataTables.dataTables.min.css">
<link rel="stylesheet" href="https://cdn.datatables.net/buttons/3.2.0/css/buttons.dataTables.min.css">
<style>
  :root {
$themeCss
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    background: var(--bg-primary);
    color: var(--text-primary);
    line-height: 1.6;
    min-height: 100vh;
  }
  .container { max-width: 98%; margin: 0 auto; padding: 1.5rem 2rem; }
  .path-cell-inner { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 0; min-width: 100%; }
  /* Active states for filters */
  .qf-btn.active-filter {
    background: var(--accent) !important; color: #fff !important; border-color: var(--accent) !important;
  }
  .chart-segment { transition: opacity 0.25s, transform 0.2s, filter 0.2s; cursor: pointer; }
  .chart-segment:hover { filter: brightness(1.15); }
  .chart-segment.active { stroke: var(--accent); stroke-width: 3px; transform: scale(1.04); filter: brightness(1.1); opacity: 1 !important; }
  .chart-segment.dimmed { opacity: 0.3; }

  /* Unified filter bar */
  .filter-bar {
    background: var(--bg-card); border: 1px solid var(--border-glass); border-radius: var(--radius-card);
    padding: 1rem 1.5rem; margin-bottom: 1.5rem;
  }
  .filter-bar-row { display: flex; gap: 0.5rem; flex-wrap: wrap; align-items: center; }
  .filter-bar-sep { width: 1px; height: 20px; background: var(--border-glass); flex-shrink: 0; }

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
    font-family: 'JetBrains Mono', 'Cascadia Code', 'Consolas', 'Courier New', monospace; font-size: 0.85rem; color: var(--text-secondary);
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
  .analysis-tabs {
    display: inline-flex;
    gap: 0.4rem;
    background: var(--bg-card);
    border: 1px solid var(--border-glass);
    border-radius: 999px;
    padding: 0.35rem;
    margin-bottom: 1.3rem;
  }
  .analysis-tab-btn {
    border: none;
    border-radius: 999px;
    background: transparent;
    color: var(--text-secondary);
    font-size: 0.85rem;
    font-weight: 600;
    padding: 0.45rem 1rem;
    cursor: pointer;
  }
  .analysis-tab-btn.active {
    background: var(--bg-secondary);
    color: var(--accent);
  }
  .analysis-pane { display: none; }
  .analysis-pane.active { display: block; }
  .chart-card {
    background: var(--bg-card); border-radius: var(--radius-card); padding: 2rem;
    border: 1px solid var(--border-glass);
  }
  .chart-card h2 { font-family: var(--font-head); font-size: 1.1rem; font-weight: 500; margin-bottom: 1.5rem; color: var(--text-primary); }
  .chart-container { display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 1rem; }
  .chart-actions { display: flex; gap: 0.5rem; margin-top: 1rem; flex-wrap: wrap; }
  .chart-card .ext-browser { margin-top: 1rem; margin-bottom: 0; }
  .legend { display: flex; flex-direction: row; flex-wrap: wrap; gap: 0.45rem 0.6rem; font-size: 0.82rem; width: 100%; max-height: none; overflow: visible; }
  .legend-item {
    display: inline-flex; align-items: center; gap: 0.45rem; color: var(--text-secondary);
    border-radius: 999px;
    padding: 0.15rem 0; white-space: nowrap; font-size: 0.78rem;
  }
  .legend-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
  .legend-count { display: none; }
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
  .qf-btn-empty { opacity: 0.45; }
  .ext-browser {
    background: var(--bg-secondary);
    border: 1px solid var(--border-glass);
    border-radius: 14px;
    padding: 0.8rem;
    margin-bottom: 1.2rem;
  }
  .ext-browser-head {
    display: flex;
    gap: 0.6rem;
    align-items: center;
    margin-bottom: 0.7rem;
  }
  .ext-browser-head input {
    flex: 1;
    padding: 0.45rem 0.9rem;
    border-radius: 99px;
    border: 1px solid var(--border-glass);
    background: var(--bg-card);
    color: var(--text-primary);
    font-size: 0.8rem;
    outline: none;
  }
  .ext-browser-head input:focus { border-color: var(--accent); }
  .ext-browser-list {
    display: flex;
    flex-wrap: wrap;
    gap: 0.45rem;
    max-height: 220px;
    overflow: auto;
    padding-right: 0.25rem;
  }
  .ext-browser-list .qf-btn { font-size: 0.78rem; padding: 0.35rem 0.8rem; }

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
  .activity-section { padding: 0.8rem 1.1rem; margin-bottom: 1rem; }
  .activity-section .section-header { margin-bottom: 0; }
  .activity-section.activity-expanded { padding: 1.5rem 1.8rem; }
  .activity-section.activity-expanded .section-header { margin-bottom: 0.8rem; }
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
  .top10-size { font-family: 'JetBrains Mono', 'Cascadia Code', 'Consolas', 'Courier New', monospace; color: var(--accent); white-space: nowrap; }

  /* Fallback: basic table styling when DataTables CDN fails */
  table.display:not(.dataTable) { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
  table.display:not(.dataTable) th, table.display:not(.dataTable) td { padding: 0.5rem 0.8rem; border-bottom: 1px solid var(--border-glass); text-align: left; }
  table.display:not(.dataTable) th { font-weight: 600; color: var(--text-primary); }
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
    .chart-section, .delta-section, .info-box, .quick-filters, .dt-buttons, .dt-search, .dt-length, .structure-section, .top10-section, .offline-banner, .docling-section { display: none !important; }
    .stat-card { border: 1px solid #ddd; box-shadow: none; }
    .stat-value { color: #333; }
    table.dataTable thead th { background: #f0f0f0 !important; color: #333 !important; }
    table.dataTable tbody td { color: #333; }
  }

  /* Docling Conversion Panel */
  .docling-section { background: var(--bg-card); border-radius: var(--radius-card); padding: 1.5rem 2rem; margin: 1.5rem 0; border: 2px solid var(--success); }
  .folder-tree { max-height: 400px; overflow-y: auto; border: 1px solid var(--border-glass); border-radius: 12px; padding: 0.5rem; background: var(--bg-secondary); }
  .ft-node { padding-left: 1.2rem; }
  .ft-node-root { padding-left: 0; }
  .ft-row { display: flex; align-items: center; gap: 0.4rem; padding: 0.25rem 0.4rem; border-radius: 6px; cursor: default; font-size: 0.82rem; user-select: none; }
  .ft-row:hover { background: var(--bg-card); }
  .ft-toggle { width: 18px; text-align: center; font-size: 0.7rem; color: var(--text-muted); cursor: pointer; flex-shrink: 0; }
  .ft-toggle.empty { visibility: hidden; }
  .ft-cb { accent-color: var(--accent); width: 15px; height: 15px; cursor: pointer; flex-shrink: 0; }
  .ft-label { flex: 1; color: var(--text-primary); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; cursor: pointer; }
  .ft-count { font-size: 0.72rem; color: var(--text-muted); font-family: 'JetBrains Mono', monospace; white-space: nowrap; }
  .ft-children { display: none; }
  .ft-children.expanded { display: block; }
  .dl-options { display: grid; grid-template-columns: 1fr 1fr; gap: 0.8rem; margin: 1rem 0; }
  @media (max-width: 768px) { .dl-options { grid-template-columns: 1fr; } }
  .dl-opt-group { background: var(--bg-secondary); border-radius: 12px; padding: 0.7rem 1rem; border: 1px solid var(--border-glass); }
  .dl-opt-group h4 { font-size: 0.72rem; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; margin: 0 0 0.4rem 0; }
  .dl-opt-row { display: flex; flex-wrap: wrap; gap: 0.3rem 0.8rem; }
  .dl-opt-row label { display: flex; align-items: center; gap: 0.25rem; font-size: 0.8rem; color: var(--text-secondary); cursor: pointer; }
  .dl-opt-row input[type="radio"], .dl-opt-row input[type="checkbox"] { accent-color: var(--accent); }
  .dl-summary { display: flex; flex-wrap: wrap; gap: 0.6rem 1.2rem; align-items: center; padding: 0.7rem 1rem; background: var(--bg-secondary); border-radius: 12px; border: 1px solid var(--border-glass); margin: 0.8rem 0; font-size: 0.82rem; }
  .dl-summary-val { font-weight: 600; color: var(--accent); }
</style>
</head>
<body>
<div class="container">

  <!-- Offline Warning -->
  <div class="offline-banner" id="offline-banner">⚠️ Einige Funktionen (Treemap, Tabellen-Export) benötigen eine Internetverbindung für CDN-Bibliotheken.</div>

  <!-- Header -->
  <div class="header">
    <h1>Folder Scan Dashboard <span>&bull; Übersicht</span></h1>
    <div class="scan-path">$ResolvedPath</div>
    <div class="scan-meta">$scanInfoText</div>
  </div>

  <!-- Stats -->
  <div class="stats-grid" id="stats-grid">
    <div class="stat-card"><div class="stat-value" id="stat-files">$totalFiles</div><div class="stat-label">Dateien</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-size">$totalSizeDisplay</div><div class="stat-label">Gesamtgröße</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-folders">$uniqueFolders</div><div class="stat-label">Ordner (mit Dateien)<br><small style="opacity:0.6">$totalFolders gesamt</small></div></div>
    <div class="stat-card"><div class="stat-value" id="stat-depth">$maxDepth</div><div class="stat-label">Max. Tiefe</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-conv">$convertibleCount</div><div class="stat-label">Konvertierbar</div></div>
    <div class="stat-card"><div class="stat-value" id="stat-types">$($extGroups.Count)</div><div class="stat-label">Dateitypen</div></div>
  </div>

  $deltaHtml

  <!-- Activity Chart -->
  <div class="structure-section activity-section" id="activity-section">
    <div class="section-header">
      <h2 onclick="toggleActivity()"><span class="collapse-arrow collapsed" id="activity-arrow">▼</span> 📅 Datei-Aktivität (Anzahl Dateien pro Änderungsdatum)</h2>
    </div>
    <div id="activity-body" data-collapsed="1" style="transition: max-height 0.35s ease, opacity 0.25s ease; max-height: 0px; opacity: 0; overflow: hidden;">
      <div id="activity-chart" style="width:100%;height:280px"></div>
    </div>
  </div>

  <!-- Top 10 Biggest Files -->
  <div class="top10-section" style="background:var(--bg-card);border-radius:var(--radius-card);padding:1.5rem 2rem;border:1px solid var(--border-glass);margin-bottom:1.5rem">
    <h2>🏋️ Top 10 größte Dateien</h2>
    <div class="top10-list" id="top10-list"></div>
  </div>

  <!-- Analysis Tabs -->
  <div class="analysis-tabs" id="analysis-tabs">
    <button class="analysis-tab-btn active" id="btn-pane-files" onclick="switchAnalysisPane('files')">📊 Datei-Analyse</button>
    <button class="analysis-tab-btn" id="btn-pane-structure" onclick="switchAnalysisPane('structure')">🧱 Struktur-Analyse</button>
  </div>

  <div class="analysis-pane active" id="pane-files">
    <!-- Charts -->
    <div class="chart-section">
      <div class="chart-card">
        <h2>📊 Verteilung nach Dateiendung</h2>
        <div class="chart-container">
          <div style="position:relative">
            <svg id="pie-ext" width="420" height="420" viewBox="-210 -210 420 420"></svg>
            <div class="pie-tooltip" id="tip-ext"></div>
          </div>
          <div class="legend" id="leg-ext"></div>
        </div>
      </div>
      <div class="chart-card">
        <h2>✅ Konvertierbarkeit</h2>
        <div class="chart-container">
          <div style="position:relative">
            <svg id="pie-conv" width="420" height="420" viewBox="-210 -210 420 420"></svg>
            <div class="pie-tooltip" id="tip-conv"></div>
          </div>
          <div class="legend" id="leg-conv"></div>
        </div>
      </div>
    </div>
  </div>

  <div class="analysis-pane" id="pane-structure">
    <!-- Structure Analysis -->
    <div class="structure-section">
      <div class="section-header">
        <h2 onclick="toggleTreemap()"><span class="collapse-arrow collapsed" id="tm-arrow">▼</span> 🧱 Struktur-Analyse</h2>
        <div class="toggle-group" id="tm-toggle-group" style="display:none">
          <button class="qf-btn active" id="btn-chart-treemap" onclick="switchChartType('treemap')">Treemap</button>
          <button class="qf-btn" id="btn-chart-sunburst" onclick="switchChartType('sunburst')">Sunburst</button>
          <span style="border-left:1px solid var(--border-glass);margin:0 4px"></span>
          <button class="qf-btn active" id="btn-tm-count" onclick="drawTreemap('count')">Anzahl</button>
          <button class="qf-btn" id="btn-tm-size" onclick="drawTreemap('size')">Größe</button>
        </div>
      </div>
      <div class="treemap-body collapsed" id="treemap-body">
        <div id="treemap-container"></div>
        <div class="treemap-hint">Klicke auf einen Ordner, um die Dateiliste unten auf diesen Ordner zu filtern. Doppelklick zum Zoomen.</div>
      </div>
    </div>
  </div>

  <!-- Global Folder Selection -->
  <details id="gf-panel" style="background:var(--bg-card);border-radius:var(--radius-card);padding:0;margin-bottom:1.5rem;border:1px solid var(--border-glass)">
    <summary style="padding:1rem 1.5rem;cursor:pointer;font-weight:600;font-size:0.95rem;list-style:none;display:flex;align-items:center;gap:0.6rem;user-select:none">
      <span style="transition:transform 0.2s;display:inline-block" id="gf-arrow">&#9654;</span>
      &#128193; Ordner-Auswahl
      <span id="gf-summary-badge" style="font-weight:400;font-size:0.78rem;color:var(--text-muted);margin-left:0.5rem">(alle Ordner)</span>
    </summary>
    <div style="padding:0 1.5rem 1.2rem 1.5rem">
      <div style="display:flex;gap:0.5rem;align-items:center;margin-bottom:0.8rem;flex-wrap:wrap">
        <button class="qf-btn" id="gf-btn-all" onclick="gfSelectAll()" style="font-size:0.78rem;padding:0.3rem 0.8rem">&#10004; Alle</button>
        <button class="qf-btn" id="gf-btn-none" onclick="gfSelectNone()" style="font-size:0.78rem;padding:0.3rem 0.8rem">&#10006; Keine</button>
        <button class="qf-btn" id="gf-btn-conv" onclick="gfSelectConvertible()" style="font-size:0.78rem;padding:0.3rem 0.8rem">&#9989; Nur konvertierbare</button>
        <input type="text" id="gf-search" placeholder="Ordner suchen..." style="flex:1;min-width:150px;padding:0.35rem 0.8rem;border-radius:8px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.8rem" oninput="gfFilterTree(this.value)">
      </div>
      <div id="gf-folder-tree" style="max-height:350px;overflow-y:auto;border:1px solid var(--border-glass);border-radius:8px;padding:0.5rem;background:var(--bg-secondary)"></div>
      <div style="margin-top:0.6rem;font-size:0.78rem;color:var(--text-muted);display:flex;gap:1.2rem;flex-wrap:wrap">
        <span>Ordner: <strong id="gf-sum-folders">0</strong></span>
        <span>Dateien: <strong id="gf-sum-files">0</strong></span>
        <span>Gr&ouml;&szlig;e: <strong id="gf-sum-size">0</strong></span>
        <span>(<strong id="gf-sum-pct">100</strong>% aller Dateien)</span>
      </div>
    </div>
  </details>

  <!-- Unified Filter Bar -->
  <div class="filter-bar" id="filter-bar">
    <div class="filter-bar-row" id="qf-container"></div>
    <div class="ext-browser" id="ext-browser" data-mode="all" style="display:none;margin-top:0.8rem">
      <div class="ext-browser-head">
        <input id="ext-browser-search" type="text" placeholder="Dateiendung suchen (z.B. .yaml, .cs, .h)" oninput="applyExtBrowserFilter()">
        <button class="qf-btn" onclick="selectAllVisibleExt()">Alle auswählen</button>
        <button class="qf-btn" onclick="deselectAllVisibleExt()">Keine</button>
        <button class="qf-btn" id="btn-ext-browser-close" onclick="toggleExtBrowser(false)">Schließen</button>
      </div>
      <div class="ext-browser-list" id="ext-browser-list"></div>
    </div>
  </div>

  <!-- Folder Filter Indicator -->
  <div id="folder-filter-indicator" style="display:none;background:var(--bg-card);border-radius:var(--radius-card);padding:1rem 1.5rem;margin-bottom:1rem;border:1px solid var(--accent);font-size:0.9rem;"></div>

  <!-- File Table -->
  <div class="table-section">
    <h2>📋 Dateiliste</h2>
    <div class="folder-search">
      <input type="text" id="folder-search-input" placeholder="🔍 Ordner filtern (z.B. Mockup-Data)..." oninput="filterByPath(this.value)">
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

  <!-- Workflow Guide -->
  <div class="structure-section" id="workflow-guide" style="padding:1.5rem 2rem;margin:1.5rem 0;border:2px solid var(--accent);border-radius:var(--radius-card);background:var(--bg-card)">
    <h2 style="margin:0 0 1rem 0;font-size:1.1rem">Dateien verarbeiten &mdash; Schritt f&uuml;r Schritt</h2>

    <!-- Step 1: Filter -->
    <div class="wf-step" style="display:flex;gap:1rem;align-items:flex-start;margin-bottom:1.2rem;padding-bottom:1.2rem;border-bottom:1px solid var(--border-glass)">
      <div style="min-width:2.2rem;height:2.2rem;border-radius:50%;background:var(--accent);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:1rem;flex-shrink:0">1</div>
      <div style="flex:1">
        <div style="font-weight:600;font-size:0.95rem;margin-bottom:0.3rem">Dateien filtern</div>
        <div style="font-size:0.82rem;color:var(--text-muted)">Nutze die Filter oben (Konvertierbar, Dateityp, Ordner), um nur die Dateien anzuzeigen, die du verarbeiten m&ouml;chtest.<br>
        <strong>Aktuell im Filter:</strong> <span id="wf-filter-count" style="color:var(--accent);font-weight:600">0</span> Dateien</div>
      </div>
    </div>

    <!-- Step 2: Export Pipeline CSV -->
    <div class="wf-step" style="display:flex;gap:1rem;align-items:flex-start;margin-bottom:1.2rem;padding-bottom:1.2rem;border-bottom:1px solid var(--border-glass)">
      <div style="min-width:2.2rem;height:2.2rem;border-radius:50%;background:var(--accent);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:1rem;flex-shrink:0">2</div>
      <div style="flex:1">
        <div style="font-weight:600;font-size:0.95rem;margin-bottom:0.3rem">Arbeitsliste exportieren &amp; verschieben</div>
        <div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:0.5rem">Exportiere die gefilterte Auswahl als CSV. Die Datei landet im Browser-Downloads-Ordner &mdash; f&uuml;hre danach den angezeigten Befehl aus, um sie in den <strong>scan-results</strong> Ordner zu verschieben.</div>
        <div style="display:flex;gap:0.5rem;flex-wrap:wrap;align-items:center">
          <button class="qf-btn qf-btn-accent" id="btn-pipeline-csv" onclick="exportPipelineCSV()" style="font-size:0.9rem;padding:0.5rem 1.4rem">📋 Arbeitsliste speichern (CSV)</button>
          <span id="wf-csv-saved" style="font-size:0.8rem;color:var(--success);opacity:0;transition:opacity 0.3s"></span>
        </div>
        <div id="wf-move-section" style="display:none;margin-top:0.6rem">
          <div style="font-size:0.78rem;color:var(--text-muted);margin-bottom:0.3rem">Diesen Befehl in PowerShell ausf&uuml;hren, um die Datei in den richtigen Ordner zu verschieben:</div>
          <div style="position:relative">
            <pre id="wf-move-cmd" style="background:var(--bg-secondary);border:1px solid var(--border-glass);border-radius:8px;padding:0.6rem 1rem;font-family:'JetBrains Mono','Cascadia Code','Consolas',monospace;font-size:0.75rem;color:var(--text-primary);white-space:pre-wrap;word-break:break-all;margin:0;cursor:pointer" onclick="copyMoveCmd()"></pre>
            <button onclick="copyMoveCmd()" style="position:absolute;top:0.3rem;right:0.3rem;background:var(--accent);color:#fff;border:none;border-radius:6px;padding:0.25rem 0.7rem;font-size:0.72rem;cursor:pointer">📋 Kopieren</button>
          </div>
          <div id="wf-move-copied" style="font-size:0.72rem;color:var(--success);margin-top:0.2rem;opacity:0;transition:opacity 0.3s"></div>
        </div>
      </div>
    </div>

    <!-- Step 3: Review -->
    <div class="wf-step" style="display:flex;gap:1rem;align-items:flex-start;margin-bottom:1.2rem;padding-bottom:1.2rem;border-bottom:1px solid var(--border-glass)">
      <div style="min-width:2.2rem;height:2.2rem;border-radius:50%;background:var(--text-muted);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:1rem;flex-shrink:0" id="wf-step3-num">3</div>
      <div style="flex:1">
        <div style="font-weight:600;font-size:0.95rem;margin-bottom:0.3rem">&Ouml;ffnen &amp; pr&uuml;fen</div>
        <div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:0.5rem">Pr&uuml;fe die Arbeitsliste bevor du kopierst. Du kannst sie in Excel &ouml;ffnen oder mit diesem Befehl interaktiv anzeigen:</div>
        <div id="wf-review-section" style="display:none">
          <div style="position:relative">
            <pre id="wf-review-cmd" style="background:var(--bg-secondary);border:1px solid var(--border-glass);border-radius:8px;padding:0.6rem 1rem;font-family:'JetBrains Mono','Cascadia Code','Consolas',monospace;font-size:0.75rem;color:var(--text-primary);white-space:pre-wrap;word-break:break-all;margin:0;cursor:pointer" onclick="copyReviewCmd()"></pre>
            <button onclick="copyReviewCmd()" style="position:absolute;top:0.3rem;right:0.3rem;background:var(--accent);color:#fff;border:none;border-radius:6px;padding:0.25rem 0.7rem;font-size:0.72rem;cursor:pointer">📋 Kopieren</button>
          </div>
          <div id="wf-review-copied" style="font-size:0.72rem;color:var(--success);margin-top:0.2rem;opacity:0;transition:opacity 0.3s"></div>
          <div style="font-size:0.75rem;color:var(--text-muted);margin-top:0.3rem">L&ouml;sche Zeilen die du nicht kopieren m&ouml;chtest. Erst wenn du zufrieden bist &rarr; weiter zu Schritt 4.</div>
        </div>
        <div id="wf-review-placeholder" style="font-size:0.78rem;color:var(--text-muted);opacity:0.5;font-style:italic">Wird nach Schritt 2 freigeschaltet</div>
      </div>
    </div>

    <!-- Step 4: Copy Command -->
    <div class="wf-step" style="display:flex;gap:1rem;align-items:flex-start;margin-bottom:1.2rem;padding-bottom:1.2rem;border-bottom:1px solid var(--border-glass)">
      <div style="min-width:2.2rem;height:2.2rem;border-radius:50%;background:var(--text-muted);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:1rem;flex-shrink:0" id="wf-step4-num">4</div>
      <div style="flex:1">
        <div style="font-weight:600;font-size:0.95rem;margin-bottom:0.3rem">In Staging kopieren</div>
        <div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:0.5rem">F&uuml;ge den Befehl in PowerShell ein. Die Ordnerstruktur wird 1:1 gespiegelt. Bereits vorhandene Dateien werden &uuml;bersprungen.</div>
        <div style="display:flex;gap:0.5rem;align-items:center;margin-bottom:0.5rem;flex-wrap:wrap">
          <label style="font-size:0.82rem;font-weight:500;color:var(--text-secondary)">Staging-Ordner:</label>
          <input type="text" id="cmd-dest" value="$(if ($StagingPath) { $StagingPath } else { 'C:\TEMP\Staging' })" style="flex:1;min-width:200px;padding:0.4rem 0.8rem;border-radius:8px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.82rem;font-family:'JetBrains Mono','Cascadia Code','Consolas',monospace" oninput="updateCopyCommand();dlSyncFromStaging()">
          <label style="font-size:0.82rem;display:flex;align-items:center;gap:0.3rem;cursor:pointer">
            <input type="checkbox" id="cmd-whatif" onchange="updateCopyCommand()"> -WhatIf (Testlauf)
          </label>
        </div>
        <div style="position:relative">
          <pre id="cmd-output" style="background:var(--bg-secondary);border:1px solid var(--border-glass);border-radius:8px;padding:0.8rem 1rem;font-family:'JetBrains Mono','Cascadia Code','Consolas',monospace;font-size:0.78rem;color:var(--text-primary);white-space:pre-wrap;word-break:break-all;margin:0;cursor:pointer;opacity:0.4" onclick="copyCmdToClipboard()" title="Erst Arbeitsliste exportieren (Schritt 2)" id="cmd-output-pre">Bitte erst Arbeitsliste exportieren (Schritt 2)</pre>
          <button onclick="copyCmdToClipboard()" style="position:absolute;top:0.3rem;right:0.3rem;background:var(--accent);color:#fff;border:none;border-radius:6px;padding:0.25rem 0.7rem;font-size:0.72rem;cursor:pointer" id="btn-copy-cmd" disabled>📋 Kopieren</button>
        </div>
        <div id="cmd-copied" style="font-size:0.75rem;color:var(--success);margin-top:0.3rem;opacity:0;transition:opacity 0.3s"></div>
      </div>
    </div>

    <!-- Step 5: Rescan -->
    <div class="wf-step" style="display:flex;gap:1rem;align-items:flex-start">
      <div style="min-width:2.2rem;height:2.2rem;border-radius:50%;background:var(--text-muted);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:1rem;flex-shrink:0" id="wf-step5-num">5</div>
      <div style="flex:1">
        <div style="font-weight:600;font-size:0.95rem;margin-bottom:0.3rem">Rescan &mdash; Status pr&uuml;fen</div>
        <div style="font-size:0.82rem;color:var(--text-muted);margin-bottom:0.5rem">F&uuml;hre den Scan erneut aus, um den Pipeline-Status zu aktualisieren. Das Dashboard zeigt dann welche Dateien bereits kopiert/veredelt sind.</div>
        <div style="position:relative">
          <pre id="cmd-rescan" style="background:var(--bg-secondary);border:1px solid var(--border-glass);border-radius:8px;padding:0.6rem 1rem;font-family:'JetBrains Mono','Cascadia Code','Consolas',monospace;font-size:0.75rem;color:var(--text-primary);white-space:pre-wrap;word-break:break-all;margin:0;cursor:pointer" onclick="copyRescanCmd()"></pre>
          <button onclick="copyRescanCmd()" style="position:absolute;top:0.3rem;right:0.3rem;background:var(--accent);color:#fff;border:none;border-radius:6px;padding:0.25rem 0.7rem;font-size:0.72rem;cursor:pointer">📋 Kopieren</button>
        </div>
        <div id="wf-rescan-copied" style="font-size:0.72rem;color:var(--success);margin-top:0.2rem;opacity:0;transition:opacity 0.3s"></div>
      </div>
    </div>

    <!-- Quick exports (less prominent, at the bottom) -->
    <div style="margin-top:1.2rem;padding-top:1rem;border-top:1px solid var(--border-glass);display:flex;gap:0.5rem;flex-wrap:wrap;align-items:center">
      <span style="font-size:0.78rem;color:var(--text-muted);margin-right:0.3rem">Sonstige Exporte:</span>
      <button class="qf-btn" id="btn-quick-csv" onclick="quickExportCSV()" style="font-size:0.78rem;padding:0.35rem 0.8rem">📄 CSV (Ansicht)</button>
      <button class="qf-btn" id="btn-quick-excel" onclick="quickExportExcel()" style="font-size:0.78rem;padding:0.35rem 0.8rem">📊 Excel (Ansicht)</button>
    </div>
  </div>

  <!-- Docling Conversion Panel -->
  <div class="docling-section" id="docling-section">
    <h2 style="margin:0 0 0.5rem 0;font-size:1.1rem">🤖 Docling Konvertierung &mdash; Dateien veredeln</h2>
    <p style="font-size:0.82rem;color:var(--text-muted);margin-bottom:1rem">Sende Dateien aus dem Staging-Ordner an die Docling API zur automatischen Konvertierung (PDF &rarr; Markdown, HTML, etc.).</p>

    <div style="display:flex;gap:0.8rem;flex-wrap:wrap;align-items:end;margin-bottom:1rem">
      <div style="flex:1;min-width:200px">
        <label style="font-size:0.78rem;font-weight:600;color:var(--text-secondary);display:block;margin-bottom:0.3rem">Docling API URL</label>
        <input type="text" id="dl-url" placeholder="http://localhost:5001" style="width:100%;padding:0.5rem 0.8rem;border-radius:8px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.82rem;font-family:'JetBrains Mono',monospace" oninput="dlSaveSettings();dlBuildCommand()">
      </div>
      <div style="flex:1;min-width:200px">
        <label style="font-size:0.78rem;font-weight:600;color:var(--text-secondary);display:block;margin-bottom:0.3rem">Staging-Ordner (Input)</label>
        <input type="text" id="dl-input" value="$(if ($StagingPath) { $StagingPath } else { 'C:\TEMP\Staging' })" placeholder="C:\TEMP\Staging" style="width:100%;padding:0.5rem 0.8rem;border-radius:8px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.82rem;font-family:'JetBrains Mono',monospace" oninput="dlOnInputManual();dlSaveSettings();dlBuildCommand()">
      </div>
      <div style="flex:1;min-width:200px">
        <label style="font-size:0.78rem;font-weight:600;color:var(--text-secondary);display:block;margin-bottom:0.3rem">Ergebnis-Ordner (Output)</label>
        <input type="text" id="dl-output" value="$(if ($ErgebnisPath) { $ErgebnisPath } else { 'C:\TEMP\Result' })" placeholder="C:\TEMP\Result" style="width:100%;padding:0.5rem 0.8rem;border-radius:8px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.82rem;font-family:'JetBrains Mono',monospace" oninput="dlOnOutputManual();dlSaveSettings();dlBuildCommand()">
      </div>
    </div>

    <div style="margin-bottom:1rem">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:0.5rem;flex-wrap:wrap;gap:0.5rem">
        <label style="font-size:0.78rem;font-weight:600;color:var(--text-secondary)">Ordner ausw&auml;hlen (rekursiv) <span id="dl-folder-count" style="font-weight:400;color:var(--text-muted)"></span></label>
        <div style="display:flex;gap:0.4rem">
          <button class="qf-btn" id="dl-btn-all" onclick="dlSelectAll()" style="font-size:0.72rem;padding:0.25rem 0.7rem">Alle</button>
          <button class="qf-btn" id="dl-btn-none" onclick="dlSelectNone()" style="font-size:0.72rem;padding:0.25rem 0.7rem">Keine</button>
          <button class="qf-btn" id="dl-btn-conv" onclick="dlSelectConvertible()" style="font-size:0.72rem;padding:0.25rem 0.7rem">Nur konvertierbare</button>
        </div>
      </div>
      <input type="text" id="dl-folder-search" placeholder="Ordner suchen..." style="width:100%;padding:0.4rem 0.8rem;border-radius:8px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.8rem;margin-bottom:0.5rem" oninput="dlFilterTree(this.value)">
      <div class="folder-tree" id="dl-folder-tree"></div>
    </div>

    <div class="dl-summary" id="dl-summary">
      <div>📂 <span class="dl-summary-val" id="dl-sum-folders">0</span> Ordner</div>
      <div>📄 <span class="dl-summary-val" id="dl-sum-files">0</span> Dateien</div>
      <div>(<span class="dl-summary-val" id="dl-sum-pct">0</span>%)</div>
      <div>💾 <span class="dl-summary-val" id="dl-sum-size">0 B</span></div>
      <div>⏱️ <span class="dl-summary-val" id="dl-sum-eta">&ndash;</span></div>
    </div>

    <details style="margin-bottom:1rem">
      <summary style="cursor:pointer;font-size:0.9rem;font-weight:600;color:var(--text-primary);padding:0.5rem 0">⚙️ Docling Optionen <span style="font-size:0.72rem;font-weight:400;color:var(--text-muted)">(siehe auch: &lt;Docling-URL&gt;/ui)</span></summary>
      <div class="dl-options" id="dl-options">
        <div class="dl-opt-group">
          <h4>To Formats</h4>
          <div class="dl-opt-row">
            <label><input type="checkbox" name="dl-fmt" value="markdown" checked onchange="dlBuildCommand()"> Markdown</label>
            <label><input type="checkbox" name="dl-fmt" value="html" onchange="dlBuildCommand()"> HTML</label>
            <label><input type="checkbox" name="dl-fmt" value="text" onchange="dlBuildCommand()"> Plain Text</label>
            <label><input type="checkbox" name="dl-fmt" value="json" onchange="dlBuildCommand()"> Docling JSON</label>
            <label><input type="checkbox" name="dl-fmt" value="doctags" onchange="dlBuildCommand()"> Doc Tags</label>
          </div>
        </div>
        <div class="dl-opt-group">
          <h4>Image Export Mode</h4>
          <div class="dl-opt-row">
            <label><input type="radio" name="dl-imgmode" value="embedded" checked onchange="dlBuildCommand()"> Embedded</label>
            <label><input type="radio" name="dl-imgmode" value="placeholder" onchange="dlBuildCommand()"> Placeholder</label>
            <label><input type="radio" name="dl-imgmode" value="referenced" onchange="dlBuildCommand()"> Referenced</label>
          </div>
        </div>
        <div class="dl-opt-group">
          <h4>Pipeline Type</h4>
          <div class="dl-opt-row">
            <label><input type="radio" name="dl-pipeline" value="legacy" onchange="dlBuildCommand()"> Legacy</label>
            <label><input type="radio" name="dl-pipeline" value="standard" checked onchange="dlBuildCommand()"> Standard</label>
            <label><input type="radio" name="dl-pipeline" value="vlm" onchange="dlBuildCommand()"> Vlm</label>
            <label><input type="radio" name="dl-pipeline" value="asr" onchange="dlBuildCommand()"> Asr</label>
          </div>
        </div>
        <div class="dl-opt-group">
          <h4>OCR</h4>
          <div class="dl-opt-row">
            <label><input type="checkbox" id="dl-ocr" checked onchange="dlBuildCommand()"> Enable OCR</label>
            <label><input type="checkbox" id="dl-forceocr" onchange="dlBuildCommand()"> Force OCR</label>
          </div>
          <div class="dl-opt-row" style="margin-top:0.3rem">
            <label><input type="radio" name="dl-ocreng" value="auto" checked onchange="dlBuildCommand()"> Auto</label>
            <label><input type="radio" name="dl-ocreng" value="easyocr" onchange="dlBuildCommand()"> EasyOCR</label>
            <label><input type="radio" name="dl-ocreng" value="tesseract" onchange="dlBuildCommand()"> Tesseract</label>
            <label><input type="radio" name="dl-ocreng" value="rapidocr" onchange="dlBuildCommand()"> RapidOCR</label>
          </div>
        </div>
        <div class="dl-opt-group">
          <h4>PDF Backend</h4>
          <div class="dl-opt-row">
            <label><input type="radio" name="dl-pdfbe" value="pypdfium2" onchange="dlBuildCommand()"> pypdfium2</label>
            <label><input type="radio" name="dl-pdfbe" value="dlparse_v1" onchange="dlBuildCommand()"> dlparse_v1</label>
            <label><input type="radio" name="dl-pdfbe" value="dlparse_v2" onchange="dlBuildCommand()"> dlparse_v2</label>
            <label><input type="radio" name="dl-pdfbe" value="dlparse_v4" checked onchange="dlBuildCommand()"> dlparse_v4</label>
          </div>
        </div>
        <div class="dl-opt-group">
          <h4>Table Mode</h4>
          <div class="dl-opt-row">
            <label><input type="radio" name="dl-tblmode" value="fast" onchange="dlBuildCommand()"> Fast</label>
            <label><input type="radio" name="dl-tblmode" value="accurate" checked onchange="dlBuildCommand()"> Accurate</label>
          </div>
        </div>
        <div class="dl-opt-group" style="grid-column:1/-1">
          <h4>Enrichment</h4>
          <div class="dl-opt-row">
            <label><input type="checkbox" id="dl-code-enrich" onchange="dlBuildCommand()"> Code enrichment</label>
            <label><input type="checkbox" id="dl-formula-enrich" onchange="dlBuildCommand()"> Formula enrichment</label>
            <label><input type="checkbox" id="dl-pic-class" onchange="dlBuildCommand()"> Picture classification</label>
            <label><input type="checkbox" id="dl-pic-desc" onchange="dlBuildCommand()"> Picture description</label>
          </div>
        </div>
      </div>
    </details>

    <div style="display:flex;gap:1.2rem;align-items:center;margin-bottom:1rem;flex-wrap:wrap">
      <label style="font-size:0.82rem;display:flex;align-items:center;gap:0.4rem;cursor:pointer"><input type="checkbox" id="dl-skip" checked onchange="dlBuildCommand()" style="accent-color:var(--success)"> <strong>Skip existing</strong> <span style="font-size:0.72rem;color:var(--text-muted)">(bereits konvertierte Dateien &uuml;berspringen)</span></label>
      <label style="font-size:0.82rem;display:flex;align-items:center;gap:0.4rem;cursor:pointer"><input type="checkbox" id="dl-abort" onchange="dlBuildCommand()" style="accent-color:var(--danger)"> <strong>Abort on Error</strong> <span style="font-size:0.72rem;color:var(--text-muted)">(bei Fehler sofort abbrechen)</span></label>
      <label style="font-size:0.82rem;display:flex;align-items:center;gap:0.4rem"><strong>Retries</strong> <input type="number" id="dl-retries" value="3" min="0" max="10" step="1" style="width:3.5rem;padding:0.3rem 0.5rem;border-radius:6px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.82rem;font-family:'JetBrains Mono',monospace;text-align:center" oninput="dlBuildCommand()"> <span style="font-size:0.72rem;color:var(--text-muted)">(Wiederholungen bei Fehler)</span></label>
      <label style="font-size:0.82rem;display:flex;align-items:center;gap:0.4rem"><strong>Timeout</strong> <input type="number" id="dl-timeout" value="900" min="30" max="7200" step="30" style="width:4.5rem;padding:0.3rem 0.5rem;border-radius:6px;border:1px solid var(--border-glass);background:var(--bg-secondary);color:var(--text-primary);font-size:0.82rem;font-family:'JetBrains Mono',monospace;text-align:center" oninput="dlBuildCommand()"> <span style="font-size:0.72rem;color:var(--text-muted)">Sek. pro Datei (grosse PDFs: 900+)</span></label>
    </div>

    <div style="margin-top:1rem">
      <label style="font-size:0.78rem;font-weight:600;color:var(--text-secondary);display:block;margin-bottom:0.3rem">Generierter Befehl</label>
      <div style="position:relative">
        <pre id="dl-cmd" style="background:var(--bg-secondary);border:1px solid var(--border-glass);border-radius:8px;padding:0.8rem 1rem;font-family:'JetBrains Mono','Cascadia Code','Consolas',monospace;font-size:0.75rem;color:var(--text-primary);white-space:pre-wrap;word-break:break-all;margin:0;cursor:pointer;min-height:3rem" onclick="dlCopyCmd()"></pre>
        <button onclick="dlCopyCmd()" style="position:absolute;top:0.3rem;right:0.3rem;background:var(--accent);color:#fff;border:none;border-radius:6px;padding:0.25rem 0.7rem;font-size:0.72rem;cursor:pointer">📋 Kopieren</button>
      </div>
      <div id="dl-cmd-copied" style="font-size:0.72rem;color:var(--success);margin-top:0.2rem;opacity:0;transition:opacity 0.3s"></div>
    </div>
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
const CSV_PATH = "$($csvPath -replace '\\', '\\\\')";
const csvPath = CSV_PATH.replace(/\\\\/g, '\\');
const SCRIPT_ROOT = "$($PSScriptRoot -replace '\\', '\\\\')";
const scriptRoot = SCRIPT_ROOT.replace(/\\\\/g, '\\');
const HAS_PIPELINE = $(if ($hasPipeline) { 'true' } else { 'false' });
const STAGING_PATH = "$(if ($StagingPath) { $StagingPath -replace '\\', '\\\\' } else { '' })";
const stagingPath = STAGING_PATH.replace(/\\\\/g, '\\');
const ERGEBNIS_PATH = "$(if ($ErgebnisPath) { $ErgebnisPath -replace '\\', '\\\\' } else { '' })";
const ergebnisPath = ERGEBNIS_PATH.replace(/\\\\/g, '\\');
const DOWNLOADS_PATH = "$(([Environment]::GetFolderPath('UserProfile') + '\Downloads') -replace '\\', '\\\\')";
const downloadsPath = DOWNLOADS_PATH.replace(/\\\\/g, '\\');
const pipelineStats = { offen: $($pipelineStats.offen), kopiert: $($pipelineStats.kopiert), veredelt: $($pipelineStats.veredelt) };
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
  folder: null, // absolute folder path for treemap click filter
  pipeline: null // null, 'offen', 'kopiert', 'veredelt'
};
let currentOtherExtLabels = [];
let extSelectionBeforeConvAuto = null;

// Build set of convertible extensions from data
var convertibleExtSet = new Set();
(allFiles || []).forEach(function(f) { if (f.convertible && f.ext) convertibleExtSet.add(f.ext); });

function buildExtStats(files) {
  var map = {};
  (files || []).forEach(function(f) {
    var ext = (f && f.ext) ? f.ext : '(none)';
    map[ext] = (map[ext] || 0) + 1;
  });
  return Object.keys(map).map(function(k) {
    return { label: k, count: map[k] };
  }).sort(function(a, b) {
    if (b.count !== a.count) return b.count - a.count;
    return a.label.localeCompare(b.label);
  });
}

const allExtStats = buildExtStats(allFiles);

function applyFilters() {
  // Extension Filter (OR logic)
  if (activeFilters.ext.size > 0) {
    // Escape special chars for regex (e.g. . ++ etc)
    const regex = Array.from(activeFilters.ext).map(e => '^' + e.replace(/[.*+?^\x24{}()|[\x5D\x5C]/g, '\\$&') + '$').join('|');
    table.column(2).search(regex, true, false);
  } else {
    table.column(2).search('');
  }

  // Convertible Filter
  if (activeFilters.conv) {
    table.column(11).search(activeFilters.conv);
  } else {
    table.column(11).search('');
  }

  // Folder Filter (from treemap click)
  if (activeFilters.folder) {
    table.column(3).search('^' + escRx(activeFilters.folder), true, false);
  } else {
    table.column(3).search('');
  }

  // Pipeline Filter
  if (activeFilters.pipeline) {
    table.column(13).search('^' + activeFilters.pipeline + '$', true, false);
  } else {
    table.column(13).search('');
  }

  table.draw();
  updateVisuals();
  updateFolderIndicator();
}

function getFolderScopedFiles() {
  var base = allFiles;
  // Apply global folder selection first
  if (!gfAllSelected) {
    base = base.filter(function(f) { return gfIsFileIncluded(f); });
  }
  if (!activeFilters.folder) return base;
  var folder = activeFilters.folder;
  return base.filter(function(f) {
    var dir = f.dir ? f.dir.replace(/\\\\/g, '\\') : '';
    return dir === folder || dir.indexOf(folder + '\\') === 0;
  });
}

function getFilteredFiles() {
  var scoped = getFolderScopedFiles();
  return scoped.filter(function(f) {
    var ext = f.ext || '(none)';
    if (activeFilters.ext.size > 0 && !activeFilters.ext.has(ext)) return false;
    if (activeFilters.conv === 'Ja' && !f.convertible) return false;
    if (activeFilters.conv === 'Nein' && f.convertible) return false;
    if (activeFilters.pipeline && f.pipeline !== activeFilters.pipeline) return false;
    return true;
  });
}

function togglePipeline(val) {
  activeFilters.pipeline = (activeFilters.pipeline === val) ? null : val;
  applyFilters();
}

function updateVisuals() {
  var filtered = getFilteredFiles();

  // Update quick-filter button counts
  var extCountScope = getFolderScopedFiles().filter(function(f) {
    if (activeFilters.conv === 'Ja' && !f.convertible) return false;
    if (activeFilters.conv === 'Nein' && f.convertible) return false;
    return true;
  });
  var extCounts = {};
  var convCount = 0;
  extCountScope.forEach(function(f) {
    var ext = f.ext || '(none)';
    extCounts[ext] = (extCounts[ext] || 0) + 1;
  });
  filtered.forEach(function(f) {
    if (f.convertible) convCount++;
  });

  document.querySelectorAll('.qf-btn[data-ext]').forEach(function(b) {
    var ext = b.dataset.ext;
    var cnt = extCounts[ext] || 0;
    // Preserve color dot if present
    var dotEl = b.querySelector('span');
    var dotHtml = dotEl ? dotEl.outerHTML : '';
    b.innerHTML = dotHtml + ext + ' (' + cnt + ')';
    b.classList.toggle('active-filter', activeFilters.ext.has(ext));
    b.classList.toggle('qf-btn-empty', cnt === 0);
  });

  // Conv buttons
  var btnConv = document.getElementById('btn-conv-yes');
  if (btnConv) {
    btnConv.textContent = '\u2705 Nur Konvertierbare (' + convCount + ')';
    btnConv.classList.toggle('active-filter', activeFilters.conv === 'Ja');
  }
  var btnConvNo = document.getElementById('btn-conv-no');
  if (btnConvNo) {
    var nonConvCount = filtered.length - convCount;
    btnConvNo.textContent = '\u274c Nur Nicht-Konvertierbare (' + nonConvCount + ')';
    btnConvNo.classList.toggle('active-filter', activeFilters.conv === 'Nein');
  }

  // Pipeline buttons
  if (HAS_PIPELINE) {
    var pCounts = { offen: 0, kopiert: 0, veredelt: 0 };
    filtered.forEach(function(f) { if (f.pipeline && pCounts.hasOwnProperty(f.pipeline)) pCounts[f.pipeline]++; });
    var bO = document.getElementById('btn-pipe-offen');
    if (bO) { bO.textContent = '\u23F3 Offen (' + pCounts.offen + ')'; bO.classList.toggle('active-filter', activeFilters.pipeline === 'offen'); }
    var bK = document.getElementById('btn-pipe-kopiert');
    if (bK) { bK.textContent = '\ud83d\udce6 Kopiert (' + pCounts.kopiert + ')'; bK.classList.toggle('active-filter', activeFilters.pipeline === 'kopiert'); }
    var bV = document.getElementById('btn-pipe-veredelt');
    if (bV) { bV.textContent = '\u2705 Veredelt (' + pCounts.veredelt + ')'; bV.classList.toggle('active-filter', activeFilters.pipeline === 'veredelt'); }
  }

  // Workflow guide: update filter count
  var wfCount = document.getElementById('wf-filter-count');
  if (wfCount) wfCount.textContent = filtered.length;

  // Update stat cards
  updateStatCards(filtered);

  // Update Top-10 biggest files
  try { buildTop10(); } catch(e) {}

  // Update pie charts with FULL folder-scoped data (not ext-filtered)
  // so the pie always shows all extensions; highlighting marks the active ones
  var scopedFiles = getFolderScopedFiles();
  updatePieCharts(scopedFiles);

  // Activity chart uses the fully-filtered set
  drawActivityChart(filtered);

  // Pie Segments highlight (AFTER redraw so classes are not wiped)
  var hasExtFilter = activeFilters.ext.size > 0;
  document.querySelectorAll('.chart-segment[data-ext]').forEach(function(p) {
    var ext = p.dataset.ext;
    var isOtherSelected = ext === 'Other' && currentOtherExtLabels.some(function(label) { return activeFilters.ext.has(label); });
    var isActive = activeFilters.ext.has(ext) || isOtherSelected;
    p.classList.toggle('active', isActive);
    p.classList.toggle('dimmed', hasExtFilter && !isActive);
  });
  var hasConvFilter = !!activeFilters.conv;
  document.querySelectorAll('.chart-segment[data-conv]').forEach(function(p) {
    var isActive = activeFilters.conv === p.dataset.conv;
    p.classList.toggle('active', isActive);
    p.classList.toggle('dimmed', hasConvFilter && !isActive);
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
  
  // Sort by count DESC
  var sortedExts = Object.values(extMap).sort(function(a, b) { return b.count - a.count; });
  var total = filtered.length || 1;
  
  // Group Top 10 + Other
  var top10 = sortedExts.slice(0, 10);
  var others = sortedExts.slice(10);
  currentOtherExtLabels = others.map(function(d) { return d.label; });
  updateOtherExtButton();
  applyExtBrowserFilter();
  
  if (others.length > 0) {
    var otherCount = others.reduce((s, d) => s + d.count, 0);
    var otherSize = others.reduce((s, d) => s + d.sizeBytes, 0);
    top10.push({ label: 'Other', count: otherCount, sizeBytes: otherSize });
  }

  top10.forEach(function(d) {
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

  drawPie('pie-ext','tip-ext','leg-ext', top10, CC, 180, 'ext');
  drawPie('pie-conv','tip-conv','leg-conv', dynConvData, ['#4ade80','#475569'], 180, 'conv');
}

function scrollToTable() {
  var el = document.getElementById('file-table');
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function toggleExt(ext) {
  if (activeFilters.ext.has(ext)) activeFilters.ext.delete(ext);
  else activeFilters.ext.add(ext);
  // If user manually changes ext selection during conv auto-mode, clear conv filter
  if (activeFilters.conv) {
    activeFilters.conv = null;
    extSelectionBeforeConvAuto = null;
  }
  applyFilters();
  scrollToTable();
}

function toggleConv(val) {
  // Toggle off
  if (activeFilters.conv === val) {
    activeFilters.conv = null;
    if (extSelectionBeforeConvAuto) {
      activeFilters.ext = new Set(extSelectionBeforeConvAuto);
      extSelectionBeforeConvAuto = null;
    }
    applyFilters();
    return;
  }

  // Switching away from previous auto-mode restores prior manual ext-selection
  if (activeFilters.conv && extSelectionBeforeConvAuto) {
    activeFilters.ext = new Set(extSelectionBeforeConvAuto);
    extSelectionBeforeConvAuto = null;
  }

  activeFilters.conv = val;
  extSelectionBeforeConvAuto = new Set(activeFilters.ext);

  // Auto-select matching extensions for the chosen filter
  var matchExts = new Set();
  getFolderScopedFiles().forEach(function(f) {
    var match = (val === 'Ja') ? f.convertible : !f.convertible;
    if (match) matchExts.add(f.ext || '(none)');
  });
  activeFilters.ext = matchExts;

  applyFilters();
}

function clearAllFilters() {
  activeFilters.ext.clear();
  activeFilters.conv = null;
  activeFilters.pipeline = null;
  extSelectionBeforeConvAuto = null;
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
  var normalizedRoot = (scanRoot || '').replace(/\\\\/g, '\\').replace(/[\\/]+$/, '').toLowerCase();
  var normalizedTarget = (absPath || '').replace(/\\\\/g, '\\').replace(/[\\/]+$/, '').toLowerCase();

  // Root path clears filter. Same folder no longer toggles off (prevents accidental clear via pathbar clicks).
  if (!normalizedTarget || normalizedTarget === normalizedRoot) {
    activeFilters.folder = null;
  } else {
    activeFilters.folder = (absPath || '').replace(/\\\\/g, '\\').replace(/[\\/]+$/, '');
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
  var current = activeFilters.folder.replace(/\\\\/g, '\\').replace(/[\\/]+$/, '');
  var idx = current.lastIndexOf('\\');
  if (idx <= 0) { clearAllFilters(); return; }
  var parent = current.substring(0, idx);
  var normParent = parent.toLowerCase();
  var normRoot = (scanRoot || '').replace(/\\\\/g, '\\').replace(/[\\/]+$/, '').toLowerCase();
  if (!normParent || normParent.length < normRoot.length || normParent === normRoot) {
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

function switchAnalysisPane(pane) {
  var filesPane = document.getElementById('pane-files');
  var structurePane = document.getElementById('pane-structure');
  var btnFiles = document.getElementById('btn-pane-files');
  var btnStructure = document.getElementById('btn-pane-structure');
  var showFiles = pane !== 'structure';

  if (filesPane) filesPane.classList.toggle('active', showFiles);
  if (structurePane) structurePane.classList.toggle('active', !showFiles);
  if (btnFiles) btnFiles.classList.toggle('active', showFiles);
  if (btnStructure) btnStructure.classList.toggle('active', !showFiles);

  if (!showFiles && window.fileTable && window.fileTable.columns) {
    setTimeout(function() {
      try { window.fileTable.columns.adjust().draw(false); } catch (e) { console.error(e); }
    }, 80);
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

  if (!data || data.length === 0 || total <= 0) {
    var strokeWidth = Math.max(10, Math.round(r * 0.18));
    var emptyR = Math.max(8, Math.round(r * 0.72));
    svg.innerHTML = '<circle cx="0" cy="0" r="' + emptyR + '" fill="none" stroke="rgba(127,127,127,0.35)" stroke-width="' + strokeWidth + '"></circle>' +
      '<text x="0" y="6" text-anchor="middle" fill="rgba(127,127,127,0.85)" font-size="14">Keine Daten</text>';

    const li = document.createElement('div');
    li.className = 'legend-item';
    li.style.cursor = 'default';
    li.style.opacity = '0.8';
    li.textContent = 'Keine Daten im aktuellen Filter';
    leg.appendChild(li);
    return;
  }

  var drawItems = data.map(function(d, i) {
    return {
      label: d.label,
      count: d.count,
      percent: d.percent,
      sizeMB: d.sizeMB,
      colorIndex: i
    };
  }).filter(function(d) { return d.count > 0; });

  function handleSegmentToggle(d) {
    if (type === 'ext') {
      if (d.label === 'Other') { openOtherExtensions(); return; }
      toggleExt(d.label);
    }
    if (type === 'conv') toggleConv(d.label === 'Konvertierbar' ? 'Ja' : 'Nein');
  }

  function bindSegmentEvents(el, d) {
    if (type === 'ext') el.setAttribute('data-ext', d.label);
    if (type === 'conv') el.setAttribute('data-conv', d.label === 'Konvertierbar' ? 'Ja' : 'Nein');

    el.addEventListener('click', function() { handleSegmentToggle(d); });
    el.addEventListener('mouseenter', function() {
      tip.style.opacity = '1';
      tip.innerHTML = '<strong>'+d.label+'</strong><br>'+d.count+' ('+d.percent+'%)'+(d.sizeMB!==undefined?'<br>'+d.sizeMB+' MB':'');
    });
    el.addEventListener('mousemove', function(e) {
      const rect = svg.closest('.chart-container').getBoundingClientRect();
      tip.style.left = (e.clientX-rect.left+12)+'px'; tip.style.top = (e.clientY-rect.top-8)+'px';
    });
    el.addEventListener('mouseleave', function() { tip.style.opacity='0'; });
  }

  if (drawItems.length === 1) {
    var only = drawItems[0];
    var cOnly = colors[only.colorIndex % colors.length];
    var ring = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    ring.setAttribute('cx', '0');
    ring.setAttribute('cy', '0');
    ring.setAttribute('r', ((r + ri) / 2).toString());
    ring.setAttribute('fill', 'none');
    ring.setAttribute('stroke', cOnly);
    ring.setAttribute('stroke-width', (r - ri).toString());
    ring.setAttribute('class', 'chart-segment');
    bindSegmentEvents(ring, only);
    svg.appendChild(ring);
  } else {
    drawItems.forEach((d) => {
      const i = d.colorIndex;
    const ang = (d.count / total) * 2 * Math.PI;
    const x1 = Math.cos(a)*r, y1 = Math.sin(a)*r, x1i = Math.cos(a)*ri, y1i = Math.sin(a)*ri;
    a += ang;
    const x2 = Math.cos(a)*r, y2 = Math.sin(a)*r, x2i = Math.cos(a)*ri, y2i = Math.sin(a)*ri;
    const la = ang > Math.PI ? 1 : 0, c = colors[i % colors.length];
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    p.setAttribute('d', 'M'+x1i+' '+y1i+'L'+x1+' '+y1+'A'+r+' '+r+' 0 '+la+' 1 '+x2+' '+y2+'L'+x2i+' '+y2i+'A'+ri+' '+ri+' 0 '+la+' 0 '+x1i+' '+y1i+'Z');
    p.setAttribute('fill', c);
    p.setAttribute('class', 'chart-segment');
    bindSegmentEvents(p, d);
    svg.appendChild(p);
  });
  }

  data.forEach((d, i) => {
    const c = colors[i % colors.length];
    // Legend click
    const li = document.createElement('div');
    li.className = 'legend-item';
    li.style.cursor = 'pointer';
    var convMark = (type === 'ext' && d.label !== 'Other' && convertibleExtSet.has(d.label)) ? ' ✅' : '';
    li.innerHTML = '<span class="legend-dot" style="background:'+c+'"></span><span>'+d.label+' ('+d.count+')'+convMark+'</span><span class="legend-count">'+d.count+'</span>';
    li.onclick = () => { handleSegmentToggle(d); };
    leg.appendChild(li);
  });
}

// === ACTIVITY CHART ===
function drawActivityChart(files) {
  var target = document.getElementById('activity-chart');
  if (!target) return;

  if (typeof Plotly === 'undefined') {
    target.innerHTML = '<div style="padding:1rem;text-align:center;opacity:0.65">Aktivitäts-Chart benötigt Internet (Plotly CDN).</div>';
    return;
  }

  var buckets = {};
  (files || []).forEach(function(f) {
    if (!f || !f.modified) return;
    var key = String(f.modified).substring(0, 10); // yyyy-MM-dd
    if (!key || key.length < 10) return;
    buckets[key] = (buckets[key] || 0) + 1;
  });

  var days = Object.keys(buckets).sort();
  var values = days.map(function(d) { return buckets[d]; });

  if (days.length === 0) {
    target.innerHTML = '<div style="padding:1rem;text-align:center;opacity:0.65">Keine Daten für Aktivitäts-Chart verfügbar.</div>';
    return;
  }

  var bodyStyle = getComputedStyle(document.body);
  var textColor = bodyStyle.getPropertyValue('--text-secondary').trim();
  var accentColor = bodyStyle.getPropertyValue('--accent').trim() || '#4285F4';

  var trace = {
    x: days,
    y: values,
    type: 'bar',
    marker: { color: accentColor },
    hovertemplate: '%{x}<br>%{y} Dateien<extra></extra>'
  };

  var layout = {
    margin: { t: 8, l: 40, r: 12, b: 48 },
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: textColor },
    xaxis: { title: 'Datum', tickangle: -35, showgrid: false },
    yaxis: { title: 'Dateien', rangemode: 'tozero', gridcolor: 'rgba(127,127,127,0.18)' }
  };

  Plotly.newPlot(target, [trace], layout, { displayModeBar: false, responsive: true });
}

function toggleActivity() {
  var body = document.getElementById('activity-body');
  var arrow = document.getElementById('activity-arrow');
  var section = document.getElementById('activity-section');
  if (!body) return;

  var isCollapsed = body.getAttribute('data-collapsed') === '1';
  if (isCollapsed) {
    body.style.maxHeight = '460px';
    body.style.opacity = '1';
    body.setAttribute('data-collapsed', '0');
    if (arrow) arrow.classList.remove('collapsed');
    if (section) section.classList.add('activity-expanded');
  } else {
    body.style.maxHeight = '0px';
    body.style.opacity = '0';
    body.setAttribute('data-collapsed', '1');
    if (arrow) arrow.classList.add('collapsed');
    if (section) section.classList.remove('activity-expanded');
  }
}

function toggleExtBrowser(forceOpen, mode) {
  var panel = document.getElementById('ext-browser');
  var search = document.getElementById('ext-browser-search');
  if (!panel) return;

  if (mode) {
    panel.setAttribute('data-mode', mode);
    if (search) search.value = '';
  }

  var currentlyOpen = panel.style.display !== 'none';
  var shouldOpen = (typeof forceOpen === 'boolean') ? forceOpen : !currentlyOpen;
  panel.style.display = shouldOpen ? 'block' : 'none';

  if (shouldOpen) {
    applyExtBrowserFilter();
  }
}

function applyExtBrowserFilter() {
  var panel = document.getElementById('ext-browser');
  var list = document.getElementById('ext-browser-list');
  var search = document.getElementById('ext-browser-search');
  if (!panel || !list) return;

  var mode = panel.getAttribute('data-mode') || 'all';
  var query = search ? search.value.trim().toLowerCase() : '';
  var otherMap = {};
  currentOtherExtLabels.forEach(function(ext) { otherMap[String(ext).toLowerCase()] = true; });

  list.querySelectorAll('.qf-btn[data-ext]').forEach(function(btn) {
    var ext = btn.dataset.ext || '';
    var extLower = ext.toLowerCase();
    var inMode = (mode !== 'other') || !!otherMap[extLower];
    var inSearch = (query.length === 0) || (extLower.indexOf(query) >= 0);
    btn.style.display = (inMode && inSearch) ? '' : 'none';
  });
}

function selectAllVisibleExt() {
  var list = document.getElementById('ext-browser-list');
  if (!list) return;
  list.querySelectorAll('.qf-btn[data-ext]').forEach(function(btn) {
    if (btn.style.display !== 'none' && btn.dataset.ext) {
      activeFilters.ext.add(btn.dataset.ext);
    }
  });
  if (activeFilters.conv === 'Ja' && extSelectionBeforeConvAuto) extSelectionBeforeConvAuto = null;
  applyFilters();
}

function deselectAllVisibleExt() {
  var list = document.getElementById('ext-browser-list');
  if (!list) return;
  list.querySelectorAll('.qf-btn[data-ext]').forEach(function(btn) {
    if (btn.style.display !== 'none' && btn.dataset.ext) {
      activeFilters.ext.delete(btn.dataset.ext);
    }
  });
  applyFilters();
}

function openOtherExtensions() {
  toggleExtBrowser(true, 'other');
}

function updateOtherExtButton() {
  var count = currentOtherExtLabels.length;
  ['btn-show-other-ext'].forEach(function(id) {
    var btn = document.getElementById(id);
    if (!btn) return;
    btn.textContent = '🧩 Other anzeigen (' + count + ')';
    btn.style.display = count > 0 ? '' : 'none';
  });
}

drawPie('pie-ext','tip-ext','leg-ext', extData, CC, 180, 'ext');
drawPie('pie-conv','tip-conv','leg-conv', convData, ['#4ade80','#475569'], 180, 'conv');

// === UNIFIED FILTER BUTTONS ===
(function(){
  const c = document.getElementById('qf-container');
  const extList = document.getElementById('ext-browser-list');
  c.innerHTML = '';

  // Reset Button
  const btnReset = document.createElement('button');
  btnReset.className = 'qf-btn qf-btn-accent';
  btnReset.textContent = '\u2716 Filter zur\u00fccksetzen';
  btnReset.onclick = clearAllFilters;
  c.appendChild(btnReset);

  // Conv Button
  const btnConv = document.createElement('button');
  btnConv.className = 'qf-btn qf-btn-accent'; 
  btnConv.id = 'btn-conv-yes';
  btnConv.textContent = '\u2705 Nur Konvertierbare';
  btnConv.onclick = () => toggleConv('Ja');
  c.appendChild(btnConv);

  const btnConvNo = document.createElement('button');
  btnConvNo.className = 'qf-btn qf-btn-accent';
  btnConvNo.id = 'btn-conv-no';
  btnConvNo.textContent = '\u274c Nur Nicht-Konvertierbare';
  btnConvNo.onclick = () => toggleConv('Nein');
  c.appendChild(btnConvNo);

  // Pipeline filter buttons (only if pipeline paths were provided)
  if (HAS_PIPELINE) {
    var sepP = document.createElement('span');
    sepP.className = 'filter-bar-sep';
    c.appendChild(sepP);

    var btnOffen = document.createElement('button');
    btnOffen.className = 'qf-btn';
    btnOffen.id = 'btn-pipe-offen';
    btnOffen.textContent = '\u23F3 Offen (' + pipelineStats.offen + ')';
    btnOffen.onclick = function() { togglePipeline('offen'); };
    c.appendChild(btnOffen);

    var btnKopiert = document.createElement('button');
    btnKopiert.className = 'qf-btn';
    btnKopiert.id = 'btn-pipe-kopiert';
    btnKopiert.textContent = '\ud83d\udce6 Kopiert (' + pipelineStats.kopiert + ')';
    btnKopiert.onclick = function() { togglePipeline('kopiert'); };
    c.appendChild(btnKopiert);

    var btnVeredelt = document.createElement('button');
    btnVeredelt.className = 'qf-btn';
    btnVeredelt.id = 'btn-pipe-veredelt';
    btnVeredelt.textContent = '\u2705 Veredelt (' + pipelineStats.veredelt + ')';
    btnVeredelt.onclick = function() { togglePipeline('veredelt'); };
    c.appendChild(btnVeredelt);
  }

  // Separator
  var sep = document.createElement('span');
  sep.className = 'filter-bar-sep';
  c.appendChild(sep);

  // Top ext buttons with color dots
  const topExts = allExtStats.slice(0, 10);
  topExts.forEach((d, i) => {
    const b = document.createElement('button');
    b.className = 'qf-btn';
    b.dataset.ext = d.label;
    var dot = '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + CC[i % CC.length] + ';margin-right:4px;vertical-align:middle"></span>';
    b.innerHTML = dot + d.label + ' (' + d.count + ')';
    b.addEventListener('click', () => toggleExt(d.label));
    c.appendChild(b);
  });

  // Separator before Other/Alle
  var sep2 = document.createElement('span');
  sep2.className = 'filter-bar-sep';
  c.appendChild(sep2);

  const btnAllExt = document.createElement('button');
  btnAllExt.className = 'qf-btn';
  btnAllExt.textContent = '\ud83e\udde9 Alle Endungen';
  btnAllExt.onclick = () => toggleExtBrowser(true, 'all');
  c.appendChild(btnAllExt);

  const btnOtherExt = document.createElement('button');
  btnOtherExt.className = 'qf-btn';
  btnOtherExt.id = 'btn-show-other-ext';
  btnOtherExt.textContent = '\ud83e\udde9 Other anzeigen';
  btnOtherExt.onclick = () => openOtherExtensions();
  c.appendChild(btnOtherExt);

  // Ext browser list (all extensions)
  if (extList) {
    extList.innerHTML = '';
    allExtStats.forEach((d, i) => {
      const b = document.createElement('button');
      b.className = 'qf-btn';
      b.dataset.ext = d.label;
      var dotColor = i < 10 ? CC[i % CC.length] : '#888';
      var dot = '<span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:' + dotColor + ';margin-right:4px;vertical-align:middle"></span>';
      b.innerHTML = dot + d.label + ' (' + d.count + ')';
      b.addEventListener('click', () => toggleExt(d.label));
      extList.appendChild(b);
    });
  }

  updateOtherExtButton();
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
  order: [[1, 'asc']],
  layout: {
    topStart: 'buttons',
    topEnd: 'search',
    bottomStart: ['pageLength', 'info'],
    bottomEnd: 'paging'
  },
  buttons: [
    'copy',
    { extend: 'colvis', text: 'Spalten' },
    { extend: 'csvHtml5', title: 'FolderScan', exportOptions: { columns: ':visible:not(.no-export)', orthogonal: 'export' } },
    { extend: 'excelHtml5', title: 'FolderScan', exportOptions: { columns: ':visible:not(.no-export)', orthogonal: 'export' } },
    { extend: 'print', title: 'Folder Scan', exportOptions: { columns: ':visible:not(.no-export)' } }
  ],
  columns: [
    {
      data: 'ext',
      title: '',
      orderable: false,
      searchable: false,
      className: 'no-export',
      width: '28px',
      render: function(data) {
        var icon = '📄';
        var ext = data ? data.toLowerCase() : '';
        if (['.jpg','.png','.gif','.svg','.webp','.ico','.bmp','.tiff'].includes(ext)) icon = '🖼️';
        else if (['.mp4','.mov','.avi','.mkv','.wmv'].includes(ext)) icon = '🎬';
        else if (['.mp3','.wav','.flac','.ogg','.aac'].includes(ext)) icon = '🎵';
        else if (['.zip','.rar','.7z','.tar','.gz','.bz2'].includes(ext)) icon = '📦';
        else if (['.exe','.msi','.bat','.ps1','.sh','.cmd'].includes(ext)) icon = '⚙️';
        else if (['.pdf'].includes(ext)) icon = '📕';
        else if (['.docx','.doc','.txt','.md','.rtf','.odt'].includes(ext)) icon = '📝';
        else if (['.xlsx','.xls','.csv','.ods'].includes(ext)) icon = '📊';
        else if (['.pptx','.ppt','.odp'].includes(ext)) icon = '📽️';
        else if (['.html','.css','.js','.json','.xml','.yaml','.yml','.ts','.jsx','.tsx'].includes(ext)) icon = '💻';
        return '<span style="font-size:1rem">'+icon+'</span>';
      }
    },
    {
      data: 'name',
      title: 'Name',
      render: function(data, type) { 
        if (type !== 'display') return data;
        return '<span class="cell-name" title="'+escH(data)+'">'+escH(data)+'</span>'; 
      }
    },
    {
      data: 'ext',
      title: 'Typ',
      render: function(data, type) { if (type !== 'display') return data; return '<span class="cell-ext">'+escH(data)+'</span>'; }
    },
    {
      data: 'dir',
      title: 'Verzeichnis',
      render: function(data, type) {
        const cleanPath = data.replace(/\\\\/g, '\\');
        if (type !== 'display') return cleanPath;
        return '<div class="path-cell-inner" title="'+escH(cleanPath)+'">'+escH(cleanPath)+'</div>';
      }
    },
    {
      data: 'fullPath',
      title: 'Vollst. Pfad',
      visible: false,
      render: function(data, type) {
        const cleanPath = data.replace(/\\\\/g, '\\');
        if (type !== 'display') return cleanPath;
        return '<div class="path-cell-inner" title="'+escH(cleanPath)+'">'+escH(cleanPath)+'</div>';
      }
    },
    {
      data: 'sizeBytes',
      title: 'Größe',
      render: function(data, type) {
        if (type !== 'display') return data;
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
    { data: 'modified', title: 'Geändert', render: function(data, type) { if (type !== 'display') return data; return '<span style="white-space:nowrap">' + escH(data) + '</span>'; } },
    
    { 
      data: 'readOnly', title: '🔒', visible: false,
      render: function(d) { return d ? '🔒' : ''; }
    },
    {
      data: 'convertible',
      title: 'Konv.',
      render: function(data, type) {
        if (type !== 'display') return data ? 'Ja' : 'Nein';
        return data ? '<span class="badge badge-yes">Ja</span>' : '<span class="badge badge-no">Nein</span>';
      }
    },
    {
      data: 'status',
      title: 'Status',
      render: function(data, type) {
        if (type !== 'display') return data || '';
        if (data === 'new') return '<span class="badge badge-new">🟢 Neu</span>';
        if (data === 'modified') return '<span class="badge badge-modified">🟡 Geändert</span>';
        return '';
      },
      visible: HAS_DELTA
    },
    {
      data: 'pipeline',
      title: 'Pipeline',
      render: function(data, type) {
        if (type !== 'display') return data || '';
        if (data === 'veredelt') return '<span class="badge badge-yes">Veredelt</span>';
        if (data === 'kopiert') return '<span class="badge badge-modified">Kopiert</span>';
        if (data === 'offen') return '<span class="badge badge-no">Offen</span>';
        return '';
      },
      visible: HAS_PIPELINE
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

window.fileTable = table;

function quickExportCSV() {
  try { table.button('.buttons-csv').trigger(); } catch(e) { console.error('CSV export failed:', e); }
}
function quickExportExcel() {
  try { table.button('.buttons-excel').trigger(); } catch(e) { console.error('Excel export failed:', e); }
}

function exportPipelineCSV() {
  var filtered = getFilteredFiles();
  if (filtered.length === 0) { alert('Keine Dateien im aktuellen Filter.'); return; }

  var cols = ['Name','Extension','FullPath','DirectoryName','SizeBytes','SizeKB','SizeMB','CreationTime','LastWriteTime','IsReadOnly','IsConvertible','Status','PipelineStatus'];
  var rows = [cols.join(',')];

  filtered.forEach(function(f) {
    var dir = f.dir ? f.dir.replace(/\\\\/g, '\\') : '';
    var fp = f.fullPath ? f.fullPath.replace(/\\\\/g, '\\') : '';
    var line = [
      '"' + (f.name || '').replace(/"/g, '""') + '"',
      '"' + (f.ext || '') + '"',
      '"' + fp.replace(/"/g, '""') + '"',
      '"' + dir.replace(/"/g, '""') + '"',
      f.sizeBytes || 0,
      f.sizeKB || 0,
      f.sizeMB || 0,
      '"' + (f.created || '') + '"',
      '"' + (f.modified || '') + '"',
      f.readOnly ? 'True' : 'False',
      f.convertible ? 'True' : 'False',
      '"' + (f.status || 'unchanged') + '"',
      '"' + (f.pipeline || '') + '"'
    ];
    rows.push(line.join(','));
  });

  // Generate timestamp filename
  var now = new Date();
  var ts = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0') + '-' + String(now.getDate()).padStart(2,'0') + '_' + String(now.getHours()).padStart(2,'0') + String(now.getMinutes()).padStart(2,'0');
  var fileName = 'Pipeline_' + ts + '.csv';

  var csv = '\uFEFF' + rows.join('\r\n');
  var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  // Track exported CSV and unlock step 4
  lastPipelineCsvName = fileName;
  var scanResultsDir = csvPath.replace(/[^\\]+$/, '');
  lastPipelineCsvFullPath = scanResultsDir + fileName;

  var savedEl = document.getElementById('wf-csv-saved');
  if (savedEl) { savedEl.textContent = 'Download gestartet: ' + fileName; savedEl.style.opacity = '1'; }

  // Show move command
  var moveSection = document.getElementById('wf-move-section');
  var moveCmd = document.getElementById('wf-move-cmd');
  if (moveSection && moveCmd) {
    moveCmd.textContent = 'Move-Item "' + downloadsPath + '\\' + fileName + '" "' + scanResultsDir.slice(0, -1) + '"';
    moveSection.style.display = 'block';
  }

  // Show review command (Step 3)
  var reviewSection = document.getElementById('wf-review-section');
  var reviewCmd = document.getElementById('wf-review-cmd');
  var reviewPlaceholder = document.getElementById('wf-review-placeholder');
  if (reviewSection && reviewCmd) {
    reviewCmd.textContent = 'Import-Csv "' + lastPipelineCsvFullPath + '" | Out-GridView -Title "Arbeitsliste pruefen"';
    reviewSection.style.display = 'block';
    if (reviewPlaceholder) reviewPlaceholder.style.display = 'none';
  }

  // Activate steps 3-5
  ['wf-step3-num','wf-step4-num','wf-step5-num'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.style.background = 'var(--accent)';
  });

  // Unlock copy command
  var cmdEl = document.getElementById('cmd-output');
  if (cmdEl) { cmdEl.style.opacity = '1'; cmdEl.title = 'Klicken zum Kopieren'; }
  var btnCopy = document.getElementById('btn-copy-cmd');
  if (btnCopy) btnCopy.disabled = false;

  updateCopyCommand();
}

// === WORKFLOW STATE ===
var lastPipelineCsvName = null;
var lastPipelineCsvFullPath = null;

function copyMoveCmd() {
  var el = document.getElementById('wf-move-cmd');
  if (!el) return;
  navigator.clipboard.writeText(el.textContent).then(function() {
    var msg = document.getElementById('wf-move-copied');
    if (msg) { msg.textContent = 'In Zwischenablage kopiert!'; msg.style.opacity = '1'; setTimeout(function() { msg.style.opacity = '0'; }, 2000); }
  });
}

function copyReviewCmd() {
  var el = document.getElementById('wf-review-cmd');
  if (!el) return;
  navigator.clipboard.writeText(el.textContent).then(function() {
    var msg = document.getElementById('wf-review-copied');
    if (msg) { msg.textContent = 'In Zwischenablage kopiert!'; msg.style.opacity = '1'; setTimeout(function() { msg.style.opacity = '0'; }, 2000); }
  });
}

function updateRescanCommand() {
  var staging = (document.getElementById('cmd-dest') || {}).value || '';
  var result = (document.getElementById('dl-output') || {}).value || '';
  var cmd = 'Invoke-FolderScan -Path "' + scanRoot + '" -Recurse';
  if (staging) cmd += ' -StagingPath "' + staging + '"';
  if (result) cmd += ' -ErgebnisPath "' + result + '"';
  var el = document.getElementById('cmd-rescan');
  if (el) el.textContent = cmd;
}

function copyRescanCmd() {
  var el = document.getElementById('cmd-rescan');
  if (!el) return;
  navigator.clipboard.writeText(el.textContent).then(function() {
    var msg = document.getElementById('wf-rescan-copied');
    if (msg) { msg.textContent = 'In Zwischenablage kopiert!'; msg.style.opacity = '1'; setTimeout(function() { msg.style.opacity = '0'; }, 2000); }
  });
}

function updateCopyCommand() {
  if (!lastPipelineCsvFullPath) return;
  var dest = document.getElementById('cmd-dest').value || 'D:\\Staging';
  var whatIf = document.getElementById('cmd-whatif').checked;
  var scriptDir = scriptRoot.replace(/[\\\/]+$/, '') + '\\';

  var cmd = '. "' + scriptDir + 'Copy-ScannedFiles.ps1"\n';
  cmd += 'Copy-ScannedFiles -CsvPath "' + lastPipelineCsvFullPath + '" -Destination "' + dest + '"';
  if (whatIf) cmd += ' -WhatIf';

  var el = document.getElementById('cmd-output');
  if (el) el.textContent = cmd;
}

function copyCmdToClipboard() {
  if (!lastPipelineCsvFullPath) {
    alert('Bitte erst Arbeitsliste exportieren (Schritt 2).');
    return;
  }
  var el = document.getElementById('cmd-output');
  if (!el) return;
  navigator.clipboard.writeText(el.textContent).then(function() {
    var msg = document.getElementById('cmd-copied');
    if (msg) {
      msg.textContent = 'In Zwischenablage kopiert!';
      msg.style.opacity = '1';
      setTimeout(function() { msg.style.opacity = '0'; }, 2000);
    }
  });
}

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

// === CHART SWITCHER (Treemap/Sunburst) ===
var currentChartType = 'treemap'; // 'treemap' or 'sunburst'

function switchChartType(type) {
  currentChartType = type;
  document.getElementById('btn-chart-treemap').classList.toggle('active', type === 'treemap');
  document.getElementById('btn-chart-sunburst').classList.toggle('active', type === 'sunburst');
  drawTreemap(currentTreemapMetric); // Re-draws with current type
}

function drawTreemap(metric) {
  const container = document.getElementById('treemap-container');
  if (!container) return;

  if (typeof Plotly === 'undefined') {
    container.innerHTML = '<div style="padding:2rem;text-align:center;opacity:0.6">Charts need internet (Plotly CDN).</div>';
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
  
  // Colorscale logic
  var colorscale = metric === 'size' ? 'Blues' : 'Greens';
  if (currentChartType === 'sunburst') colorscale = 'Viridis';

  const trace = {
    type: currentChartType,
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
      colorscale: colorscale,
      showscale: false,
      line: { width: lineWidths, color: lineColors }
    },
    pathbar: { visible: true, thickness: 36, textfont: { size: 14 } }
  };
  
  const layout = {
    margin: { t: 44, l: 0, r: 0, b: 0 },
    paper_bgcolor: 'transparent',
    font: { family: fontFamily, color: textColor }
  };

  Plotly.newPlot('treemap-container', [trace], layout).then(function() {
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
    table.column(3).search('');
  } else {
    table.column(3).search(query);
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

// === GLOBAL FOLDER SELECTION ===
var gfSelectedFolders = new Set();
var gfAllSelected = true;

function gfBuildTree() {
  if (!dlTreeData) return;
  var container = document.getElementById('gf-folder-tree');
  if (!container) return;
  container.innerHTML = '';
  container.appendChild(gfCreateNode(scanRoot, true));
  gfSelectAll();
}

function gfCreateNode(path, isRoot) {
  var node = dlTreeData[path];
  if (!node) return document.createDocumentFragment();
  var childKeys = Object.keys(node.children).sort();
  var hasChildren = childKeys.length > 0;

  var div = document.createElement('div');
  div.className = 'ft-node' + (isRoot ? ' ft-node-root' : '');
  div.dataset.path = path;

  var row = document.createElement('div');
  row.className = 'ft-row';

  var toggle = document.createElement('span');
  toggle.className = 'ft-toggle' + (hasChildren ? '' : ' empty');
  toggle.textContent = hasChildren ? '\u25B6' : '';
  toggle.onclick = function(e) { e.stopPropagation(); gfToggleExpand(div); };
  row.appendChild(toggle);

  var cb = document.createElement('input');
  cb.type = 'checkbox';
  cb.className = 'gf-cb';
  cb.dataset.path = path;
  cb.checked = true;
  cb.onchange = function() { gfOnCheck(path, cb.checked); };
  row.appendChild(cb);

  var label = document.createElement('span');
  label.className = 'ft-label';
  label.textContent = node.name;
  label.onclick = function() { if (hasChildren) gfToggleExpand(div); };
  row.appendChild(label);

  var count = document.createElement('span');
  count.className = 'ft-count';
  count.textContent = node.totalFiles + ' Dateien \u00B7 ' + fmtSize(node.totalSize);
  row.appendChild(count);

  div.appendChild(row);

  if (hasChildren) {
    var childDiv = document.createElement('div');
    childDiv.className = 'ft-children';
    childDiv.dataset.loaded = '0';
    div.appendChild(childDiv);
  }
  return div;
}

function gfToggleExpand(nodeDiv) {
  var childDiv = nodeDiv.querySelector(':scope > .ft-children');
  if (!childDiv) return;
  var isExpanded = childDiv.classList.contains('expanded');
  var toggle = nodeDiv.querySelector(':scope > .ft-row > .ft-toggle');

  if (isExpanded) {
    childDiv.classList.remove('expanded');
    if (toggle) toggle.textContent = '\u25B6';
  } else {
    if (childDiv.dataset.loaded === '0') {
      var path = nodeDiv.dataset.path;
      var node = dlTreeData[path];
      if (node) {
        Object.keys(node.children).sort().forEach(function(ck) {
          childDiv.appendChild(gfCreateNode(ck, false));
        });
      }
      childDiv.dataset.loaded = '1';
      gfApplyCheckState(childDiv);
    }
    childDiv.classList.add('expanded');
    if (toggle) toggle.textContent = '\u25BC';
  }
}

function gfOnCheck(path, checked) {
  function setDesc(p, val) {
    if (val) gfSelectedFolders.add(p); else gfSelectedFolders.delete(p);
    var n = dlTreeData[p];
    if (n) Object.keys(n.children).forEach(function(ck) { setDesc(ck, val); });
  }
  setDesc(path, checked);
  gfBubbleUp(path);
  gfRefreshAllCbs();
  gfUpdateSummary();
  gfClearBtnActive();
  applyFilters();
}

function gfBubbleUp(path) {
  var current = path;
  while (true) {
    var idx = current.lastIndexOf('\\');
    if (idx <= 0) break;
    var parent = current.substring(0, idx);
    if (!dlTreeData[parent]) break;
    var childKeys = Object.keys(dlTreeData[parent].children);
    var allChecked = childKeys.every(function(ck) { return gfIsFullyChecked(ck); });
    if (allChecked) gfSelectedFolders.add(parent); else gfSelectedFolders.delete(parent);
    current = parent;
  }
}

function gfIsFullyChecked(path) {
  if (!gfSelectedFolders.has(path)) return false;
  var node = dlTreeData[path];
  if (!node) return false;
  return Object.keys(node.children).every(function(ck) { return gfIsFullyChecked(ck); });
}

function gfIsIndet(path) {
  var node = dlTreeData[path];
  if (!node) return false;
  var ck = Object.keys(node.children);
  if (ck.length === 0) return false;
  var hasC = false, hasU = false;
  function walk(p) {
    if (gfSelectedFolders.has(p)) hasC = true; else hasU = true;
    var n = dlTreeData[p];
    if (n) Object.keys(n.children).forEach(walk);
  }
  ck.forEach(walk);
  return hasC && hasU;
}

function gfRefreshAllCbs() {
  document.querySelectorAll('#gf-folder-tree .gf-cb').forEach(function(cb) {
    var p = cb.dataset.path;
    cb.checked = gfSelectedFolders.has(p);
    cb.indeterminate = gfIsIndet(p);
  });
}

function gfApplyCheckState(container) {
  container.querySelectorAll('.gf-cb').forEach(function(cb) {
    var p = cb.dataset.path;
    cb.checked = gfSelectedFolders.has(p);
    cb.indeterminate = gfIsIndet(p);
  });
}

function gfClearBtnActive() {
  ['gf-btn-all','gf-btn-none','gf-btn-conv'].forEach(function(id) {
    var b = document.getElementById(id);
    if (b) b.classList.remove('active');
  });
  gfAllSelected = false;
}

function gfSelectAll() {
  if (!dlTreeData) return;
  Object.keys(dlTreeData).forEach(function(p) { gfSelectedFolders.add(p); });
  gfRefreshAllCbs();
  gfUpdateSummary();
  gfAllSelected = true;
  var b = document.getElementById('gf-btn-all');
  if (b) { gfClearBtnActive(); b.classList.add('active'); }
  gfAllSelected = true;
  applyFilters();
}

function gfSelectNone() {
  gfSelectedFolders.clear();
  gfRefreshAllCbs();
  gfUpdateSummary();
  gfClearBtnActive();
  var b = document.getElementById('gf-btn-none');
  if (b) b.classList.add('active');
  gfAllSelected = false;
  applyFilters();
}

function gfSelectConvertible() {
  if (!dlTreeData) return;
  gfSelectedFolders.clear();
  Object.keys(dlTreeData).forEach(function(p) {
    if (dlTreeData[p].totalConv > 0) gfSelectedFolders.add(p);
  });
  gfRefreshAllCbs();
  gfUpdateSummary();
  gfClearBtnActive();
  var b = document.getElementById('gf-btn-conv');
  if (b) b.classList.add('active');
  gfAllSelected = false;
  applyFilters();
}

function gfFilterTree(query) {
  query = (query || '').toLowerCase();
  document.querySelectorAll('#gf-folder-tree .ft-node').forEach(function(node) {
    var path = (node.dataset.path || '').toLowerCase();
    var label = node.querySelector('.ft-label');
    var name = label ? label.textContent.toLowerCase() : '';
    node.style.display = (!query || path.indexOf(query) >= 0 || name.indexOf(query) >= 0) ? '' : 'none';
  });
}

function gfUpdateSummary() {
  if (!dlTreeData) return;
  var totalFiles = 0, totalSize = 0, folderCount = 0;
  var allTotal = allFiles ? allFiles.length : 0;

  gfSelectedFolders.forEach(function(p) {
    var node = dlTreeData[p];
    if (node) { totalFiles += node.fileCount; totalSize += node.sizeBytes; folderCount++; }
  });

  var pct = allTotal > 0 ? Math.round((totalFiles / allTotal) * 100) : 0;
  var elF = document.getElementById('gf-sum-folders');
  var elFi = document.getElementById('gf-sum-files');
  var elS = document.getElementById('gf-sum-size');
  var elP = document.getElementById('gf-sum-pct');

  if (elF) elF.textContent = folderCount;
  if (elFi) elFi.textContent = totalFiles;
  if (elS) elS.textContent = fmtSize(totalSize);
  if (elP) elP.textContent = pct;

  var badge = document.getElementById('gf-summary-badge');
  if (badge) {
    if (gfAllSelected || folderCount === Object.keys(dlTreeData).length) {
      badge.textContent = '(alle Ordner)';
    } else if (folderCount === 0) {
      badge.textContent = '(keine Ordner ausgew\u00E4hlt)';
    } else {
      badge.textContent = '(' + folderCount + ' Ordner, ' + totalFiles + ' Dateien)';
    }
  }
}

function gfIsFileIncluded(f) {
  if (gfAllSelected) return true;
  if (gfSelectedFolders.size === 0) return false;
  var dir = f.dir ? f.dir.replace(/\\\\/g, '\\') : scanRoot;
  return gfSelectedFolders.has(dir);
}

// Register DataTables custom filter for global folder selection
if (typeof jQuery !== 'undefined' && jQuery.fn.dataTable) {
  jQuery.fn.dataTable.ext.search.push(function(settings, data, dataIndex) {
    if (gfAllSelected) return true;
    if (gfSelectedFolders.size === 0) return false;
    var dir = data[3] || '';
    return gfSelectedFolders.has(dir);
  });
}

// Arrow toggle for details panel
(function() {
  var panel = document.getElementById('gf-panel');
  var arrow = document.getElementById('gf-arrow');
  if (panel && arrow) {
    panel.addEventListener('toggle', function() {
      arrow.style.transform = panel.open ? 'rotate(90deg)' : 'rotate(0deg)';
    });
  }
})();

// === DOCLING CONVERSION PANEL ===
var dlTreeData = null;
var dlSelectedFolders = new Set();
var dlAvgPerFile = $doclingAvgPerFile;

function dlBuildTree() {
  var root = scanRoot;
  var folders = {};
  folders[root] = { path: root, name: root.split('\\').pop() || root, children: {}, fileCount: 0, totalFiles: 0, sizeBytes: 0, totalSize: 0, convCount: 0, totalConv: 0 };

  allFiles.forEach(function(f) {
    var dir = f.dir ? f.dir.replace(/\\\\/g, '\\') : root;
    if (!folders[dir]) {
      var rel = dir.substring(root.length).replace(/^\\/, '');
      var parts = rel.split('\\');
      var current = root;
      for (var i = 0; i < parts.length; i++) {
        var next = current + '\\' + parts[i];
        if (!folders[next]) {
          folders[next] = { path: next, name: parts[i], children: {}, fileCount: 0, totalFiles: 0, sizeBytes: 0, totalSize: 0, convCount: 0, totalConv: 0 };
          folders[current].children[next] = true;
        }
        current = next;
      }
    }
    folders[dir].fileCount++;
    folders[dir].sizeBytes += f.sizeBytes;
    if (f.convertible) folders[dir].convCount++;
  });

  function propagate(path) {
    var node = folders[path];
    node.totalFiles = node.fileCount;
    node.totalSize = node.sizeBytes;
    node.totalConv = node.convCount;
    Object.keys(node.children).forEach(function(ck) {
      propagate(ck);
      node.totalFiles += folders[ck].totalFiles;
      node.totalSize += folders[ck].totalSize;
      node.totalConv += folders[ck].totalConv;
    });
  }
  propagate(root);
  dlTreeData = folders;
}

function dlRenderTree() {
  var container = document.getElementById('dl-folder-tree');
  if (!container || !dlTreeData) return;
  container.innerHTML = '';
  container.appendChild(dlCreateNode(scanRoot, true));
}

function dlCreateNode(path, isRoot) {
  var node = dlTreeData[path];
  if (!node) return document.createDocumentFragment();
  var childKeys = Object.keys(node.children).sort();
  var hasChildren = childKeys.length > 0;

  var div = document.createElement('div');
  div.className = 'ft-node' + (isRoot ? ' ft-node-root' : '');
  div.dataset.path = path;

  var row = document.createElement('div');
  row.className = 'ft-row';

  var toggle = document.createElement('span');
  toggle.className = 'ft-toggle' + (hasChildren ? '' : ' empty');
  toggle.textContent = hasChildren ? '\u25B6' : '';
  toggle.onclick = function(e) { e.stopPropagation(); dlToggleExpand(div); };
  row.appendChild(toggle);

  var cb = document.createElement('input');
  cb.type = 'checkbox';
  cb.className = 'ft-cb';
  cb.dataset.path = path;
  cb.onchange = function() { dlOnCheck(path, cb.checked); };
  row.appendChild(cb);

  var label = document.createElement('span');
  label.className = 'ft-label';
  label.textContent = node.name;
  label.onclick = function() { if (hasChildren) dlToggleExpand(div); };
  row.appendChild(label);

  var count = document.createElement('span');
  count.className = 'ft-count';
  count.textContent = node.totalFiles + ' Dateien \u00B7 ' + fmtSize(node.totalSize);
  row.appendChild(count);

  div.appendChild(row);

  if (hasChildren) {
    var childDiv = document.createElement('div');
    childDiv.className = 'ft-children';
    childDiv.dataset.loaded = '0';
    div.appendChild(childDiv);
  }
  return div;
}

function dlToggleExpand(nodeDiv) {
  var childDiv = nodeDiv.querySelector(':scope > .ft-children');
  if (!childDiv) return;
  var isExpanded = childDiv.classList.contains('expanded');
  var toggle = nodeDiv.querySelector(':scope > .ft-row > .ft-toggle');

  if (isExpanded) {
    childDiv.classList.remove('expanded');
    if (toggle) toggle.textContent = '\u25B6';
  } else {
    if (childDiv.dataset.loaded === '0') {
      var path = nodeDiv.dataset.path;
      var node = dlTreeData[path];
      if (node) {
        Object.keys(node.children).sort().forEach(function(ck) {
          childDiv.appendChild(dlCreateNode(ck, false));
        });
      }
      childDiv.dataset.loaded = '1';
      dlApplyCheckState(childDiv);
    }
    childDiv.classList.add('expanded');
    if (toggle) toggle.textContent = '\u25BC';
  }
}

function dlOnCheck(path, checked) {
  function setDesc(p, val) {
    if (val) dlSelectedFolders.add(p); else dlSelectedFolders.delete(p);
    var n = dlTreeData[p];
    if (n) Object.keys(n.children).forEach(function(ck) { setDesc(ck, val); });
  }
  setDesc(path, checked);
  dlUpdateParents(path);
  dlRefreshAllCbs();
  dlUpdateSummary();
  dlBuildCommand();
  dlClearBtnActive();
}

function dlIsIndet(path) {
  var node = dlTreeData[path];
  if (!node) return false;
  var ck = Object.keys(node.children);
  if (ck.length === 0) return false;
  var hasC = false, hasU = false;
  function walk(p) {
    if (dlSelectedFolders.has(p)) hasC = true; else hasU = true;
    var n = dlTreeData[p];
    if (n) Object.keys(n.children).forEach(walk);
  }
  ck.forEach(walk);
  return hasC && hasU;
}

function dlIsFullyChecked(path) {
  if (!dlSelectedFolders.has(path)) return false;
  var node = dlTreeData[path];
  if (!node) return false;
  return Object.keys(node.children).every(function(ck) { return dlIsFullyChecked(ck); });
}

function dlUpdateParents(path) {
  var root = scanRoot;
  var current = path;
  while (current !== root) {
    var idx = current.lastIndexOf('\\');
    if (idx <= 0) break;
    var parent = current.substring(0, idx);
    if (!dlTreeData[parent]) break;
    var childKeys = Object.keys(dlTreeData[parent].children);
    var allChecked = childKeys.every(function(ck) { return dlIsFullyChecked(ck); });
    if (allChecked) dlSelectedFolders.add(parent); else dlSelectedFolders.delete(parent);
    current = parent;
  }
}

function dlRefreshAllCbs() {
  document.querySelectorAll('#dl-folder-tree .ft-cb').forEach(function(cb) {
    var p = cb.dataset.path;
    cb.checked = dlSelectedFolders.has(p);
    cb.indeterminate = dlIsIndet(p);
  });
}

function dlApplyCheckState(container) {
  container.querySelectorAll('.ft-cb').forEach(function(cb) {
    var p = cb.dataset.path;
    cb.checked = dlSelectedFolders.has(p);
    cb.indeterminate = dlIsIndet(p);
  });
}

function dlClearBtnActive() {
  ['dl-btn-all','dl-btn-none','dl-btn-conv'].forEach(function(id) {
    var b = document.getElementById(id);
    if (b) b.classList.remove('active');
  });
}

function dlSetBtnActive(id) {
  dlClearBtnActive();
  var b = document.getElementById(id);
  if (b) b.classList.add('active');
}

function dlSelectAll() {
  Object.keys(dlTreeData).forEach(function(p) { dlSelectedFolders.add(p); });
  dlRefreshAllCbs();
  dlUpdateSummary();
  dlBuildCommand();
  dlSetBtnActive('dl-btn-all');
}

function dlSelectNone() {
  dlSelectedFolders.clear();
  dlRefreshAllCbs();
  dlUpdateSummary();
  dlBuildCommand();
  dlSetBtnActive('dl-btn-none');
}

function dlSelectConvertible() {
  dlSelectedFolders.clear();
  Object.keys(dlTreeData).forEach(function(p) {
    if (dlTreeData[p].totalConv > 0) dlSelectedFolders.add(p);
  });
  dlRefreshAllCbs();
  dlUpdateSummary();
  dlBuildCommand();
  dlSetBtnActive('dl-btn-conv');
}

function dlFilterTree(query) {
  query = (query || '').toLowerCase();
  document.querySelectorAll('#dl-folder-tree .ft-node').forEach(function(node) {
    var path = (node.dataset.path || '').toLowerCase();
    var label = node.querySelector('.ft-label');
    var name = label ? label.textContent.toLowerCase() : '';
    node.style.display = (!query || path.indexOf(query) >= 0 || name.indexOf(query) >= 0) ? '' : 'none';
  });
}

function dlUpdateSummary() {
  var totalFiles = 0, totalSize = 0, folderCount = 0;
  var allTotal = dlTreeData && dlTreeData[scanRoot] ? dlTreeData[scanRoot].totalFiles : (allFiles ? allFiles.length : 0);

  dlSelectedFolders.forEach(function(p) {
    var node = dlTreeData[p];
    if (node) { totalFiles += node.fileCount; totalSize += node.sizeBytes; folderCount++; }
  });

  var pct = allTotal > 0 ? Math.round((totalFiles / allTotal) * 100) : 0;
  var elF = document.getElementById('dl-sum-folders');
  var elFi = document.getElementById('dl-sum-files');
  var elP = document.getElementById('dl-sum-pct');
  var elS = document.getElementById('dl-sum-size');
  var elE = document.getElementById('dl-sum-eta');

  if (elF) elF.textContent = folderCount;
  if (elFi) elFi.textContent = totalFiles;
  if (elP) elP.textContent = pct;
  if (elS) elS.textContent = fmtSize(totalSize);

  if (elE) {
    if (dlAvgPerFile > 0 && totalFiles > 0) {
      var secs = Math.round(totalFiles * dlAvgPerFile);
      var h = Math.floor(secs / 3600), m = Math.floor((secs % 3600) / 60);
      elE.textContent = (h > 0 ? h + 'h ' : '') + m + 'min (\u00D8 ' + dlAvgPerFile.toFixed(1) + 's/Datei)';
    } else if (totalFiles > 0) {
      var secs2 = totalFiles;
      var h2 = Math.floor(secs2 / 3600), m2 = Math.floor((secs2 % 3600) / 60);
      elE.textContent = '~' + (h2 > 0 ? h2 + 'h ' : '') + m2 + 'min (Sch\u00E4tzung)';
    } else {
      elE.textContent = '\u2013';
    }
  }
}

function dlGetSelectedTopFolders() {
  var root = scanRoot;
  var result = [];
  function walk(path) {
    if (!dlSelectedFolders.has(path) && !dlIsIndet(path)) {
      return;
    }
    if (dlIsFullyChecked(path)) {
      if (path !== root) {
        result.push(path.substring(root.length).replace(/^\\/, ''));
      }
      return;
    }
    var node = dlTreeData[path];
    if (node) Object.keys(node.children).forEach(walk);
  }
  walk(root);
  if (dlIsFullyChecked(root)) return [];
  return result;
}

function dlBuildCommand() {
  var url = (document.getElementById('dl-url') || {}).value || '';
  var input = (document.getElementById('dl-input') || {}).value || '';
  var output = (document.getElementById('dl-output') || {}).value || '';
  var el = document.getElementById('dl-cmd');

  if (!url || !input || !output) {
    if (el) el.textContent = 'Bitte Docling URL, Input- und Output-Ordner angeben.';
    return;
  }

  var fmts = [];
  document.querySelectorAll('input[name="dl-fmt"]:checked').forEach(function(cb) { fmts.push(cb.value); });
  if (fmts.length === 0) fmts = ['markdown'];

  var pipeline = (document.querySelector('input[name="dl-pipeline"]:checked') || {}).value || 'standard';
  var ocrEng = (document.querySelector('input[name="dl-ocreng"]:checked') || {}).value || 'easyocr';
  var pdfBe = (document.querySelector('input[name="dl-pdfbe"]:checked') || {}).value || 'dlparse_v4';
  var tblMode = (document.querySelector('input[name="dl-tblmode"]:checked') || {}).value || 'accurate';
  var imgMode = (document.querySelector('input[name="dl-imgmode"]:checked') || {}).value || 'embedded';

  var ocr = document.getElementById('dl-ocr') && document.getElementById('dl-ocr').checked;
  var forceOcr = document.getElementById('dl-forceocr') && document.getElementById('dl-forceocr').checked;
  var skip = document.getElementById('dl-skip') && document.getElementById('dl-skip').checked;
  var abort = document.getElementById('dl-abort') && document.getElementById('dl-abort').checked;
  var retries = parseInt((document.getElementById('dl-retries') || {}).value || '3', 10);
  var timeout = parseInt((document.getElementById('dl-timeout') || {}).value || '900', 10);
  var codeE = document.getElementById('dl-code-enrich') && document.getElementById('dl-code-enrich').checked;
  var formulaE = document.getElementById('dl-formula-enrich') && document.getElementById('dl-formula-enrich').checked;
  var picC = document.getElementById('dl-pic-class') && document.getElementById('dl-pic-class').checked;
  var picD = document.getElementById('dl-pic-desc') && document.getElementById('dl-pic-desc').checked;

  var folders = dlGetSelectedTopFolders();

  var bt = '\x60';
  var cmd = '. "' + scriptRoot.replace(/[\\\/]+$/, '') + '\\Invoke-DoclingConversion.ps1"\n';
  cmd += 'Invoke-DoclingConversion -DoclingUrl "' + url + '" ' + bt + '\n';
  cmd += '  -InputPath "' + input + '" -OutputPath "' + output + '"';

  if (folders.length > 0) {
    cmd += ' ' + bt + '\n  -Folders ' + folders.map(function(f) { return '"' + f + '"'; }).join(',');
  }
  cmd += ' ' + bt + '\n  -ToFormats ' + fmts.join(',');

  var extras = '';
  if (pipeline !== 'standard') extras += ' -PipelineType ' + pipeline;
  if (ocr) extras += ' -EnableOcr';
  if (forceOcr) extras += ' -ForceOcr';
  if (ocr && ocrEng !== 'auto') extras += ' -OcrEngine ' + ocrEng;
  if (pdfBe !== 'dlparse_v4') extras += ' -PdfBackend ' + pdfBe;
  if (tblMode !== 'accurate') extras += ' -TableMode ' + tblMode;
  if (imgMode !== 'embedded') extras += ' -ImageExportMode ' + imgMode;
  if (abort) extras += ' -AbortOnError';
  if (codeE) extras += ' -EnableCodeEnrichment';
  if (formulaE) extras += ' -EnableFormulaEnrichment';
  if (picC) extras += ' -EnablePictureClassification';
  if (picD) extras += ' -EnablePictureDescription';
  if (!skip) extras += ' -Force';
  if (retries !== 3) extras += ' -RetryCount ' + retries;
  if (timeout !== 900) extras += ' -TimeoutSec ' + timeout;
  if (extras) cmd += ' ' + bt + '\n ' + extras.trim();

  if (el) el.textContent = cmd;
  dlSaveSettings();
}

function dlCopyCmd() {
  var el = document.getElementById('dl-cmd');
  if (!el) return;
  navigator.clipboard.writeText(el.textContent).then(function() {
    var msg = document.getElementById('dl-cmd-copied');
    if (msg) { msg.textContent = 'In Zwischenablage kopiert!'; msg.style.opacity = '1'; setTimeout(function() { msg.style.opacity = '0'; }, 2000); }
  });
}

var dlInputManual = false;
var dlOutputManual = false;

function dlDeriveResultPath(stagingPath) {
  if (!stagingPath) return '';
  var p = stagingPath.replace(/[\\\/]+$/, '');
  var idx = p.lastIndexOf('\\');
  if (idx < 0) idx = p.lastIndexOf('/');
  if (idx >= 0) return p.substring(0, idx + 1) + 'Result';
  return p + '\\Result';
}

function dlSyncFromStaging() {
  var staging = (document.getElementById('cmd-dest') || {}).value || '';
  if (!staging) return;
  var dlIn = document.getElementById('dl-input');
  var dlOut = document.getElementById('dl-output');
  if (dlIn && !dlInputManual) {
    dlIn.value = staging;
  }
  if (dlOut && !dlOutputManual) {
    var src = dlIn ? dlIn.value : staging;
    dlOut.value = dlDeriveResultPath(src);
  }
  dlSaveSettings();
  dlBuildCommand();
  updateRescanCommand();
}

function dlOnInputManual() {
  dlInputManual = true;
  var dlIn = document.getElementById('dl-input');
  var dlOut = document.getElementById('dl-output');
  if (dlIn && dlOut && !dlOutputManual) {
    dlOut.value = dlDeriveResultPath(dlIn.value);
  }
  updateRescanCommand();
}

function dlOnOutputManual() {
  dlOutputManual = true;
  updateRescanCommand();
}

function dlSaveSettings() {
  try {
    var s = {};
    // Text fields
    s.url = (document.getElementById('dl-url') || {}).value || '';
    s.input = (document.getElementById('dl-input') || {}).value || '';
    s.output = (document.getElementById('dl-output') || {}).value || '';
    s.staging = (document.getElementById('cmd-dest') || {}).value || '';
    s.retries = (document.getElementById('dl-retries') || {}).value || '3';
    s.timeout = (document.getElementById('dl-timeout') || {}).value || '900';
    // Checkboxes
    ['dl-ocr','dl-forceocr','dl-skip','dl-abort','dl-code-enrich','dl-formula-enrich','dl-pic-class','dl-pic-desc'].forEach(function(id) {
      var el = document.getElementById(id);
      if (el) s[id] = el.checked;
    });
    // Format checkboxes
    s.fmts = [];
    document.querySelectorAll('input[name="dl-fmt"]').forEach(function(cb) { if (cb.checked) s.fmts.push(cb.value); });
    // Radio groups
    ['dl-imgmode','dl-pipeline','dl-ocreng','dl-pdfbe','dl-tblmode'].forEach(function(name) {
      var el = document.querySelector('input[name="' + name + '"]:checked');
      if (el) s[name] = el.value;
    });
    // Override flags
    s.inputManual = dlInputManual;
    s.outputManual = dlOutputManual;
    localStorage.setItem('dl_settings', JSON.stringify(s));
  } catch(e) {}
}

function dlLoadSettings() {
  try {
    var raw = localStorage.getItem('dl_settings');
    if (!raw) return;
    var s = JSON.parse(raw);
    // Text fields
    if (s.url) { var e = document.getElementById('dl-url'); if (e) e.value = s.url; }
    if (s.staging) { var e2 = document.getElementById('cmd-dest'); if (e2) e2.value = s.staging; }
    if (s.input) { var e3 = document.getElementById('dl-input'); if (e3) e3.value = s.input; }
    if (s.output) { var e4 = document.getElementById('dl-output'); if (e4) e4.value = s.output; }
    if (s.retries) { var e5 = document.getElementById('dl-retries'); if (e5) e5.value = s.retries; }
    if (s.timeout) { var e6 = document.getElementById('dl-timeout'); if (e6) e6.value = s.timeout; }
    // Checkboxes
    ['dl-ocr','dl-forceocr','dl-skip','dl-abort','dl-code-enrich','dl-formula-enrich','dl-pic-class','dl-pic-desc'].forEach(function(id) {
      if (s.hasOwnProperty(id)) { var el = document.getElementById(id); if (el) el.checked = s[id]; }
    });
    // Format checkboxes
    if (s.fmts) {
      document.querySelectorAll('input[name="dl-fmt"]').forEach(function(cb) { cb.checked = s.fmts.indexOf(cb.value) >= 0; });
    }
    // Radio groups
    ['dl-imgmode','dl-pipeline','dl-ocreng','dl-pdfbe','dl-tblmode'].forEach(function(name) {
      if (s[name]) {
        var el = document.querySelector('input[name="' + name + '"][value="' + s[name] + '"]');
        if (el) el.checked = true;
      }
    });
    // Override flags
    if (s.inputManual) dlInputManual = true;
    if (s.outputManual) dlOutputManual = true;
  } catch(e) {}
}

function dlInit() {
  dlBuildTree();
  var fc = document.getElementById('dl-folder-count');
  if (fc && dlTreeData) {
    var cnt = Object.keys(dlTreeData).length;
    fc.textContent = '(' + cnt + ' Ordner gesamt)';
  }
  dlRenderTree();
  dlLoadSettings();
  // Initial sync: push Step 4 staging path into Docling fields if not manually set
  dlSyncFromStaging();
  dlBuildCommand();
  // Initialize global folder selection (reuses dlTreeData)
  gfBuildTree();
}

// Initial draw
jQuery(document).ready(function() {
  var activityBody = document.getElementById('activity-body');
  var activityArrow = document.getElementById('activity-arrow');
  var activitySection = document.getElementById('activity-section');
  if (activityBody && !activityBody.getAttribute('data-collapsed')) {
    activityBody.setAttribute('data-collapsed', '1');
  }
  if (activityBody && activityBody.getAttribute('data-collapsed') === '1') {
    activityBody.style.maxHeight = '0px';
    activityBody.style.opacity = '0';
    if (activityArrow) activityArrow.classList.add('collapsed');
    if (activitySection) activitySection.classList.remove('activity-expanded');
  }

  try {
    if (typeof updateVisuals === 'function') updateVisuals();
  } catch (e) { console.error('UpdateVisuals Error:', e); }

  try { buildTop10(); } catch(e) { console.error('Top10 Error:', e); }
  try { switchAnalysisPane('files'); } catch(e) {}
  try { checkOffline(); } catch(e) {}
  try { dlInit(); } catch(e) { console.error('Docling Init Error:', e); }
});

</script>
</body>
</html>
"@

  # ═══════════════════════════════════════════════════════════════
  # WRITE HTML & OPEN
  # ═══════════════════════════════════════════════════════════════
  $htmlPath = Join-Path $OutputDir "Dashboard.html"
  [System.IO.File]::WriteAllText($htmlPath, $htmlContent, [System.Text.UTF8Encoding]::new($true))
  Write-Host "[Invoke-FolderScan] Dashboard: $htmlPath" -ForegroundColor Green

  Start-Process $htmlPath
  Write-Host "[Invoke-FolderScan] Dashboard opened in browser." -ForegroundColor Cyan

  Write-Host "`n[Invoke-FolderScan] Done! $($FileList.Count) files. Results: $OutputDir" -ForegroundColor Green
  if (-not $PassThru) {
    Write-Host "[Invoke-FolderScan] Tip: Invoke-FolderScan ... -PassThru um Objekte zurueckzubekommen" -ForegroundColor DarkGray
  }

  if ($PassThru) { return $FileList }
}
