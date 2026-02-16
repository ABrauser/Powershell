<#
.SYNOPSIS
    Copies files listed in a FolderScan CSV to a target directory, preserving folder structure.

.DESCRIPTION
    Reads a FullScan.csv (or Pipeline CSV) exported by Invoke-FolderScan and copies the listed
    files to a destination folder. The relative folder structure from the original scan root is
    preserved under the destination.

    By default, files that already exist in the destination are SKIPPED (incremental copy).
    Use -Force to overwrite existing files.

    Supports filtering by convertibility and file extension, so you can e.g. copy only
    convertible files for downstream processing (Docling, Markdown conversion, etc.).

.PARAMETER CsvPath
    Path to the FullScan.csv or Pipeline CSV file.

.PARAMETER Destination
    Target folder where files will be copied to. Created automatically if it doesn't exist.
    The relative folder structure from the scan root is preserved.

.PARAMETER ScanRoot
    The original scan root path. Used to calculate relative paths.
    If omitted, the script tries to read it from FullScan.meta.json next to the CSV.

.PARAMETER OnlyConvertible
    If set, only files where IsConvertible = True are copied.

.PARAMETER Extensions
    Optional array of extensions to include (e.g. '.pdf', '.docx').
    If omitted, all extensions are included.

.PARAMETER Force
    Overwrite files that already exist in the destination.
    Without this flag, existing files are skipped.

.PARAMETER WhatIf
    Shows what would be copied without actually copying.

.EXAMPLE
    . .\Copy-ScannedFiles.ps1
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\Staging" -OnlyConvertible

.EXAMPLE
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\Staging" -Extensions '.pdf','.docx'

.EXAMPLE
    # Incremental: only new files are copied, existing are skipped
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\Staging" -OnlyConvertible

.EXAMPLE
    # Force overwrite all files
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\Staging" -OnlyConvertible -Force

.EXAMPLE
    # Dry run: see what would be copied
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\Staging" -OnlyConvertible -WhatIf
#>

function Copy-ScannedFiles {
  [CmdletBinding(SupportsShouldProcess)]
  param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [string]$Destination,

    [string]$ScanRoot,

    [switch]$OnlyConvertible,

    [string[]]$Extensions,

    [switch]$Force
  )

  # ── Resolve paths ──
  $CsvPath = (Resolve-Path $CsvPath).Path
  $csvDir = Split-Path $CsvPath -Parent

  # ── Try to detect ScanRoot from metadata ──
  if (-not $ScanRoot) {
    $metaPath = Join-Path $csvDir "FullScan.meta.json"
    if (Test-Path $metaPath) {
      try {
        $meta = Get-Content -Path $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($meta.ScanRoot) {
          $ScanRoot = $meta.ScanRoot
          Write-Host "[Copy-ScannedFiles] Scan root from metadata: $ScanRoot" -ForegroundColor Cyan
        }
      }
      catch {
        Write-Warning "[Copy-ScannedFiles] Could not read metadata file: $_"
      }
    }
  }

  if (-not $ScanRoot) {
    Write-Error "[Copy-ScannedFiles] Could not determine scan root. Please provide -ScanRoot parameter."
    return
  }

  $ScanRoot = $ScanRoot.TrimEnd('\')

  # ── Safety check: refuse FullScan.csv (raw inventory, not a curated work list) ──
  $csvFileName = [System.IO.Path]::GetFileName($CsvPath)
  if ($csvFileName -eq 'FullScan.csv') {
    Write-Error @"
[Copy-ScannedFiles] FullScan.csv ist ein Inventar aller Dateien und darf nicht direkt als Kopierliste verwendet werden.

Workflow:
  1. Dashboard oeffnen und Dateien filtern
  2. 'Arbeitsliste speichern (CSV)' klicken -> erzeugt eine Pipeline_*.csv
  3. Diese CSV pruefen (oeffnen, Zeilen entfernen/ergaenzen)
  4. Copy-ScannedFiles -CsvPath 'Pipeline_*.csv' -Destination '...'

Tipp: Die Pipeline-CSV enthaelt nur die gefilterten Dateien aus dem Dashboard.
"@
    return
  }

  # ── Read CSV ──
  Write-Host "[Copy-ScannedFiles] Reading CSV: $CsvPath" -ForegroundColor Cyan
  $files = Import-Csv -Path $CsvPath -Encoding UTF8

  if (-not $files -or $files.Count -eq 0) {
    Write-Warning "[Copy-ScannedFiles] CSV is empty."
    return
  }

  Write-Host "[Copy-ScannedFiles] Total files in CSV: $($files.Count)" -ForegroundColor Gray

  # ── Filter ──
  if ($OnlyConvertible) {
    $files = $files | Where-Object { $_.IsConvertible -eq 'True' }
    Write-Host "[Copy-ScannedFiles] After convertible filter: $($files.Count) files" -ForegroundColor Gray
  }

  if ($Extensions -and $Extensions.Count -gt 0) {
    $extSet = $Extensions | ForEach-Object { $_.ToLower() }
    $files = $files | Where-Object { $extSet -contains $_.Extension.ToLower() }
    Write-Host "[Copy-ScannedFiles] After extension filter ($($Extensions -join ', ')): $($files.Count) files" -ForegroundColor Gray
  }

  if ($files.Count -eq 0) {
    Write-Warning "[Copy-ScannedFiles] No files match the filter criteria."
    return
  }

  # ── Create destination ──
  if (-not (Test-Path $Destination)) {
    if ($PSCmdlet.ShouldProcess($Destination, "Create directory")) {
      New-Item -ItemType Directory -Path $Destination -Force | Out-Null
      Write-Host "[Copy-ScannedFiles] Created destination: $Destination" -ForegroundColor Green
    }
  }
  $Destination = if (Test-Path $Destination) { (Resolve-Path $Destination).Path } else { $Destination }

  # ── Copy files ──
  $copied = 0
  $skippedExists = 0
  $skippedNotFound = 0
  $errors = 0
  $totalSize = 0
  $total = $files.Count
  $current = 0
  $notFoundLog = @()
  $errorLog = @()
  $copyStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  $totalCopyTime = 0

  Write-Host "`n[Copy-ScannedFiles] Starting copy of $total files..." -ForegroundColor Cyan
  if (-not $Force) {
    Write-Host "[Copy-ScannedFiles] Mode: Incremental (existing files will be skipped, use -Force to overwrite)" -ForegroundColor Gray
  }
  else {
    Write-Host "[Copy-ScannedFiles] Mode: Force (existing files will be overwritten)" -ForegroundColor Yellow
  }

  foreach ($f in $files) {
    $current++
    # Normalize path: trim whitespace, remove BOM/invisible chars, normalize separators
    $sourcePath = ($f.FullPath -replace '[\u200B\u200C\u200D\uFEFF]', '').Trim()
    $sourcePath = $sourcePath -replace '/', '\'
    $pctComplete = [math]::Round(($current / $total) * 100)

    # Calculate relative path from scan root
    if ($sourcePath.StartsWith($ScanRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
      $relativePath = $sourcePath.Substring($ScanRoot.Length).TrimStart('\')
    }
    else {
      $relativePath = $f.Name
      Write-Warning "[Copy-ScannedFiles] File outside scan root, copying flat: $sourcePath"
    }

    $destPath = Join-Path $Destination $relativePath
    $destDir = Split-Path $destPath -Parent

    # Progress bar with ETA
    $eta = ''
    if ($copied -gt 0) {
      $avgSec = $totalCopyTime / $copied
      $remaining = ($total - $current + 1) * $avgSec
      $etaTs = [TimeSpan]::FromSeconds([math]::Max(0, $remaining))
      $elapsed = $copyStopwatch.Elapsed
      $eta = " | ETA: $($etaTs.ToString('hh\:mm\:ss')) | Elapsed: $($elapsed.ToString('hh\:mm\:ss'))"
    }
    Write-Progress -Activity "Dateien kopieren" `
      -Status "[$current/$total] $relativePath$eta" `
      -PercentComplete $pctComplete

    # Check source exists (with LiteralPath for special characters)
    if (-not (Test-Path -LiteralPath $sourcePath)) {
      $notFoundLog += [PSCustomObject]@{
        Index        = $current
        FullPath     = $sourcePath
        OriginalPath = $f.FullPath
        Name         = $f.Name
        Extension    = $f.Extension
        SizeBytes    = $f.SizeBytes
        Directory    = Split-Path $sourcePath -Parent
        DirExists    = (Test-Path -LiteralPath (Split-Path $sourcePath -Parent))
        PathLength   = $sourcePath.Length
      }
      $skippedNotFound++
      continue
    }

    # Check if destination already exists (incremental skip)
    if ((Test-Path -LiteralPath $destPath) -and -not $Force) {
      $skippedExists++
      continue
    }

    $fileStart = [System.Diagnostics.Stopwatch]::StartNew()

    if ($PSCmdlet.ShouldProcess($sourcePath, "Copy to $destPath")) {
      try {
        if (-not (Test-Path -LiteralPath $destDir)) {
          New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Copy-Item -LiteralPath $sourcePath -Destination $destPath -Force
        $copied++
        $totalSize += [long]($f.SizeBytes)
      }
      catch {
        Write-Warning "[Copy-ScannedFiles] Failed to copy '$sourcePath': $_"
        $errorLog += [PSCustomObject]@{
          FullPath = $sourcePath
          Name     = $f.Name
          Error    = $_.Exception.Message
        }
        $errors++
      }
    }

    $fileStart.Stop()
    $totalCopyTime += $fileStart.Elapsed.TotalSeconds
  }

  Write-Progress -Activity "Dateien kopieren" -Completed
  $copyStopwatch.Stop()

  # ── Write not-found log ──
  $logDir = $PSScriptRoot
  if ($notFoundLog.Count -gt 0) {
    $notFoundCsvPath = Join-Path $logDir "copy_notfound.csv"
    $notFoundLog | Export-Csv -Path $notFoundCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`n[Copy-ScannedFiles] Not-found log: $notFoundCsvPath" -ForegroundColor Yellow
    Write-Host "[Copy-ScannedFiles] First 5 not-found paths:" -ForegroundColor Yellow
    $notFoundLog | Select-Object -First 5 | ForEach-Object {
      Write-Host "  $($_.FullPath)" -ForegroundColor DarkYellow
      if ($_.OriginalPath -ne $_.FullPath) {
        Write-Host "    (original CSV value: '$($_.OriginalPath)')" -ForegroundColor DarkGray
      }
      Write-Host "    Dir exists: $($_.DirExists) | Path length: $($_.PathLength)" -ForegroundColor DarkGray
    }
  }

  # ── Write error log ──
  if ($errorLog.Count -gt 0) {
    $errorCsvPath = Join-Path $logDir "copy_errors.csv"
    $errorLog | Export-Csv -Path $errorCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "[Copy-ScannedFiles] Error log: $errorCsvPath" -ForegroundColor Red
  }

  # ── Summary ──
  $sizeMB = [math]::Round($totalSize / 1MB, 2)
  $elapsed = $copyStopwatch.Elapsed
  Write-Host "`n[Copy-ScannedFiles] ════════════════════════════════" -ForegroundColor Green
  Write-Host "[Copy-ScannedFiles] Copied:     $copied files ($sizeMB MB)" -ForegroundColor Green
  Write-Host "[Copy-ScannedFiles] Duration:   $($elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Green
  if ($skippedExists -gt 0) { Write-Host "[Copy-ScannedFiles] Skipped:    $skippedExists (already in destination)" -ForegroundColor DarkGray }
  if ($skippedNotFound -gt 0) { Write-Host "[Copy-ScannedFiles] Not found:  $skippedNotFound (source missing) -> see copy_notfound.csv" -ForegroundColor Yellow }
  if ($errors -gt 0) { Write-Host "[Copy-ScannedFiles] Errors:     $errors -> see copy_errors.csv" -ForegroundColor Red }
  Write-Host "[Copy-ScannedFiles] Target:     $Destination" -ForegroundColor Green
  Write-Host "[Copy-ScannedFiles] ════════════════════════════════" -ForegroundColor Green
}
