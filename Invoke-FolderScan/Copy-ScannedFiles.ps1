<#
.SYNOPSIS
    Copies files listed in a FolderScan CSV to a target directory, preserving folder structure.

.DESCRIPTION
    Reads a FullScan.csv (or Pipeline CSV) exported by Invoke-FolderScan and copies the listed
    files to a destination folder. The relative folder structure from the original scan root is
    preserved under the destination.

    Supports filtering by convertibility and file extension, so you can e.g. copy only
    convertible files for downstream processing (Ducling, Markdown conversion, etc.).

.PARAMETER CsvPath
    Path to the FullScan.csv or Pipeline CSV file.

.PARAMETER Destination
    Target folder where files will be copied to. Created automatically if it doesn't exist.

.PARAMETER ScanRoot
    The original scan root path. Used to calculate relative paths.
    If omitted, the script tries to read it from FullScan.meta.json next to the CSV.

.PARAMETER OnlyConvertible
    If set, only files where IsConvertible = True are copied.

.PARAMETER Extensions
    Optional array of extensions to include (e.g. '.pdf', '.docx').
    If omitted, all extensions are included.

.PARAMETER WhatIf
    Shows what would be copied without actually copying.

.EXAMPLE
    . .\Copy-ScannedFiles.ps1
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\ConvertibleFiles"

.EXAMPLE
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\Output" -OnlyConvertible

.EXAMPLE
    Copy-ScannedFiles -CsvPath ".\scan-results\FullScan.csv" -Destination "D:\Output" -Extensions '.pdf','.docx' -WhatIf
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

    [string[]]$Extensions
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
  $skipped = 0
  $errors = 0
  $totalSize = 0

  foreach ($f in $files) {
    $sourcePath = $f.FullPath

    # Calculate relative path from scan root
    if ($sourcePath.StartsWith($ScanRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
      $relativePath = $sourcePath.Substring($ScanRoot.Length).TrimStart('\')
    }
    else {
      # Fallback: use DirectoryName relative part + filename
      $relativePath = $f.Name
      Write-Warning "[Copy-ScannedFiles] File outside scan root, copying flat: $sourcePath"
    }

    $destPath = Join-Path $Destination $relativePath
    $destDir = Split-Path $destPath -Parent

    # Check source exists
    if (-not (Test-Path $sourcePath)) {
      Write-Warning "[Copy-ScannedFiles] Source not found, skipping: $sourcePath"
      $skipped++
      continue
    }

    if ($PSCmdlet.ShouldProcess($sourcePath, "Copy to $destPath")) {
      try {
        if (-not (Test-Path $destDir)) {
          New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Copy-Item -Path $sourcePath -Destination $destPath -Force
        $copied++
        $totalSize += [long]($f.SizeBytes)
      }
      catch {
        Write-Warning "[Copy-ScannedFiles] Failed to copy '$sourcePath': $_"
        $errors++
      }
    }
  }

  # ── Summary ──
  $sizeMB = [math]::Round($totalSize / 1MB, 2)
  Write-Host "`n[Copy-ScannedFiles] ════════════════════════════════" -ForegroundColor Green
  Write-Host "[Copy-ScannedFiles] Copied:  $copied files ($sizeMB MB)" -ForegroundColor Green
  if ($skipped -gt 0) { Write-Host "[Copy-ScannedFiles] Skipped: $skipped (source not found)" -ForegroundColor Yellow }
  if ($errors -gt 0)  { Write-Host "[Copy-ScannedFiles] Errors:  $errors" -ForegroundColor Red }
  Write-Host "[Copy-ScannedFiles] Target:  $Destination" -ForegroundColor Green
  Write-Host "[Copy-ScannedFiles] ════════════════════════════════" -ForegroundColor Green
}
