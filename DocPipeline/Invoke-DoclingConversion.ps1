<#
.SYNOPSIS
    Sends files to a Docling Serve API for document conversion (PDF → Markdown, HTML, etc.).

.DESCRIPTION
    Invoke-DoclingConversion processes files from an input folder (typically the Staging folder
    from the Invoke-FolderScan pipeline) and sends them to a Docling Serve REST API for conversion.

    Features:
    - Bulk processing with Write-Progress
    - Skip already-converted files (incremental / resume-safe)
    - Retry with exponential backoff on API errors
    - Concurrent API calls with -MaxConcurrency
    - Detailed per-file log (docling_log.csv)
    - Run summary log (docling_runs.json) for time estimation in the dashboard
    - Configurable output formats, OCR, pipeline type, PDF backend, etc.
    - API health check before starting batch

    Designed to work with the Invoke-FolderScan pipeline:
      Source (Scan) → Pipeline_*.csv → Staging (Copy-ScannedFiles) → Docling → Ergebnis

.PARAMETER DoclingUrl
    Base URL of the Docling Serve API (e.g. 'http://localhost:5001').

.PARAMETER InputPath
    Folder containing files to convert (typically the Staging folder).

.PARAMETER OutputPath
    Target folder for conversion results. Folder structure from InputPath is preserved.
    When multiple formats are selected, subfolders per format are created.

.PARAMETER Folders
    Optional array of subfolder names (relative to InputPath) to process.
    If omitted, all files in InputPath are processed (recursively).

.PARAMETER Extensions
    Optional array of file extensions to include (e.g. '.pdf', '.docx').
    If omitted, all files are included.

.PARAMETER ToFormats
    Output formats to generate. Default: 'markdown'.
    Valid values: 'markdown', 'html', 'text', 'json', 'doctags'.

.PARAMETER PipelineType
    Docling pipeline type. Default: 'standard'.
    Valid values: 'legacy', 'standard', 'vlm', 'asr'.

.PARAMETER EnableOcr
    Enable OCR processing.

.PARAMETER ForceOcr
    Force OCR on all pages (even those with selectable text).

.PARAMETER OcrEngine
    OCR engine to use. Default: 'auto'.
    Valid values: 'auto', 'easyocr', 'tesseract', 'rapidocr'.

.PARAMETER PdfBackend
    PDF parsing backend. Default: 'docling_parse'.
    Valid values: 'docling_parse', 'pypdfium2'.

.PARAMETER TableMode
    Table extraction mode. Default: 'accurate'.
    Valid values: 'fast', 'accurate'.

.PARAMETER ImageExportMode
    How images are handled in output. Default: 'embedded'.
    Valid values: 'embedded', 'placeholder', 'referenced'.

.PARAMETER AbortOnError
    If set, stop processing on first conversion error.

.PARAMETER EnableCodeEnrichment
    Enable code block enrichment in output.

.PARAMETER EnableFormulaEnrichment
    Enable formula/math enrichment in output.

.PARAMETER EnablePictureClassification
    Enable automatic picture classification.

.PARAMETER EnablePictureDescription
    Enable automatic picture description generation.

.PARAMETER SkipExisting
    Skip files that already have output in the target folder. Default: $true.
    This makes the script resume-safe — just re-run after a crash.

.PARAMETER Force
    Overwrite existing output files (opposite of SkipExisting).

.PARAMETER MaxConcurrency
    Maximum number of parallel API calls. Default: 1.
    Increase if the Docling server has capacity for concurrent processing.

.PARAMETER TimeoutSec
    Timeout in seconds per API call. Default: 300 (5 minutes).
    Large files may need more time.

.PARAMETER RetryCount
    Number of retries on API failure. Default: 3.

.PARAMETER CsvPath
    Optional: Use a Pipeline_*.csv as input instead of -InputPath.
    The CSV must contain a 'FullPath' column.

.PARAMETER LogDir
    Directory for log files. Defaults to the script directory ($PSScriptRoot).

.NOTES
    Name:    Invoke-DoclingConversion
    Author:  Benjamin Rauser
    Contact: Benjamin.Rauser@outlook.com
    Version: 1.0 - 2026-02-16

.EXAMPLE
    . .\Invoke-DoclingConversion.ps1
    Invoke-DoclingConversion -DoclingUrl "http://localhost:5001" -InputPath "D:\Staging" -OutputPath "D:\Ergebnis"

.EXAMPLE
    # Only specific subfolders, multiple output formats
    Invoke-DoclingConversion -DoclingUrl "http://localhost:5001" -InputPath "D:\Staging" -OutputPath "D:\Ergebnis" `
      -Folders "Vertraege","Rechnungen\Eingang" -ToFormats markdown,html

.EXAMPLE
    # Force re-conversion of all files, with OCR
    Invoke-DoclingConversion -DoclingUrl "http://localhost:5001" -InputPath "D:\Staging" -OutputPath "D:\Ergebnis" `
      -Force -EnableOcr -OcrEngine easyocr

.EXAMPLE
    # From Pipeline CSV instead of folder
    Invoke-DoclingConversion -DoclingUrl "http://localhost:5001" -CsvPath ".\scan-results\Pipeline_2026-02-16_1430.csv" `
      -OutputPath "D:\Ergebnis" -ToFormats markdown
#>

function Invoke-DoclingConversion {
  [CmdletBinding(DefaultParameterSetName = 'Folder')]
  param(
    [Parameter(Mandatory)]
    [string]$DoclingUrl,

    [Parameter(Mandatory, ParameterSetName = 'Folder')]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$InputPath,

    [Parameter(Mandatory)]
    [string]$OutputPath,

    [Parameter(ParameterSetName = 'Folder')]
    [string[]]$Folders,

    [string[]]$Extensions,

    [ValidateSet('markdown', 'html', 'text', 'json', 'doctags')]
    [string[]]$ToFormats = @('markdown'),

    [ValidateSet('legacy', 'standard', 'vlm', 'asr')]
    [string]$PipelineType = 'standard',

    [switch]$EnableOcr,

    [switch]$ForceOcr,

    [ValidateSet('auto', 'easyocr', 'tesseract', 'rapidocr')]
    [string]$OcrEngine = 'auto',

    [ValidateSet('docling_parse', 'pypdfium2')]
    [string]$PdfBackend = 'docling_parse',

    [ValidateSet('fast', 'accurate')]
    [string]$TableMode = 'accurate',

    [ValidateSet('embedded', 'placeholder', 'referenced')]
    [string]$ImageExportMode = 'embedded',

    [switch]$AbortOnError,

    [switch]$EnableCodeEnrichment,

    [switch]$EnableFormulaEnrichment,

    [switch]$EnablePictureClassification,

    [switch]$EnablePictureDescription,

    [bool]$SkipExisting = $true,

    [switch]$Force,

    [ValidateRange(1, 20)]
    [int]$MaxConcurrency = 1,

    [ValidateRange(30, 7200)]
    [int]$TimeoutSec = 900,

    [ValidateRange(0, 10)]
    [int]$RetryCount = 3,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [string]$LogDir
  )

  # ================================================================
  # INIT
  # ================================================================
  $DoclingUrl = $DoclingUrl.TrimEnd('/')

  if ($Force) { $SkipExisting = $false }

  if (-not $LogDir) { $LogDir = $PSScriptRoot }

  # Ensure output & log directories exist
  foreach ($dir in @($OutputPath, $LogDir)) {
    if (-not (Test-Path $dir)) {
      New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
  }
  $OutputPath = (Resolve-Path $OutputPath).Path
  $LogDir = (Resolve-Path $LogDir).Path

  # Format → file extension mapping
  $formatExtMap = @{
    'markdown' = '.md'
    'html'     = '.html'
    'text'     = '.txt'
    'json'     = '.json'
    'doctags'  = '.doctags.xml'
  }

  # Format → Docling API "to_formats" value mapping
  $formatApiMap = @{
    'markdown' = 'md'
    'html'     = 'html'
    'text'     = 'text'
    'json'     = 'json'
    'doctags'  = 'doctags'
  }

  $multiFormat = ($ToFormats.Count -gt 1)
  $logCsvPath = Join-Path $LogDir "docling_log.csv"
  $runsJsonPath = Join-Path $LogDir "docling_runs.json"

  Write-Host "`n[Invoke-DoclingConversion] ================================" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Docling URL:  $DoclingUrl" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Formats:      $($ToFormats -join ', ')" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Pipeline:     $PipelineType" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] OCR:          $(if ($EnableOcr) { "$OcrEngine$(if ($ForceOcr) { ' (forced)' })" } else { 'disabled' })" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] PDF Backend:  $PdfBackend" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Table Mode:   $TableMode" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Skip existing: $SkipExisting" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Concurrency:  $MaxConcurrency" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Timeout:      ${TimeoutSec}s" -ForegroundColor Cyan

  # Normalize URL: remove trailing slashes
  $DoclingUrl = $DoclingUrl.TrimEnd('/')

  # ================================================================
  # API HEALTH CHECK (with IPv6->IPv4 fallback for .local hostnames)
  # ================================================================
  Write-Host "`n[Invoke-DoclingConversion] Checking API health..." -ForegroundColor Yellow
  $apiReachable = $false
  $urlsToTry = @($DoclingUrl)

  # If hostname resolves to IPv6 link-local, PowerShell HTTP clients often fail.
  # Try resolving to IPv4 as fallback.
  try {
    $uri = [System.Uri]$DoclingUrl
    if ($uri.HostNameType -eq [System.UriHostNameType]::Dns) {
      $addresses = [System.Net.Dns]::GetHostAddresses($uri.Host)
      $ipv4 = $addresses | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | Select-Object -First 1
      if ($ipv4) {
        $ipv4Url = "$($uri.Scheme)://$($ipv4.ToString()):$($uri.Port)"
        if ($ipv4Url -ne $DoclingUrl) {
          $urlsToTry += $ipv4Url
        }
      }
    }
  }
  catch { }

  foreach ($baseUrl in $urlsToTry) {
    foreach ($ep in @("$baseUrl/health", "$baseUrl/docs", "$baseUrl/")) {
      try {
        $null = Invoke-WebRequest -Uri $ep -Method GET -TimeoutSec 10 -ErrorAction Stop -UseBasicParsing
        $apiReachable = $true
        if ($baseUrl -ne $DoclingUrl) {
          Write-Host "[Invoke-DoclingConversion] Hostname resolved to IPv6; using IPv4 ($baseUrl) instead." -ForegroundColor Yellow
          $DoclingUrl = $baseUrl
        }
        Write-Host "[Invoke-DoclingConversion] API is reachable ($ep)." -ForegroundColor Green
        break
      }
      catch { }
    }
    if ($apiReachable) { break }
  }
  if (-not $apiReachable) {
    Write-Error "[Invoke-DoclingConversion] Cannot reach Docling API at '$DoclingUrl'. Please check the URL and ensure the service is running."
    return
  }

  # Detect async endpoint availability (avoids proxy timeout for large files)
  $useAsync = $false
  $convertEndpoint = "$DoclingUrl/v1/convert/file"
  $asyncEndpoint = "$DoclingUrl/v1/convert/file/async"
  $pollEndpoint = "$DoclingUrl/v1/status/poll"
  $resultEndpoint = "$DoclingUrl/v1/result"

  try {
    # OPTIONS or a quick test to see if async endpoint exists
    $null = Invoke-WebRequest -Uri "$DoclingUrl/v1/convert/file/async" -Method Options -TimeoutSec 5 -ErrorAction Stop -UseBasicParsing
    $useAsync = $true
  }
  catch {
    # Try HEAD as fallback (some servers don't support OPTIONS)
    try {
      $null = Invoke-WebRequest -Uri "$DoclingUrl/v1/convert/file/async" -Method Head -TimeoutSec 5 -ErrorAction Stop -UseBasicParsing
      $useAsync = $true
    }
    catch {
      # 405 Method Not Allowed means the endpoint exists but doesn't accept HEAD/OPTIONS
      if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 405) {
        $useAsync = $true
      }
    }
  }

  if ($useAsync) {
    Write-Host "[Invoke-DoclingConversion] Mode:         ASYNC (resilient against proxy timeouts)" -ForegroundColor Green
  }
  else {
    Write-Host "[Invoke-DoclingConversion] Mode:         SYNC (async endpoint not available)" -ForegroundColor Yellow
    Write-Host "[Invoke-DoclingConversion] WARNING:      Large files may fail with 504 Gateway Timeout if a reverse proxy is in front of Docling." -ForegroundColor Yellow
  }

  # ================================================================
  # COLLECT FILES
  # ================================================================
  $filesToProcess = @()
  $inputRoot = ''

  if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    # CSV mode
    Write-Host "[Invoke-DoclingConversion] Reading CSV: $CsvPath" -ForegroundColor Cyan
    $csvData = Import-Csv -Path $CsvPath -Encoding UTF8
    if (-not $csvData -or $csvData.Count -eq 0) {
      Write-Warning "[Invoke-DoclingConversion] CSV is empty."
      return
    }

    # Try to determine input root from CSV paths
    $firstPath = $csvData[0].FullPath
    if (-not $firstPath) {
      Write-Error "[Invoke-DoclingConversion] CSV must contain a 'FullPath' column."
      return
    }

    # Use ScanRoot from metadata if available
    $csvDir = Split-Path (Resolve-Path $CsvPath).Path -Parent
    $metaPath = Join-Path $csvDir "FullScan.meta.json"
    if (Test-Path $metaPath) {
      try {
        $meta = Get-Content -Path $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($meta.StagingPath) { $inputRoot = $meta.StagingPath.TrimEnd('\') }
        elseif ($meta.ScanRoot) { $inputRoot = $meta.ScanRoot.TrimEnd('\') }
      }
      catch { }
    }

    if (-not $inputRoot) {
      $inputRoot = Split-Path $firstPath -Parent
      Write-Warning "[Invoke-DoclingConversion] Could not determine root from metadata. Using: $inputRoot"
    }

    foreach ($row in $csvData) {
      if ($row.FullPath -and (Test-Path $row.FullPath)) {
        $filesToProcess += [PSCustomObject]@{
          FullPath = $row.FullPath
          Name     = [System.IO.Path]::GetFileName($row.FullPath)
          SizeMB   = if ($row.SizeMB) { [double]$row.SizeMB } else { [math]::Round((Get-Item $row.FullPath).Length / 1MB, 2) }
        }
      }
    }
  }
  else {
    # Folder mode
    $InputPath = (Resolve-Path $InputPath).Path
    $inputRoot = $InputPath.TrimEnd('\')
    Write-Host "[Invoke-DoclingConversion] Input:        $InputPath" -ForegroundColor Cyan

    $gciParams = @{ Path = $InputPath; File = $true; Recurse = $true; ErrorAction = 'SilentlyContinue' }
    $allInputFiles = Get-ChildItem @gciParams

    # Filter by folders
    if ($Folders -and $Folders.Count -gt 0) {
      Write-Host "[Invoke-DoclingConversion] Folder filter: $($Folders -join ', ')" -ForegroundColor Cyan
      $filteredFiles = @()
      foreach ($f in $allInputFiles) {
        $relPath = ''
        if ($f.FullName.StartsWith($inputRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
          $relPath = $f.FullName.Substring($inputRoot.Length).TrimStart('\')
        }
        foreach ($folder in $Folders) {
          $folderNorm = $folder.TrimStart('\').TrimEnd('\')
          if ($relPath.StartsWith($folderNorm + '\', [System.StringComparison]::OrdinalIgnoreCase) -or
            $relPath.StartsWith($folderNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
            $filteredFiles += $f
            break
          }
        }
      }
      $allInputFiles = $filteredFiles
    }

    # Filter by extensions
    if ($Extensions -and $Extensions.Count -gt 0) {
      $extSet = $Extensions | ForEach-Object { $_.ToLower() }
      $allInputFiles = @($allInputFiles | Where-Object { $extSet -contains $_.Extension.ToLower() })
      Write-Host "[Invoke-DoclingConversion] Extension filter: $($Extensions -join ', ')" -ForegroundColor Cyan
    }

    foreach ($f in $allInputFiles) {
      $filesToProcess += [PSCustomObject]@{
        FullPath = $f.FullName
        Name     = $f.Name
        SizeMB   = [math]::Round($f.Length / 1MB, 2)
      }
    }
  }

  # Filter out unsupported input formats (per Docling API from_formats)
  $supportedExtensions = @(
    '.pdf', '.docx', '.pptx', '.html', '.htm', '.md', '.csv', '.xlsx',
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif', '.webp',
    '.asciidoc', '.adoc', '.xml', '.json', '.vtt',
    '.mp3', '.wav', '.flac', '.m4a', '.ogg', '.wma'
  )
  $unsupported = @($filesToProcess | Where-Object {
      $ext = [System.IO.Path]::GetExtension($_.FullPath).ToLower()
      $supportedExtensions -notcontains $ext
    })
  if ($unsupported.Count -gt 0) {
    $skippedExts = ($unsupported | ForEach-Object { [System.IO.Path]::GetExtension($_.FullPath).ToLower() } | Sort-Object -Unique) -join ', '
    Write-Warning "[Invoke-DoclingConversion] Skipping $($unsupported.Count) file(s) with unsupported format(s): $skippedExts"
    foreach ($u in $unsupported) {
      Write-Host "  [SKIP] $($u.Name) (unsupported format)" -ForegroundColor DarkYellow
    }
    $filesToProcess = @($filesToProcess | Where-Object {
        $ext = [System.IO.Path]::GetExtension($_.FullPath).ToLower()
        $supportedExtensions -contains $ext
      })
  }

  Write-Host "[Invoke-DoclingConversion] Output:       $OutputPath" -ForegroundColor Cyan
  Write-Host "[Invoke-DoclingConversion] Files found:  $($filesToProcess.Count)" -ForegroundColor Green

  if ($filesToProcess.Count -eq 0) {
    Write-Warning "[Invoke-DoclingConversion] No files to process."
    return
  }

  # ================================================================
  # SKIP EXISTING CHECK
  # ================================================================
  $filesToConvert = @()
  $skippedExisting = 0

  foreach ($f in $filesToProcess) {
    $relPath = ''
    if ($f.FullPath.StartsWith($inputRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
      $relPath = $f.FullPath.Substring($inputRoot.Length).TrimStart('\')
    }
    else {
      $relPath = $f.Name
    }

    $shouldSkip = $false
    if ($SkipExisting) {
      # Check if output exists for the PRIMARY format (first in list)
      $primaryFormat = $ToFormats[0]
      $baseName = [System.IO.Path]::GetFileNameWithoutExtension($relPath)
      $relDir = [System.IO.Path]::GetDirectoryName($relPath)
      $outExt = $formatExtMap[$primaryFormat]

      $relFile = if ([string]::IsNullOrEmpty($relDir)) { "$baseName$outExt" } else { Join-Path $relDir "$baseName$outExt" }

      if ($multiFormat) {
        $outFile = Join-Path $OutputPath (Join-Path $primaryFormat $relFile)
      }
      else {
        $outFile = Join-Path $OutputPath $relFile
      }

      if (Test-Path $outFile) {
        $shouldSkip = $true
        $skippedExisting++
      }
    }

    if (-not $shouldSkip) {
      $filesToConvert += [PSCustomObject]@{
        FullPath     = $f.FullPath
        Name         = $f.Name
        SizeMB       = $f.SizeMB
        RelativePath = $relPath
      }
    }
  }

  if ($skippedExisting -gt 0) {
    Write-Host "[Invoke-DoclingConversion] Skipped (existing): $skippedExisting" -ForegroundColor DarkGray
  }

  Write-Host "[Invoke-DoclingConversion] To convert:   $($filesToConvert.Count)" -ForegroundColor Green

  if ($filesToConvert.Count -eq 0) {
    Write-Host "[Invoke-DoclingConversion] All files already converted. Nothing to do." -ForegroundColor Green
    return
  }

  $totalSizeMB = [math]::Round(($filesToConvert | Measure-Object -Property SizeMB -Sum).Sum, 2)
  Write-Host "[Invoke-DoclingConversion] Total size:   $totalSizeMB MB" -ForegroundColor Cyan

  # ================================================================
  # PREPARE LOG
  # ================================================================
  $logEntries = @()
  $batchStart = Get-Date

  # ================================================================
  # CONVERSION LOOP
  # ================================================================
  $converted = 0
  $failed = 0
  $totalFiles = $filesToConvert.Count
  $current = 0
  $totalConvertTime = 0
  $manifestEntries = @{}

  Write-Host "`n[Invoke-DoclingConversion] Starting conversion..." -ForegroundColor Cyan

  # Detect PowerShell version for HTTP method selection
  $useDotNetHttp = $PSVersionTable.PSVersion.Major -ge 7
  $httpClient = $null
  $httpHandler = $null

  if ($useDotNetHttp) {
    Write-Host "[Invoke-DoclingConversion] HTTP engine:  HttpClient (PS $($PSVersionTable.PSVersion))" -ForegroundColor DarkGray
    $httpHandler = [System.Net.Http.HttpClientHandler]::new()
    $httpClient = [System.Net.Http.HttpClient]::new($httpHandler)
    $httpClient.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
    $httpClient.DefaultRequestHeaders.Add('Accept', 'application/json')
  }
  else {
    Write-Host "[Invoke-DoclingConversion] HTTP engine:  Invoke-RestMethod (PS $($PSVersionTable.PSVersion))" -ForegroundColor DarkGray
  }

  $batchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

  try {

    foreach ($file in $filesToConvert) {
      $current++
      $pctComplete = [math]::Round(($current / $totalFiles) * 100)
      $fileStart = [System.Diagnostics.Stopwatch]::StartNew()

      # Progress bar with ETA
      $eta = ''
      if ($converted -gt 0) {
        $avgSec = $totalConvertTime / $converted
        $remaining = ($totalFiles - $current + 1) * $avgSec
        $etaTs = [TimeSpan]::FromSeconds([math]::Max(0, $remaining))
        $elapsed = $batchStopwatch.Elapsed
        $eta = " | ETA: $($etaTs.ToString('hh\:mm\:ss')) | Elapsed: $($elapsed.ToString('hh\:mm\:ss')) | Avg: $([math]::Round($avgSec,1))s/file"
      }

      Write-Progress -Activity "Docling Conversion" `
        -Status "[$current/$totalFiles] $($file.Name) ($($file.SizeMB) MB)$eta" `
        -PercentComplete $pctComplete

      $convertSuccess = $false
      $errorText = ''
      $responseContent = $null

      for ($attempt = 1; $attempt -le ($RetryCount + 1); $attempt++) {
        try {
          if ($current -eq 1 -and $attempt -eq 1) {
            $targetEp = if ($useAsync) { $asyncEndpoint } else { $convertEndpoint }
            Write-Host "  [INFO] Sending first file to API ($targetEp)..." -ForegroundColor DarkGray
          }

          $fileBytes = [System.IO.File]::ReadAllBytes($file.FullPath)
          $fileName = [System.IO.Path]::GetFileName($file.FullPath)

          # ---- Build multipart body (shared by all code paths) ----
          $boundary = [System.Guid]::NewGuid().ToString('N')
          $LF = "`r`n"
          $enc = [System.Text.Encoding]::UTF8
          $bodyParts = [System.Collections.Generic.List[byte[]]]::new()

          # File part
          $fileHeader = "--$boundary$LF" +
          "Content-Disposition: form-data; name=`"files`"; filename=`"$fileName`"$LF" +
          "Content-Type: application/octet-stream$LF$LF"
          $bodyParts.Add($enc.GetBytes($fileHeader))
          $bodyParts.Add($fileBytes)
          $bodyParts.Add($enc.GetBytes($LF))

          # to_formats (multiple values with same field name)
          foreach ($fmt in $ToFormats) {
            $part = "--$boundary$LF" +
            "Content-Disposition: form-data; name=`"to_formats`"$LF$LF" +
            "$($formatApiMap[$fmt])$LF"
            $bodyParts.Add($enc.GetBytes($part))
          }

          # Single-value form fields
          $singleFields = [ordered]@{
            'image_export_mode'         = $ImageExportMode
            'do_ocr'                    = $EnableOcr.ToString().ToLower()
            'force_ocr'                 = $ForceOcr.ToString().ToLower()
            'ocr_engine'                = $OcrEngine
            'pdf_backend'               = $PdfBackend
            'table_mode'                = $TableMode
            'pipeline'                  = $PipelineType
            'abort_on_error'            = $AbortOnError.ToString().ToLower()
            'do_code_enrichment'        = $EnableCodeEnrichment.ToString().ToLower()
            'do_formula_enrichment'     = $EnableFormulaEnrichment.ToString().ToLower()
            'do_picture_classification' = $EnablePictureClassification.ToString().ToLower()
            'do_picture_description'    = $EnablePictureDescription.ToString().ToLower()
          }
          foreach ($key in $singleFields.Keys) {
            $part = "--$boundary$LF" +
            "Content-Disposition: form-data; name=`"$key`"$LF$LF" +
            "$($singleFields[$key])$LF"
            $bodyParts.Add($enc.GetBytes($part))
          }
          $bodyParts.Add($enc.GetBytes("--$boundary--$LF"))

          # Merge into single byte array
          $totalLen = 0; foreach ($p in $bodyParts) { $totalLen += $p.Length }
          $bodyBytes = [byte[]]::new($totalLen)
          $bOffset = 0
          foreach ($p in $bodyParts) {
            [System.Array]::Copy($p, 0, $bodyBytes, $bOffset, $p.Length)
            $bOffset += $p.Length
          }
          $contentType = "multipart/form-data; boundary=$boundary"

          # ---- Helper: send multipart and get response text (UTF-8 safe) ----
          # Used by both sync and async submit
          function Send-MultipartRequest {
            param([string]$Uri, [byte[]]$Body, [string]$CT, [int]$Timeout)
            if ($useDotNetHttp) {
              $rawContent = $null; $postTask2 = $null; $resp2 = $null
              try {
                $rawContent = [System.Net.Http.ByteArrayContent]::new($Body)
                $rawContent.Headers.Remove('Content-Type') | Out-Null
                $rawContent.Headers.TryAddWithoutValidation('Content-Type', $CT) | Out-Null
                $postTask2 = $httpClient.PostAsync($Uri, $rawContent)
                if (-not $postTask2.Wait([int]($Timeout * 1000))) {
                  throw "HTTP request timed out after ${Timeout}s"
                }
                $resp2 = $postTask2.Result
                $readTask2 = $resp2.Content.ReadAsStringAsync()
                $readTask2.Wait(30000) | Out-Null
                $respText = $readTask2.Result
                if (-not $resp2.IsSuccessStatusCode) {
                  throw "HTTP $([int]$resp2.StatusCode) $($resp2.ReasonPhrase): $($respText.Substring(0, [math]::Min(500, $respText.Length)))"
                }
                return $respText
              }
              finally {
                if ($resp2) { try { $resp2.Dispose() } catch {} }
                if ($postTask2) { try { $postTask2.Dispose() } catch {} }
                if ($rawContent) { try { $rawContent.Dispose() } catch {} }
              }
            }
            else {
              $wr = Invoke-WebRequest -Uri $Uri -Method Post -ContentType $CT -Body $Body `
                -TimeoutSec $Timeout -ErrorAction Stop -UseBasicParsing
              if ($wr.RawContentStream) {
                $ms2 = $wr.RawContentStream; $ms2.Position = 0
                $rd2 = [System.IO.StreamReader]::new($ms2, [System.Text.Encoding]::UTF8)
                $txt = $rd2.ReadToEnd(); $rd2.Dispose()
                return $txt
              }
              else {
                $rb = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetBytes($wr.Content)
                return [System.Text.Encoding]::UTF8.GetString($rb)
              }
            }
          }

          # ---- Helper: GET request with UTF-8 response ----
          function Get-Utf8Response {
            param([string]$Uri, [int]$Timeout)
            if ($useDotNetHttp) {
              $getTask = $null; $resp3 = $null
              try {
                $getTask = $httpClient.GetAsync($Uri)
                if (-not $getTask.Wait([int]($Timeout * 1000))) {
                  throw "HTTP GET timed out after ${Timeout}s"
                }
                $resp3 = $getTask.Result
                $readTask3 = $resp3.Content.ReadAsStringAsync()
                $readTask3.Wait(30000) | Out-Null
                $txt3 = $readTask3.Result
                if (-not $resp3.IsSuccessStatusCode) {
                  throw "HTTP $([int]$resp3.StatusCode): $($txt3.Substring(0, [math]::Min(500, $txt3.Length)))"
                }
                return $txt3
              }
              finally {
                if ($resp3) { try { $resp3.Dispose() } catch {} }
                if ($getTask) { try { $getTask.Dispose() } catch {} }
              }
            }
            else {
              $wr3 = Invoke-WebRequest -Uri $Uri -Method Get -TimeoutSec $Timeout -ErrorAction Stop -UseBasicParsing
              if ($wr3.RawContentStream) {
                $ms3 = $wr3.RawContentStream; $ms3.Position = 0
                $rd3 = [System.IO.StreamReader]::new($ms3, [System.Text.Encoding]::UTF8)
                $txt3 = $rd3.ReadToEnd(); $rd3.Dispose()
                return $txt3
              }
              else {
                $rb3 = [System.Text.Encoding]::GetEncoding('ISO-8859-1').GetBytes($wr3.Content)
                return [System.Text.Encoding]::UTF8.GetString($rb3)
              }
            }
          }

          if ($useAsync) {
            # ---- ASYNC path: submit, poll, fetch result ----
            $submitText = Send-MultipartRequest -Uri $asyncEndpoint -Body $bodyBytes -CT $contentType -Timeout 60
            $taskInfo = $submitText | ConvertFrom-Json
            $taskId = $taskInfo.task_id

            if (-not $taskId) {
              throw "Async submit did not return a task_id. Response: $($submitText.Substring(0, [math]::Min(300, $submitText.Length)))"
            }

            # Poll until done
            $pollInterval = 5
            $pollDeadline = [DateTime]::UtcNow.AddSeconds($TimeoutSec)
            $taskStatus = $taskInfo.task_status

            while ($taskStatus -notin @('success', 'failure')) {
              if ([DateTime]::UtcNow -gt $pollDeadline) {
                throw "Async conversion timed out after ${TimeoutSec}s (task_id: $taskId, last status: $taskStatus)"
              }
              Start-Sleep -Seconds $pollInterval
              $pollText = Get-Utf8Response -Uri "$pollEndpoint/$taskId" -Timeout 30
              $pollResult = $pollText | ConvertFrom-Json
              $taskStatus = $pollResult.task_status

              # Update progress with polling info + ETA
              $elapsedFile = $fileStart.Elapsed
              $pollEta = ''
              if ($converted -gt 0) {
                $avgSec2 = $totalConvertTime / $converted
                $remaining2 = ($totalFiles - $current) * $avgSec2 + [math]::Max(0, $avgSec2 - $elapsedFile.TotalSeconds)
                $etaTs2 = [TimeSpan]::FromSeconds([math]::Max(0, $remaining2))
                $elapsed2 = $batchStopwatch.Elapsed
                $pollEta = " | ETA: $($etaTs2.ToString('hh\:mm\:ss')) | Elapsed: $($elapsed2.ToString('hh\:mm\:ss')) | Avg: $([math]::Round($avgSec2,1))s/file"
              }
              Write-Progress -Activity "Docling Conversion" `
                -Status "[$current/$totalFiles] $($file.Name) - converting... ($([math]::Round($elapsedFile.TotalSeconds))s)$pollEta" `
                -PercentComplete $pctComplete
            }

            if ($taskStatus -eq 'failure') {
              throw "Async conversion failed on server (task_id: $taskId)"
            }

            # Fetch result
            $resultText = Get-Utf8Response -Uri "$resultEndpoint/$taskId" -Timeout 60
            $responseContent = $resultText | ConvertFrom-Json
          }
          else {
            # ---- SYNC path: direct POST, wait for response ----
            $syncText = Send-MultipartRequest -Uri $convertEndpoint -Body $bodyBytes -CT $contentType -Timeout $TimeoutSec
            $responseContent = $syncText | ConvertFrom-Json
          }

          $convertSuccess = $true
          break
        }
        catch {
          $errorText = $_.Exception.Message
          # Provide helpful hint for 504 errors
          if ($errorText -match '504') {
            $errorText += " [HINT: A reverse proxy is timing out. The Docling async endpoint (/v1/convert/file/async) would avoid this, but may not be available on your server version.]"
          }
          if ($attempt -le $RetryCount) {
            $waitSec = [math]::Pow(2, $attempt)
            Write-Warning "[Invoke-DoclingConversion] Attempt $attempt failed for '$($file.Name)': $errorText. Retrying in ${waitSec}s..."
            Start-Sleep -Seconds $waitSec
          }
          else {
            Write-Warning "[Invoke-DoclingConversion] FAILED after $($RetryCount + 1) attempts: '$($file.Name)': $errorText"
          }
        }
      }

      $fileStart.Stop()
      $fileDuration = [math]::Round($fileStart.Elapsed.TotalSeconds, 2)

      if ($convertSuccess) {
        $converted++
        $totalConvertTime += $fileDuration

        # Save output files
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.RelativePath)
        $relDir = [System.IO.Path]::GetDirectoryName($file.RelativePath)

        foreach ($fmt in $ToFormats) {
          $outExt = $formatExtMap[$fmt]
          $apiKey = $formatApiMap[$fmt]

          $relFile = if ([string]::IsNullOrEmpty($relDir)) { "$baseName$outExt" } else { Join-Path $relDir "$baseName$outExt" }

          if ($multiFormat) {
            $outFilePath = Join-Path $OutputPath (Join-Path $fmt $relFile)
          }
          else {
            $outFilePath = Join-Path $OutputPath $relFile
          }

          $outDir = Split-Path $outFilePath -Parent
          if (-not (Test-Path $outDir)) {
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
          }

          # Extract content from response
          $content = $null

          # Docling Serve response structure: check for document content
          if ($responseContent.document) {
            $doc = $responseContent.document
            switch ($fmt) {
              'markdown' {
                $content = if ($doc.md_content) { $doc.md_content }
                elseif ($doc.export_to_markdown) { $doc.export_to_markdown }
                elseif ($doc.content) { $doc.content }
              }
              'html' {
                $content = if ($doc.html_content) { $doc.html_content }
                elseif ($doc.export_to_html) { $doc.export_to_html }
              }
              'text' {
                $content = if ($doc.text_content) { $doc.text_content }
                elseif ($doc.export_to_text) { $doc.export_to_text }
              }
              'json' {
                $content = $doc | ConvertTo-Json -Depth 20
              }
              'doctags' {
                $content = if ($doc.doctags_content) { $doc.doctags_content }
                elseif ($doc.export_to_document_tokens) { $doc.export_to_document_tokens }
              }
            }
          }
          elseif ($responseContent.$apiKey) {
            $content = $responseContent.$apiKey
          }
          elseif ($responseContent.content) {
            $content = $responseContent.content
          }

          if ($content) {
            [System.IO.File]::WriteAllText($outFilePath, $content, [System.Text.Encoding]::UTF8)
          }
          else {
            # Fallback: save entire response as JSON
            $fallbackContent = $responseContent | ConvertTo-Json -Depth 20
            [System.IO.File]::WriteAllText($outFilePath, $fallbackContent, [System.Text.Encoding]::UTF8)
            Write-Warning "[Invoke-DoclingConversion] Could not extract '$fmt' content for '$($file.Name)'. Saved raw response."
          }
        }

        Write-Host "  [OK] $($file.Name) (${fileDuration}s)" -ForegroundColor Green

        # Track for manifest
        $manifestEntries[$file.RelativePath] = @{
          status       = 'veredelt'
          processedAt  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
          sourceScript = 'Invoke-DoclingConversion'
        }
      }
      else {
        $failed++

        $logEntries += [PSCustomObject]@{
          Timestamp    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
          File         = $file.Name
          RelativePath = $file.RelativePath
          SizeMB       = $file.SizeMB
          Status       = 'FAILED'
          DurationSec  = $fileDuration
          Error        = $errorText
        }

        Write-Host "  [FAIL] $($file.Name) (${fileDuration}s) - $errorText" -ForegroundColor Red

        if ($AbortOnError) {
          Write-Error "[Invoke-DoclingConversion] Aborting due to -AbortOnError."
          break
        }

        continue
      }

      # Log entry for success
      $logEntries += [PSCustomObject]@{
        Timestamp    = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        File         = $file.Name
        RelativePath = $file.RelativePath
        SizeMB       = $file.SizeMB
        Status       = 'OK'
        DurationSec  = $fileDuration
        Error        = ''
      }
    }

    Write-Progress -Activity "Docling Conversion" -Completed
    $batchStopwatch.Stop()

  } # end try
  finally {
    # Cleanup HttpClient (PS 7+ only) - always runs, even on exceptions
    if ($httpClient) { try { $httpClient.Dispose() } catch {} }
    if ($httpHandler) { try { $httpHandler.Dispose() } catch {} }
  }

  # ================================================================
  # WRITE LOGS
  # ================================================================
  $batchEnd = Get-Date
  $batchDuration = ($batchEnd - $batchStart).TotalSeconds

  # Per-file log (append)
  if ($logEntries.Count -gt 0) {
    $logExists = Test-Path $logCsvPath
    $logEntries | Export-Csv -Path $logCsvPath -NoTypeInformation -Encoding UTF8 -Append:$logExists
    Write-Host "`n[Invoke-DoclingConversion] Per-file log: $logCsvPath" -ForegroundColor DarkGray
  }

  # Run summary log (for dashboard time estimation)
  $runEntry = @{
    date         = $batchStart.ToString('yyyy-MM-dd HH:mm:ss')
    files        = $converted + $failed
    converted    = $converted
    failed       = $failed
    skipped      = $skippedExisting
    totalSeconds = [math]::Round($batchDuration, 2)
    avgPerFile   = if ($converted -gt 0) { [math]::Round($totalConvertTime / $converted, 2) } else { 0 }
    avgPerMB     = if ($converted -gt 0 -and $totalSizeMB -gt 0) { [math]::Round($totalConvertTime / $totalSizeMB, 2) } else { 0 }
    totalSizeMB  = $totalSizeMB
    formats      = ($ToFormats -join ',')
    pipelineType = $PipelineType
    ocrEngine    = if ($EnableOcr) { $OcrEngine } else { 'disabled' }
  }

  $existingRuns = @()
  if (Test-Path $runsJsonPath) {
    try {
      $runsData = Get-Content -Path $runsJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($runsData.runs) {
        $existingRuns = @($runsData.runs)
      }
    }
    catch { }
  }

  $existingRuns += $runEntry
  $runsJson = @{ runs = $existingRuns } | ConvertTo-Json -Depth 5
  [System.IO.File]::WriteAllText($runsJsonPath, $runsJson, [System.Text.Encoding]::UTF8)
  Write-Host "[Invoke-DoclingConversion] Run summary:  $runsJsonPath" -ForegroundColor DarkGray

  # Write pipeline manifest
  if ($manifestEntries.Count -gt 0) {
    $manifestPath = Join-Path $OutputPath 'pipeline_manifest.json'
    $manifest = @{ manifestVersion = 1; entries = @{} }
    if (Test-Path $manifestPath) {
      try {
        $existing = Get-Content -Path $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($existing.entries) {
          foreach ($prop in $existing.entries.PSObject.Properties) {
            $manifest.entries[$prop.Name] = @{
              status       = $prop.Value.status
              processedAt  = $prop.Value.processedAt
              sourceScript = $prop.Value.sourceScript
            }
          }
        }
      }
      catch { Write-Warning "[Invoke-DoclingConversion] Could not read existing manifest: $_" }
    }
    foreach ($key in $manifestEntries.Keys) {
      $manifest.entries[$key] = $manifestEntries[$key]
    }
    $manifestJson = $manifest | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.Encoding]::UTF8)
    Write-Host "[Invoke-DoclingConversion] Manifest:     $manifestPath ($($manifestEntries.Count) entries)" -ForegroundColor Green
  }

  # ================================================================
  # SUMMARY
  # ================================================================
  $avgTime = if ($converted -gt 0) { [math]::Round($totalConvertTime / $converted, 2) } else { 0 }

  Write-Host "`n[Invoke-DoclingConversion] ================================" -ForegroundColor Green
  Write-Host "[Invoke-DoclingConversion] Converted:    $converted files" -ForegroundColor Green
  if ($failed -gt 0) { Write-Host "[Invoke-DoclingConversion] Failed:       $failed files" -ForegroundColor Red }
  if ($skippedExisting -gt 0) { Write-Host "[Invoke-DoclingConversion] Skipped:      $skippedExisting (already exist)" -ForegroundColor DarkGray }
  Write-Host "[Invoke-DoclingConversion] Duration:     $([math]::Round($batchDuration, 1))s (avg ${avgTime}s/file)" -ForegroundColor Green
  Write-Host "[Invoke-DoclingConversion] Output:       $OutputPath" -ForegroundColor Green
  Write-Host "[Invoke-DoclingConversion] ================================" -ForegroundColor Green
}
