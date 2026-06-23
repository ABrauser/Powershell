<#
.SYNOPSIS
  Vergleicht die hochgeladenen Dokumente zwischen zwei AnythingLLM-Instanzen und stellt das Ergebnis in einer interaktiven HTML-GUI dar.

.DESCRIPTION
  Dieses Skript analysiert die Dokumenten-Ordnerstrukturen von zwei AnythingLLM-Systemen (Quelle und Ziel).
  Es generiert einen interaktiven HTML-Bericht, in dem fehlende Dateien identifiziert, Ordner verglichen
  und gezielt PowerShell-Befehle generiert werden können, um fehlende Dateien aus einem lokalen Quellverzeichnis
  auf das Zielsystem zu übertragen.

  Das Skript unterstuetzt zwei Modi:
  1. Berichtsmodus (Standard): Query der Instanzen, Scan des optionalen lokalen Ordners und Oeffnen der interaktiven HTML-GUI.
  2. Sync-Modus (-ExecuteSync): Fuehrt den eigentlichen API-Upload fuer ausgewaehlte Dateien/Ordner durch.

.PARAMETER SourceUrl
  Die Basis-URL des Quell-AnythingLLM-Portals (z. B. https://rag.lua.rlp.de).

.PARAMETER TargetUrl
  Die Basis-URL des Ziel-AnythingLLM-Portals (z. B. http://srvkoap0055:3002).

.PARAMETER SourceApiKey
  API-Key fuer das Quell-System. Wenn nicht angegeben, wird der Session-Cache verwendet oder danach gefragt.

.PARAMETER TargetApiKey
  API-Key fuer das Ziel-System. Wenn nicht angegeben, wird der Session-Cache verwendet oder danach gefragt.

.PARAMETER LocalPath
  Der lokale Pfad, auf dem sich die physischen Originaldateien befinden (z. B. ein Desktop-Ordner oder Netzlaufwerk).
  Erleichtert den Sync, da Dateien von hier direkt hochgeladen werden koennen.

.PARAMETER SyncFiles
  Array von relativen Dateipfaden, die im Sync-Modus synchronisiert werden sollen (z. B. "Silo_1/doc.pdf").

.PARAMETER SyncFolders
  Array von Ordnernamen, die im Sync-Modus synchronisiert werden sollen. Alle darin fehlenden Dateien werden uebertragen.

.PARAMETER WorkspaceSlug
  Optionaler Workspace-Slug im Ziel-System. Wenn angegeben, werden die hochgeladenen Dateien auch an diesen Workspace angepinnt und eingebettet (Embeddings-Update).

.PARAMETER ExecuteSync
  Switch. Wenn gesetzt, fuehrt das Skript die Synchronisation aus, anstatt die HTML-GUI zu generieren.

.PARAMETER OutputPath
  Der Pfad, unter dem der HTML-Bericht gespeichert werden soll. Standard: AnythingLLM-Compare.html im Skriptverzeichnis.

.EXAMPLE
  # Standard-Aufruf (Vergleich ausfuehren und HTML-GUI oeffnen):
  .\Compare-AnythingLLMWorkspaces.ps1 -SourceUrl "https://rag.lua.rlp.de" -TargetUrl "http://srvkoap0055:3002" -LocalPath "\\lua.rlp.de\system$\Ordnerumleitung\koazd\rauserb\Desktop\Dokumente"

.EXAMPLE
  # Sync-Aufruf (wird i.d.R. aus der GUI kopiert und ausgefuehrt):
  .\Compare-AnythingLLMWorkspaces.ps1 -TargetUrl "http://srvkoap0055:3002" -LocalPath "\\lua.rlp.de\system$\Ordnerumleitung\koazd\rauserb\Desktop\Dokumente" -SyncFiles @("Silo_1/bericht.pdf") -ExecuteSync
#>
[CmdletBinding()]
param(
  [string]$SourceUrl,
  [string]$TargetUrl,
  [string]$SourceApiKey,
  [string]$TargetApiKey,
  [string]$LocalPath,
  [string[]]$SyncFiles,
  [string[]]$SyncFolders,
  [string]$WorkspaceSlug,
  [switch]$ExecuteSync,
  [string]$OutputPath
)

$ErrorActionPreference = "Stop"

# Cache-Pfade fuer API-Keys
$sourceKeyPath = [System.IO.Path]::Combine($env:TEMP, 'allm_compare_source_apikey.xml')
$targetKeyPath = [System.IO.Path]::Combine($env:TEMP, 'allm_compare_target_apikey.xml')

# Cache-Pfade fuer URLs
$sourceUrlCachePath = [System.IO.Path]::Combine($env:TEMP, 'allm_compare_source_url.txt')
$targetUrlCachePath = [System.IO.Path]::Combine($env:TEMP, 'allm_compare_target_url.txt')

# Zeitstempel-Hilfsfunktion
$ts = { Get-Date -Format 'HH:mm:ss' }

# ── API-KEY COOLDOWN & ABRUFMETHODE ──
function Get-ApiKey {
  param(
    [string]$Key,
    [string]$CachePath,
    [string]$SystemName
  )
  if ($Key) { return $Key }
  if (Test-Path $CachePath) {
    try {
      $secure = Import-Clixml $CachePath
      $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
      $decrypted = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
      [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
      Write-Host "[& $ts] [API] API-Key fuer $SystemName geladen." -ForegroundColor DarkGray
      return $decrypted
    } catch {
      Write-Warning "Konnte API-Key fuer $SystemName nicht aus Cache laden."
    }
  }
  
  $secKey = Read-Host "API-Key fuer $SystemName eingeben" -AsSecureString
  $secKey | Export-Clixml $CachePath
  $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secKey)
  $decrypted = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
  [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
  Write-Host "[& $ts] [API] API-Key fuer $SystemName gespeichert." -ForegroundColor Green
  return $decrypted
}

# ── REC-FUNCTION ZUM EXTRAHIEREN AUS DER LOCALFILES-STRUKTUR ──
function Get-FlatFileList {
  param(
    $Node,
    [string]$CurrentPath = ""
  )
  $files = @()
  if (-not $Node) { return $files }
  
  if ($Node.type -eq "folder") {
    foreach ($item in $Node.items) {
      $newPath = if ($CurrentPath -eq "") { $item.name } else { "$CurrentPath/$($item.name)" }
      if ($item.type -eq "folder") {
        $files += Get-FlatFileList -Node $item -CurrentPath $newPath
      } else {
        $files += [PSCustomObject]@{
          Name         = $item.name
          RelativePath = $newPath
          Folder       = $CurrentPath
          Size         = if ($item.size) { [long]$item.size } else { 0 }
          Id           = $item.id
        }
      }
    }
  } else {
    $files += [PSCustomObject]@{
      Name         = $Node.name
      RelativePath = $CurrentPath
      Folder       = [System.IO.Path]::GetDirectoryName($CurrentPath).Replace('\', '/')
      Size         = if ($Node.size) { [long]$Node.size } else { 0 }
      Id           = $Node.id
    }
  }
  return $files
}

# ── SCHNELLER LOKALER DATEISCAN MIT FORTSCHRITT ──
function Get-LocalFilesWithProgress {
  param(
    [string]$RootPath
  )
  $files = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
  $queue = [System.Collections.Generic.Queue[string]]::new()
  $queue.Enqueue($RootPath)
  
  $folderCount = 0
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  
  while ($queue.Count -gt 0) {
    $currentDir = $queue.Dequeue()
    $folderCount++
    
    if ($folderCount % 5 -eq 0) {
      $elapsed = $sw.Elapsed.ToString('mm\:ss')
      Write-Progress -Activity "Scanne lokale Dateien" `
        -Status "Ordner analysiert: $folderCount | Dateien gefunden: $($files.Count) | Dauer: $elapsed" `
        -PercentComplete -1
    }
    
    try {
      # Subdirectories holen
      $subDirs = [System.IO.Directory]::GetDirectories($currentDir)
      foreach ($sd in $subDirs) {
        $queue.Enqueue($sd)
      }
      
      # Dateien holen
      $dirFiles = [System.IO.Directory]::GetFiles($currentDir)
      foreach ($f in $dirFiles) {
        $files.Add([System.IO.FileInfo]::new($f))
      }
    }
    catch {
      # Ignoriere Berechtigungsfehler etc.
    }
  }
  
  $sw.Stop()
  Write-Progress -Activity "Scanne lokale Dateien" -Completed
  return $files
}

# ── SINGLE-FILE UPLOAD IN ANYTHINLLM ──
function Upload-FileToAnythingLLM {
  param(
    [string]$Url,
    [string]$ApiKey,
    [string]$FilePath,
    [string]$DocumentFolder
  )
  
  $cleanUrl = $Url.TrimEnd('/')
  $uploadUrl = if ($DocumentFolder) {
    "$cleanUrl/api/v1/document/upload/$DocumentFolder"
  } else {
    "$cleanUrl/api/v1/document/upload"
  }
  
  try {
    $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
    $fileName = [System.IO.Path]::GetFileName($FilePath)

    $boundary = [System.Guid]::NewGuid().ToString('N')
    $LF = "`r`n"
    $enc = [System.Text.Encoding]::UTF8
    $bodyParts = [System.Collections.Generic.List[byte[]]]::new()

    # File-Abschnitt erstellen
    $fileHeader = "--$boundary$LF" +
                  "Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"$LF" +
                  "Content-Type: application/octet-stream$LF$LF"
    $bodyParts.Add($enc.GetBytes($fileHeader))
    $bodyParts.Add($fileBytes)
    $bodyParts.Add($enc.GetBytes($LF))

    # Boundary schliessen
    $bodyParts.Add($enc.GetBytes("--$boundary--$LF"))

    # Zu einem Byte-Array zusammenfuehren
    $totalLen = 0; foreach ($p in $bodyParts) { $totalLen += $p.Length }
    $bodyBytes = [byte[]]::new($totalLen)
    $bOffset = 0
    foreach ($p in $bodyParts) {
      [System.Array]::Copy($p, 0, $bodyBytes, $bOffset, $p.Length)
      $bOffset += $p.Length
    }
    
    $contentType = "multipart/form-data; boundary=$boundary"
    $headers = @{
      'Authorization' = "Bearer $ApiKey"
      'Accept'        = 'application/json'
    }

    $response = Invoke-WebRequest -Uri $uploadUrl -Method Post -ContentType $contentType `
      -Body $bodyBytes -Headers $headers -TimeoutSec 300 -ErrorAction Stop -UseBasicParsing

    $responseObj = $response.Content | ConvertFrom-Json

    if ($responseObj.success) {
      if ($responseObj.documents -and $responseObj.documents.Count -gt 0) {
        return $responseObj.documents[0].location
      }
      return "uploaded"
    } else {
      Write-Warning "Server hat Fehler gemeldet: $($responseObj.error)"
      return $null
    }
  }
  catch {
    Write-Warning "Fehler beim Upload der Datei $($fileName): $($_.Exception.Message)"
    return $null
  }
}

# ── EMBEDDINGS AN WORKSPACE GEBEN ──
function Update-WorkspaceEmbeddings {
  param(
    [string]$Url,
    [string]$ApiKey,
    [string]$WorkspaceSlug,
    [string]$Location
  )
  
  $cleanUrl = $Url.TrimEnd('/')
  $uri = "$cleanUrl/api/v1/workspace/$WorkspaceSlug/update-embeddings"
  
  $embedPath = $Location
  if ($embedPath -notlike "custom-documents/*") {
    $embedPath = "custom-documents/$embedPath"
  }
  
  $body = @{
    adds    = @($embedPath)
    deletes = @()
  } | ConvertTo-Json -Depth 3
  
  $headers = @{
    'Authorization' = "Bearer $ApiKey"
    'Content-Type'  = 'application/json'
  }
  
  try {
    $resp = Invoke-RestMethod -Uri $uri -Method Post -Body $body -Headers $headers -TimeoutSec 300 -ErrorAction Stop
    return $true
  }
  catch {
    Write-Warning "Fehler beim Hinzufuegen der Datei zum Workspace '$WorkspaceSlug': $($_.Exception.Message)"
    return $false
  }
}

# ── DOKUMENTEN-STRUKTUR ABRUFEN ──
function Get-AnythingLLMDocumentTree {
  param(
    [string]$Url,
    [string]$ApiKey,
    [string]$SystemName = "AnythingLLM"
  )
  $headers = @{
    'Authorization' = "Bearer $ApiKey"
    'Accept'        = 'application/json'
  }
  
  $cleanUrl = $Url.TrimEnd('/')
  $uri = "$cleanUrl/api/v1/documents"
  
  Write-Host "[$(& $ts)] Verbinde mit $SystemName ($cleanUrl)..." -ForegroundColor Cyan
  
  $scriptBlock = {
    param($uri, $headers)
    Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -TimeoutSec 45 -ErrorAction Stop
  }
  
  $runspace = [PowerShell]::Create()
  $null = $runspace.AddScript($scriptBlock).AddArgument($uri).AddArgument($headers)
  
  $sw = [System.Diagnostics.Stopwatch]::StartNew()
  $asyncResult = $runspace.BeginInvoke()
  
  $spinners = @('|', '/', '-', '\')
  $spinIndex = 0
  
  while (-not $asyncResult.IsCompleted) {
    Start-Sleep -Milliseconds 150
    $spinChar = $spinners[$spinIndex]
    $spinIndex = ($spinIndex + 1) % $spinners.Count
    $elapsed = $sw.Elapsed.ToString('ss\.ff') + "s"
    
    Write-Progress -Activity "Dokumenten-Struktur laden" `
      -Status "Lade Daten von $SystemName... [$elapsed] [$spinChar]" `
      -PercentComplete -1
  }
  
  $sw.Stop()
  Write-Progress -Activity "Dokumenten-Struktur laden" -Completed
  
  $resp = $null
  try {
    $resp = $runspace.EndInvoke($asyncResult)
  }
  catch {
    Write-Error "Fehler beim Abrufen der Dokumenten-Struktur von $cleanUrl : $($_.Exception.Message)"
  }
  
  $runspace.Dispose()
  
  if ($resp -and $resp.localFiles) {
    return $resp.localFiles
  } else {
    Write-Warning "Keine 'localFiles' Struktur im System an $cleanUrl vorhanden."
    return $null
  }
}

# ─────────────────────────────────────────────────────────
#  HAUPT-ABLAUF
# ─────────────────────────────────────────────────────────

# Bereinigung des lokalen Pfads von eventuellen Zeilenumbrüchen und Einrückungen (z.B. durch Kopieren in der PowerShell-Konsole)
if ($LocalPath) {
  $LocalPath = $LocalPath.Replace("`r", "").Replace("`n", "").Trim()
  $LocalPath = $LocalPath -replace '\s*[\r\n]+\s*', ''
  $LocalPath = $LocalPath.Trim()
}

# URLs laden, abfragen und cachen
if (-not $SourceUrl) {
  if (Test-Path $sourceUrlCachePath) {
    $SourceUrl = Get-Content $sourceUrlCachePath -Raw
  } else {
    $SourceUrl = Read-Host "Quell-AnythingLLM URL eingeben (z.B. https://rag.lua.rlp.de)"
  }
}
$SourceUrl = $SourceUrl.Trim().TrimEnd('/')
$SourceUrl | Out-File -FilePath $sourceUrlCachePath -Encoding utf8 -NoNewline

if (-not $TargetUrl) {
  if (Test-Path $targetUrlCachePath) {
    $TargetUrl = Get-Content $targetUrlCachePath -Raw
  } else {
    $TargetUrl = Read-Host "Ziel-AnythingLLM URL eingeben (z.B. http://srvkoap0055:3002)"
  }
}
$TargetUrl = $TargetUrl.Trim().TrimEnd('/')
$TargetUrl | Out-File -FilePath $targetUrlCachePath -Encoding utf8 -NoNewline

if ($ExecuteSync) {
  # ==========================================
  #  MODUS 2: DOCKER-SYNC-AUSFUEHRUNG
  # ==========================================
  
  if (-not $LocalPath -or -not (Test-Path $LocalPath)) {
    Write-Error "Der lokale Pfad (-LocalPath) ist ungueltig oder wurde nicht angegeben: '$LocalPath'"
    return
  }
  
  $TargetApiKey = Get-ApiKey -Key $TargetApiKey -CachePath $targetKeyPath -SystemName "Ziel-System ($TargetUrl)"
  if (-not $TargetApiKey) {
    Write-Error "Ziel-API-Key erforderlich fuer den Sync."
    return
  }
  
  $localRoot = (Resolve-Path $LocalPath).Path
  $filesToSync = @()
  
  if ($SyncFiles -and $SyncFiles.Count -gt 0) {
    $filesToSync += $SyncFiles
  }
  
  # Falls ganze Ordner angegeben wurden, vergleichen wir sie live
  if ($SyncFolders -and $SyncFolders.Count -gt 0) {
    Write-Host "[$(& $ts)] Analysiere Ordner fuer Sync: $($SyncFolders -join ', ')..." -ForegroundColor Cyan
    $SourceApiKey = Get-ApiKey -Key $SourceApiKey -CachePath $sourceKeyPath -SystemName "Quell-System ($SourceUrl)"
    
    $sourceTree = Get-AnythingLLMDocumentTree -Url $SourceUrl -ApiKey $SourceApiKey -SystemName "Quell-System"
    $targetTree = Get-AnythingLLMDocumentTree -Url $TargetUrl -ApiKey $TargetApiKey -SystemName "Ziel-System"
    
    $sourceFiles = Get-FlatFileList -Node $sourceTree
    $targetFiles = Get-FlatFileList -Node $targetTree
    
    $targetPaths = @{}
    foreach ($tf in $targetFiles) {
      $targetPaths[$tf.RelativePath.ToLower()] = $true
    }
    
    foreach ($sf in $sourceFiles) {
      $inSyncFolder = $false
      foreach ($folder in $SyncFolders) {
        $folderNormalized = if ($folder -eq "(Wurzelverzeichnis)") { "" } else { $folder }
        $sfFolder = if ($sf.Folder) { $sf.Folder } else { "" }
        if (($folderNormalized -ne "" -and $sf.RelativePath.StartsWith($folderNormalized + "/", [System.StringComparison]::OrdinalIgnoreCase)) -or 
            ($folderNormalized -eq "" -and -not $sf.RelativePath.Contains("/")) -or
            $sfFolder -eq $folderNormalized) {
          $inSyncFolder = $true
          break
        }
      }
      
      if ($inSyncFolder) {
        if (-not $targetPaths.ContainsKey($sf.RelativePath.ToLower())) {
          $filesToSync += $sf.RelativePath
        }
      }
    }
  }
  
  if ($filesToSync.Count -eq 0) {
    Write-Host "[$(& $ts)] Keine fehlenden Dateien zum Synchronisieren gefunden." -ForegroundColor Yellow
    return
  }
  
  Write-Host "[$(& $ts)] Starte Synchronisation von $($filesToSync.Count) Dateien..." -ForegroundColor Cyan
  
  $successCount = 0
  $failCount = 0
  $totalFiles = $filesToSync.Count
  $currentIndex = 0
  
  foreach ($relPath in $filesToSync) {
    $currentIndex++
    $pct = [int]($currentIndex / $totalFiles * 100)
    Write-Progress -Activity "Synchronisiere Dateien" -Status "Upload: $currentIndex von $totalFiles ($relPath) [$pct%]" -PercentComplete $pct
    
    $localFilePath = Join-Path $localRoot ($relPath.Replace('/', '\'))
    if (-not (Test-Path $localFilePath)) {
      Write-Warning "Lokale Datei existiert nicht: $localFilePath"
      $failCount++
      continue
    }
    
    # Ordner ermitteln (relativer Pfad minus Dateiname)
    $fileName = [System.IO.Path]::GetFileName($localFilePath)
    $docFolder = ""
    if ($relPath.Contains("/")) {
      $docFolder = $relPath.Substring(0, $relPath.Length - $fileName.Length - 1)
    }
    
    Write-Host "[$(& $ts)] Upload: '$relPath'..." -ForegroundColor Cyan
    $loc = Upload-FileToAnythingLLM -Url $TargetUrl -ApiKey $TargetApiKey -FilePath $localFilePath -DocumentFolder $docFolder
    
    if ($loc) {
      $successCount++
      Write-Host "  [OK] Datei hochgeladen -> $loc" -ForegroundColor Green
      
      # Falls gewuenscht, auch in einen Workspace einbetten
      if ($WorkspaceSlug) {
        Write-Host "  Füge zu Workspace '$WorkspaceSlug' hinzu..." -ForegroundColor DarkGray
        $embedOk = Update-WorkspaceEmbeddings -Url $TargetUrl -ApiKey $TargetApiKey -WorkspaceSlug $WorkspaceSlug -Location $loc
        if ($embedOk) {
          Write-Host "  [OK] Workspace-Embeddings aktualisiert." -ForegroundColor Green
        } else {
          Write-Warning "  [WARN] Embeddings konnten nicht aktualisiert werden."
        }
      }
    } else {
      $failCount++
      Write-Warning "  [FEHLER] Upload fehlgeschlagen: $relPath"
    }
  }
  Write-Progress -Activity "Synchronisiere Dateien" -Completed
  
  Write-Host "[$(& $ts)] Synchronisation abgeschlossen. Erfolg: $successCount, Fehler: $failCount" -ForegroundColor Green
  return
}

# ==========================================
#  MODUS 1: VERGLEICH & HTML-GUI GENERIERUNG
# ==========================================

Write-Host "[$(& $ts)] Starte AnythingLLM Workspace Vergleich..." -ForegroundColor Cyan

$SourceApiKey = Get-ApiKey -Key $SourceApiKey -CachePath $sourceKeyPath -SystemName "Quell-System ($SourceUrl)"
$TargetApiKey = Get-ApiKey -Key $TargetApiKey -CachePath $targetKeyPath -SystemName "Ziel-System ($TargetUrl)"

# Dokumenten-Struktur holen
$sourceTree = Get-AnythingLLMDocumentTree -Url $SourceUrl -ApiKey $SourceApiKey -SystemName "Quell-System"
$targetTree = Get-AnythingLLMDocumentTree -Url $TargetUrl -ApiKey $TargetApiKey -SystemName "Ziel-System"

if (-not $sourceTree) {
  Write-Error "Dokumentenstruktur der Quelle konnte nicht geladen werden."
  return
}
if (-not $targetTree) {
  Write-Error "Dokumentenstruktur des Ziels konnte nicht geladen werden."
  return
}

# Flache Listen generieren
$sourceFiles = Get-FlatFileList -Node $sourceTree
$targetFiles = Get-FlatFileList -Node $targetTree

Write-Host "[$(& $ts)] Quelle: $($sourceFiles.Count) Dateien gefunden." -ForegroundColor Green
Write-Host "[$(& $ts)] Ziel: $($targetFiles.Count) Dateien gefunden." -ForegroundColor Green

# Lokale Dateien scannen falls Pfad angegeben
$localFilesHash = @{}
$localRoot = ""
if ($LocalPath) {
  if (Test-Path $LocalPath) {
    $localRoot = (Resolve-Path $LocalPath).Path
    Write-Host "[$(& $ts)] Scanne lokale Dateien in '$localRoot'..." -ForegroundColor Cyan
    
    $allLocal = Get-LocalFilesWithProgress -RootPath $localRoot
    
    $totalLocal = $allLocal.Count
    $processedLocal = 0
    foreach ($lf in $allLocal) {
      $processedLocal++
      if ($processedLocal % 500 -eq 0 -or $processedLocal -eq $totalLocal) {
        $pct = [int]($processedLocal / $totalLocal * 100)
        Write-Host -NoNewline "`r[Compare] Verarbeite lokale Dateien... $processedLocal / $totalLocal ($pct%)"
      }
      $rel = ""
      if ($lf.FullName.StartsWith($localRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $rel = $lf.FullName.Substring($localRoot.Length).TrimStart('\').Replace('\', '/')
      } else {
        $rel = $lf.Name
      }
      $localFilesHash[$rel.ToLower()] = $lf.FullName
    }
    Write-Host -NoNewline "`r[Compare] Verarbeite lokale Dateien... Fertig! ($processedLocal Dateien)                     `n"
    Write-Host "[$(& $ts)] $($_ = $localFilesHash.Count) lokale Dateien gefunden." -ForegroundColor Green
  } else {
    Write-Warning "Lokaler Pfad existiert nicht: $LocalPath"
  }
}

# Verzeichnis-Vergleich aufbauen
$sourceHash = @{}
foreach ($sf in $sourceFiles) { $sourceHash[$sf.RelativePath.ToLower()] = $sf }

$targetHash = @{}
foreach ($tf in $targetFiles) { $targetHash[$tf.RelativePath.ToLower()] = $tf }

$allKeys = @{}
foreach ($k in $sourceHash.Keys) { $allKeys[$k] = $true }
foreach ($k in $targetHash.Keys) { $allKeys[$k] = $true }

$totalKeys = $allKeys.Count
$counter = 0
$comparisonList = foreach ($key in ($allKeys.Keys | Sort-Object)) {
  $counter++
  if ($counter % 500 -eq 0 -or $counter -eq $totalKeys) {
    $pct = [int]($counter / $totalKeys * 100)
    Write-Host -NoNewline "`r[Compare] Erstelle Vergleichsdaten... $counter / $totalKeys ($pct%)"
  }
  
  $sf = $sourceHash[$key]
  $tf = $targetHash[$key]
  
  $name = if ($sf) { $sf.Name } else { $tf.Name }
  $folder = if ($sf) { $sf.Folder } else { $tf.Folder }
  if ($folder -eq "") { $folder = "(Wurzelverzeichnis)" }
  
  $inLocal = $false
  $localSize = 0
  if ($localFilesHash.ContainsKey($key)) {
    $inLocal = $true
    try {
      $localSize = (Get-Item $localFilesHash[$key]).Length
    } catch { }
  }
  
  [PSCustomObject]@{
    RelativePath = $key
    Name         = $name
    Folder       = $folder
    InSource     = [bool]$sf
    InTarget     = [bool]$tf
    InLocal      = $inLocal
    SourceSize   = if ($sf) { $sf.Size } else { 0 }
    TargetSize   = if ($tf) { $tf.Size } else { 0 }
    LocalSize    = $localSize
  }
}
Write-Host -NoNewline "`r[Compare] Erstelle Vergleichsdaten... Fertig! ($totalKeys Dokumente)                      `n"

$comparisonJson = $comparisonList | ConvertTo-Json -Depth 3 -Compress
$localPathJs = if ($localRoot) { $localRoot } else { '' }

# HTML Generieren (Template mit Platzhaltern)
if (-not $OutputPath) {
  $OutputPath = Join-Path $PSScriptRoot "AnythingLLM-Compare.html"
}

$htmlTemplate = @'
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AnythingLLM Workspace Vergleich</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet" onerror="this.remove()">
<link rel="stylesheet" href="https://cdn.datatables.net/2.3.7/css/dataTables.dataTables.min.css">
<style>
  :root {
    --bg-primary: #0E0E11;
    --bg-secondary: #131314;
    --bg-card: #1E1F20;
    --bg-glass: rgba(255, 255, 255, 0.05);
    --border-glass: rgba(255, 255, 255, 0.1);
    --text-primary: #E3E3E3;
    --text-secondary: #C4C7C5;
    --text-muted: #8E918F;
    --accent: #A8C7FA;
    --accent-glow: rgba(168, 199, 250, 0.1);
    --success: #6DD58C;
    --success-glow: rgba(109, 213, 140, 0.12);
    --warning: #FFB74D;
    --warning-glow: rgba(255, 183, 77, 0.12);
    --danger: #E25C5C;
    --danger-glow: rgba(226, 92, 92, 0.12);
    --radius-card: 24px;
    --radius-btn: 20px;
    --font-head: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  }
  
  * { margin: 0; padding: 0; box-sizing: border-box; }
  
  body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg-primary); color: var(--text-primary); line-height: 1.6; min-height: 100vh;
  }
  
  .container { max-width: 1400px; margin: 0 auto; padding: 2rem; }
  
  header {
    background: var(--bg-card); border: 1px solid var(--border-glass);
    border-radius: var(--radius-card); padding: 1.8rem 2rem; margin-bottom: 1.5rem;
  }
  
  header h1 {
    font-size: 1.8rem; font-weight: 500; letter-spacing: -0.02em; margin-bottom: 0.3rem;
  }
  
  header h1 span {
    background: linear-gradient(90deg, #4285F4, #9B72CB, #D96570);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; font-weight: 600;
  }
  
  header .sub { font-size: 0.85rem; color: var(--text-muted); }
  
  .systems-panel {
    display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1.5rem;
  }
  
  @media (max-width: 768px) {
    .systems-panel { grid-template-columns: 1fr; }
  }
  
  .sys-card {
    background: var(--bg-card); border: 1px solid var(--border-glass);
    border-radius: 16px; padding: 1.2rem;
  }
  
  .sys-card h3 { font-size: 0.9rem; color: var(--text-muted); text-transform: uppercase; margin-bottom: 0.4rem; letter-spacing: 0.05em; }
  .sys-card .url { font-family: 'JetBrains Mono', monospace; font-size: 0.85rem; color: var(--accent); word-break: break-all; }
  
  .stats-grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-bottom: 1.5rem;
  }
  
  .stat-card {
    background: var(--bg-card); border: 1px solid var(--border-glass);
    border-radius: 16px; padding: 1.2rem; text-align: center;
    transition: transform 0.2s, border-color 0.2s;
  }
  
  .stat-card:hover { transform: translateY(-2px); border-color: rgba(168, 199, 250, 0.2); }
  .stat-card .val { font-size: 1.8rem; font-weight: 700; color: var(--accent); }
  .stat-card.ok .val { color: var(--success); }
  .stat-card.warn .val { color: var(--warning); }
  .stat-card.danger .val { color: var(--danger); }
  .stat-card .lbl { font-size: 0.72rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; margin-top: 0.2rem; }
  
  .main-layout {
    display: flex;
    flex-direction: column;
    gap: 1.5rem;
    margin-bottom: 1.5rem;
  }
  
  .sidebar-panel {
    background: var(--bg-card); border: 1px solid var(--border-glass);
    border-radius: var(--radius-card); padding: 1.5rem;
  }
  
  .sidebar-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 0.8rem;
    flex-wrap: wrap;
    gap: 0.5rem;
  }
  
  .sidebar-panel h2 { font-size: 1.05rem; font-weight: 500; margin: 0; }
  
  .sidebar-actions {
    display: flex; gap: 0.5rem; margin: 0;
  }
  
  .btn-sm {
    padding: 0.3rem 0.6rem; border-radius: 6px; border: 1px solid var(--border-glass);
    background: transparent; color: var(--text-secondary); cursor: pointer; font-size: 0.72rem;
    transition: all 0.2s;
  }
  .btn-sm:hover { border-color: var(--accent); color: var(--accent); }
  
  .folder-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 0.6rem;
    max-height: 180px;
    overflow-y: auto;
    border: 1px solid var(--border-glass);
    background: var(--bg-secondary);
    border-radius: 8px;
    padding: 0.8rem;
  }
  
  .folder-item {
    display: inline-flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.4rem 0.8rem;
    border-radius: 20px;
    background: var(--bg-glass);
    border: 1px solid var(--border-glass);
    font-size: 0.8rem;
    cursor: pointer;
    transition: all 0.2s;
    user-select: none;
  }
  .folder-item:hover {
    background: var(--table-hover);
    border-color: var(--accent);
  }
  .folder-item.active {
    background: var(--accent-glow);
    border-color: var(--accent);
    color: var(--accent);
  }
  
  .folder-item input[type="checkbox"] { accent-color: var(--accent); cursor: pointer; }
  .folder-item .folder-name { max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .folder-item .folder-badges { display: flex; gap: 3px; align-items: center; }
  .folder-item .f-badge { font-size: 0.65rem; padding: 1px 5px; border-radius: 10px; }
  .folder-item .f-badge.src { background: rgba(168, 199, 250, 0.1); color: var(--accent); }
  .folder-item .f-badge.trg { background: rgba(255, 255, 255, 0.05); color: var(--text-muted); }
  .folder-item .f-badge.miss { background: var(--danger-glow); color: var(--danger); font-weight: bold; }
  
  .content-panel {
    background: var(--bg-card); border: 1px solid var(--border-glass);
    border-radius: var(--radius-card); padding: 1.5rem;
  }
  
  .content-panel h2 { font-size: 1.05rem; font-weight: 500; margin-bottom: 1rem; }
  
  .filters {
    display: flex; gap: 0.6rem; margin-bottom: 1rem; flex-wrap: wrap; align-items: center;
  }
  
  .filters input[type="text"] {
    flex: 1; min-width: 200px; padding: 0.5rem 0.8rem; border-radius: 8px;
    border: 1px solid var(--border-glass); background: var(--bg-secondary);
    color: var(--text-primary); font-size: 0.82rem; outline: none;
    transition: border-color 0.2s;
  }
  
  .filters input[type="text"]:focus { border-color: var(--accent); }
  
  .fbtn {
    padding: 0.4rem 0.8rem; border-radius: 8px; border: 1px solid var(--border-glass);
    background: transparent; color: var(--text-secondary); cursor: pointer; font-size: 0.78rem;
    transition: all 0.2s;
  }
  
  .fbtn:hover { border-color: var(--accent); color: var(--accent); }
  .fbtn.on { background: var(--accent); color: #111; border-color: var(--accent); font-weight: 500; }
  
  .tbl-container { overflow-x: auto; margin-top: 0.5rem; border: 1px solid var(--border-glass); border-radius: 8px; background: var(--bg-secondary); }
  
  table { width: 100%; border-collapse: collapse; font-size: 0.82rem; text-align: left; }
  
  th {
    padding: 0.7rem 0.8rem; border-bottom: 2px solid var(--border-glass);
    color: var(--text-muted); font-weight: 600; text-transform: uppercase; font-size: 0.7rem;
    letter-spacing: 0.05em; cursor: pointer; user-select: none;
    background: var(--bg-secondary);
    position: sticky; top: 0; z-index: 10;
  }
  th:hover { color: var(--accent); }
  
  td { padding: 0.6rem 0.8rem; border-bottom: 1px solid rgba(255,255,255,0.03); }
  
  tr:hover { background: var(--accent-glow); }
  tr.missing-trg { background: rgba(226, 92, 92, 0.04); }
  tr.missing-trg:hover { background: rgba(226, 92, 92, 0.07); }
  tr.missing-src { background: rgba(255, 183, 77, 0.03); }
  tr.missing-src:hover { background: rgba(255, 183, 77, 0.06); }
  
  .badge {
    display: inline-block; padding: 2px 6px; border-radius: 4px;
    font-size: 0.65rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em;
  }
  .badge.ok { background: var(--success-glow); color: var(--success); }
  .badge.warn { background: var(--warning-glow); color: var(--warning); }
  .badge.danger { background: var(--danger-glow); color: var(--danger); }
  
  .mono { font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: var(--text-muted); }
  
  .sync-panel {
    background: var(--bg-card); border: 1px solid var(--border-glass);
    border-radius: var(--radius-card); padding: 1.8rem 2rem; margin-top: 1.5rem;
  }
  
  .sync-panel h2 { font-size: 1.1rem; font-weight: 500; margin-bottom: 0.8rem; }
  .sync-panel h2 span { color: var(--accent); }
  
  .sync-tabs {
    display: flex; gap: 0.5rem; margin-bottom: 1rem; border-bottom: 1px solid var(--border-glass); padding-bottom: 0.5rem;
  }
  
  .sync-tab {
    padding: 0.4rem 1rem; border-radius: 6px; border: 1px solid transparent;
    background: transparent; color: var(--text-secondary); cursor: pointer; font-size: 0.8rem;
    transition: all 0.2s;
  }
  .sync-tab.active { background: var(--bg-secondary); border-color: var(--border-glass); color: var(--accent); font-weight: 500; }
  
  .field { margin-bottom: 1.2rem; }
  .field label {
    font-size: 0.78rem; font-weight: 600; color: var(--text-secondary); display: block; margin-bottom: 0.4rem;
  }
  .field input[type="text"] {
    width: 100%; padding: 0.6rem 0.8rem; border-radius: 8px;
    border: 1px solid var(--border-glass); background: var(--bg-secondary);
    color: var(--text-primary); font-size: 0.82rem;
    font-family: 'JetBrains Mono', monospace; outline: none;
  }
  .field input[type="text"]:focus { border-color: var(--accent); }
  
  .cmd-section { margin-top: 1.2rem; }
  .cmd-section label {
    font-size: 0.78rem; font-weight: 600; color: var(--text-secondary); display: block; margin-bottom: 0.3rem;
  }
  .cmd-wrap { position: relative; }
  .cmd-pre {
    background: var(--bg-secondary); border: 1px solid var(--border-glass); border-radius: 8px;
    padding: 1.2rem;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem; color: var(--text-primary); white-space: pre-wrap;
    word-break: break-all; margin: 0; cursor: pointer; min-height: 4.5rem;
  }
  
  .btn-copy {
    position: absolute; top: 0.4rem; right: 0.4rem; background: var(--accent); color: #111;
    border: none; border-radius: 6px; padding: 0.3rem 0.8rem; font-size: 0.72rem;
    cursor: pointer; font-weight: 600; transition: opacity 0.2s;
  }
  .btn-copy:hover { opacity: 0.85; }
  .copy-msg {
    font-size: 0.72rem; color: var(--success); margin-top: 0.3rem;
    opacity: 0; transition: opacity 0.3s;
  }
  
  .col-chk { width: 40px; text-align: center; }
  .col-chk input[type="checkbox"] { accent-color: var(--accent); cursor: pointer; }
  .col-name { max-width: 500px; word-break: break-all; }
  .col-folder { max-width: 200px; word-break: break-all; }
  .col-badge { width: 110px; text-align: center; }
  
  .footer {
    text-align: center; font-size: 0.75rem; color: var(--text-muted);
    margin-top: 2rem; padding-top: 1rem; border-top: 1px solid var(--border-glass);
  }
  
  /* DataTables Theme Override */
  .dt-container {
    color: var(--text-secondary);
  }
  
  table.dataTable.compact {
    color: var(--text-secondary) !important;
    border-collapse: separate !important;
    border-spacing: 0;
    width: 100% !important;
    background: transparent !important;
    border: none !important;
    margin-top: 1rem !important;
    margin-bottom: 1rem !important;
  }
  
  table.dataTable.compact thead th {
    background: var(--bg-secondary) !important;
    color: var(--text-muted) !important;
    border-bottom: 2px solid var(--border-glass) !important;
    font-size: 0.72rem !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    padding: 6px 10px !important;
    font-weight: 600 !important;
  }
  
  table.dataTable.compact tbody td {
    padding: 5px 10px !important;
    border-bottom: 1px solid rgba(255,255,255,0.03) !important;
    font-size: 0.78rem !important;
    vertical-align: middle !important;
    background: transparent !important;
    color: var(--text-secondary) !important;
  }
  
  table.dataTable tbody tr {
    background: transparent !important;
  }
  
  table.dataTable tbody tr:hover {
    background: var(--accent-glow) !important;
  }
  
  table.dataTable tbody tr.missing-trg {
    background: rgba(226, 92, 92, 0.03) !important;
  }
  table.dataTable tbody tr.missing-trg:hover {
    background: rgba(226, 92, 92, 0.06) !important;
  }
  
  table.dataTable tbody tr.missing-src {
    background: rgba(255, 183, 77, 0.02) !important;
  }
  table.dataTable tbody tr.missing-src:hover {
    background: rgba(255, 183, 77, 0.05) !important;
  }

  /* DataTables controls */
  .dt-container .dt-length, .dt-container .dt-search {
    margin-bottom: 1rem;
    font-size: 0.8rem;
  }
  
  .dt-container .dt-search input {
    background: var(--bg-secondary) !important;
    border: 1px solid var(--border-glass) !important;
    border-radius: 8px !important;
    color: var(--text-primary) !important;
    padding: 0.4rem 0.8rem !important;
    outline: none;
    margin-left: 0.5rem;
  }
  
  .dt-container .dt-search input:focus {
    border-color: var(--accent) !important;
  }
  
  .dt-container .dt-length select {
    background: var(--bg-secondary) !important;
    border: 1px solid var(--border-glass) !important;
    border-radius: 8px !important;
    color: var(--text-primary) !important;
    padding: 0.3rem 0.6rem !important;
    outline: none;
    margin-left: 0.3rem;
    margin-right: 0.3rem;
  }
  
  /* Pagination styling */
  .dt-container .dt-paging {
    margin-top: 1rem;
    font-size: 0.8rem;
    display: flex;
    justify-content: flex-end;
    align-items: center;
    gap: 4px;
  }
  
  .dt-container .dt-paging button,
  .dt-container .dt-paging .dt-paging-button {
    background: transparent !important;
    color: var(--text-secondary) !important;
    border: 1px solid var(--border-glass) !important;
    border-radius: 8px !important;
    padding: 0.4rem 0.8rem !important;
    cursor: pointer;
    transition: all 0.2s;
    outline: none;
    font-family: inherit;
  }
  
  .dt-container .dt-paging button:hover:not(.disabled):not(.current),
  .dt-container .dt-paging .dt-paging-button:hover:not(.disabled):not(.current) {
    background: var(--border-glass) !important;
    color: var(--accent) !important;
    border-color: var(--accent) !important;
  }
  
  .dt-container .dt-paging button.current,
  .dt-container .dt-paging .dt-paging-button.current {
    background: var(--accent) !important;
    color: #111 !important;
    border-color: var(--accent) !important;
    font-weight: 600;
  }
  
  .dt-container .dt-paging button.disabled,
  .dt-container .dt-paging .dt-paging-button.disabled {
    opacity: 0.4 !important;
    cursor: not-allowed !important;
  }
  
  /* Remove sorting column markers from checkboxes */
  table.dataTable thead th.col-chk::before,
  table.dataTable thead th.col-chk::after {
    display: none !important;
  }
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>AnythingLLM <span>Workspace &bull; Vergleich</span></h1>
    <div class="sub">Interaktive GUI zum Abgleich hochgeladener Dokumente und Generierung von PowerShell-Sync-Befehlen.</div>
  </header>

  <div class="systems-panel">
    <div class="sys-card">
      <h3>Quelle (System A)</h3>
      <div class="url" id="lbl-source">https://rag.lua.rlp.de</div>
    </div>
    <div class="sys-card">
      <h3>Ziel (System B)</h3>
      <div class="url" id="lbl-target">http://srvkoap0055:3002</div>
    </div>
  </div>

  <div class="stats-grid">
    <div class="stat-card">
      <div class="val" id="st-total">0</div>
      <div class="lbl">Dateien Gesamt</div>
    </div>
    <div class="stat-card ok">
      <div class="val" id="st-match">0</div>
      <div class="lbl">Identisch (In Beiden)</div>
    </div>
    <div class="stat-card danger">
      <div class="val" id="st-miss-trg">0</div>
      <div class="lbl">Fehlen im Ziel (Sync n&ouml;tig)</div>
    </div>
    <div class="stat-card warn">
      <div class="val" id="st-miss-src">0</div>
      <div class="lbl">Fehlen in Quelle</div>
    </div>
    <div class="stat-card">
      <div class="val" id="st-local">0</div>
      <div class="lbl">Lokal verf&uuml;gbar</div>
    </div>
  </div>

  <div class="main-layout">
    <!-- Ordner-Filter oben positioniert -->
    <div class="sidebar-panel">
      <div class="sidebar-header">
        <h2>Unterordner (Silos)</h2>
        <div class="sidebar-actions">
          <button class="btn-sm" onclick="setAllFolders(true)">Alle</button>
          <button class="btn-sm" onclick="setAllFolders(false)">Keine</button>
          <button class="btn-sm" onclick="setFoldersWithMissingOnly()">Nur fehlende</button>
        </div>
      </div>
      <div class="folder-grid" id="folderList">
        <!-- JS-generierte Ordnerliste -->
      </div>
    </div>

    <!-- Content: Dateiliste -->
    <div class="content-panel">
      <h2>Dokumenten-Tabelle</h2>
      
      <div class="filters">
        <input type="text" id="txtSearch" placeholder="Dateiname suchen..." oninput="filterData()">
        <button class="fbtn on" id="btn-f-all" onclick="setFilter('all')">Alle</button>
        <button class="fbtn" id="btn-f-miss-trg" onclick="setFilter('missing-target')">Fehlen im Ziel</button>
        <button class="fbtn" id="btn-f-miss-src" onclick="setFilter('missing-source')">Fehlen in Quelle</button>
        <button class="fbtn" id="btn-f-match" onclick="setFilter('identical')">Identisch</button>
      </div>
      
      <div class="tbl-container">
        <table id="tblFiles" class="display compact" style="width:100%">
          <thead>
            <tr>
              <th class="col-chk"><input type="checkbox" id="chkAllFiles" onclick="toggleSelectAllFiles(this)"></th>
              <th class="col-name">Dateiname</th>
              <th class="col-folder">Unterordner</th>
              <th class="col-badge">Quelle</th>
              <th class="col-badge">Ziel</th>
              <th class="col-badge">Lokal</th>
            </tr>
          </thead>
          <tbody id="tblBody">
            <!-- JS-generierte Tabellenzeilen -->
          </tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- Sync Command Builder -->
  <div class="sync-panel">
    <h2>PowerShell <span>Sync Command Builder</span></h2>
    
    <div class="sync-tabs">
      <button class="sync-tab active" id="tab-files" onclick="setSyncTab('files')">Selektierte Dateien syncen</button>
      <button class="sync-tab" id="tab-folders" onclick="setSyncTab('folders')">Selektierte Ordner syncen</button>
    </div>

    <div class="field">
      <label>Lokaler Quell-Pfad (Verzeichnis der Originaldateien auf Festplatte / Netzlaufwerk)</label>
      <input type="text" id="local-path-input" placeholder="Z.B. \\lua.rlp.de\system$\Ordnerumleitung\koazd\rauserb\Desktop\Dokumente" oninput="updateLocalPath(this.value)">
    </div>

    <div class="field" id="workspace-field">
      <label>Ziel Workspace Slug (optional - wenn angegeben, werden die Dateien im Ziel an diesen Workspace angepinnt und eingebettet)</label>
      <input type="text" id="workspace-slug-input" placeholder="Z.B. wein-knowledge (leerlassen fuer nur Upload in Storage)" oninput="updateWorkspaceSlug(this.value)">
    </div>

    <div class="cmd-section">
      <label id="cmd-label">Auszuf&uuml;hrender PowerShell-Befehl</label>
      <div class="cmd-wrap">
        <pre class="cmd-pre" id="cmd-output" onclick="copyCmd()"></pre>
        <button class="btn-copy" onclick="copyCmd()">&#128203; Kopieren</button>
      </div>
      <div class="copy-msg" id="copy-msg">In Zwischenablage kopiert!</div>
    </div>
  </div>

  <div class="footer">AnythingLLM Workspace Compare &mdash; Generiert am {{GENERATED_TIME}}</div>
</div>

<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script src="https://cdn.datatables.net/2.3.7/js/dataTables.min.js"></script>
<script>
// Datenplatzhalter (vom PS-Skript befuellt)
const SOURCE_URL = '{{SOURCE_URL}}';
const TARGET_URL = '{{TARGET_URL}}';
const SCRIPT_PATH = '{{SCRIPT_PATH}}';
const INITIAL_LOCAL_PATH = '{{LOCAL_PATH}}';
const comparisonData = {{COMPARISON_JSON}};

// UI-State
let activeFilter = 'all'; // all, missing-target, missing-source, identical
let selectedFolders = new Set();
let selectedFiles = new Set();
let localPath = INITIAL_LOCAL_PATH || localStorage.getItem('allm_compare_localpath') || '';
let workspaceSlug = localStorage.getItem('allm_compare_workspace') || '';
let activeSyncTab = 'files'; // files, folders
let folderList = [];
let sortCol = 1;
let sortAsc = true;

// Init
document.getElementById('lbl-source').textContent = SOURCE_URL;
document.getElementById('lbl-target').textContent = TARGET_URL;
document.getElementById('local-path-input').value = localPath;
document.getElementById('workspace-slug-input').value = workspaceSlug;

// Ordner analysieren
const folderMap = {};
comparisonData.forEach(item => {
  const f = item.Folder;
  if (!folderMap[f]) {
    folderMap[f] = { name: f, total: 0, source: 0, target: 0, missingTarget: 0 };
  }
  folderMap[f].total++;
  if (item.InSource) folderMap[f].source++;
  if (item.InTarget) folderMap[f].target++;
  if (item.InSource && !item.InTarget) folderMap[f].missingTarget++;
});

folderList = Object.values(folderMap).sort((a,b) => a.name.localeCompare(b.name, 'de'));

// Standardmaessig alle Ordner selektieren
folderList.forEach(f => selectedFolders.add(f.name));

// Ordnerliste rendern
function renderFolders() {
  const el = document.getElementById('folderList');
  let html = '';
  folderList.forEach((f, idx) => {
    const isChecked = selectedFolders.has(f.name) ? 'checked' : '';
    const activeClass = selectedFolders.has(f.name) ? 'active' : '';
    const missBadge = f.missingTarget > 0 ? `<span class="f-badge miss" title="Fehlen im Ziel">${f.missingTarget}</span>` : '';
    
    html += `
      <div class="folder-item ${activeClass}" onclick="toggleFolderClick(event, '${f.name}', 'chk-fold-${idx}')">
        <input type="checkbox" id="chk-fold-${idx}" ${isChecked} onclick="event.stopPropagation(); toggleFolderCheck('${f.name}', this.checked)">
        <span class="folder-name" title="${f.name}">${f.name}</span>
        <div class="folder-badges">
          <span class="f-badge src" title="In Quelle">${f.source}</span>
          <span class="f-badge trg" title="In Ziel">${f.target}</span>
          ${missBadge}
        </div>
      </div>
    `;
  });
  el.innerHTML = html;
}

function toggleFolderClick(event, folderName, checkboxId) {
  if (event.target.tagName === 'INPUT') return;
  const chk = document.getElementById(checkboxId);
  chk.checked = !chk.checked;
  toggleFolderCheck(folderName, chk.checked);
}

function toggleFolderCheck(folderName, isChecked) {
  if (isChecked) {
    selectedFolders.add(folderName);
  } else {
    selectedFolders.delete(folderName);
  }
  renderFolders();
  // Auto-Select visible files when folder selection changes
  updateSelectAllCheckboxState();
  renderTable();
  buildCmd();
}

function setAllFolders(val) {
  selectedFolders.clear();
  if (val) {
    folderList.forEach(f => selectedFolders.add(f.name));
  }
  renderFolders();
  renderTable();
  buildCmd();
}

function setFoldersWithMissingOnly() {
  selectedFolders.clear();
  folderList.forEach(f => {
    if (f.missingTarget > 0) {
      selectedFolders.add(f.name);
    }
  });
  renderFolders();
  renderTable();
  buildCmd();
}

// Stats berechnen
function updateStats() {
  document.getElementById('st-total').textContent = comparisonData.length;
  document.getElementById('st-match').textContent = comparisonData.filter(d => d.InSource && d.InTarget).length;
  document.getElementById('st-miss-trg').textContent = comparisonData.filter(d => d.InSource && !d.InTarget).length;
  document.getElementById('st-miss-src').textContent = comparisonData.filter(d => !d.InSource && d.InTarget).length;
  document.getElementById('st-local').textContent = comparisonData.filter(d => d.InLocal).length;
}

// Filter setzen
function setFilter(f) {
  activeFilter = f;
  document.querySelectorAll('.filters .fbtn').forEach(btn => btn.classList.remove('on'));
  document.getElementById('btn-f-' + f).classList.add('on');
  renderTable();
}

function filterData() {
  renderTable();
}

// Dateigroesse formatieren
function formatSize(bytes) {
  if (bytes === 0) return '--';
  if (bytes < 1024) return bytes + ' B';
  const kb = bytes / 1024;
  if (kb < 1024) return kb.toFixed(1) + ' KB';
  const mb = kb / 1024;
  return mb.toFixed(2) + ' MB';
}

let table;

function getFilteredData() {
  const q = document.getElementById('txtSearch').value.toLowerCase();
  
  return comparisonData.filter(item => {
    // 1. Ordner-Filter
    if (!selectedFolders.has(item.Folder)) return false;
    
    // 2. Suche
    if (q && !item.Name.toLowerCase().includes(q)) return false;
    
    // 3. Status-Filter
    if (activeFilter === 'missing-target' && (!item.InSource || item.InTarget)) return false;
    if (activeFilter === 'missing-source' && (item.InSource || !item.InTarget)) return false;
    if (activeFilter === 'identical' && (!item.InSource || !item.InTarget)) return false;
    
    return true;
  });
}

function initDataTable() {
  table = new DataTable('#tblFiles', {
    data: getFilteredData(),
    deferRender: true,
    pageLength: 25,
    lengthMenu: [[15, 25, 50, 100, -1], [15, 25, 50, 100, 'Alle']],
    order: [[1, 'asc']],
    autoWidth: false,
    columns: [
      {
        data: 'RelativePath',
        title: '<input type="checkbox" id="chkAllFiles" onclick="toggleSelectAllFiles(this)">',
        orderable: false,
        searchable: false,
        className: 'col-chk',
        render: function(data, type, row) {
          const isChecked = selectedFiles.has(data) ? 'checked' : '';
          const canSync = row.InSource && !row.InTarget;
          const chkDisabled = canSync ? '' : 'disabled';
          return `<input type="checkbox" ${chkDisabled} ${isChecked} onchange="toggleFileCheck('${data}', this.checked)">`;
        }
      },
      {
        data: 'Name',
        title: 'Dateiname',
        className: 'col-name',
        render: function(data, type, row) {
          const size = row.SourceSize || row.TargetSize || row.LocalSize;
          const sizeStr = formatSize(size);
          return `
            <div style="font-weight:500;" title="${row.RelativePath}">${data}</div>
            <div class="mono" style="font-size:0.7rem;margin-top:2px;">Gr\u00f6\u00dfe: ${sizeStr}</div>
          `;
        }
      },
      {
        data: 'Folder',
        title: 'Unterordner',
        className: 'col-folder mono',
        render: function(data) {
          return data;
        }
      },
      {
        data: 'InSource',
        title: 'Quelle',
        className: 'col-badge',
        render: function(data) {
          return data ? '<span class="badge ok">Vorhanden</span>' : '<span class="badge danger">Fehlt</span>';
        }
      },
      {
        data: 'InTarget',
        title: 'Ziel',
        className: 'col-badge',
        render: function(data) {
          return data ? '<span class="badge ok">Vorhanden</span>' : '<span class="badge danger">Fehlt</span>';
        }
      },
      {
        data: 'InLocal',
        title: 'Lokal',
        className: 'col-badge',
        render: function(data) {
          return data ? '<span class="badge ok">Ja</span>' : (INITIAL_LOCAL_PATH ? '<span class="badge danger">Nein</span>' : '<span class="badge warn">--</span>');
        }
      }
    ],
    createdRow: function(row, data, dataIndex) {
      const cls = (!data.InTarget && data.InSource) ? 'missing-trg' : ((!data.InSource && data.InTarget) ? 'missing-src' : '');
      if (cls) {
        row.classList.add(cls);
      }
    },
    layout: {
      topStart: null,
      topEnd: null,
      bottomStart: ['pageLength', 'info'],
      bottomEnd: 'paging'
    },
    language: {
      search: "Suchen:",
      lengthMenu: "_MENU_ Einträge anzeigen",
      info: "Zeige _START_ bis _END_ von _TOTAL_ Einträgen",
      infoEmpty: "Zeige 0 bis 0 von 0 Einträgen",
      infoFiltered: "(gefiltert aus _MAX_ Gesamteinträgen)",
      zeroRecords: "Keine passenden Einträge gefunden",
      paginate: {
        first: "Erste",
        previous: "Zurück",
        next: "Nächste",
        last: "Letzte"
      }
    }
  });
  
  table.on('draw', function() {
    updateSelectAllCheckboxState();
  });
}

function renderTable() {
  if (table) {
    table.clear();
    table.rows.add(getFilteredData());
    table.draw(false);
  }
}

function toggleFileCheck(relPath, isChecked) {
  if (isChecked) {
    selectedFiles.add(relPath);
  } else {
    selectedFiles.delete(relPath);
  }
  buildCmd();
  updateSelectAllCheckboxState();
}

function updateSelectAllCheckboxState() {
  const chkAll = document.getElementById('chkAllFiles');
  if (!chkAll) return;
  
  const visibleCheckboxes = Array.from(document.querySelectorAll('#tblFiles tbody input[type="checkbox"]:not(:disabled)'));
  if (visibleCheckboxes.length === 0) {
    chkAll.checked = false;
    chkAll.disabled = true;
    return;
  }
  chkAll.disabled = false;
  const allChecked = visibleCheckboxes.every(c => c.checked);
  chkAll.checked = allChecked;
}

function toggleSelectAllFiles(chk) {
  const visibleRows = table.rows({ search: 'applied' }).data().toArray();
  visibleRows.forEach(row => {
    const canSync = row.InSource && !row.InTarget;
    if (canSync) {
      if (chk.checked) {
        selectedFiles.add(row.RelativePath);
      } else {
        selectedFiles.delete(row.RelativePath);
      }
    }
  });
  
  table.draw(false);
  buildCmd();
}

// Sync Modus Tabs
function setSyncTab(tab) {
  activeSyncTab = tab;
  document.querySelectorAll('.sync-tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  
  if (tab === 'folders') {
    document.getElementById('cmd-label').textContent = 'Auszuführender PowerShell Befehl (Ordner-basiert)';
  } else {
    document.getElementById('cmd-label').textContent = 'Auszuführender PowerShell Befehl (Datei-basiert)';
  }
  buildCmd();
}

function updateLocalPath(val) {
  localPath = val.trim();
  localStorage.setItem('allm_compare_localpath', localPath);
  buildCmd();
}

function updateWorkspaceSlug(val) {
  workspaceSlug = val.trim();
  localStorage.setItem('allm_compare_workspace', workspaceSlug);
  buildCmd();
}

// Befehl generieren
function buildCmd() {
  const el = document.getElementById('cmd-output');
  if (!localPath) {
    el.textContent = 'Bitte gib einen lokalen Quell-Pfad an, um den PowerShell-Befehl zu generieren.';
    return;
  }
  
  const bt = '`';
  let cmd = `. "${SCRIPT_PATH}" ${bt}\n`;
  cmd += `  -TargetUrl "${TARGET_URL}" ${bt}\n`;
  cmd += `  -LocalPath "${localPath}"`;
  
  if (workspaceSlug) {
    cmd += ` ${bt}\n  -WorkspaceSlug "${workspaceSlug}"`;
  }
  
  if (activeSyncTab === 'folders') {
    // Ordner-basiert
    const foldersArray = Array.from(selectedFolders);
    // Nur Ordner nehmen, die auch fehlende Dateien haben
    const foldersWithMissing = foldersArray.filter(f => {
      const fObj = folderList.find(x => x.name === f);
      return fObj && fObj.missingTarget > 0;
    });
    
    if (foldersWithMissing.length === 0) {
      el.textContent = 'Keine der ausgewählten Ordner hat fehlende Dateien im Ziel.';
      return;
    }
    
    cmd += ` ${bt}\n  -SyncFolders ${foldersWithMissing.map(f => `"${f}"`).join(',')} ${bt}\n  -ExecuteSync`;
  } else {
    // Datei-basiert
    if (selectedFiles.size === 0) {
      el.textContent = 'Bitte selektiere mindestens eine fehlende Datei in der Tabelle (Häkchen setzen).';
      return;
    }
    
    const filesArray = Array.from(selectedFiles);
    cmd += ` ${bt}\n  -SyncFiles ${filesArray.map(f => `"${f}"`).join(',')} ${bt}\n  -ExecuteSync`;
  }
  
  el.textContent = cmd;
}

function copyCmd() {
  const el = document.getElementById('cmd-output');
  if (el.textContent.startsWith('Bitte') || el.textContent.startsWith('Keine')) return;
  
  navigator.clipboard.writeText(el.textContent).then(function() {
    const msg = document.getElementById('copy-msg');
    msg.style.opacity = '1';
    setTimeout(function() { msg.style.opacity = '0'; }, 2000);
  });
}

// Start
updateStats();
renderFolders();
initDataTable();
buildCmd();

</script>
</body>
</html>
'@

$html = $htmlTemplate
$html = $html.Replace('{{SOURCE_URL}}', $SourceUrl)
$html = $html.Replace('{{TARGET_URL}}', $TargetUrl)
$html = $html.Replace('{{SCRIPT_PATH}}', ($PSCommandPath -replace '\\', '\\\\'))
$html = $html.Replace('{{LOCAL_PATH}}', ($localPathJs -replace '\\', '\\\\'))
$html = $html.Replace('{{COMPARISON_JSON}}', $comparisonJson)
$html = $html.Replace('{{GENERATED_TIME}}', (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))

# UTF-8 mit BOM schreiben
$encoding = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutputPath, $html, $encoding)

Write-Host "[$(& $ts)] HTML-Bericht erfolgreich generiert: $OutputPath" -ForegroundColor Green
Write-Host "[$(& $ts)] Oeffne Bericht im Browser..." -ForegroundColor Cyan

Start-Process $OutputPath
