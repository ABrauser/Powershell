<#
.SYNOPSIS
  Generates an interactive HTML report of AnythingLLM documents with optional local file comparison.

.DESCRIPTION
  Queries the AnythingLLM API and creates an HTML report showing:
  - Server documents with embedding status
  - Duplicate detection
  - Optional: Tree view comparing local files against server status

.PARAMETER AnythingLLMUrl
  Base URL of the AnythingLLM instance.

.PARAMETER DocumentFolder
  Document folder name in AnythingLLM.

.PARAMETER WorkspaceSlug
  Workspace slug.

.PARAMETER InputPath
  Optional. Local folder to compare against server. Enables tree view.

.PARAMETER Folders
  Optional. Subfolders to include (same as Invoke-AnythingLLMUpload).

.PARAMETER Extensions
  Optional. File extensions to include. Default: .md

.PARAMETER ApiKey
  API key. Uses session cache if not provided.

.PARAMETER OutputPath
  Path for the HTML report.

.EXAMPLE
  # Server-only view
  .\Get-AnythingLLMReport.ps1 -AnythingLLMUrl "https://rag.example.com" -DocumentFolder "Wein" -WorkspaceSlug "wein"

  # With local comparison
  .\Get-AnythingLLMReport.ps1 -AnythingLLMUrl "https://rag.example.com" -DocumentFolder "Wein" -WorkspaceSlug "wein" -InputPath "C:\TEMP\Result" -Folders "2_Begutachtung","3_Untersuchungsverfahren"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string]$AnythingLLMUrl,
  [Parameter(Mandatory)] [string]$DocumentFolder,
  [Parameter(Mandatory)] [string]$WorkspaceSlug,
  [string]$InputPath,
  [string[]]$Folders,
  [string[]]$Extensions = @('.md'),
  [string]$ApiKey,
  [string]$OutputPath
)

$AnythingLLMUrl = $AnythingLLMUrl.TrimEnd('/')

# ── API Key handling (session cache) ──
$keyPath = [System.IO.Path]::Combine($env:TEMP, 'anythingllm_apikey.xml')
if (-not $ApiKey) {
  if (Test-Path $keyPath) {
    $secure = Import-Clixml $keyPath
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    $ApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    Write-Host "[Report] API-Key aus Cache geladen." -ForegroundColor DarkGray
  }
  else {
    $secKey = Read-Host "API Key eingeben" -AsSecureString
    $secKey | Export-Clixml $keyPath
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secKey)
    $ApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    Write-Host "[Report] API-Key gespeichert." -ForegroundColor Green
  }
}
$headers = @{ Authorization = "Bearer $ApiKey" }

# ── Fetch server documents ──
Write-Host "[Report] Lade Dokumente aus '$DocumentFolder'..." -ForegroundColor Cyan
try {
  $resp = Invoke-RestMethod -Uri "$AnythingLLMUrl/api/v1/documents/folder/$DocumentFolder" `
    -Method Get -Headers $headers -TimeoutSec 30 -ErrorAction Stop
}
catch {
  Write-Error "Konnte Dokumente nicht laden: $($_.Exception.Message)"
  return
}

$serverDocs = @()
if ($resp.documents) {
  foreach ($doc in $resp.documents) {
    $serverDocs += [PSCustomObject]@{
      Title       = if ($doc.title) { $doc.title } else { $doc.name }
      ServerName  = $doc.name
      Cached      = [bool]$doc.cached
      Description = if ($doc.description) { $doc.description } else { '' }
      WordCount   = if ($doc.wordCount) { [int]$doc.wordCount } else { 0 }
      Tokens      = if ($doc.token_count_estimate) { [int]$doc.token_count_estimate } else { 0 }
    }
  }
}
Write-Host "[Report] $($serverDocs.Count) Server-Dokumente gefunden." -ForegroundColor Green

# ── Build server lookup by title ──
$serverByTitle = @{}
foreach ($sd in $serverDocs) {
  if (-not $serverByTitle.ContainsKey($sd.Title)) {
    $serverByTitle[$sd.Title] = @()
  }
  $serverByTitle[$sd.Title] += $sd
}

# ── Duplicate detection ──
$dupTitles = @{}
foreach ($key in $serverByTitle.Keys) {
  if ($serverByTitle[$key].Count -gt 1) {
    $dupTitles[$key] = $serverByTitle[$key].Count
  }
}

# ── Local file scan (optional) ──
$localFiles = @()
$hasLocal = $false
if ($InputPath -and (Test-Path $InputPath)) {
  $hasLocal = $true
  $InputPath = (Resolve-Path $InputPath).Path
  $inputRoot = $InputPath.TrimEnd('\')
  Write-Host "[Report] Scanne lokale Dateien in '$InputPath'..." -ForegroundColor Cyan

  $gciParams = @{ Path = $InputPath; File = $true; Recurse = $true; ErrorAction = 'SilentlyContinue' }
  $allFiles = Get-ChildItem @gciParams

  # Folder filter
  if ($Folders -and $Folders.Count -gt 0) {
    $filtered = @()
    foreach ($f in $allFiles) {
      $rel = ''
      if ($f.FullName.StartsWith($inputRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        $rel = $f.FullName.Substring($inputRoot.Length).TrimStart('\')
      }
      foreach ($folder in $Folders) {
        $fn = $folder.TrimStart('\').TrimEnd('\')
        if ($rel.StartsWith($fn + '\', [System.StringComparison]::OrdinalIgnoreCase) -or
            $rel.StartsWith($fn, [System.StringComparison]::OrdinalIgnoreCase)) {
          $filtered += $f
          break
        }
      }
    }
    $allFiles = $filtered
  }

  # Extension filter
  if ($Extensions -and $Extensions.Count -gt 0) {
    $extSet = $Extensions | ForEach-Object { $_.ToLower() }
    $allFiles = @($allFiles | Where-Object { $extSet -contains $_.Extension.ToLower() })
  }

  foreach ($f in $allFiles) {
    $relPath = ''
    if ($f.FullName.StartsWith($inputRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
      $relPath = $f.FullName.Substring($inputRoot.Length).TrimStart('\')
    } else { $relPath = $f.Name }

    # Check server status
    $serverMatch = $serverByTitle[$f.Name]
    $status = 'missing'
    if ($serverMatch -and $serverMatch.Count -gt 0) {
      $anyCached = $serverMatch | Where-Object { $_.Cached }
      if ($anyCached) { $status = 'embedded' }
      else { $status = 'uploaded' }
    }

    $localFiles += [PSCustomObject]@{
      Name         = $f.Name
      RelativePath = $relPath
      Folder       = [System.IO.Path]::GetDirectoryName($relPath)
      SizeMB       = [math]::Round($f.Length / 1MB, 2)
      Status       = $status
    }
  }
  Write-Host "[Report] $($localFiles.Count) lokale Dateien gefunden." -ForegroundColor Green
}

# ── Statistics ──
$totalServer = $serverDocs.Count
$cachedCount = ($serverDocs | Where-Object { $_.Cached }).Count
$notCachedCount = $totalServer - $cachedCount
$dupCount = ($dupTitles.Values | Measure-Object -Sum).Sum - $dupTitles.Count
$uniqueTitles = ($serverDocs | Select-Object -Property Title -Unique).Count
$totalWords = ($serverDocs | Measure-Object -Property WordCount -Sum).Sum
$totalTokens = ($serverDocs | Measure-Object -Property Tokens -Sum).Sum

$localTotal = $localFiles.Count
$localEmbedded = ($localFiles | Where-Object { $_.Status -eq 'embedded' }).Count
$localUploaded = ($localFiles | Where-Object { $_.Status -eq 'uploaded' }).Count
$localMissing = ($localFiles | Where-Object { $_.Status -eq 'missing' }).Count

# ── Generate HTML ──
Add-Type -AssemblyName System.Web
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$dateStr = Get-Date -Format 'yyyy-MM-dd'
$e = [System.Web.HttpUtility]

if (-not $OutputPath) {
  $folderSafe = $DocumentFolder -replace '[^a-zA-Z0-9_-]', '_'
  $OutputPath = Join-Path $PSScriptRoot "report_${folderSafe}_${dateStr}.html"
}

# Build server document table rows
$serverRowsHtml = ""
foreach ($doc in ($serverDocs | Sort-Object Title)) {
  $isDup = $dupTitles.ContainsKey($doc.Title)
  $rc = if ($isDup) { 'dup' } elseif (-not $doc.Cached) { 'warn' } else { '' }
  $badge = if ($doc.Cached) { '<span class="badge ok">Embedded</span>' } else { '<span class="badge warn">Ausstehend</span>' }
  $dupBadge = if ($isDup) { "<span class=""badge dup"">$($dupTitles[$doc.Title])x</span>" } else { '' }

  $serverRowsHtml += "<tr class=""$rc"" data-title=""$($e::HtmlEncode($doc.Title).ToLower())"" data-st=""$(if($doc.Cached){'c'}else{'n'})"" data-dp=""$(if($isDup){'y'}else{'n'})"">"
  $serverRowsHtml += "<td>$($e::HtmlEncode($doc.Title)) $dupBadge</td><td>$badge</td>"
  $serverRowsHtml += "<td>$($doc.WordCount.ToString('N0'))</td><td>$($doc.Tokens.ToString('N0'))</td>"
  $serverRowsHtml += "<td class=""mono"">$($e::HtmlEncode($doc.ServerName))</td></tr>`n"
}

# Build tree view data (JSON) for local files
$treeJson = '[]'
if ($hasLocal) {
  $treeData = @()
  $folderGroups = $localFiles | Group-Object Folder | Sort-Object Name
  foreach ($g in $folderGroups) {
    $folderName = if ($g.Name) { $g.Name } else { '(Stammverzeichnis)' }
    $files = @()
    foreach ($f in ($g.Group | Sort-Object Name)) {
      $files += @{
        name   = $f.Name
        size   = $f.SizeMB
        status = $f.Status
      }
    }
    $treeData += @{
      folder = $folderName
      files  = $files
      total  = $g.Count
      embedded = ($g.Group | Where-Object { $_.Status -eq 'embedded' }).Count
      uploaded = ($g.Group | Where-Object { $_.Status -eq 'uploaded' }).Count
      missing  = ($g.Group | Where-Object { $_.Status -eq 'missing' }).Count
    }
  }
  $treeJson = $treeData | ConvertTo-Json -Depth 5 -Compress
}

# ── Duplicate rows ──
$dupRowsHtml = ""
foreach ($key in ($dupTitles.Keys | Sort-Object)) {
  $entries = $serverByTitle[$key]
  $badges = ($entries | ForEach-Object { if ($_.Cached) { '<span class="badge ok">E</span>' } else { '<span class="badge warn">A</span>' } }) -join ' '
  $dupRowsHtml += "<tr><td>$($e::HtmlEncode($key))</td><td class=""dup-n"">$($dupTitles[$key])x</td><td>$badges</td></tr>`n"
}

$html = @"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AnythingLLM Report - $($e::HtmlEncode($DocumentFolder))</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet" onerror="this.remove()">
<style>
:root {
  --bg-primary: #0E0E11;
  --bg-secondary: #131314;
  --bg-card: #1E1F20;
  --bg-glass: rgba(255,255,255,0.05);
  --border-glass: rgba(255,255,255,0.1);
  --text-primary: #E3E3E3;
  --text-secondary: #C4C7C5;
  --text-muted: #8E918F;
  --accent: #A8C7FA;
  --accent-glow: rgba(168,199,250,0.1);
  --success: #6DD58C;
  --warning: #FFB74D;
  --danger: #E25C5C;
  --radius-card: 24px;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: var(--bg-primary); color: var(--text-primary); line-height: 1.6; min-height: 100vh;
}
.container { max-width: 1400px; margin: 0 auto; padding: 24px; }
header {
  background: var(--bg-card); border: 1px solid var(--border-glass);
  border-radius: var(--radius-card); padding: 28px 32px; margin-bottom: 20px;
}
header h1 { font-size: 1.7rem; font-weight: 500; margin-bottom: 2px; }
header h1 span {
  background: linear-gradient(90deg, #4285F4, #9B72CB, #D96570);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 600;
}
header .meta { color: var(--text-muted); font-size: 0.82rem; }
.stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(155px, 1fr)); gap: 14px; margin-bottom: 20px; }
.stat {
  background: var(--bg-card); border: 1px solid var(--border-glass);
  border-radius: 16px; padding: 18px; text-align: center;
  transition: transform 0.2s, border-color 0.2s;
}
.stat:hover { transform: translateY(-2px); border-color: rgba(168,199,250,0.2); }
.stat .val { font-size: 1.8rem; font-weight: 700; color: var(--accent); }
.stat.ok .val { color: var(--success); }
.stat.warn .val { color: var(--warning); }
.stat.danger .val { color: var(--danger); }
.stat .lbl { font-size: 0.72rem; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; margin-top: 2px; }
.panel {
  background: var(--bg-card); border: 1px solid var(--border-glass);
  border-radius: var(--radius-card); padding: 24px; margin-bottom: 20px;
}
.panel h2 { font-size: 1.1rem; font-weight: 500; margin-bottom: 14px; }
.filters { display: flex; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; align-items: center; }
.filters input {
  flex: 1; min-width: 220px; padding: 8px 14px; border-radius: 8px;
  border: 1px solid var(--border-glass); background: var(--bg-secondary);
  color: var(--text-primary); font-size: 0.85rem; outline: none;
}
.filters input:focus { border-color: var(--accent); }
.fbtn {
  padding: 6px 14px; border-radius: 8px; border: 1px solid var(--border-glass);
  background: transparent; color: var(--text-muted); cursor: pointer; font-size: 0.8rem;
  transition: all 0.2s;
}
.fbtn:hover { border-color: var(--accent); color: var(--accent); }
.fbtn.on { background: var(--accent); color: #111; border-color: var(--accent); }
table { width: 100%; border-collapse: collapse; font-size: 0.82rem; }
th {
  text-align: left; padding: 8px 10px; border-bottom: 2px solid var(--border-glass);
  color: var(--text-muted); font-weight: 600; text-transform: uppercase; font-size: 0.7rem;
  letter-spacing: 0.4px; cursor: pointer; user-select: none;
}
th:hover { color: var(--accent); }
td { padding: 8px 10px; border-bottom: 1px solid rgba(255,255,255,0.03); }
tr:hover { background: var(--accent-glow); }
tr.dup { background: rgba(226,92,92,0.06); }
tr.warn { background: rgba(255,183,77,0.05); }
.hidden { display: none !important; }
.badge {
  display: inline-block; padding: 2px 7px; border-radius: 5px;
  font-size: 0.65rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.3px;
}
.badge.ok { background: rgba(109,213,140,0.12); color: var(--success); }
.badge.warn { background: rgba(255,183,77,0.12); color: var(--warning); }
.badge.dup { background: rgba(226,92,92,0.12); color: var(--danger); margin-left: 6px; }
.badge.miss { background: rgba(226,92,92,0.12); color: var(--danger); }
.mono { font-family: 'JetBrains Mono', monospace; font-size: 0.7rem; color: var(--text-muted); max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.dup-n { font-weight: 700; color: var(--danger); font-size: 1rem; }
.dup-border { border-left: 3px solid var(--danger); padding-left: 20px; }
.cnt { color: var(--text-muted); font-size: 0.82rem; margin-bottom: 6px; }

/* Tree View */
.tree {
  font-size: 0.85rem;
  background: var(--bg-secondary); border: 1px solid var(--border-glass);
  border-radius: 12px; overflow: hidden;
}
.tree:empty { display: none; }
.tree-folder { border-bottom: 1px solid var(--border-glass); }
.tree-folder:last-child { border-bottom: none; }
.tree-head {
  display: flex; align-items: center; gap: 10px; padding: 10px 14px;
  cursor: pointer; user-select: none;
}
.tree-head:hover { background: var(--accent-glow); }
.tree-arrow { transition: transform 0.2s; font-size: 0.7rem; color: var(--text-muted); }
.tree-folder.open .tree-arrow { transform: rotate(90deg); }
.tree-fname { font-weight: 500; flex: 1; }
.tree-stats { display: flex; gap: 8px; font-size: 0.72rem; }
.tree-stats span { padding: 2px 6px; border-radius: 4px; }
.tree-stats .ts-ok { background: rgba(109,213,140,0.1); color: var(--success); }
.tree-stats .ts-warn { background: rgba(255,183,77,0.1); color: var(--warning); }
.tree-stats .ts-miss { background: rgba(226,92,92,0.1); color: var(--danger); }
.tree-files { display: none; padding: 2px 14px 10px 36px; }
.tree-folder.open .tree-files { display: block; }
.tree-file {
  display: flex; align-items: center; gap: 8px; padding: 5px 0;
  border-bottom: 1px solid rgba(255,255,255,0.03); font-size: 0.82rem;
}
.tree-file:last-child { border-bottom: none; }
.tree-icon { width: 18px; text-align: center; font-size: 0.9rem; }
.tree-file-name { flex: 1; }
.tree-file-size { color: var(--text-muted); font-size: 0.75rem; font-family: 'JetBrains Mono', monospace; }

.progress-bar {
  height: 8px; background: var(--bg-primary); border-radius: 4px;
  overflow: hidden; margin-bottom: 16px;
}
.progress-bar-inner { height: 100%; border-radius: 4px; transition: width 0.3s; }
.footer {
  text-align: center; font-size: 0.72rem; color: var(--text-muted);
  margin-top: 20px; padding-top: 12px; border-top: 1px solid var(--border-glass);
}
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>AnythingLLM <span>Report</span></h1>
    <div class="meta">
      <strong>Server:</strong> $($e::HtmlEncode($AnythingLLMUrl)) &nbsp;&bull;&nbsp;
      <strong>Ordner:</strong> $($e::HtmlEncode($DocumentFolder)) &nbsp;&bull;&nbsp;
      <strong>Workspace:</strong> $($e::HtmlEncode($WorkspaceSlug)) &nbsp;&bull;&nbsp;
      <strong>Stand:</strong> $timestamp
      $(if ($hasLocal) { "<br><strong>Lokaler Pfad:</strong> $($e::HtmlEncode($InputPath))" })
    </div>
  </header>

  <div class="stats">
    <div class="stat"><div class="val">$totalServer</div><div class="lbl">Server-Dokumente</div></div>
    <div class="stat ok"><div class="val">$cachedCount</div><div class="lbl">Embedded</div></div>
    <div class="stat$(if($notCachedCount -gt 0){' warn'}else{''})"><div class="val">$notCachedCount</div><div class="lbl">Ausstehend</div></div>
    <div class="stat$(if($dupCount -gt 0){' danger'}else{''})"><div class="val">$dupCount</div><div class="lbl">Duplikate</div></div>
    <div class="stat"><div class="val">$uniqueTitles</div><div class="lbl">Unique Titel</div></div>
    <div class="stat"><div class="val">$($totalTokens.ToString('N0'))</div><div class="lbl">Tokens</div></div>
    $(if ($hasLocal) { @"
    <div class="stat"><div class="val">$localTotal</div><div class="lbl">Lokal gesamt</div></div>
    <div class="stat$(if($localMissing -gt 0){' danger'}else{' ok'})"><div class="val">$localMissing</div><div class="lbl">Fehlend</div></div>
"@ })
  </div>

  $(if ($hasLocal) {
    $pctDone = if ($localTotal -gt 0) { [math]::Round(($localEmbedded / $localTotal) * 100) } else { 0 }
    $pctUpl = if ($localTotal -gt 0) { [math]::Round(($localUploaded / $localTotal) * 100) } else { 0 }
    @"
  <div class="panel">
    <h2>&#128194; Lokal &harr; Server Abgleich <span style="color:var(--text-muted);font-size:0.85rem;font-weight:400">($localEmbedded/$localTotal embedded = ${pctDone}%)</span></h2>
    <div class="progress-bar">
      <div class="progress-bar-inner" style="width:${pctDone}%;background:linear-gradient(90deg,var(--success),#38f9d7);$(if($pctUpl -gt 0){"box-shadow:${pctDone}% 0 0 0 var(--warning);"})"></div>
    </div>
    <div class="filters">
      <input type="text" id="treeSearch" placeholder="Ordner/Datei suchen..." oninput="filterTree()">
      <button class="fbtn on" onclick="treeFilter('all',this)">Alle</button>
      <button class="fbtn" onclick="treeFilter('missing',this)">&#10060; Fehlend ($localMissing)</button>
      <button class="fbtn" onclick="treeFilter('uploaded',this)">&#11014;&#65039; Ausstehend ($localUploaded)</button>
      <button class="fbtn" onclick="treeFilter('embedded',this)">&#9989; Embedded ($localEmbedded)</button>
    </div>
    <div class="tree" id="treeView"></div>
  </div>
"@
  })

  $(if ($dupTitles.Count -gt 0) { @"
  <div class="panel dup-border">
    <h2>&#9888;&#65039; Duplikate ($($dupTitles.Count) Titel betroffen)</h2>
    <table><thead><tr><th>Titel</th><th>Anzahl</th><th>Status</th></tr></thead>
    <tbody>$dupRowsHtml</tbody></table>
  </div>
"@ })

  <div class="panel">
    <h2>Server-Dokumente</h2>
    <div class="filters">
      <input type="text" id="search" placeholder="Suche..." oninput="filterTbl()">
      <button class="fbtn on" onclick="setF('all',this)">Alle</button>
      <button class="fbtn" onclick="setF('c',this)">Embedded</button>
      <button class="fbtn" onclick="setF('n',this)">Ausstehend</button>
      <button class="fbtn" onclick="setF('dup',this)">Duplikate</button>
    </div>
    <div class="cnt" id="cnt">$totalServer Dokumente</div>
    <table id="tbl">
      <thead><tr>
        <th onclick="srt(0)">Titel &#8597;</th><th onclick="srt(1)">Status &#8597;</th>
        <th onclick="srt(2)">W&ouml;rter &#8597;</th><th onclick="srt(3)">Tokens &#8597;</th>
        <th>Server-Name</th>
      </tr></thead>
      <tbody>$serverRowsHtml</tbody>
    </table>
  </div>

  <div class="footer">Get-AnythingLLMReport v2.0 &mdash; $timestamp</div>
</div>

<script>
// Server table filter
var cf='all';
function setF(f,b){cf=f;document.querySelectorAll('.panel:last-of-type .fbtn').forEach(x=>x.classList.remove('on'));b.classList.add('on');filterTbl();}
function filterTbl(){
  var s=document.getElementById('search').value.toLowerCase(),vis=0;
  document.querySelectorAll('#tbl tbody tr').forEach(function(r){
    var t=r.dataset.title||'',st=r.dataset.st||'',dp=r.dataset.dp||'';
    var ms=!s||t.includes(s),mf=cf==='all'||(cf==='c'&&st==='c')||(cf==='n'&&st==='n')||(cf==='dup'&&dp==='y');
    r.classList.toggle('hidden',!(ms&&mf));if(ms&&mf)vis++;
  });
  document.getElementById('cnt').textContent=vis+' von $totalServer Dokumente';
}
var sd={};
function srt(c){
  var tb=document.querySelector('#tbl tbody'),rows=Array.from(tb.querySelectorAll('tr'));
  sd[c]=!sd[c];var d=sd[c]?1:-1;
  rows.sort(function(a,b){
    var va=a.cells[c].textContent.trim(),vb=b.cells[c].textContent.trim();
    var na=parseFloat(va.replace(/\./g,'').replace(',','.')),nb=parseFloat(vb.replace(/\./g,'').replace(',','.'));
    if(!isNaN(na)&&!isNaN(nb))return(na-nb)*d;
    return va.localeCompare(vb,'de')*d;
  });
  rows.forEach(function(r){tb.appendChild(r);});
}

// Tree view
var treeData=$treeJson;
var treeCf='all';
function renderTree(){
  var el=document.getElementById('treeView');
  if(!el)return;
  var s=(document.getElementById('treeSearch')||{}).value||'';
  s=s.toLowerCase();
  var html='';
  treeData.forEach(function(g,i){
    var fn=g.folder,files=g.files;
    // Filter
    var ffiles=files.filter(function(f){
      var matchS=!s||f.name.toLowerCase().includes(s)||fn.toLowerCase().includes(s);
      var matchF=treeCf==='all'||f.status===treeCf;
      return matchS&&matchF;
    });
    if(ffiles.length===0)return;
    var eC=ffiles.filter(function(f){return f.status==='embedded';}).length;
    var uC=ffiles.filter(function(f){return f.status==='uploaded';}).length;
    var mC=ffiles.filter(function(f){return f.status==='missing';}).length;
    html+='<div class="tree-folder" id="tf'+i+'">';
    html+='<div class="tree-head" onclick="document.getElementById(\'tf'+i+'\').classList.toggle(\'open\')">';
    html+='<span class="tree-arrow">&#9654;</span>';
    html+='<span class="tree-fname">&#128194; '+fn+'</span>';
    html+='<span class="tree-stats">';
    if(eC>0)html+='<span class="ts-ok">&#9989; '+eC+'</span>';
    if(uC>0)html+='<span class="ts-warn">&#11014;&#65039; '+uC+'</span>';
    if(mC>0)html+='<span class="ts-miss">&#10060; '+mC+'</span>';
    html+='<span style="color:var(--text-muted)">'+ffiles.length+' Dateien</span>';
    html+='</span></div>';
    html+='<div class="tree-files">';
    ffiles.forEach(function(f){
      var icon=f.status==='embedded'?'&#9989;':f.status==='uploaded'?'&#11014;&#65039;':'&#10060;';
      var cls=f.status==='missing'?' style="color:var(--danger)"':'';
      html+='<div class="tree-file"><span class="tree-icon">'+icon+'</span>';
      html+='<span class="tree-file-name"'+cls+'>'+f.name+'</span>';
      html+='<span class="tree-file-size">'+f.size.toFixed(2)+' MB</span></div>';
    });
    html+='</div></div>';
  });
  if(!html)html='<div style="color:var(--text-muted);padding:20px;text-align:center">Keine Dateien gefunden.</div>';
  el.innerHTML=html;
}
function filterTree(){renderTree();}
function treeFilter(f,b){
  treeCf=f;
  document.querySelectorAll('.panel:nth-of-type(2) .fbtn').forEach(function(x){x.classList.remove('on');});
  b.classList.add('on');
  renderTree();
}
renderTree();
</script>
</body>
</html>
"@

# ── Write HTML with UTF-8 BOM ──
$encoding = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutputPath, $html, $encoding)
Write-Host "[Report] HTML-Report erstellt: $OutputPath" -ForegroundColor Green
Write-Host "[Report] $totalServer Server-Dokumente | $cachedCount embedded | $dupCount Duplikate" -ForegroundColor Cyan
if ($hasLocal) {
  Write-Host "[Report] $localTotal lokal | $localEmbedded embedded | $localMissing fehlend" -ForegroundColor Cyan
}

# Open in browser
Start-Process $OutputPath
