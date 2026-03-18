<#
.SYNOPSIS
  Generates an interactive HTML report of AnythingLLM documents.

.DESCRIPTION
  Queries the AnythingLLM API for all documents in a folder and workspace,
  detects duplicates, and creates a filterable HTML report with statistics.

.PARAMETER AnythingLLMUrl
  Base URL of the AnythingLLM instance.

.PARAMETER DocumentFolder
  Document folder name in AnythingLLM.

.PARAMETER WorkspaceSlug
  Workspace slug to check embedding status.

.PARAMETER ApiKey
  API key (will prompt if not provided, uses cached key if available).

.PARAMETER OutputPath
  Path for the HTML report. Default: report_<folder>_<date>.html

.EXAMPLE
  .\Get-AnythingLLMReport.ps1 -AnythingLLMUrl "https://rag.example.com" -DocumentFolder "Infoplattform Wein" -WorkspaceSlug "infoplattform-wein"
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string]$AnythingLLMUrl,
  [Parameter(Mandatory)] [string]$DocumentFolder,
  [Parameter(Mandatory)] [string]$WorkspaceSlug,
  [string]$ApiKey,
  [string]$OutputPath
)

$AnythingLLMUrl = $AnythingLLMUrl.TrimEnd('/')

# ── API Key handling ──
if (-not $ApiKey) {
  $keyPath = [System.IO.Path]::Combine($env:TEMP, 'anythingllm_apikey.xml')
  if (Test-Path $keyPath) {
    $secure = Import-Clixml $keyPath
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    $ApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    Write-Host "[Report] API-Key aus Cache geladen." -ForegroundColor DarkGray
  }
  else {
    $ApiKey = Read-Host "API Key eingeben"
  }
}

$headers = @{ Authorization = "Bearer $ApiKey" }

# ── Fetch documents ──
Write-Host "[Report] Lade Dokumente aus '$DocumentFolder'..." -ForegroundColor Cyan
try {
  $resp = Invoke-RestMethod -Uri "$AnythingLLMUrl/api/v1/documents/folder/$DocumentFolder" `
    -Method Get -Headers $headers -TimeoutSec 30 -ErrorAction Stop
}
catch {
  Write-Error "Konnte Dokumente nicht laden: $($_.Exception.Message)"
  return
}

$docs = @()
if ($resp.documents) {
  foreach ($doc in $resp.documents) {
    $docs += [PSCustomObject]@{
      Title       = if ($doc.title) { $doc.title } else { $doc.name }
      ServerName  = $doc.name
      Cached      = [bool]$doc.cached
      Description = if ($doc.description) { $doc.description } else { '' }
      Author      = if ($doc.docAuthor) { $doc.docAuthor } else { '' }
      Source       = if ($doc.docSource) { $doc.docSource } else { '' }
      Published   = if ($doc.published) { $doc.published } else { '' }
      WordCount   = if ($doc.wordCount) { [int]$doc.wordCount } else { 0 }
      Tokens      = if ($doc.token_count_estimate) { [int]$doc.token_count_estimate } else { 0 }
      Location    = if ($doc.location) { $doc.location } else { '' }
    }
  }
}

Write-Host "[Report] $($docs.Count) Dokumente gefunden." -ForegroundColor Green

# ── Detect duplicates ──
$titleGroups = $docs | Group-Object Title
$duplicates = @{}
foreach ($g in $titleGroups) {
  if ($g.Count -gt 1) {
    $duplicates[$g.Name] = $g.Count
  }
}

# ── Statistics ──
$totalDocs = $docs.Count
$cachedCount = ($docs | Where-Object { $_.Cached }).Count
$notCachedCount = $totalDocs - $cachedCount
$duplicateCount = ($duplicates.Values | Measure-Object -Sum).Sum - $duplicates.Count
$uniqueTitles = $titleGroups.Count
$totalWords = ($docs | Measure-Object -Property WordCount -Sum).Sum
$totalTokens = ($docs | Measure-Object -Property Tokens -Sum).Sum

# ── Generate HTML ──
Add-Type -AssemblyName System.Web
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$dateStr = Get-Date -Format 'yyyy-MM-dd'

if (-not $OutputPath) {
  $folderSafe = $DocumentFolder -replace '[^a-zA-Z0-9_-]', '_'
  $OutputPath = Join-Path $PSScriptRoot "report_${folderSafe}_${dateStr}.html"
}

$rowsHtml = ""
foreach ($doc in ($docs | Sort-Object Title)) {
  $isDup = $duplicates.ContainsKey($doc.Title)
  $rowClass = if ($isDup) { 'duplicate' } elseif (-not $doc.Cached) { 'not-cached' } else { '' }
  $cachedBadge = if ($doc.Cached) { '<span class="badge cached">Embedded</span>' } else { '<span class="badge not-cached">Nicht embedded</span>' }
  $dupBadge = if ($isDup) { "<span class=""badge dup"">$($duplicates[$doc.Title])x Duplikat</span>" } else { '' }

  $rowsHtml += @"
        <tr class="$rowClass" data-title="$([System.Web.HttpUtility]::HtmlEncode($doc.Title).ToLower())" data-status="$(if($doc.Cached){'cached'}else{'notcached'})" data-dup="$(if($isDup){'yes'}else{'no'})">
          <td>$([System.Web.HttpUtility]::HtmlEncode($doc.Title)) $dupBadge</td>
          <td>$cachedBadge</td>
          <td>$($doc.WordCount.ToString('N0'))</td>
          <td>$($doc.Tokens.ToString('N0'))</td>
          <td class="desc">$([System.Web.HttpUtility]::HtmlEncode($doc.Description))</td>
          <td class="servername">$([System.Web.HttpUtility]::HtmlEncode($doc.ServerName))</td>
        </tr>
"@
}

$dupRowsHtml = ""
foreach ($g in ($titleGroups | Where-Object { $_.Count -gt 1 } | Sort-Object Name)) {
  $dupRowsHtml += @"
        <tr>
          <td>$([System.Web.HttpUtility]::HtmlEncode($g.Name))</td>
          <td class="dup-count">$($g.Count)x</td>
          <td>$(($g.Group | ForEach-Object { if($_.Cached){'✅'}else{'❌'} }) -join ' ')</td>
        </tr>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AnythingLLM Report - $([System.Web.HttpUtility]::HtmlEncode($DocumentFolder))</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Inter', sans-serif;
    background: #0a0a0f;
    color: #e0e0e8;
    min-height: 100vh;
  }
  .container { max-width: 1400px; margin: 0 auto; padding: 24px; }

  header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px;
    padding: 32px;
    margin-bottom: 24px;
  }
  header h1 {
    font-size: 1.8rem;
    font-weight: 700;
    background: linear-gradient(135deg, #667eea, #764ba2);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 4px;
  }
  header .meta { color: #888; font-size: 0.85rem; }

  .stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
  }
  .stat-card {
    background: linear-gradient(135deg, #1a1a2e, #1e1e3a);
    border: 1px solid rgba(255,255,255,0.06);
    border-radius: 12px;
    padding: 20px;
    text-align: center;
    transition: transform 0.2s, border-color 0.2s;
  }
  .stat-card:hover { transform: translateY(-2px); border-color: rgba(102,126,234,0.3); }
  .stat-card .value {
    font-size: 2rem;
    font-weight: 700;
    background: linear-gradient(135deg, #667eea, #764ba2);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }
  .stat-card.warn .value {
    background: linear-gradient(135deg, #f093fb, #f5576c);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }
  .stat-card.ok .value {
    background: linear-gradient(135deg, #43e97b, #38f9d7);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }
  .stat-card .label { font-size: 0.8rem; color: #888; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.5px; }

  .panel {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 24px;
  }
  .panel h2 {
    font-size: 1.2rem;
    font-weight: 600;
    margin-bottom: 16px;
    color: #c0c0d0;
  }

  .filters { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }
  .filters input[type="text"] {
    flex: 1;
    min-width: 250px;
    padding: 10px 16px;
    border-radius: 8px;
    border: 1px solid rgba(255,255,255,0.12);
    background: rgba(0,0,0,0.3);
    color: #e0e0e8;
    font-size: 0.9rem;
    outline: none;
    transition: border-color 0.2s;
  }
  .filters input:focus { border-color: #667eea; }
  .filter-btn {
    padding: 8px 16px;
    border-radius: 8px;
    border: 1px solid rgba(255,255,255,0.12);
    background: transparent;
    color: #aaa;
    cursor: pointer;
    font-size: 0.85rem;
    transition: all 0.2s;
  }
  .filter-btn:hover { border-color: #667eea; color: #667eea; }
  .filter-btn.active { background: #667eea; color: #fff; border-color: #667eea; }

  table { width: 100%; border-collapse: collapse; font-size: 0.85rem; }
  th {
    text-align: left;
    padding: 10px 12px;
    border-bottom: 2px solid rgba(255,255,255,0.1);
    color: #888;
    font-weight: 600;
    text-transform: uppercase;
    font-size: 0.75rem;
    letter-spacing: 0.5px;
    cursor: pointer;
    user-select: none;
  }
  th:hover { color: #667eea; }
  td {
    padding: 10px 12px;
    border-bottom: 1px solid rgba(255,255,255,0.04);
    vertical-align: top;
  }
  tr:hover { background: rgba(102,126,234,0.05); }
  tr.duplicate { background: rgba(245,87,108,0.08); }
  tr.duplicate:hover { background: rgba(245,87,108,0.14); }
  tr.not-cached { background: rgba(255,193,7,0.06); }
  tr.not-cached:hover { background: rgba(255,193,7,0.12); }
  .hidden { display: none !important; }

  .badge {
    display: inline-block;
    padding: 3px 8px;
    border-radius: 6px;
    font-size: 0.7rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.3px;
  }
  .badge.cached { background: rgba(67,233,123,0.15); color: #43e97b; }
  .badge.not-cached { background: rgba(255,193,7,0.15); color: #ffc107; }
  .badge.dup { background: rgba(245,87,108,0.15); color: #f5576c; margin-left: 8px;}

  .desc { max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #666; }
  .servername { max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #555; font-size: 0.75rem; font-family: monospace; }
  .dup-count { font-weight: 700; color: #f5576c; font-size: 1.1rem; }

  .dup-section { border-left: 3px solid #f5576c; padding-left: 20px; }
  .count-display { color: #888; font-size: 0.85rem; margin-bottom: 8px; }
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>📄 AnythingLLM Dokument-Report</h1>
    <div class="meta">
      <strong>Server:</strong> $([System.Web.HttpUtility]::HtmlEncode($AnythingLLMUrl)) &nbsp;|&nbsp;
      <strong>Ordner:</strong> $([System.Web.HttpUtility]::HtmlEncode($DocumentFolder)) &nbsp;|&nbsp;
      <strong>Workspace:</strong> $([System.Web.HttpUtility]::HtmlEncode($WorkspaceSlug)) &nbsp;|&nbsp;
      <strong>Stand:</strong> $timestamp
    </div>
  </header>

  <div class="stats-grid">
    <div class="stat-card"><div class="value">$totalDocs</div><div class="label">Dokumente gesamt</div></div>
    <div class="stat-card ok"><div class="value">$cachedCount</div><div class="label">Embedded</div></div>
    <div class="stat-card$(if($notCachedCount -gt 0){' warn'}else{''})"><div class="value">$notCachedCount</div><div class="label">Nicht embedded</div></div>
    <div class="stat-card$(if($duplicateCount -gt 0){' warn'}else{''})"><div class="value">$duplicateCount</div><div class="label">Duplikate</div></div>
    <div class="stat-card"><div class="value">$uniqueTitles</div><div class="label">Unique Titel</div></div>
    <div class="stat-card"><div class="value">$($totalWords.ToString('N0'))</div><div class="label">Wörter gesamt</div></div>
    <div class="stat-card"><div class="value">$($totalTokens.ToString('N0'))</div><div class="label">Tokens gesamt</div></div>
  </div>

  $(if ($duplicates.Count -gt 0) { @"
  <div class="panel dup-section">
    <h2>⚠️ Duplikate ($($duplicates.Count) Titel betroffen)</h2>
    <table>
      <thead><tr><th>Titel</th><th>Anzahl</th><th>Embedded-Status</th></tr></thead>
      <tbody>$dupRowsHtml</tbody>
    </table>
  </div>
"@ })

  <div class="panel">
    <h2>Alle Dokumente</h2>
    <div class="filters">
      <input type="text" id="search" placeholder="Suche nach Dateiname..." oninput="filterTable()">
      <button class="filter-btn active" onclick="setFilter('all',this)">Alle</button>
      <button class="filter-btn" onclick="setFilter('cached',this)">Embedded</button>
      <button class="filter-btn" onclick="setFilter('notcached',this)">Nicht embedded</button>
      <button class="filter-btn" onclick="setFilter('dup',this)">Duplikate</button>
    </div>
    <div class="count-display" id="countDisplay">$totalDocs Dokumente</div>
    <table id="docTable">
      <thead>
        <tr>
          <th onclick="sortTable(0)">Titel ↕</th>
          <th onclick="sortTable(1)">Status ↕</th>
          <th onclick="sortTable(2)">Wörter ↕</th>
          <th onclick="sortTable(3)">Tokens ↕</th>
          <th>Beschreibung</th>
          <th>Server-Name</th>
        </tr>
      </thead>
      <tbody>$rowsHtml</tbody>
    </table>
  </div>
</div>

<script>
let currentFilter = 'all';
function setFilter(f, btn) {
  currentFilter = f;
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  filterTable();
}
function filterTable() {
  const search = document.getElementById('search').value.toLowerCase();
  const rows = document.querySelectorAll('#docTable tbody tr');
  let visible = 0;
  rows.forEach(row => {
    const title = row.dataset.title || '';
    const status = row.dataset.status || '';
    const dup = row.dataset.dup || '';
    const matchSearch = !search || title.includes(search);
    let matchFilter = true;
    if (currentFilter === 'cached') matchFilter = status === 'cached';
    else if (currentFilter === 'notcached') matchFilter = status === 'notcached';
    else if (currentFilter === 'dup') matchFilter = dup === 'yes';
    const show = matchSearch && matchFilter;
    row.classList.toggle('hidden', !show);
    if (show) visible++;
  });
  document.getElementById('countDisplay').textContent = visible + ' von $totalDocs Dokumente';
}
let sortDir = {};
function sortTable(col) {
  const table = document.getElementById('docTable');
  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  sortDir[col] = !sortDir[col];
  const dir = sortDir[col] ? 1 : -1;
  rows.sort((a, b) => {
    let va = a.cells[col].textContent.trim();
    let vb = b.cells[col].textContent.trim();
    const na = parseFloat(va.replace(/\./g,'').replace(',','.'));
    const nb = parseFloat(vb.replace(/\./g,'').replace(',','.'));
    if (!isNaN(na) && !isNaN(nb)) return (na - nb) * dir;
    return va.localeCompare(vb, 'de') * dir;
  });
  rows.forEach(r => tbody.appendChild(r));
}
</script>
</body>
</html>
"@

# ── Write HTML ──
$html | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
Write-Host "[Report] HTML-Report erstellt: $OutputPath" -ForegroundColor Green
Write-Host "[Report] $totalDocs Dokumente | $cachedCount embedded | $duplicateCount Duplikate" -ForegroundColor Cyan

# Open in browser
Start-Process $OutputPath
