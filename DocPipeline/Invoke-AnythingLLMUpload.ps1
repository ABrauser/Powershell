<#
.SYNOPSIS
    Uploads processed files to AnythingLLM via API and embeds them into a workspace.

.DESCRIPTION
    Invoke-AnythingLLMUpload processes files from an input folder (typically the Result/Ergebnis
    folder from the Invoke-DoclingConversion pipeline) and uploads them to an AnythingLLM instance
    via its REST API.

    Features:
    - Two-phase processing: Upload first, then batch-wise Embedding
    - Configurable batch size with pause between batches (prevents Vector-DB overload)
    - Embedding status polling (waits until batch is fully embedded before continuing)
    - Skip already-uploaded files (incremental / resume-safe)
    - Retry with exponential backoff on API errors
    - SecureString API key caching (persists in PowerShell session)
    - Detailed per-file log (upload_log.csv)
    - Run summary log (upload_runs.json)
    - Pipeline manifest integration (writes "hochgeladen" status)
    - Configurable folder and extension filters

    Designed to work with the Invoke-FolderScan pipeline:
      Source (Scan) → Pipeline_*.csv → Staging (Copy) → Docling → Ergebnis → AnythingLLM (this script)

.PARAMETER AnythingLLMUrl
    Base URL of the AnythingLLM instance (e.g. 'http://localhost:3001').

.PARAMETER ApiKey
    AnythingLLM API key (Bearer token). If omitted, the script prompts for it
    and caches it as a SecureString for the current PowerShell session.

.PARAMETER InputPath
    Folder containing files to upload (typically the Ergebnis/Result folder).

.PARAMETER WorkspaceSlug
    Target workspace slug in AnythingLLM.

.PARAMETER DocumentFolder
    Target folder name inside AnythingLLM for organizing uploads.
    Defaults to 'pipeline-upload'.

.PARAMETER Folders
    Optional array of subfolder names (relative to InputPath) to process.
    If omitted, all files in InputPath are processed (recursively).

.PARAMETER Extensions
    Optional array of file extensions to include (e.g. '.md', '.txt').
    Defaults to @('.md').

.PARAMETER BatchSize
    Number of files to upload per batch before pausing. Default: 10.

.PARAMETER BatchPauseSec
    Seconds to pause between batches (gives Vector-DB time). Default: 5.

.PARAMETER EmbeddingTimeoutSec
    Maximum seconds to wait for embedding to complete per batch. Default: 300.

.PARAMETER EmbeddingPollIntervalSec
    Seconds between embedding status polls. Default: 10.

.PARAMETER SkipExisting
    Skip files that have already been uploaded (based on upload_log.csv). Default: $true.

.PARAMETER Force
    Overwrite / re-upload existing files (opposite of SkipExisting).

.PARAMETER UploadOnly
    Only upload files, skip the embedding step.
    Useful for separating upload from embedding.

.PARAMETER TimeoutSec
    Timeout in seconds per API upload call. Default: 120.

.PARAMETER RetryCount
    Number of retries on API failure per file. Default: 3.

.PARAMETER CsvPath
    Optional: Use a Pipeline_*.csv or custom CSV as input instead of -InputPath.
    The CSV must contain a 'FullPath' column.

.PARAMETER LogDir
    Directory for log files. Defaults to the script directory ($PSScriptRoot).

.NOTES
    Name:    Invoke-AnythingLLMUpload
    Author:  Benjamin Rauser
    Contact: Benjamin.Rauser@outlook.com
    Version: 1.1 - 2026-03-11
             1.0 - 2026-03-11 - Initial version

.EXAMPLE
    . .\Invoke-AnythingLLMUpload.ps1
    Invoke-AnythingLLMUpload -Gui
    # Opens standalone configuration GUI in the browser

.EXAMPLE
    Invoke-AnythingLLMUpload -Gui -GuiInputPath "D:\Ergebnis"
    # Opens GUI with pre-filled input path

.EXAMPLE
    Invoke-AnythingLLMUpload -AnythingLLMUrl "http://myserver:3001" -InputPath "D:\Ergebnis" -WorkspaceSlug "knowledge-base"

.EXAMPLE
    # Only specific subfolders, with custom batch size
    Invoke-AnythingLLMUpload -AnythingLLMUrl "http://myserver:3001" -InputPath "D:\Ergebnis" `
      -WorkspaceSlug "knowledge-base" -Folders "Vertraege","Rechnungen" -BatchSize 5 -BatchPauseSec 10

.EXAMPLE
    # Upload only (no embedding), for later batch embedding
    Invoke-AnythingLLMUpload -AnythingLLMUrl "http://myserver:3001" -InputPath "D:\Ergebnis" `
      -WorkspaceSlug "knowledge-base" -UploadOnly

.EXAMPLE
    # Re-upload all files, force overwrite
    Invoke-AnythingLLMUpload -AnythingLLMUrl "http://myserver:3001" -InputPath "D:\Ergebnis" `
      -WorkspaceSlug "knowledge-base" -Force
#>

function Invoke-AnythingLLMUpload {
  [CmdletBinding(DefaultParameterSetName = 'Folder')]
  param(
    [Parameter(Mandatory, ParameterSetName = 'Folder')]
    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$AnythingLLMUrl,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [string]$ApiKey,

    [Parameter(Mandatory, ParameterSetName = 'Folder')]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$InputPath,

    [Parameter(Mandatory, ParameterSetName = 'Folder')]
    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$WorkspaceSlug,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [string]$DocumentFolder = 'pipeline-upload',

    [Parameter(ParameterSetName = 'Folder')]
    [string[]]$Folders,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [string[]]$Extensions = @('.md'),

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [ValidateRange(1, 500)]
    [int]$BatchSize = 10,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [ValidateRange(0, 300)]
    [int]$BatchPauseSec = 5,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [ValidateRange(30, 3600)]
    [int]$EmbeddingTimeoutSec = 300,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [ValidateRange(3, 120)]
    [int]$EmbeddingPollIntervalSec = 10,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [bool]$SkipExisting = $true,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [switch]$Force,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [switch]$UploadOnly,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [ValidateRange(10, 3600)]
    [int]$TimeoutSec = 120,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [ValidateRange(0, 10)]
    [int]$RetryCount = 3,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter(ParameterSetName = 'Folder')]
    [Parameter(ParameterSetName = 'Csv')]
    [string]$LogDir,

    [Parameter(Mandatory, ParameterSetName = 'Gui')]
    [switch]$Gui,

    [Parameter(ParameterSetName = 'Gui')]
    [string]$GuiInputPath
  )

  # ================================================================
  # GUI MODE
  # ================================================================
  if ($PSCmdlet.ParameterSetName -eq 'Gui') {
    $guiInputHtml = if ($GuiInputPath) { $GuiInputPath } else { '' }
    $guiInputJs = if ($GuiInputPath) { $GuiInputPath -replace '\\', '\\\\' } else { '' }
    $scriptDirJs = ($PSScriptRoot -replace '\\', '\\\\').Replace("'", "\\'")
    $guiHtml = @"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AnythingLLM Upload</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet" onerror="this.remove()">
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
    --warning: #FFB74D;
    --danger: #E25C5C;
    --radius-card: 24px;
    --radius-btn: 20px;
    --font-head: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
  }
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: var(--bg-primary); color: var(--text-primary); line-height: 1.6; min-height: 100vh;
  }
  .container { max-width: 900px; margin: 0 auto; padding: 2rem; }
  .header { margin-bottom: 2rem; }
  .header h1 {
    font-family: var(--font-head); font-size: 2rem; font-weight: 500; letter-spacing: -0.02em; margin-bottom: 0.3rem;
  }
  .header h1 span {
    background: linear-gradient(90deg, #4285F4, #9B72CB, #D96570);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; font-weight: 600;
  }
  .header .sub { font-size: 0.85rem; color: var(--text-muted); }
  .panel {
    background: var(--bg-card); border-radius: var(--radius-card); padding: 1.8rem 2rem;
    border: 1px solid var(--border-glass); margin-bottom: 1.5rem;
  }
  .panel h2 { font-family: var(--font-head); font-size: 1.05rem; font-weight: 500; margin-bottom: 1rem; }
  .field-row { display: flex; gap: 0.8rem; flex-wrap: wrap; align-items: end; margin-bottom: 1rem; }
  .field { flex: 1; min-width: 180px; }
  .field label {
    font-size: 0.78rem; font-weight: 600; color: var(--text-secondary); display: block; margin-bottom: 0.3rem;
  }
  .field input[type="text"], .field input[type="number"] {
    width: 100%; padding: 0.5rem 0.8rem; border-radius: 8px;
    border: 1px solid var(--border-glass); background: var(--bg-secondary);
    color: var(--text-primary); font-size: 0.82rem;
    font-family: 'JetBrains Mono', 'Cascadia Code', 'Consolas', monospace; outline: none;
    transition: border-color 0.2s;
  }
  .field input:focus { border-color: var(--accent); }
  .field-sm { flex: 0 0 auto; min-width: 80px; }
  .field-sm input { width: 4rem; text-align: center; }
  .key-hint {
    display: flex; align-items: flex-start; gap: 0.6rem; padding: 0.7rem 1rem;
    background: rgba(168, 199, 250, 0.08); border: 1px solid rgba(168, 199, 250, 0.15);
    border-radius: 10px; font-size: 0.78rem; color: var(--text-secondary); margin-bottom: 1.2rem;
  }
  .key-hint .icon { font-size: 1.1rem; flex-shrink: 0; }
  details { margin-bottom: 1rem; }
  summary {
    cursor: pointer; font-size: 0.9rem; font-weight: 600; color: var(--text-primary);
    padding: 0.5rem 0; user-select: none;
  }
  .opt-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0.8rem; margin-top: 0.8rem; }
  @media (max-width: 600px) { .opt-grid { grid-template-columns: 1fr; } }
  .opt-group {
    background: var(--bg-secondary); border-radius: 12px; padding: 0.7rem 1rem;
    border: 1px solid var(--border-glass);
  }
  .opt-group h4 {
    font-size: 0.72rem; font-weight: 600; color: var(--text-muted);
    text-transform: uppercase; letter-spacing: 0.05em; margin: 0 0 0.4rem 0;
  }
  .opt-row { display: flex; flex-wrap: wrap; gap: 0.3rem 0.8rem; }
  .opt-row label {
    display: flex; align-items: center; gap: 0.3rem; font-size: 0.8rem;
    color: var(--text-secondary); cursor: pointer;
  }
  .opt-row input[type="checkbox"] { accent-color: var(--accent); }
  .opt-row input[type="number"] {
    width: 3.5rem; padding: 0.3rem 0.5rem; border-radius: 6px;
    border: 1px solid var(--border-glass); background: var(--bg-secondary);
    color: var(--text-primary); font-size: 0.82rem;
    font-family: 'JetBrains Mono', monospace; text-align: center;
  }
  .cmd-section { margin-top: 1.2rem; }
  .cmd-section label {
    font-size: 0.78rem; font-weight: 600; color: var(--text-secondary); display: block; margin-bottom: 0.3rem;
  }
  .cmd-wrap { position: relative; }
  .cmd-pre {
    background: var(--bg-secondary); border: 1px solid var(--border-glass); border-radius: 8px;
    padding: 0.8rem 1rem;
    font-family: 'JetBrains Mono', 'Cascadia Code', 'Consolas', monospace;
    font-size: 0.75rem; color: var(--text-primary); white-space: pre-wrap;
    word-break: break-all; margin: 0; cursor: pointer; min-height: 3rem;
  }
  .btn-copy {
    position: absolute; top: 0.3rem; right: 0.3rem; background: var(--accent); color: #111;
    border: none; border-radius: 6px; padding: 0.25rem 0.7rem; font-size: 0.72rem;
    cursor: pointer; font-weight: 600; transition: opacity 0.2s;
  }
  .btn-copy:hover { opacity: 0.85; }
  .copy-msg {
    font-size: 0.72rem; color: var(--success); margin-top: 0.2rem;
    opacity: 0; transition: opacity 0.3s;
  }
  .footer {
    text-align: center; font-size: 0.75rem; color: var(--text-muted);
    margin-top: 2rem; padding-top: 1rem; border-top: 1px solid var(--border-glass);
  }
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>AnythingLLM Upload <span>&bull; Konfiguration</span></h1>
    <div class="sub">Standalone GUI &mdash; Konfiguriere den Upload-Befehl und kopiere ihn in PowerShell.</div>
  </div>

  <div class="panel">
    <h2>&#128279; Verbindung</h2>
    <div class="key-hint">
      <span class="icon">&#128273;</span>
      <div>Der <strong>API Key</strong> wird beim Ausf&uuml;hren des Befehls abgefragt und als SecureString in der PowerShell-Session gecached. Er wird <em>nicht</em> hier gespeichert.</div>
    </div>
    <div class="field-row">
      <div class="field">
        <label>AnythingLLM URL</label>
        <input type="text" id="allm-url" placeholder="http://localhost:3001" oninput="save();build()">
      </div>
      <div class="field">
        <label>Workspace Slug</label>
        <input type="text" id="allm-workspace" placeholder="knowledge-base" oninput="save();build()">
      </div>
      <div class="field">
        <label>Document Folder</label>
        <input type="text" id="allm-docfolder" value="pipeline-upload" placeholder="pipeline-upload" oninput="save();build()">
      </div>
    </div>
  </div>

  <div class="panel">
    <h2>&#128194; Dateien</h2>
    <div class="field-row">
      <div class="field" style="flex:2">
        <label>Input-Ordner</label>
        <input type="text" id="allm-input" value="$guiInputHtml" placeholder="C:\TEMP\Result" oninput="save();build()">
      </div>
      <div class="field" style="flex:0.7;min-width:100px">
        <label>Extensions</label>
        <input type="text" id="allm-ext" value=".md" placeholder=".md,.txt" oninput="save();build()">
      </div>
    </div>
    <div class="field-row" style="margin-bottom:0">
      <div class="field" style="flex:2">
        <label>Ordner-Filter (optional, kommasepariert)</label>
        <input type="text" id="allm-folders" placeholder="Vertraege, Rechnungen (leer = alle)" oninput="save();build()">
      </div>
    </div>
  </div>

  <div class="panel">
    <details>
      <summary>&#9881;&#65039; Upload &amp; Embedding Optionen</summary>
      <div class="opt-grid">
        <div class="opt-group">
          <h4>Batch</h4>
          <div class="opt-row">
            <label>Size <input type="number" id="allm-batch" value="10" min="1" max="500" oninput="save();build()"></label>
            <label>Pause <input type="number" id="allm-pause" value="5" min="0" max="300" oninput="save();build()"> Sek.</label>
          </div>
        </div>
        <div class="opt-group">
          <h4>Embedding</h4>
          <div class="opt-row">
            <label>Timeout <input type="number" id="allm-emb-timeout" value="300" min="30" max="3600" oninput="save();build()"> Sek.</label>
            <label>Poll <input type="number" id="allm-emb-poll" value="10" min="3" max="120" oninput="save();build()"> Sek.</label>
          </div>
        </div>
        <div class="opt-group">
          <h4>Retries &amp; Timeout</h4>
          <div class="opt-row">
            <label>Retries <input type="number" id="allm-retries" value="3" min="0" max="10" oninput="save();build()"></label>
            <label>Timeout <input type="number" id="allm-timeout" value="120" min="10" max="3600" oninput="save();build()"> Sek./Upload</label>
          </div>
        </div>
        <div class="opt-group">
          <h4>Modus</h4>
          <div class="opt-row">
            <label><input type="checkbox" id="allm-skip" checked onchange="save();build()"> Skip existing</label>
            <label><input type="checkbox" id="allm-uploadonly" onchange="save();build()"> Upload only</label>
          </div>
        </div>
      </div>
    </details>

    <div class="cmd-section">
      <label>Generierter Befehl</label>
      <div class="cmd-wrap">
        <pre class="cmd-pre" id="cmd" onclick="copyCmd()"></pre>
        <button class="btn-copy" onclick="copyCmd()">&#128203; Kopieren</button>
      </div>
      <div class="copy-msg" id="copy-msg"></div>
    </div>
  </div>

  <div class="footer">Invoke-AnythingLLMUpload v1.1 &mdash; Standalone GUI</div>
</div>

<script>
var SCRIPT_ROOT = '$scriptDirJs';

function g(id) { return document.getElementById(id); }
function v(id) { var e = g(id); return e ? (e.type === 'checkbox' ? e.checked : e.value) : ''; }

function build() {
  var url = v('allm-url'), ws = v('allm-workspace'), input = v('allm-input');
  var el = g('cmd');
  if (!url || !ws || !input) { el.textContent = 'Bitte URL, Workspace und Input-Ordner angeben.'; return; }

  var docFolder = v('allm-docfolder') || 'pipeline-upload';
  var ext = v('allm-ext') || '.md';
  var folders = v('allm-folders') || '';
  var batch = parseInt(v('allm-batch') || '10', 10);
  var pause = parseInt(v('allm-pause') || '5', 10);
  var embTimeout = parseInt(v('allm-emb-timeout') || '300', 10);
  var embPoll = parseInt(v('allm-emb-poll') || '10', 10);
  var retries = parseInt(v('allm-retries') || '3', 10);
  var timeout = parseInt(v('allm-timeout') || '120', 10);
  var skip = v('allm-skip');
  var uploadOnly = v('allm-uploadonly');

  var bt = String.fromCharCode(96);
  var sr = SCRIPT_ROOT.replace(/\\\\/g, '\\');
  var cmd = '. "' + sr + '\\Invoke-AnythingLLMUpload.ps1"\n';
  cmd += 'Invoke-AnythingLLMUpload -AnythingLLMUrl "' + url + '" ' + bt + '\n';
  cmd += '  -InputPath "' + input + '" ' + bt + '\n';
  cmd += '  -WorkspaceSlug "' + ws + '"';

  if (docFolder && docFolder !== 'pipeline-upload') cmd += ' ' + bt + '\n  -DocumentFolder "' + docFolder + '"';

  // Folders filter
  var fArr = folders.split(',').map(function(f){return f.trim();}).filter(function(f){return f;});
  if (fArr.length > 0) cmd += ' ' + bt + '\n  -Folders ' + fArr.map(function(f){return '"'+f+'"';}).join(',');

  // Extensions
  var exts = ext.split(',').map(function(e){return e.trim();}).filter(function(e){return e;});
  if (exts.length > 0 && !(exts.length === 1 && exts[0] === '.md')) {
    cmd += ' ' + bt + '\n  -Extensions ' + exts.map(function(e){return "'"+e+"'";}).join(',');
  }

  var extras = '';
  if (batch !== 10) extras += ' -BatchSize ' + batch;
  if (pause !== 5) extras += ' -BatchPauseSec ' + pause;
  if (embTimeout !== 300) extras += ' -EmbeddingTimeoutSec ' + embTimeout;
  if (embPoll !== 10) extras += ' -EmbeddingPollIntervalSec ' + embPoll;
  if (!skip) extras += ' -Force';
  if (uploadOnly) extras += ' -UploadOnly';
  if (retries !== 3) extras += ' -RetryCount ' + retries;
  if (timeout !== 120) extras += ' -TimeoutSec ' + timeout;
  if (extras) cmd += ' ' + bt + '\n ' + extras.trim();

  el.textContent = cmd;
}

function copyCmd() {
  var el = g('cmd'); if (!el) return;
  navigator.clipboard.writeText(el.textContent).then(function() {
    var msg = g('copy-msg');
    msg.textContent = 'In Zwischenablage kopiert!';
    msg.style.opacity = '1';
    setTimeout(function(){ msg.style.opacity = '0'; }, 2000);
  });
}

var FIELDS = ['allm-url','allm-workspace','allm-docfolder','allm-input','allm-ext','allm-folders',
  'allm-batch','allm-pause','allm-emb-timeout','allm-emb-poll','allm-retries','allm-timeout'];
var CHECKS = ['allm-skip','allm-uploadonly'];

function save() {
  try {
    var s = {};
    FIELDS.forEach(function(id){ var e = g(id); if (e) s[id] = e.value; });
    CHECKS.forEach(function(id){ var e = g(id); if (e) s[id] = e.checked; });
    localStorage.setItem('allm_gui_settings', JSON.stringify(s));
  } catch(e) {}
}

function load() {
  try {
    var raw = localStorage.getItem('allm_gui_settings'); if (!raw) return;
    var s = JSON.parse(raw);
    FIELDS.forEach(function(id){ if (s[id] !== undefined) { var e = g(id); if (e) e.value = s[id]; } });
    CHECKS.forEach(function(id){ if (s.hasOwnProperty(id)) { var e = g(id); if (e) e.checked = s[id]; } });
  } catch(e) {}
}

// Init
load();
// Override input if passed via PowerShell
var initInput = '$guiInputJs'.replace(/\\\\/g, '\\');
if (initInput) { var e = g('allm-input'); if (e && !e.value) e.value = initInput; }
build();
</script>
</body>
</html>
"@

    $guiDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
    $guiPath = Join-Path $guiDir 'AnythingLLM-Upload.html'
    [System.IO.File]::WriteAllText($guiPath, $guiHtml, [System.Text.UTF8Encoding]::new($true))
    Write-Host "[Invoke-AnythingLLMUpload] GUI: $guiPath" -ForegroundColor Green
    Start-Process $guiPath
    Write-Host "[Invoke-AnythingLLMUpload] GUI opened in browser." -ForegroundColor Cyan
    return
  }

  # ================================================================
  # INIT
  # ================================================================
  $AnythingLLMUrl = $AnythingLLMUrl.TrimEnd('/')
  if ($Force) { $SkipExisting = $false }
  if (-not $LogDir) { $LogDir = $PSScriptRoot }

  foreach ($dir in @($LogDir)) {
    if (-not (Test-Path $dir)) {
      New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
  }
  $LogDir = (Resolve-Path $LogDir).Path

  $logCsvPath = Join-Path $LogDir 'upload_log.csv'
  $runsJsonPath = Join-Path $LogDir 'upload_runs.json'

  # ================================================================
  # API KEY HANDLING (SecureString session cache)
  # ================================================================
  if (-not $ApiKey) {
    if (-not $script:AnythingLLMApiKey) {
      Write-Host "[Invoke-AnythingLLMUpload] Kein API Key angegeben." -ForegroundColor Yellow
      $script:AnythingLLMApiKey = Read-Host "AnythingLLM API Key eingeben" -AsSecureString
    }
    # SecureString -> plaintext for HTTP header
    $ApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
      [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($script:AnythingLLMApiKey)
    )
  }

  $authHeaders = @{
    'Authorization' = "Bearer $ApiKey"
    'Accept'        = 'application/json'
  }

  Write-Host "`n[Invoke-AnythingLLMUpload] ================================" -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] URL:          $AnythingLLMUrl" -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] Workspace:    $WorkspaceSlug" -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] Folder:       $DocumentFolder" -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] Batch:        $BatchSize files, ${BatchPauseSec}s pause" -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] Extensions:   $($Extensions -join ', ')" -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] Skip existing: $SkipExisting" -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] Upload only:  $UploadOnly" -ForegroundColor Cyan

  # ================================================================
  # API HEALTH CHECK
  # ================================================================
  Write-Host "`n[Invoke-AnythingLLMUpload] Checking API health..." -ForegroundColor Yellow
  $apiReachable = $false

  foreach ($ep in @("$AnythingLLMUrl/api/v1/auth", "$AnythingLLMUrl/api/docs")) {
    try {
      $null = Invoke-WebRequest -Uri $ep -Method GET -Headers $authHeaders -TimeoutSec 10 -ErrorAction Stop -UseBasicParsing
      $apiReachable = $true
      Write-Host "[Invoke-AnythingLLMUpload] API is reachable ($ep)." -ForegroundColor Green
      break
    }
    catch {
      # 403 means API is reachable but auth failed or not needed for this endpoint
      if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 403) {
        $apiReachable = $true
        Write-Host "[Invoke-AnythingLLMUpload] API is reachable ($ep)." -ForegroundColor Green
        break
      }
    }
  }

  if (-not $apiReachable) {
    Write-Error "[Invoke-AnythingLLMUpload] Cannot reach AnythingLLM API at '$AnythingLLMUrl'. Please check the URL."
    return
  }

  # Verify API key works
  try {
    $null = Invoke-RestMethod -Uri "$AnythingLLMUrl/api/v1/auth" -Method GET -Headers $authHeaders -TimeoutSec 10 -ErrorAction Stop
    Write-Host "[Invoke-AnythingLLMUpload] API Key valid." -ForegroundColor Green
  }
  catch {
    $statusCode = if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 }
    if ($statusCode -eq 403) {
      Write-Error "[Invoke-AnythingLLMUpload] API Key is invalid (403 Forbidden). Please check your key."
      $script:AnythingLLMApiKey = $null
      return
    }
  }

  # ================================================================
  # COLLECT FILES
  # ================================================================
  $filesToProcess = @()
  $inputRoot = ''

  if ($PSCmdlet.ParameterSetName -eq 'Csv') {
    # CSV mode
    Write-Host "[Invoke-AnythingLLMUpload] Reading CSV: $CsvPath" -ForegroundColor Cyan
    $csvData = Import-Csv -Path $CsvPath -Encoding UTF8
    if (-not $csvData -or $csvData.Count -eq 0) {
      Write-Warning "[Invoke-AnythingLLMUpload] CSV is empty."
      return
    }

    $firstPath = $csvData[0].FullPath
    if (-not $firstPath) {
      Write-Error "[Invoke-AnythingLLMUpload] CSV must contain a 'FullPath' column."
      return
    }

    $inputRoot = Split-Path $firstPath -Parent
    Write-Warning "[Invoke-AnythingLLMUpload] Input root from CSV: $inputRoot"

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
    Write-Host "[Invoke-AnythingLLMUpload] Input:        $InputPath" -ForegroundColor Cyan

    $gciParams = @{ Path = $InputPath; File = $true; Recurse = $true; ErrorAction = 'SilentlyContinue' }
    $allInputFiles = Get-ChildItem @gciParams

    # Filter by folders
    if ($Folders -and $Folders.Count -gt 0) {
      Write-Host "[Invoke-AnythingLLMUpload] Folder filter: $($Folders -join ', ')" -ForegroundColor Cyan
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
      Write-Host "[Invoke-AnythingLLMUpload] Extension filter: $($Extensions -join ', ')" -ForegroundColor Cyan
    }

    foreach ($f in $allInputFiles) {
      $filesToProcess += [PSCustomObject]@{
        FullPath = $f.FullName
        Name     = $f.Name
        SizeMB   = [math]::Round($f.Length / 1MB, 2)
      }
    }
  }

  Write-Host "[Invoke-AnythingLLMUpload] Files found:  $($filesToProcess.Count)" -ForegroundColor Green

  if ($filesToProcess.Count -eq 0) {
    Write-Warning "[Invoke-AnythingLLMUpload] No files to process."
    return
  }

  # ================================================================
  # CALCULATE RELATIVE PATHS
  # ================================================================
  $filesToUpload = @()
  foreach ($f in $filesToProcess) {
    $relPath = ''
    if ($f.FullPath.StartsWith($inputRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
      $relPath = $f.FullPath.Substring($inputRoot.Length).TrimStart('\')
    }
    else {
      $relPath = $f.Name
    }
    $filesToUpload += [PSCustomObject]@{
      FullPath     = $f.FullPath
      Name         = $f.Name
      SizeMB       = $f.SizeMB
      RelativePath = $relPath
    }
  }

  # ================================================================
  # SKIP EXISTING CHECK (based on upload log)
  # ================================================================
  $skippedExisting = 0
  if ($SkipExisting -and (Test-Path $logCsvPath)) {
    try {
      $previousLog = Import-Csv -Path $logCsvPath -Encoding UTF8
      $uploadedPaths = @{}
      foreach ($entry in $previousLog) {
        if ($entry.Status -eq 'OK') {
          $uploadedPaths[$entry.RelativePath] = $true
        }
      }

      if ($uploadedPaths.Count -gt 0) {
        $filtered = @()
        foreach ($f in $filesToUpload) {
          if ($uploadedPaths.ContainsKey($f.RelativePath)) {
            $skippedExisting++
          }
          else {
            $filtered += $f
          }
        }
        $filesToUpload = $filtered
      }
    }
    catch { Write-Warning "[Invoke-AnythingLLMUpload] Could not read upload log: $_" }
  }

  if ($skippedExisting -gt 0) {
    Write-Host "[Invoke-AnythingLLMUpload] Skipped (existing): $skippedExisting" -ForegroundColor DarkGray
  }

  Write-Host "[Invoke-AnythingLLMUpload] To upload:    $($filesToUpload.Count)" -ForegroundColor Green

  if ($filesToUpload.Count -eq 0) {
    Write-Host "[Invoke-AnythingLLMUpload] All files already uploaded. Nothing to do." -ForegroundColor Green
    return
  }

  $totalSizeMB = [math]::Round(($filesToUpload | Measure-Object -Property SizeMB -Sum).Sum, 2)
  Write-Host "[Invoke-AnythingLLMUpload] Total size:   $totalSizeMB MB" -ForegroundColor Cyan

  # ================================================================
  # UPLOAD + EMBEDDING LOOP
  # ================================================================
  $logEntries = @()
  $manifestEntries = @{}
  $batchStart = Get-Date
  $uploaded = 0
  $embedded = 0
  $failed = 0
  $totalFiles = $filesToUpload.Count
  $current = 0
  $totalUploadTime = 0
  $batchDocLocations = @()  # Collected document locations for batch embedding
  $batchCounter = 0
  $batchNum = 0

  $batchStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

  Write-Host "`n[Invoke-AnythingLLMUpload] Starting upload..." -ForegroundColor Cyan
  Write-Host "[Invoke-AnythingLLMUpload] Phase 1: Upload | Phase 2: $(if ($UploadOnly) { 'SKIPPED (-UploadOnly)' } else { 'Embedding' })" -ForegroundColor Cyan

  foreach ($file in $filesToUpload) {
    $current++
    $batchCounter++
    $pctComplete = [math]::Round(($current / $totalFiles) * 100)
    $fileStart = [System.Diagnostics.Stopwatch]::StartNew()

    # Progress bar with ETA
    $eta = ''
    if ($uploaded -gt 0) {
      $avgSec = $totalUploadTime / $uploaded
      $remaining = ($totalFiles - $current + 1) * $avgSec
      $etaTs = [TimeSpan]::FromSeconds([math]::Max(0, $remaining))
      $elapsed = $batchStopwatch.Elapsed
      $eta = " | ETA: $($etaTs.ToString('hh\:mm\:ss')) | Elapsed: $($elapsed.ToString('hh\:mm\:ss')) | Avg: $([math]::Round($avgSec,1))s/file"
    }

    Write-Progress -Activity "AnythingLLM Upload" `
      -Status "[$current/$totalFiles] $($file.Name) ($($file.SizeMB) MB)$eta" `
      -PercentComplete $pctComplete

    # ── Phase 1: Upload ──
    $uploadSuccess = $false
    $errorText = ''
    $documentLocation = ''

    for ($attempt = 1; $attempt -le ($RetryCount + 1); $attempt++) {
      try {
        $uploadUrl = if ($DocumentFolder) {
          "$AnythingLLMUrl/api/v1/document/upload/$DocumentFolder"
        }
        else {
          "$AnythingLLMUrl/api/v1/document/upload"
        }

        # Build multipart form data
        $fileBytes = [System.IO.File]::ReadAllBytes($file.FullPath)
        $fileName = [System.IO.Path]::GetFileName($file.FullPath)

        $boundary = [System.Guid]::NewGuid().ToString('N')
        $LF = "`r`n"
        $enc = [System.Text.Encoding]::UTF8
        $bodyParts = [System.Collections.Generic.List[byte[]]]::new()

        # File part
        $fileHeader = "--$boundary$LF" +
          "Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"$LF" +
          "Content-Type: application/octet-stream$LF$LF"
        $bodyParts.Add($enc.GetBytes($fileHeader))
        $bodyParts.Add($fileBytes)
        $bodyParts.Add($enc.GetBytes($LF))

        # Close boundary
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

        $uploadHeaders = @{
          'Authorization' = "Bearer $ApiKey"
          'Accept'        = 'application/json'
        }

        $response = Invoke-WebRequest -Uri $uploadUrl -Method Post -ContentType $contentType `
          -Body $bodyBytes -Headers $uploadHeaders -TimeoutSec $TimeoutSec -ErrorAction Stop -UseBasicParsing

        $responseObj = $response.Content | ConvertFrom-Json

        if ($responseObj.success) {
          $uploadSuccess = $true
          if ($responseObj.documents -and $responseObj.documents.Count -gt 0) {
            $documentLocation = $responseObj.documents[0].location
          }
          break
        }
        else {
          $errorText = if ($responseObj.error) { $responseObj.error } else { 'Unknown API error' }
          throw $errorText
        }
      }
      catch {
        $errorText = $_.Exception.Message
        if ($attempt -le $RetryCount) {
          $waitSec = [math]::Pow(2, $attempt)
          Write-Warning "[Invoke-AnythingLLMUpload] Attempt $attempt failed for '$($file.Name)': $errorText. Retrying in ${waitSec}s..."
          Start-Sleep -Seconds $waitSec
        }
        else {
          Write-Warning "[Invoke-AnythingLLMUpload] FAILED after $($RetryCount + 1) attempts: '$($file.Name)': $errorText"
        }
      }
    }

    $fileStart.Stop()
    $fileDuration = [math]::Round($fileStart.Elapsed.TotalSeconds, 2)

    if ($uploadSuccess) {
      $uploaded++
      $totalUploadTime += $fileDuration
      $batchDocLocations += $documentLocation

      # Track for manifest
      $manifestEntries[$file.RelativePath] = @{
        status           = 'hochgeladen'
        processedAt      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        sourceScript     = 'Invoke-AnythingLLMUpload'
        workspaceSlug    = $WorkspaceSlug
        documentLocation = $documentLocation
      }

      $logEntries += [PSCustomObject]@{
        Timestamp        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        File             = $file.Name
        RelativePath     = $file.RelativePath
        SizeMB           = $file.SizeMB
        Status           = 'OK'
        DurationSec      = $fileDuration
        DocumentLocation = $documentLocation
        Error            = ''
      }

      Write-Host "  [OK] $($file.Name) (${fileDuration}s) -> $documentLocation" -ForegroundColor Green
    }
    else {
      $failed++

      $logEntries += [PSCustomObject]@{
        Timestamp        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        File             = $file.Name
        RelativePath     = $file.RelativePath
        SizeMB           = $file.SizeMB
        Status           = 'FAILED'
        DurationSec      = $fileDuration
        DocumentLocation = ''
        Error            = $errorText
      }

      Write-Host "  [FAIL] $($file.Name) (${fileDuration}s) - $errorText" -ForegroundColor Red
      continue
    }

    # ── Phase 2: Batch Embedding ──
    if (-not $UploadOnly -and $batchCounter -ge $BatchSize) {
      $batchNum++
      $validLocations = @($batchDocLocations | Where-Object { $_ })

      if ($validLocations.Count -gt 0) {
        Write-Host "`n  [EMBED] Batch ${batchNum}: Embedding $($validLocations.Count) documents into workspace '$WorkspaceSlug'..." -ForegroundColor Yellow

        try {
          $embedBody = @{
            adds    = $validLocations
            deletes = @()
          } | ConvertTo-Json -Depth 3

          $null = Invoke-RestMethod -Uri "$AnythingLLMUrl/api/v1/workspace/$WorkspaceSlug/update-embeddings" `
            -Method Post -Body $embedBody -ContentType 'application/json' `
            -Headers $authHeaders -TimeoutSec $TimeoutSec -ErrorAction Stop

          # Poll embedding status
          $pollDeadline = [DateTime]::UtcNow.AddSeconds($EmbeddingTimeoutSec)
          $allCached = $false

          while (-not $allCached -and [DateTime]::UtcNow -lt $pollDeadline) {
            Start-Sleep -Seconds $EmbeddingPollIntervalSec

            try {
              $docsResponse = Invoke-RestMethod -Uri "$AnythingLLMUrl/api/v1/documents/folder/$DocumentFolder" `
                -Method Get -Headers $authHeaders -TimeoutSec 30 -ErrorAction Stop

              if ($docsResponse.documents) {
                $locationNames = $validLocations | ForEach-Object { [System.IO.Path]::GetFileName($_) }
                $matchingDocs = @($docsResponse.documents | Where-Object { $locationNames -contains $_.name })
                $cachedCount = @($matchingDocs | Where-Object { $_.cached -eq $true }).Count

                Write-Progress -Activity "AnythingLLM Embedding" `
                  -Status "Batch ${batchNum}: $cachedCount/$($validLocations.Count) embedded" `
                  -PercentComplete ([math]::Min(100, [math]::Round(($cachedCount / [math]::Max(1, $validLocations.Count)) * 100)))

                if ($cachedCount -ge $validLocations.Count) {
                  $allCached = $true
                  $embedded += $cachedCount
                  Write-Host "  [EMBED] Batch ${batchNum}: All $cachedCount documents embedded successfully." -ForegroundColor Green
                }
              }
            }
            catch {
              Write-Warning "[Invoke-AnythingLLMUpload] Embedding poll error: $($_.Exception.Message)"
            }
          }

          if (-not $allCached) {
            Write-Warning "[Invoke-AnythingLLMUpload] Embedding timeout for Batch ${batchNum} after ${EmbeddingTimeoutSec}s. Some documents may not be fully embedded yet."
          }

          Write-Progress -Activity "AnythingLLM Embedding" -Completed
        }
        catch {
          Write-Warning "[Invoke-AnythingLLMUpload] Embedding request failed for Batch ${batchNum}: $($_.Exception.Message)"
        }
      }

      # Reset batch counters
      $batchDocLocations = @()
      $batchCounter = 0

      # Batch pause
      if ($current -lt $totalFiles -and $BatchPauseSec -gt 0) {
        Write-Host "  [PAUSE] ${BatchPauseSec}s (Vector-DB verarbeitet)..." -ForegroundColor DarkGray
        Start-Sleep -Seconds $BatchPauseSec
      }
    }
  }

  # ── Handle remaining files in last incomplete batch ──
  if (-not $UploadOnly -and $batchDocLocations.Count -gt 0) {
    $batchNum++
    $validLocations = @($batchDocLocations | Where-Object { $_ })

    if ($validLocations.Count -gt 0) {
      Write-Host "`n  [EMBED] Batch ${batchNum} (final): Embedding $($validLocations.Count) documents..." -ForegroundColor Yellow

      try {
        $embedBody = @{
          adds    = $validLocations
          deletes = @()
        } | ConvertTo-Json -Depth 3

        $null = Invoke-RestMethod -Uri "$AnythingLLMUrl/api/v1/workspace/$WorkspaceSlug/update-embeddings" `
          -Method Post -Body $embedBody -ContentType 'application/json' `
          -Headers $authHeaders -TimeoutSec $TimeoutSec -ErrorAction Stop

        $pollDeadline = [DateTime]::UtcNow.AddSeconds($EmbeddingTimeoutSec)
        $allCached = $false

        while (-not $allCached -and [DateTime]::UtcNow -lt $pollDeadline) {
          Start-Sleep -Seconds $EmbeddingPollIntervalSec

          try {
            $docsResponse = Invoke-RestMethod -Uri "$AnythingLLMUrl/api/v1/documents/folder/$DocumentFolder" `
              -Method Get -Headers $authHeaders -TimeoutSec 30 -ErrorAction Stop

            if ($docsResponse.documents) {
              $locationNames = $validLocations | ForEach-Object { [System.IO.Path]::GetFileName($_) }
              $matchingDocs = @($docsResponse.documents | Where-Object { $locationNames -contains $_.name })
              $cachedCount = @($matchingDocs | Where-Object { $_.cached -eq $true }).Count

              Write-Progress -Activity "AnythingLLM Embedding" `
                -Status "Batch ${batchNum}: $cachedCount/$($validLocations.Count) embedded" `
                -PercentComplete ([math]::Min(100, [math]::Round(($cachedCount / [math]::Max(1, $validLocations.Count)) * 100)))

              if ($cachedCount -ge $validLocations.Count) {
                $allCached = $true
                $embedded += $cachedCount
                Write-Host "  [EMBED] Batch ${batchNum}: All $cachedCount documents embedded successfully." -ForegroundColor Green
              }
            }
          }
          catch {
            Write-Warning "[Invoke-AnythingLLMUpload] Embedding poll error: $($_.Exception.Message)"
          }
        }

        if (-not $allCached) {
          Write-Warning "[Invoke-AnythingLLMUpload] Embedding timeout for final batch after ${EmbeddingTimeoutSec}s."
        }

        Write-Progress -Activity "AnythingLLM Embedding" -Completed
      }
      catch {
        Write-Warning "[Invoke-AnythingLLMUpload] Embedding request failed for final batch: $($_.Exception.Message)"
      }
    }
  }

  Write-Progress -Activity "AnythingLLM Upload" -Completed
  $batchStopwatch.Stop()

  # ================================================================
  # WRITE LOGS
  # ================================================================
  $batchEnd = Get-Date
  $batchDuration = ($batchEnd - $batchStart).TotalSeconds

  # Per-file log (append)
  if ($logEntries.Count -gt 0) {
    $logExists = Test-Path $logCsvPath
    $logEntries | Export-Csv -Path $logCsvPath -NoTypeInformation -Encoding UTF8 -Append:$logExists
    Write-Host "`n[Invoke-AnythingLLMUpload] Per-file log: $logCsvPath" -ForegroundColor DarkGray
  }

  # Error log
  $failedEntries = @($logEntries | Where-Object { $_.Status -eq 'FAILED' })
  if ($failedEntries.Count -gt 0) {
    $errorCsvPath = Join-Path $LogDir 'upload_errors.csv'
    $failedEntries | Export-Csv -Path $errorCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "[Invoke-AnythingLLMUpload] Error log:    $errorCsvPath" -ForegroundColor Red
  }

  # Run summary
  $runEntry = @{
    date            = $batchStart.ToString('yyyy-MM-dd HH:mm:ss')
    files           = $uploaded + $failed
    uploaded        = $uploaded
    embedded        = $embedded
    failed          = $failed
    skipped         = $skippedExisting
    totalSeconds    = [math]::Round($batchDuration, 2)
    avgPerFile      = if ($uploaded -gt 0) { [math]::Round($totalUploadTime / $uploaded, 2) } else { 0 }
    totalSizeMB     = $totalSizeMB
    workspaceSlug   = $WorkspaceSlug
    documentFolder  = $DocumentFolder
    uploadOnly      = [bool]$UploadOnly
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
  Write-Host "[Invoke-AnythingLLMUpload] Run summary:  $runsJsonPath" -ForegroundColor DarkGray

  # Write pipeline manifest
  if ($manifestEntries.Count -gt 0) {
    $manifestDir = if ($PSCmdlet.ParameterSetName -eq 'Folder') { $InputPath } else { $LogDir }
    $manifestPath = Join-Path $manifestDir 'pipeline_manifest.json'
    $manifest = @{ manifestVersion = 1; entries = @{} }
    if (Test-Path $manifestPath) {
      try {
        $existing = Get-Content -Path $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($existing.entries) {
          foreach ($prop in $existing.entries.PSObject.Properties) {
            $manifest.entries[$prop.Name] = @{}
            foreach ($subProp in $prop.Value.PSObject.Properties) {
              $manifest.entries[$prop.Name][$subProp.Name] = $subProp.Value
            }
          }
        }
      }
      catch { Write-Warning "[Invoke-AnythingLLMUpload] Could not read existing manifest: $_" }
    }
    foreach ($key in $manifestEntries.Keys) {
      $manifest.entries[$key] = $manifestEntries[$key]
    }
    $manifestJson = $manifest | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.Encoding]::UTF8)
    Write-Host "[Invoke-AnythingLLMUpload] Manifest:     $manifestPath ($($manifestEntries.Count) entries)" -ForegroundColor Green
  }

  # ================================================================
  # SUMMARY
  # ================================================================
  $avgTime = if ($uploaded -gt 0) { [math]::Round($totalUploadTime / $uploaded, 2) } else { 0 }

  Write-Host "`n[Invoke-AnythingLLMUpload] ================================" -ForegroundColor Green
  Write-Host "[Invoke-AnythingLLMUpload] Uploaded:     $uploaded files" -ForegroundColor Green
  if (-not $UploadOnly) { Write-Host "[Invoke-AnythingLLMUpload] Embedded:     $embedded files" -ForegroundColor Green }
  if ($failed -gt 0) { Write-Host "[Invoke-AnythingLLMUpload] Failed:       $failed files" -ForegroundColor Red }
  if ($skippedExisting -gt 0) { Write-Host "[Invoke-AnythingLLMUpload] Skipped:      $skippedExisting (already uploaded)" -ForegroundColor DarkGray }
  Write-Host "[Invoke-AnythingLLMUpload] Duration:     $([math]::Round($batchDuration, 1))s (avg ${avgTime}s/file)" -ForegroundColor Green
  Write-Host "[Invoke-AnythingLLMUpload] Workspace:    $WorkspaceSlug" -ForegroundColor Green
  Write-Host "[Invoke-AnythingLLMUpload] ================================" -ForegroundColor Green
}
