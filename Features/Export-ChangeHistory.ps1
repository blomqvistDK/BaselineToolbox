<#
.SYNOPSIS
    Extracts the "Appendix: Change History" from a benchmark-style PDF and writes
    a CSV file and a dark-mode interactive HTML report of the changes.

.DESCRIPTION
    Works with any PDF that contains an "Appendix: Change History" section whose
    entries follow the CIS pattern, e.g.:

        Date: 06/25/2026 Version: 5.0.0
        REMOVE - 25 Ensure 'Minimum Password Age' is set to '1 or more day(s)'
        Ticket #28056

    By default the LATEST version block is exported. Both outputs include the
    change type plus metadata columns for the recommendation profile and Intune
    policy name when they can be resolved.

    Metadata enrichment uses a previous-version markdown export when provided.
    Rows are matched by CIS reference first, then normalized recommendation
    text. For unmatched Add rows, the script also scans the current PDF's main
    recommendation content to derive Profile and Intune Policy using the same
    naming approach as the PDF-to-Markdown workflow in Invoke-BaselineToolbox.

    PDF text is extracted with Poppler's 'pdftotext' when available (best
    fidelity); otherwise the script falls back to Microsoft Word automation,
    which needs no third-party install on a machine that has Word.

.PARAMETER PdfPath
    Path to the input PDF. Required.

.PARAMETER Version
    Specific version to export, e.g. '5.0.0'. Defaults to the latest version.

.PARAMETER OutputDir
    Destination folder for the .csv and .html files. Defaults to the folder the
    script lives in (or the current directory when run interactively).

.PARAMETER PreviousMarkdownPath
  Optional path to the previous-version markdown or txt export. When provided,
  the script uses it to resolve Profile and Intune Policy for matching change
  rows before falling back to current-PDF lookup for unmatched Add rows.

.PARAMETER PdftotextPath
    Optional full path to pdftotext.exe if it is not on your PATH.

.EXAMPLE
    .\Export-ChangeHistory.ps1 -PdfPath '.\CIS_Microsoft_Intune_for_Windows_11_Benchmark_v5.0.0.pdf'

.EXAMPLE
    .\Export-ChangeHistory.ps1 -PdfPath '.\benchmark.pdf' -Version '4.0.0' -OutputDir 'C:\Reports'

.EXAMPLE
  .\Export-ChangeHistory.ps1 -PdfPath '.\benchmark.pdf' -PreviousMarkdownPath '.\benchmark-v4.md'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$PdfPath,
    [string]$Version,
    [string]$OutputDir,
    [string]$PreviousMarkdownPath,
    [string]$PdftotextPath,
  [double]$SimilarityThreshold = 0.80,
    [int]$FirstPage = 0,
    [int]$LastPage  = 0,
    [switch]$DiagnosticDump
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $PdfPath)) { throw "PDF not found: $PdfPath" }
$PdfPath = (Resolve-Path -LiteralPath $PdfPath).Path

if ($PreviousMarkdownPath) {
  if (-not (Test-Path -LiteralPath $PreviousMarkdownPath)) {
    throw "Previous markdown file not found: $PreviousMarkdownPath"
  }
  $PreviousMarkdownPath = (Resolve-Path -LiteralPath $PreviousMarkdownPath).Path
}

if (-not $OutputDir -or $OutputDir.Trim() -eq '') {
    if ($PSScriptRoot) { $OutputDir = $PSScriptRoot } else { $OutputDir = (Get-Location).Path }
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

if ($SimilarityThreshold -lt 0.0 -or $SimilarityThreshold -gt 1.0) {
  throw 'SimilarityThreshold must be between 0 and 1, for example 0.90.'
}

# --- Dark interactive HTML template (placeholders replaced below) ------------
$HtmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__ - Change History v__VERSION__</title>
<style>
:root{
  --bg:#0d1117;--panel:#161b22;--panel-2:#0d1117;--border:#30363d;
  --text:#e6edf3;--muted:#8b949e;--accent:#58a6ff;
  --add:#137333;--add-bg:#e6f4ea;
  --remove:#b3261e;--remove-bg:#fce8e6;
  --update:#1a56c4;--update-bg:#e8f0fe;
  --rename:#9a5b00;--rename-bg:#feefc3;
  --other:#5f6368;--other-bg:#e8eaed;
}
*{box-sizing:border-box}
html{color-scheme:dark}
body{font-family:-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;margin:0;background:var(--bg);color:var(--text);}
.wrap{max-width:1120px;margin:0 auto;padding:24px;}
header{background:var(--panel);border:1px solid var(--border);border-radius:12px;padding:22px 26px;}
header h1{margin:0 0 4px;font-size:19px;letter-spacing:-.01em;color:var(--text);}
.sub{margin:0;color:var(--muted);font-size:14px;}
.sub strong{color:var(--text);}
.sub .date{color:var(--muted);}
.stats{display:flex;flex-wrap:wrap;gap:8px;margin-top:16px;}
.stat{background:rgba(139,148,158,.12);border:1px solid var(--border);border-radius:16px;padding:4px 12px;font-size:12.5px;color:var(--muted);}
.stat b{font-weight:700;color:var(--text);}
.stat.b-add{background:var(--add-bg);border-color:transparent;color:var(--add);}
.stat.b-add b{color:var(--add);}
.stat.b-remove{background:var(--remove-bg);border-color:transparent;color:var(--remove);}
.stat.b-remove b{color:var(--remove);}
.stat.b-update{background:var(--update-bg);border-color:transparent;color:var(--update);}
.stat.b-update b{color:var(--update);}
.stat.b-rename{background:var(--rename-bg);border-color:transparent;color:var(--rename);}
.stat.b-rename b{color:var(--rename);}
.stat.b-other{background:var(--other-bg);border-color:transparent;color:var(--other);}
.stat.b-other b{color:var(--other);}
.toolbar{display:flex;flex-wrap:wrap;gap:12px;align-items:center;margin:20px 0 12px;}
#search{flex:1;min-width:240px;padding:10px 16px;border:1px solid var(--border);border-radius:24px;font-size:14px;outline:none;background:var(--panel);color:var(--text);}
#search::placeholder{color:var(--muted);}
#search:focus{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent);}
.filters{display:flex;flex-wrap:wrap;gap:6px;}
.chip{border:1px solid var(--border);background:var(--panel);border-radius:16px;padding:6px 12px;font-size:13px;cursor:pointer;color:var(--muted);}
.chip:hover{background:#1c2330;color:var(--text);}
.chip.active{background:var(--accent);color:#0d1117;border-color:var(--accent);font-weight:600;}
.btn{border:1px solid var(--accent);background:transparent;border-radius:20px;padding:8px 14px;font-size:13px;cursor:pointer;color:var(--accent);font-weight:600;}
.btn:hover{background:var(--update-bg);}
.count{font-size:13px;color:var(--muted);margin-bottom:8px;}
table{width:100%;border-collapse:collapse;background:var(--panel);border:1px solid var(--border);border-radius:12px;overflow:hidden;}
thead th{text-align:left;background:var(--panel-2);padding:12px 16px;font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);border-bottom:1px solid var(--border);}
th.sortable{cursor:pointer;user-select:none;white-space:nowrap;}
th.sortable:hover{color:var(--text);}
th.change-col{width:180px;}
th.profile-col{width:120px;}
th.policy-col{width:260px;}
.arrow{color:var(--accent);font-size:11px;}
tbody td{padding:11px 16px;font-size:14px;border-bottom:1px solid var(--border);vertical-align:top;line-height:1.45;color:var(--text);}
tbody tr:hover{background:#1c2330;}
tbody tr:last-child td{border-bottom:none;}
.badge{display:inline-block;padding:3px 11px;border-radius:12px;font-size:12px;font-weight:700;white-space:nowrap;}
.b-add{background:var(--add-bg);color:var(--add);}
.b-remove{background:var(--remove-bg);color:var(--remove);}
.b-update{background:var(--update-bg);color:var(--update);}
.b-rename{background:var(--rename-bg);color:var(--rename);}
.b-other{background:var(--other-bg);color:var(--other);}
.empty{text-align:center;padding:44px;color:var(--muted);font-size:14px;}
footer{text-align:center;color:var(--muted);font-size:12px;margin:22px 0 6px;}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>__TITLE__</h1>
    <p class="sub">Appendix: Change History &mdash; Version <strong>__VERSION__</strong> <span class="date">(__DATE__)</span></p>
    <div class="stats" id="stats"></div>
  </header>

  <div class="toolbar">
    <input id="search" type="search" placeholder="Search settings or change type...">
    <div class="filters" id="filters"></div>
    <button id="downloadCsv" class="btn" type="button">Download CSV (current view)</button>
  </div>

  <div class="count" id="count"></div>

  <table id="tbl">
    <thead>
      <tr>
        <th class="sortable" data-key="Setting">Setting <span class="arrow" data-for="Setting"></span></th>
        <th class="sortable change-col" data-key="Change">Change <span class="arrow" data-for="Change"></span></th>
        <th class="sortable profile-col" data-key="Profile">Profile <span class="arrow" data-for="Profile"></span></th>
        <th class="sortable policy-col" data-key="IntunePolicy">Intune Policy <span class="arrow" data-for="IntunePolicy"></span></th>
      </tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>

  <footer>Generated __GENERATED__ &middot; __COUNT__ total changes in v__VERSION__</footer>
</div>

<script>
const DATA = __DATA__;
const VERSION = "__VERSION__";
const state = { q: "", type: "All", sortKey: null, sortDir: 1 };

function changeClass(c){
  const k = (c || "").toLowerCase();
  if (k === "add") return "b-add";
  if (k === "remove") return "b-remove";
  if (k === "update") return "b-update";
  if (k === "rename") return "b-rename";
  return "b-other";
}
function distinctTypes(){
  const s = new Set(DATA.map(r => r.Change).filter(Boolean));
  return ["All"].concat(Array.from(s).sort());
}
function applyFilters(){
  let rows = DATA.slice();
  if (state.type !== "All") rows = rows.filter(r => r.Change === state.type);
  if (state.q){
    const q = state.q.toLowerCase();
    rows = rows.filter(r => ((r.Setting || "") + " " + (r.Change || "") + " " + (r.Profile || "") + " " + (r.IntunePolicy || "")).toLowerCase().indexOf(q) !== -1);
  }
  if (state.sortKey){
    rows.sort(function(a, b){
      const av = (a[state.sortKey] || "").toLowerCase();
      const bv = (b[state.sortKey] || "").toLowerCase();
      if (av < bv) return -1 * state.sortDir;
      if (av > bv) return 1 * state.sortDir;
      return 0;
    });
  }
  return rows;
}
function render(){
  const rows = applyFilters();
  const tb = document.getElementById("tbody");
  tb.innerHTML = "";
  if (rows.length === 0){
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.className = "empty"; td.colSpan = 4;
    td.textContent = "No changes match your filters.";
    tr.appendChild(td); tb.appendChild(tr);
  } else {
    rows.forEach(function(r){
      const tr = document.createElement("tr");
      const td1 = document.createElement("td");
      td1.textContent = r.Setting;
      const td2 = document.createElement("td");
      const span = document.createElement("span");
      span.className = "badge " + changeClass(r.Change);
      span.textContent = r.Change;
      td2.appendChild(span);
      const td3 = document.createElement("td");
      td3.textContent = r.Profile || "";
      const td4 = document.createElement("td");
      td4.textContent = r.IntunePolicy || "";
      tr.appendChild(td1); tr.appendChild(td2); tr.appendChild(td3); tr.appendChild(td4);
      tb.appendChild(tr);
    });
  }
  document.getElementById("count").textContent =
    "Showing " + rows.length + " of " + DATA.length + " changes";
  window.__view = rows;
}
function buildStats(){
  const counts = {};
  DATA.forEach(function(r){ counts[r.Change] = (counts[r.Change] || 0) + 1; });
  const el = document.getElementById("stats");
  const total = document.createElement("span");
  total.className = "stat";
  total.innerHTML = "<b>" + DATA.length + "</b> total";
  el.appendChild(total);
  Object.keys(counts).sort().forEach(function(k){
    const s = document.createElement("span");
    s.className = "stat " + changeClass(k);
    s.innerHTML = "<b>" + counts[k] + "</b> " + k;
    el.appendChild(s);
  });
}
function buildFilters(){
  const wrap = document.getElementById("filters");
  distinctTypes().forEach(function(t){
    const b = document.createElement("button");
    b.type = "button";
    b.className = "chip" + (t === state.type ? " active" : "");
    const n = (t === "All") ? DATA.length : DATA.filter(function(r){ return r.Change === t; }).length;
    b.textContent = t + " (" + n + ")";
    b.onclick = function(){
      state.type = t;
      Array.prototype.forEach.call(document.querySelectorAll(".chip"), function(c){ c.classList.remove("active"); });
      b.classList.add("active");
      render();
    };
    wrap.appendChild(b);
  });
}
function csvEscape(v){
  v = (v == null) ? "" : String(v);
  if (/[",\n]/.test(v)) v = '"' + v.replace(/"/g, '""') + '"';
  return v;
}
function downloadCsv(){
  const rows = window.__view || DATA;
  let out = "Setting,Change,Profile,Intune Policy\r\n";
  rows.forEach(function(r){ out += csvEscape(r.Setting) + "," + csvEscape(r.Change) + "," + csvEscape(r.Profile || "") + "," + csvEscape(r.IntunePolicy || "") + "\r\n"; });
  const blob = new Blob([out], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "ChangeHistory_v" + VERSION + "_filtered.csv";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}
document.getElementById("search").addEventListener("input", function(e){ state.q = e.target.value; render(); });
Array.prototype.forEach.call(document.querySelectorAll("th.sortable"), function(th){
  th.addEventListener("click", function(){
    const k = th.getAttribute("data-key");
    if (state.sortKey === k){ state.sortDir *= -1; } else { state.sortKey = k; state.sortDir = 1; }
    Array.prototype.forEach.call(document.querySelectorAll(".arrow"), function(a){ a.innerHTML = ""; });
    th.querySelector(".arrow").innerHTML = state.sortDir === 1 ? "&#9650;" : "&#9660;";
    render();
  });
});
document.getElementById("downloadCsv").addEventListener("click", downloadCsv);
buildStats(); buildFilters(); render();
</script>
</body>
</html>
'@

function Get-PdfText {
    param([string]$Path, [string]$PdftotextExe, [int]$First = 0, [int]$Last = 0)

    # 1) Poppler pdftotext (preferred - best text fidelity)
    $exe = $null
    if ($PdftotextExe -and (Test-Path -LiteralPath $PdftotextExe)) {
        $exe = (Resolve-Path -LiteralPath $PdftotextExe).Path
    } else {
        $cmd = Get-Command 'pdftotext' -ErrorAction SilentlyContinue
        if ($cmd) { $exe = $cmd.Source }
    }
    if ($exe) {
        $txt = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.IO.Path]::GetRandomFileName() + '.txt'))
        try {
            $pdfToTextArgs = @('-layout', '-enc', 'UTF-8')
            if ($First -gt 0) { $pdfToTextArgs += '-f'; $pdfToTextArgs += $First.ToString() }
            if ($Last  -gt 0) { $pdfToTextArgs += '-l'; $pdfToTextArgs += $Last.ToString() }
            $pdfToTextArgs += '--'; $pdfToTextArgs += $Path; $pdfToTextArgs += $txt
            & $exe @pdfToTextArgs 2>$null | Out-Null
            if (Test-Path -LiteralPath $txt) {
                $content = Get-Content -LiteralPath $txt -Raw -Encoding UTF8
                if ($content -and $content.Trim().Length -gt 0) {
                    Write-Verbose "Extracted text using pdftotext: $exe"
                    return $content
                }
            }
        } finally {
            if (Test-Path -LiteralPath $txt) { Remove-Item -LiteralPath $txt -Force -ErrorAction SilentlyContinue }
        }
    }

    # 2) Python with PyMuPDF (pip install pymupdf)
    $python = $null
    foreach ($candidate in @('python', 'python3')) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($cmd) { $python = $cmd.Source; break }
    }
    if ($python) {
        $txt = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.IO.Path]::GetRandomFileName() + '.txt'))
        $pyFirst = if ($First -gt 0) { $First - 1 } else { 0 }           # pymupdf is 0-indexed
        $pyLast  = if ($Last  -gt 0) { $Last } else { -1 }               # -1 = all pages
        $pyScript = @"
import sys
try:
    import fitz
except ImportError:
    sys.exit(99)
doc = fitz.open(sys.argv[1])
first = int(sys.argv[2])
last  = int(sys.argv[3])
out   = sys.argv[4]
end = doc.page_count if last < 0 else min(last, doc.page_count)
text_parts = []
for i in range(first, end):
    text_parts.append(doc[i].get_text())
with open(out, 'w', encoding='utf-8') as f:
    f.write('\n'.join(text_parts))
"@
        $pyFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.IO.Path]::GetRandomFileName() + '.py'))
        try {
            Set-Content -LiteralPath $pyFile -Value $pyScript -Encoding UTF8
            & $python $pyFile $Path $pyFirst.ToString() $pyLast.ToString() $txt 2>$null
            $pyExit = $LASTEXITCODE
            if ($pyExit -eq 99) {
                Write-Verbose "Python found but pymupdf not installed."
            } elseif ((Test-Path -LiteralPath $txt)) {
                $content = Get-Content -LiteralPath $txt -Raw -Encoding UTF8
                if ($content -and $content.Trim().Length -gt 0) {
                    Write-Verbose "Extracted text using Python (PyMuPDF): $python"
                    return $content
                }
            }
        } finally {
            if (Test-Path -LiteralPath $pyFile) { Remove-Item -LiteralPath $pyFile -Force -ErrorAction SilentlyContinue }
            if (Test-Path -LiteralPath $txt)    { Remove-Item -LiteralPath $txt -Force -ErrorAction SilentlyContinue }
        }
    }

    throw ("Could not extract text from the PDF. Install one of the following:`n" +
           "  1) Poppler 'pdftotext' on your PATH (or pass -PdftotextPath)`n" +
           "  2) Python with PyMuPDF: pip install pymupdf`n" +
           "Both are free and handle large PDFs reliably.")
}

function Get-ChangeHistorySection {
    param([string]$Text)
    # Case-insensitive search for the change history section marker
    $marker = 'Appendix: Change History'
    $idx = $Text.LastIndexOf($marker, [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -lt 0) {
        # Try looser pattern: just "Change History" as a standalone heading
        $looseMatch = [regex]::Match($Text, '(?im)^\s*(?:Appendix\s*[A-Z]?:?\s*)?Change\s+History\s*$')
        if ($looseMatch.Success) {
            $idx = $looseMatch.Index
        }
    }
  if ($idx -lt 0) {
    # Some PDFs only label this section as "Changes for this version".
    $fallbackMatch = [regex]::Match($Text, '(?im)^\s*Changes\s+for\s+this\s+version\s*$')
    if ($fallbackMatch.Success) {
      $idx = $fallbackMatch.Index
    }
  }
    if ($idx -lt 0) {
    throw "This PDF does not contain a recognizable change history section ('Appendix: Change History' or 'Changes for this version')."
    }

  $tail = $Text.Substring($idx)
  # Stop at the next appendix heading to avoid accidentally parsing subsequent appendices.
  $nextAppendix = [regex]::Match($tail, '(?im)^\s*Appendix:\s+(?!Change\s+History\b).+$')
  if ($nextAppendix.Success) {
    return $tail.Substring(0, $nextAppendix.Index)
  }
  return $tail
}

function Normalize-ChangeLine {
  param([string]$Line)

  if ($null -eq $Line) { return '' }

  $t = [string]$Line
  $t = $t -replace '[\u200B-\u200D\uFEFF]', ''
  $t = $t -replace '[\u2018\u2019\u201A\u201B\u2032\u02BC\uFF07]', "'"
  $t = $t -replace '[\u201C\u201D\u201E\u201F\u2033\uFF02]', '"'
  $t = $t -replace '\s+', ' '
  $t = $t.Trim()
  return $t
}

function Repair-RecommendationText {
    param([string]$Text)

    if (-not $Text) { return '' }

    $repaired = [string]$Text
    $repaired = [regex]::Replace(
        $repaired,
        "^(Ensure\s+'[^']+?)\s+is\s+set\s+to\s+'",
        '$1'' is set to ''',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    return $repaired
}

function Get-LevenshteinSimilarity {
  param(
    [string]$Left,
    [string]$Right
  )

  if ($null -eq $Left) { $Left = '' }
  if ($null -eq $Right) { $Right = '' }
  if ($Left -eq $Right) { return 1.0 }
  if ($Left.Length -eq 0 -or $Right.Length -eq 0) { return 0.0 }

  $n = $Left.Length
  $m = $Right.Length

  $previous = New-Object 'int[]' ($m + 1)
  $current  = New-Object 'int[]' ($m + 1)

  for ($j = 0; $j -le $m; $j++) {
    $previous[$j] = $j
  }

  for ($i = 1; $i -le $n; $i++) {
    $current[0] = $i
    $leftChar = $Left[$i - 1]
    for ($j = 1; $j -le $m; $j++) {
      $cost = if ($leftChar -eq $Right[$j - 1]) { 0 } else { 1 }
      $deletion = $previous[$j] + 1
      $insertion = $current[$j - 1] + 1
      $substitute = $previous[$j - 1] + $cost
      $current[$j] = [Math]::Min([Math]::Min($deletion, $insertion), $substitute)
    }

    $tmp = $previous
    $previous = $current
    $current = $tmp
  }

  $distance = $previous[$m]
  $maxLen = [Math]::Max($n, $m)
  return 1.0 - ($distance / [double]$maxLen)
}

  function Get-PlainText {
    param([string]$Cell)
    if ($null -eq $Cell) { return '' }
    $t = [string]$Cell
    $t = $t.Trim()
    $t = [regex]::Replace($t, '\*\*([^*]*)\*\*', '$1')
    $t = [regex]::Replace($t, '\[([^\]]*)\]\([^)]*\)', '$1')
    $t = $t -replace '<br\s*/?>', ' '
    return (Normalize-ChangeLine $t)
  }

  function Split-PipeRow {
    param([string]$Line)
    $s = $Line.Trim()
    if ($s.StartsWith('|')) { $s = $s.Substring(1) }
    if ($s.EndsWith('|'))   { $s = $s.Substring(0, $s.Length - 1) }
    return @($s -split '\|' | ForEach-Object { $_.Trim() })
  }

  function Normalize-RecommendationKey {
    param([string]$Text)

    $normalized = Get-PlainText $Text
    $normalized = $normalized -replace '\s+TO\s+.+$', ''
    $normalized = $normalized -replace '\s+FROM\s+.+$', ''
    $normalized = $normalized -replace [char]0x00E0, [char]0x2026
    $normalized = $normalized -replace '\.\.\.', [char]0x2026
    $normalized = $normalized.ToLowerInvariant().Trim()
    return $normalized
  }

  function Get-SettingNameKey {
    param([string]$Text)

    $plain = Get-PlainText $Text
    $match = [regex]::Match($plain, "Ensure\s+'(?<name>[^']+)'", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) {
      $match = [regex]::Match(
        $plain,
        "Ensure\s+'(?<name>.+?)\s+is\s+set\s+to\s+'",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    if (-not $match.Success) {
      $match = [regex]::Match(
        $plain,
        "Ensure\s+'(?<name>.+?)\s+TO\s+",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    if (-not $match.Success) { return '' }

    return $match.Groups['name'].Value.Trim().ToLowerInvariant()
  }

  function Try-GuessProductName {
    param([string]$FileName)

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $base = $base -replace '^CIS_', '' -replace '_Benchmark.*$', '' -replace '_GPO', ''
    if ($base -match 'Intune_for_(.+)') { $base = "Microsoft $($Matches[1])" }
    return ($base -replace '_', ' ').Trim()
  }

  function Get-DefaultIntunePrefix {
    param([string]$PdfFileName)

    $productName = Try-GuessProductName -FileName $PdfFileName
    if (-not $productName) { $productName = [System.IO.Path]::GetFileNameWithoutExtension($PdfFileName) }
    $cleanName = $productName -replace '\s', ''
    return "Prod-Win-All-SecurityBaseline-$cleanName-"
  }

  function Get-IntunePolicySuffix {
    param([string]$SectionName)

    if (-not $SectionName) { return 'Other' }
    $suffix = $SectionName -replace '\([^)]*\)', '' -replace '\s+', ' '
    $suffix = $suffix -replace '[^A-Za-z0-9\s]', '' -replace '\s+', ' '
    $suffix = ($suffix.Trim().Split(' ') | ForEach-Object {
      if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) }
    }) -join ''
    if (-not $suffix) { return 'Other' }
    return $suffix
  }

  function Get-PolicyNameFromGroup {
    param([string]$PolicyGroup, [string]$IntunePrefix, [string]$Fallback)

    if ($PolicyGroup) {
      $pg = $PolicyGroup -replace '[^A-Za-z0-9\s]','' -replace '\s+',' '
      $pgPascal = ($pg.Trim().Split(' ') | ForEach-Object {
        if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) }
      }) -join ''
      if ($pgPascal) { return "$IntunePrefix$pgPascal" }
    }

    return $Fallback
  }

  function Read-PreviousMarkdownMetadata {
    param([string]$Path)

    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $lines = $content -split "`r?`n"
    $items = [System.Collections.Generic.List[object]]::new()

    foreach ($line in $lines) {
      if (-not $line.TrimStart().StartsWith('|')) { continue }

      $cells = @(Split-PipeRow $line)
      if ($cells.Count -lt 10) { continue }
      if ($cells[0] -match '^\*\*' -or (Get-PlainText $cells[0]) -ieq 'Recommendation' -or $cells[0] -match '^[\-:]+$') { continue }

      $recommendation = Get-PlainText $cells[0]
      $twodayId = Get-PlainText $cells[1]
      $cisReference = Get-PlainText $cells[2]
      if ((-not $cisReference -or $cisReference -eq '-') -and $twodayId -and $twodayId -ne '-') {
        $cisReference = $twodayId
      }

      $profile = Get-PlainText $cells[3]
      $intunePolicy = Get-PlainText $cells[9]

      if (-not $recommendation) { continue }

      $items.Add([pscustomobject]@{
        Recommendation = $recommendation
        CISReference   = $cisReference
        Profile        = $profile
        IntunePolicy   = $intunePolicy
        NormalizedKey  = (Normalize-RecommendationKey $recommendation)
        SettingNameKey = (Get-SettingNameKey $recommendation)
      })
    }

    return $items.ToArray()
  }

  function Get-PythonExecutable {
    foreach ($candidate in @('python', 'python3')) {
      $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
      if ($cmd -and $cmd.Source) { return $cmd.Source }
    }

    return $null
  }

  function Get-PdfRecommendationMetadata {
    param([string]$Path)

    $pythonExe = Get-PythonExecutable
    if (-not $pythonExe) {
      throw 'Python is required to resolve Add metadata from the current PDF.'
    }

    $pythonScript = @'
  import json
  import re
  import sys
  import unicodedata

  try:
    import fitz
  except ImportError:
    sys.stderr.write('PyMuPDF (fitz) is required.\n')
    sys.exit(99)

  pdf_path = sys.argv[1]
  reader = fitz.open(pdf_path)
  total_pages = reader.page_count

  summary_start = None
  summary_end = None

  for i in range(total_pages - 1, -1, -1):
    text = reader[i].get_text() or ''
    if 'Appendix: Summary Table' in text and 'CIS Benchmark Recommendation' in text:
      summary_start = i
      break

  if summary_start is None:
    for i in range(total_pages - 1, -1, -1):
      text = reader[i].get_text() or ''
      if 'Summary Table' in text and 'CIS Benchmark Recommendation' in text:
        summary_start = i
        break

  if summary_start is None:
    sys.stderr.write('Could not find Summary Table in PDF.\n')
    sys.exit(1)

  for i in range(summary_start + 1, total_pages):
    text = reader[i].get_text() or ''
    if 'Appendix: CIS Controls' in text or 'Appendix: Change History' in text:
      summary_end = i + 1
      break

  if summary_end is None:
    summary_end = min(summary_start + 20, total_pages)

  summary_text = ''
  for i in range(summary_start, summary_end):
    text = reader[i].get_text()
    if text:
      summary_text += text + '\n'

  rec_intro_pattern = r'(?:Ensure|Enable|Disable|Require|Block|Configure|Set)\b'
  profile_token_pattern = r'(?:E[35]\s+Level\s*[12]|E[35]\s+L[12]|L[12]|BL)'

  lines = summary_text.split('\n')
  merged = []
  j = 0
  while j < len(lines):
    s = lines[j].strip()
    nxt = lines[j + 1].strip() if j + 1 < len(lines) else ''
    if re.match(r'^\d+(?:\.\d+)*\.$', s) and re.match(r'^\d', nxt):
      merged.append(s + nxt)
      j += 2
      continue
    if re.match(r'^\d+(?:\.\d+)*$', s) and re.match(r'^\.\d', nxt):
      merged.append(s + nxt)
      j += 2
      continue
    if re.match(r'^\d+(?:\.\d+)*\s+\((?:E[35]\s+Level\s*[12]|E[35]\s+L[12]|L[12]|BL)\)$', s, re.IGNORECASE) and nxt:
      merged.append(s + ' ' + nxt)
      j += 2
      continue
    if re.match(r'^\d+(?:\.\d+)*\s+(?:E[35]\s+Level\s*[12]|E[35]\s+L[12]|L[12]|BL)$', s, re.IGNORECASE) and nxt:
      merged.append(s + ' ' + nxt)
      j += 2
      continue
    if re.match(r'^\d+(?:\.\d+)*$', s) and nxt and re.match(rec_intro_pattern, nxt, re.IGNORECASE):
      merged.append(s + ' ' + nxt)
      j += 2
      continue
    if re.match(r'^\d+(?:\.\d+)*$', s) and nxt and re.match(r'^[A-Z].+', nxt) and not re.match(r'^(Yes|No|Set|Correctly)$', nxt, re.IGNORECASE):
      merged.append(s + ' ' + nxt)
      j += 2
      continue
    merged.append(lines[j])
    j += 1
  lines = merged

  recommendations = []
  section_headers = {}
  ordered_items = []

  def clean_text(t):
    t = unicodedata.normalize('NFC', t)
    t = re.sub(r'[\uf06f\uf0a8\uf0fe]', '', t)
    t = re.sub(r"[\xb4\x60\u2018\u2019\u201a\u201b\u2032\u02b9\u02bb\u02bc\u02c8\u02ca\u02cb\ua78c\uff07]", "'", t)
    t = t.replace('\u2524', "'")
    t = re.sub('[\u201c\u201d\u201e\u201f\u2033\u2036\u275d\u275e\uff02]', '"', t)
    t = t.replace('\u00f4', '"').replace('\u00f6', '"')
    t = t.replace('\u00e0', '\u2026')
    t = re.sub(r'\s+', ' ', t).strip()
    return t

  def is_skip_line(line):
    return (line.startswith('Page ') or
        line.startswith('CIS Benchmark Recommendation') or
        line in ('Yes No', 'Set', 'Correctly', ''))

  def normalize_profile(profile_text):
    if profile_text is None:
      return ''
    value = re.sub(r'\s+', ' ', str(profile_text)).strip()
    if not value:
      return ''
    if re.fullmatch(r'BL', value, re.IGNORECASE):
      return 'BL'
    m = re.search(r'\b(E[35])\s*(?:Level|L)\s*([12])\b', value, re.IGNORECASE)
    if m:
      return f"{m.group(1).upper()} L{m.group(2)}"
    m = re.search(r'\bL([12])\b', value, re.IGNORECASE)
    if m:
      return f"L{m.group(1)}"
    m = re.search(r'\bLevel\s*([12])\b', value, re.IGNORECASE)
    if m:
      return f"L{m.group(1)}"
    return value

  i = 0
  while i < len(lines):
    line = lines[i].strip()
    if is_skip_line(line):
      i += 1
      continue
    if line.startswith('Appendix:') and 'Summary Table' not in line:
      break

    rec_match = re.match(r'^(\d+(?:\.\d+)*)\s+\((' + profile_token_pattern + r')\)\s+(.+)', line, re.IGNORECASE)
    if rec_match:
      ref = rec_match.group(1)
      profile = normalize_profile(rec_match.group(2))
      rest = rec_match.group(3)
      while i + 1 < len(lines):
        next_line = lines[i + 1].strip()
        if re.match(r'^(\d+(?:\.\d+)*)\s+\(' + profile_token_pattern + r'\)', next_line, re.IGNORECASE): break
        if re.match(r'^(\d+(?:\.\d+)*)\s+' + profile_token_pattern + r'\s+', next_line, re.IGNORECASE): break
        if re.match(r'^(\d+(?:\.\d+)*)\s+' + rec_intro_pattern, next_line): break
        hdr_check = re.match(r'^(\d+(?:\.\d+)*)\s+([A-Z])', next_line)
        if hdr_check and '(L' not in next_line and not re.search(rec_intro_pattern, next_line): break
        if is_skip_line(next_line) or next_line.startswith('CIS Benchmark'):
          i += 1
          continue
        if next_line.startswith('Appendix:'): break
        rest += ' ' + next_line
        i += 1
      type_match = re.search(r'\((Automated|Manual)\)', rest)
      rec_text = rest[:type_match.start()].strip() if type_match else rest
      recommendations.append({'ref': ref, 'profile': profile, 'text': clean_text(rec_text)})
      ordered_items.append({'type': 'rec', 'ref': ref})
      i += 1
      continue

    rec_match2 = re.match(r'^(\d+(?:\.\d+)*)\s+(' + profile_token_pattern + r')\s+(.+)', line, re.IGNORECASE)
    if rec_match2:
      ref = rec_match2.group(1)
      profile = normalize_profile(rec_match2.group(2))
      rest = rec_match2.group(3)
      while i + 1 < len(lines):
        next_line = lines[i + 1].strip()
        if re.match(r'^(\d+(?:\.\d+)*)\s+\(' + profile_token_pattern + r'\)', next_line, re.IGNORECASE): break
        if re.match(r'^(\d+(?:\.\d+)*)\s+' + profile_token_pattern + r'\s+', next_line, re.IGNORECASE): break
        if re.match(r'^(\d+(?:\.\d+)*)\s+' + rec_intro_pattern, next_line): break
        hdr_check = re.match(r'^(\d+(?:\.\d+)*)\s+([A-Z])', next_line)
        if hdr_check and '(L' not in next_line and not re.search(rec_intro_pattern, next_line): break
        if is_skip_line(next_line) or next_line.startswith('CIS Benchmark'):
          i += 1
          continue
        if next_line.startswith('Appendix:'): break
        rest += ' ' + next_line
        i += 1
      type_match = re.search(r'\((Automated|Manual)\)', rest)
      rec_text = rest[:type_match.start()].strip() if type_match else rest
      recommendations.append({'ref': ref, 'profile': profile, 'text': clean_text(rec_text)})
      ordered_items.append({'type': 'rec', 'ref': ref})
      i += 1
      continue

    rec_match3 = re.match(r'^(\d+(?:\.\d+)*)\s+(' + rec_intro_pattern + r'\s.+)', line)
    if rec_match3:
      ref = rec_match3.group(1)
      rest = rec_match3.group(2)
      while i + 1 < len(lines):
        next_line = lines[i + 1].strip()
        if re.match(r'^(\d+(?:\.\d+)*)\s+', next_line): break
        if is_skip_line(next_line) or next_line.startswith('Appendix:'): break
        if next_line.startswith('CIS Benchmark'):
          i += 1
          continue
        rest += ' ' + next_line
        i += 1
      type_match = re.search(r'\((Automated|Manual)\)', rest)
      rec_text = rest[:type_match.start()].strip() if type_match else rest.strip()
      recommendations.append({'ref': ref, 'profile': 'L1', 'text': clean_text(rec_text)})
      ordered_items.append({'type': 'rec', 'ref': ref})
      i += 1
      continue

    hdr_match = re.match(r'^(\d+(?:\.\d+)*)\s+([A-Z].+?)$', line)
    if hdr_match and not re.search(r'\(' + profile_token_pattern + r'\)', line, re.IGNORECASE) and not re.search(rec_intro_pattern, line):
      hdr_num = hdr_match.group(1)
      hdr_name = clean_text(hdr_match.group(2).strip())
      section_headers[hdr_num] = hdr_name
      ordered_items.append({'type': 'header', 'num': hdr_num, 'name': hdr_name})
      i += 1
      continue

    i += 1

  def is_ancestor(header_num, rec_ref):
    return rec_ref.startswith(header_num + '.')

  rec_to_header = {}
  header_stack = []

  for item in ordered_items:
    if item['type'] == 'header':
      num = item['num']
      depth = num.count('.')
      header_stack = [(n, nm) for n, nm in header_stack if n.count('.') < depth]
      header_stack.append((num, item['name']))
    elif item['type'] == 'rec':
      ref = item['ref']
      best = None
      for hdr_num, hdr_name in reversed(header_stack):
        if is_ancestor(hdr_num, ref):
          best = hdr_num
          break
      if best is None and header_stack:
        best = header_stack[-1][0]
      rec_to_header[ref] = best

  detail_texts = []
  for page_idx in range(0, summary_start):
    detail_texts.append(reader[page_idx].get_text() or '')

  rec_profile_levels = {}
  rec_policy_groups = {}
  rec_texts_from_detail = {}

  def extract_policy_group(combined_text):
    rem_match = re.search(r'\bRemediation\s*:', combined_text, re.IGNORECASE)
    if not rem_match:
      return None
    rem_text = combined_text[rem_match.start():]
    for stop in ['References:', 'CIS Controls:', 'Additional Information:', 'Audit:']:
      idx = rem_text.find(stop)
      if 0 < idx < 2000:
        rem_text = rem_text[:idx]
    path_parts = []
    in_path = False
    for line in rem_text.split('\n'):
      stripped = line.strip()
      if not stripped:
        if in_path and path_parts:
          break
        continue
      cfg_match = re.search(r'(?:Computer|User)\s+Configuration', stripped, re.IGNORECASE)
      if cfg_match:
        in_path = True
        path_str = stripped[cfg_match.start():]
        path_parts = [p.strip() for p in re.split(r'\s*\\\s*', path_str) if p.strip()]
        continue
      if in_path:
        if re.match(r'^(?:Note|NOTE|Rationale|Audit|References|CIS|Default|Impact|Additional|Group Policy|Registry Path)', stripped):
          break
        if re.match(r'^\d+\.\d+', stripped):
          break
        if '\\' in stripped:
          sub_parts = [p.strip() for p in re.split(r'\s*\\\s*', stripped) if p.strip()]
          if sub_parts and path_parts:
            path_parts[-1] = path_parts[-1] + ' ' + sub_parts[0]
            path_parts.extend(sub_parts[1:])
        else:
          break
    path_parts = [p for p in path_parts if p and len(p) > 2]
    if len(path_parts) >= 2:
      return path_parts[-2]
    if len(path_parts) == 1:
      return path_parts[0]
    return None

  def extract_profile_level_from_applicability(combined_text):
    normalized = re.sub(r'\s+', ' ', combined_text)
    m = re.search(
      r'Profile\s+Applicability\s*:?\s*(.*?)(?:Description\s*:|Rationale\s*:|Audit\s*:|Remediation\s*:|Default\s+Value\s*:|Impact\s*:|References\s*:|$)',
      normalized,
      re.IGNORECASE
    )
    if not m:
      return None
    block = m.group(1)
    block = block.replace('\u2022', ' ').replace('\uf0b7', ' ')
    explicit_profiles = set()
    for match in re.finditer(r'\b(E[35])\s*(?:Level|L)\s*([12])\b', block, re.IGNORECASE):
      explicit_profiles.add(f"{match.group(1).upper()} L{match.group(2)}")
    if len(explicit_profiles) == 1:
      return next(iter(explicit_profiles))
    if len(explicit_profiles) > 1:
      return None
    has_l1 = re.search(r'(?:E[35]\s*)?Level\s*1', block, re.IGNORECASE) is not None
    has_l2 = re.search(r'(?:E[35]\s*)?Level\s*2', block, re.IGNORECASE) is not None
    if has_l1 and not has_l2:
      return 'L1'
    if has_l2 and not has_l1:
      return 'L2'
    return None

  for page_idx in range(len(detail_texts)):
    text = detail_texts[page_idx]
    if not text:
      continue
    page_recs = re.findall(r'(\d+(?:\.\d+)*)\s+\(' + profile_token_pattern + r'\)\s+' + rec_intro_pattern, text, re.IGNORECASE)
    if not page_recs:
      page_recs = re.findall(r'(\d+(?:\.\d+)*)\s+' + rec_intro_pattern, text)
    if not page_recs:
      continue
    combined_text = text
    for lookahead in range(1, 3):
      if page_idx + lookahead < len(detail_texts):
        next_text = detail_texts[page_idx + lookahead]
        next_rec = re.search(r'\d+(?:\.\d+)*\s+\(' + profile_token_pattern + r'\)\s+' + rec_intro_pattern, next_text, re.IGNORECASE)
        if next_rec:
          before_new_rec = next_text[:next_rec.start()]
          if before_new_rec:
            combined_text += '\n' + before_new_rec
          break
        combined_text += '\n' + next_text

    for ref in page_recs:
      if ref not in rec_texts_from_detail:
        m_detail = re.search(
          r'\b' + re.escape(ref) + r'\s+\(' + profile_token_pattern + r'\)\s+((?:Ensure|Enable|Disable|Require|Block|Configure|Set)\s.*?)\s*\((Automated|Manual)\)',
          text, re.DOTALL
        )
        if not m_detail:
          m_detail = re.search(
            r'\b' + re.escape(ref) + r'\s+((?:Ensure|Enable|Disable|Require|Block|Configure|Set)\s.*?)\s*\((Automated|Manual)\)',
            text, re.DOTALL
          )
        if m_detail:
          raw = re.sub(r'\s+', ' ', m_detail.group(1)).strip()
          rec_texts_from_detail[ref] = clean_text(raw)

    best_rec = page_recs[-1]
    if best_rec not in rec_profile_levels:
      derived_profile = extract_profile_level_from_applicability(combined_text)
      if derived_profile:
        rec_profile_levels[best_rec] = derived_profile
    if best_rec not in rec_policy_groups:
      policy_group = extract_policy_group(combined_text)
      if policy_group:
        rec_policy_groups[best_rec] = policy_group

  for rec in recommendations:
    rec['parent_header'] = rec_to_header.get(rec['ref'])
    derived_profile = rec_profile_levels.get(rec['ref'])
    if derived_profile:
      rec['profile'] = derived_profile
    if rec['ref'] in rec_policy_groups:
      rec['policy_group'] = rec_policy_groups[rec['ref']]
    else:
      rec['policy_group'] = None
    if rec['ref'] in rec_texts_from_detail:
      rec['text'] = rec_texts_from_detail[rec['ref']]

  print(json.dumps({'recommendations': recommendations, 'section_headers': section_headers}, ensure_ascii=True))
'@

    $pythonScript = (($pythonScript -split "`r?`n") | ForEach-Object {
      if ($_ -match '^\s{2,}') { $_ -replace '^\s{2}', '' } else { $_ }
    }) -join "`n"

    $pyFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ([System.IO.Path]::GetRandomFileName() + '.py'))
    try {
      Set-Content -LiteralPath $pyFile -Value $pythonScript -Encoding UTF8
      $output = & $pythonExe $pyFile $Path 2>&1
      $exitCode = $LASTEXITCODE
      if ($exitCode -eq 99) {
        throw 'PyMuPDF is required to resolve Add metadata from the current PDF.'
      }
      if ($exitCode -ne 0) {
        throw ((@($output) -join [Environment]::NewLine).Trim())
      }

      $parsed = (@($output) -join [Environment]::NewLine) | ConvertFrom-Json
      $sectionHeaders = @{}
      foreach ($prop in $parsed.section_headers.PSObject.Properties) {
        $sectionHeaders[$prop.Name] = [string]$prop.Value
      }

      $intunePrefix = Get-DefaultIntunePrefix -PdfFileName (Split-Path -Path $Path -Leaf)
      $metadata = [System.Collections.Generic.List[object]]::new()
      foreach ($rec in $parsed.recommendations) {
        $sectionKey = [string]$rec.parent_header
        $sectionName = if ($sectionKey -and $sectionHeaders.ContainsKey($sectionKey)) { $sectionHeaders[$sectionKey] } else { 'Other' }
        $fallbackPolicy = if ($sectionKey) { "$intunePrefix$(Get-IntunePolicySuffix -SectionName $sectionName)" } else { "${intunePrefix}Other" }
        $policyName = Get-PolicyNameFromGroup -PolicyGroup ([string]$rec.policy_group) -IntunePrefix $intunePrefix -Fallback $fallbackPolicy
        $recommendation = Get-PlainText ([string]$rec.text)
        $metadata.Add([pscustomobject]@{
          Recommendation = $recommendation
          CISReference   = [string]$rec.ref
          Profile        = Get-PlainText ([string]$rec.profile)
          IntunePolicy   = $policyName
          NormalizedKey  = (Normalize-RecommendationKey $recommendation)
          SettingNameKey = (Get-SettingNameKey $recommendation)
        })
      }

      return $metadata.ToArray()
    } finally {
      if (Test-Path -LiteralPath $pyFile) { Remove-Item -LiteralPath $pyFile -Force -ErrorAction SilentlyContinue }
    }
  }

  function New-MetadataLookup {
    param([object[]]$Items)

    $byReference = @{}
    $byKey = @{}
    $bySettingName = @{}
    foreach ($item in @($Items)) {
      if ($item.CISReference -and -not $byReference.ContainsKey($item.CISReference)) {
        $byReference[$item.CISReference] = $item
      }
      if ($item.NormalizedKey -and -not $byKey.ContainsKey($item.NormalizedKey)) {
        $byKey[$item.NormalizedKey] = $item
      }
      if ($item.SettingNameKey -and -not $bySettingName.ContainsKey($item.SettingNameKey)) {
        $bySettingName[$item.SettingNameKey] = $item
      }
    }

    return [pscustomobject]@{
      ByReference = $byReference
      ByKey       = $byKey
      BySettingName = $bySettingName
      Items = @($Items)
    }
  }

  function Find-MetadataMatch {
    param(
      $Row,
      $Lookup,
      [double]$MinSimilarity = 0.90
    )

    if (-not $Lookup) { return $null }
    if ($Row.CISReference -and $Lookup.ByReference.ContainsKey($Row.CISReference)) {
      return $Lookup.ByReference[$Row.CISReference]
    }
    if ($Row.NormalizedKey -and $Lookup.ByKey.ContainsKey($Row.NormalizedKey)) {
      return $Lookup.ByKey[$Row.NormalizedKey]
    }
    if ($Row.SettingNameKey -and $Lookup.BySettingName.ContainsKey($Row.SettingNameKey)) {
      return $Lookup.BySettingName[$Row.SettingNameKey]
    }

    $rowKey = [string]$Row.NormalizedKey
    if (-not $rowKey) { return $null }

    $candidates = @($Lookup.Items)
    if ($Row.SettingNameKey) {
      $sameNameCandidates = @($candidates | Where-Object { $_.SettingNameKey -and $_.SettingNameKey -eq $Row.SettingNameKey })
      if ($sameNameCandidates.Count -gt 0) {
        $candidates = $sameNameCandidates
      }
    }

    $best = $null
    $bestScore = 0.0
    foreach ($candidate in $candidates) {
      if (-not $candidate.NormalizedKey) { continue }
      $score = Get-LevenshteinSimilarity -Left $rowKey -Right ([string]$candidate.NormalizedKey)
      if ($score -gt $bestScore) {
        $bestScore = $score
        $best = $candidate
      }
    }

    if ($best -and $bestScore -ge $MinSimilarity) {
      return $best
    }

    return $null
  }

function Repair-WrappedChangeLines {
  param([string[]]$Lines)

  $fixed = [System.Collections.Generic.List[string]]::new()
  for ($i = 0; $i -lt $Lines.Count; $i++) {
    $line = Normalize-ChangeLine $Lines[$i]
    if (-not $line) { continue }

    # Merge lines where the action and the content are split across lines.
    if ($line -match '^(?i)((?:ADD|REMOVE|UPDATE|RENAME|MODIFY|DELETE|DEPRECATE|NEW|MOVE)(?:\s*[,&]\s*(?:ADD|REMOVE|UPDATE|RENAME|MODIFY|DELETE|DEPRECATE|NEW|MOVE))*(?:\s+(?:&|AND)\s+(?:ADD|REMOVE|UPDATE|RENAME|MODIFY|DELETE|DEPRECATE|NEW|MOVE))*?)\s*-\s*$' -and ($i + 1) -lt $Lines.Count) {
      $next = Normalize-ChangeLine $Lines[$i + 1]
      if ($next) {
        $fixed.Add("$line $next")
        $i++
        continue
      }
    }

    $fixed.Add($line)
  }

  return ,$fixed.ToArray()
}

function Get-ChangeData {
    param([string]$Section)

    $ti = (Get-Culture).TextInfo
  $dateInlineRe  = [regex]'(?i)\bDate\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})\b.*?\bVersion\s*:?\s*([0-9]+(?:\.[0-9]+)*)\b'
  $dateOnlyRe    = [regex]'(?i)\bDate\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})\b'
  $versionOnlyRe = [regex]'(?i)\bVersion\s*:?\s*([0-9]+(?:\.[0-9]+)*)\b'

  $actionRe = [regex]'^(?i)[^\p{L}]*((?:ADD|REMOVE|UPDATE|RENAME|MODIFY|DELETE|DEPRECATE|NEW|MOVE)(?:\s*[,&]\s*(?:ADD|REMOVE|UPDATE|RENAME|MODIFY|DELETE|DEPRECATE|NEW|MOVE))*)\s*-\s*(.+)$'
  $ticketRe = [regex]'^(?i)Ticket\s*#?\s*\d+'
  $pageRe = [regex]'^(?i)Page\s+\d+(?:\s+of\s+\d+)?$'

  $diag = [ordered]@{
    LinesScanned   = 0
    VersionHeaders = 0
    ChangesParsed  = 0
    TicketLines    = 0
    UnmatchedLines = 0
    UnmatchedSample = [System.Collections.Generic.List[string]]::new()
  }

  $lines = $Section -split "`r?`n"
  $lines = Repair-WrappedChangeLines -Lines $lines

  $blocks = [System.Collections.Generic.List[object]]::new()
  $currentBlock = $null
  $lastRow = $null
  $pendingDate = $null
  $foundFallbackHeading = $false

  foreach ($rawLine in $lines) {
    $diag.LinesScanned++
    $line = Normalize-ChangeLine $rawLine

    if (-not $line) { continue }
    if ($line -match '^(?i)Appendix:\s+') { continue }
    if ($line -match '^(?i)Change\s+History\b') { continue }
    if ($line -match '^(?i)Changes\s+for\s+this\s+version\b') {
      $foundFallbackHeading = $true
      continue
    }
    if ($pageRe.IsMatch($line)) { continue }

    $dm = $dateInlineRe.Match($line)
    if ($dm.Success) {
      if ($currentBlock) { $blocks.Add($currentBlock) }
      $diag.VersionHeaders++
      $currentBlock = [pscustomobject]@{
        Date    = $dm.Groups[1].Value
        Version = $dm.Groups[2].Value
        Rows    = [System.Collections.Generic.List[object]]::new()
      }
      $lastRow = $null
      $pendingDate = $null
      continue
    }

    $dd = $dateOnlyRe.Match($line)
    if ($dd.Success) {
      $pendingDate = $dd.Groups[1].Value
      continue
    }

    if ($pendingDate) {
      $vv = $versionOnlyRe.Match($line)
      if ($vv.Success) {
        if ($currentBlock) { $blocks.Add($currentBlock) }
        $diag.VersionHeaders++
        $currentBlock = [pscustomobject]@{
          Date    = $pendingDate
          Version = $vv.Groups[1].Value
          Rows    = [System.Collections.Generic.List[object]]::new()
        }
        $lastRow = $null
        $pendingDate = $null
        continue
      }
    }

    # Some benchmark PDFs only have "Changes for this version" and no Date/Version.
    # In that case, create one implicit block as soon as we encounter a change row.
    if (-not $currentBlock -and $foundFallbackHeading -and $actionRe.IsMatch($line)) {
      $diag.VersionHeaders++
      $currentBlock = [pscustomobject]@{
        Date    = 'Unknown Date'
        Version = 'Current'
        Rows    = [System.Collections.Generic.List[object]]::new()
      }
    }

    # Ignore payload until we identify a version header or fallback block.
    if (-not $currentBlock) { continue }

    if ($ticketRe.IsMatch($line)) {
      $diag.TicketLines++
      continue
    }

    $am = $actionRe.Match($line)
    if ($am.Success) {
      $changeRaw = $am.Groups[1].Value.Trim()
      $rest = $am.Groups[2].Value.Trim()

      $cisReference = ''
      $refMatch = [regex]::Match($rest, '^(?<ref>\d+(?:\.\d+)*)\s*(?:\([^)]*\))?\s*(?<text>.+)$')
      if ($refMatch.Success) {
        $cisReference = $refMatch.Groups['ref'].Value
        $rest = $refMatch.Groups['text'].Value
      }
      $rest = Normalize-ChangeLine $rest
      $rest = Repair-RecommendationText $rest

      $row = [pscustomobject]@{
        CISReference  = $cisReference
        Setting       = $rest
        Change        = $ti.ToTitleCase($changeRaw.ToLower())
        Profile       = ''
        IntunePolicy  = ''
        MatchSource   = ''
        NormalizedKey = (Normalize-RecommendationKey $rest)
        SettingNameKey = (Get-SettingNameKey $rest)
      }

      $currentBlock.Rows.Add($row)
      $lastRow = $row
      $diag.ChangesParsed++
      continue
    }

    # Treat non-structural lines as wrapped setting continuation for the last parsed row.
    if ($lastRow -and $line -notmatch '^(?i)(Date:\s*\d{1,2}/\d{1,2}/\d{4}|Version:\s*\d+)') {
      $lastRow.Setting = Normalize-ChangeLine ("$($lastRow.Setting) $line")
      $lastRow.Setting = Repair-RecommendationText $lastRow.Setting
      $lastRow.NormalizedKey = Normalize-RecommendationKey $lastRow.Setting
      $lastRow.SettingNameKey = Get-SettingNameKey $lastRow.Setting
      continue
    }

    $diag.UnmatchedLines++
    if ($diag.UnmatchedSample.Count -lt 5) {
      $diag.UnmatchedSample.Add($line)
    }
  }

  if ($currentBlock) { $blocks.Add($currentBlock) }

  if ($blocks.Count -eq 0) {
    throw "No change entries were found in the change history (supported layouts: Date/Version headers or 'Changes for this version')."
  }

  $blocksOut = @($blocks | ForEach-Object {
    [pscustomobject]@{
      Date    = $_.Date
      Version = $_.Version
      Rows    = @($_.Rows)
    }
  })

  return [pscustomobject]@{
    Blocks      = $blocksOut
    Diagnostics = [pscustomobject]$diag
  }
}

Write-Host ("Reading PDF: {0}" -f $PdfPath)
$pdfText = Get-PdfText -Path $PdfPath -PdftotextExe $PdftotextPath -First $FirstPage -Last $LastPage
Write-Verbose ("Extracted {0:N0} characters from PDF." -f $pdfText.Length)

$section = Get-ChangeHistorySection -Text $pdfText
Write-Verbose ("Change history section: {0:N0} characters." -f $section.Length)

# Show first 20 lines of the section for diagnostic purposes
$previewLines = ($section -split "`r?`n") | Select-Object -First 20
Write-Verbose ("--- Section preview (first 20 lines) ---`n" + ($previewLines -join "`n"))

if ($DiagnosticDump) {
    $dumpPath = Join-Path $OutputDir ('_DiagnosticDump_ChangeHistorySection.txt')
    Set-Content -LiteralPath $dumpPath -Value $section -Encoding UTF8
    Write-Host ("  Diagnostic dump written to: {0}" -f $dumpPath) -ForegroundColor Yellow
}

$parsed  = Get-ChangeData -Section $section
$blocks  = $parsed.Blocks

Write-Verbose ("Parser diagnostics: lines={0}; versions={1}; changes={2}; tickets={3}; unmatched={4}" -f `
  $parsed.Diagnostics.LinesScanned,
  $parsed.Diagnostics.VersionHeaders,
  $parsed.Diagnostics.ChangesParsed,
  $parsed.Diagnostics.TicketLines,
  $parsed.Diagnostics.UnmatchedLines)
if ($parsed.Diagnostics.UnmatchedLines -gt 0 -and $parsed.Diagnostics.UnmatchedSample.Count -gt 0) {
  Write-Verbose ("Unmatched samples: {0}" -f ($parsed.Diagnostics.UnmatchedSample -join ' | '))
}

if ($Version) {
    $target = $blocks | Where-Object { $_.Version -eq $Version } | Select-Object -First 1
    if (-not $target) {
        $avail = ($blocks | ForEach-Object { $_.Version }) -join ', '
        throw "Version '$Version' not found. Available versions: $avail"
    }
} else {
    $target = $blocks[0]   # the document lists versions newest-first
}

$rows = @($target.Rows)
if ($rows.Count -eq 0) { throw "No changes were parsed for version $($target.Version)." }

$previousMetadata = @()
$previousLookup = $null
if ($PreviousMarkdownPath) {
  Write-Host ("Reading previous markdown: {0}" -f $PreviousMarkdownPath)
  $previousMetadata = @(Read-PreviousMarkdownMetadata -Path $PreviousMarkdownPath)
  $previousLookup = New-MetadataLookup -Items $previousMetadata
  Write-Verbose ("Loaded {0} metadata rows from previous markdown." -f $previousMetadata.Count)
}

$currentPdfLookup = $null
$currentPdfMetadataLoaded = $false
function Get-CurrentPdfLookup {
  if (-not $script:CurrentPdfMetadataLoaded) {
    Write-Host 'Resolving metadata from current PDF recommendations...'
    $pdfMetadata = @(Get-PdfRecommendationMetadata -Path $PdfPath)
    $script:CurrentPdfLookup = New-MetadataLookup -Items $pdfMetadata
    $script:CurrentPdfMetadataLoaded = $true
    Write-Verbose ("Loaded {0} metadata rows from current PDF." -f $pdfMetadata.Count)
  }

  return $script:CurrentPdfLookup
}

$script:CurrentPdfLookup = $null
$script:CurrentPdfMetadataLoaded = $false

$resolvedFromPrevious = 0
$resolvedFromPdf = 0
$unresolvedRows = [System.Collections.Generic.List[object]]::new()

foreach ($row in $rows) {
  $match = $null
  if ($previousLookup) {
    $match = Find-MetadataMatch -Row $row -Lookup $previousLookup -MinSimilarity $SimilarityThreshold
    if ($match) {
      $row.Profile = $match.Profile
      $row.IntunePolicy = $match.IntunePolicy
      $row.MatchSource = 'PreviousMarkdown'
      $resolvedFromPrevious++
      continue
    }
  }

  $shouldTryPdf = ($row.Change -match '(?i)\bAdd\b') -or ($row.Change -notmatch '(?i)^Remove$')
  if ($shouldTryPdf) {
    try {
      $currentPdfLookup = Get-CurrentPdfLookup
      $match = Find-MetadataMatch -Row $row -Lookup $currentPdfLookup -MinSimilarity $SimilarityThreshold
      if ($match) {
        $row.Profile = $match.Profile
        $row.IntunePolicy = $match.IntunePolicy
        $row.MatchSource = 'CurrentPdf'
        $resolvedFromPdf++
        continue
      }
    } catch {
      Write-Warning ("Current PDF metadata lookup failed: {0}" -f $_.Exception.Message)
      $script:CurrentPdfMetadataLoaded = $true
    }
  }

  $unresolvedRows.Add($row)
}

# Build generic output names from the PDF file name (trailing version suffix removed).
$pdfBase = [System.IO.Path]::GetFileNameWithoutExtension($PdfPath)
$pdfBase = [regex]::Replace($pdfBase, '[_\-\s]+[vV]?\d+(?:\.\d+)*$', '')
if ([string]::IsNullOrWhiteSpace($pdfBase)) { $pdfBase = 'ChangeHistory' }
$title   = ($pdfBase -replace '[_\-]+', ' ').Trim()

$verSafe  = ($target.Version -replace '[^A-Za-z0-9._-]', '')
if ([string]::IsNullOrWhiteSpace($verSafe)) { $verSafe = 'current' }
$baseName = "{0}_ChangeHistory_v{1}" -f $pdfBase, $verSafe
$csvPath  = Join-Path $OutputDir "$baseName.csv"
$htmlPath = Join-Path $OutputDir "$baseName.html"

# --- CSV (columns: Setting, Change, Profile, Intune Policy) ------------------
$rows | Select-Object Setting, Change, Profile, @{ Name = 'Intune Policy'; Expression = { $_.IntunePolicy } } |
    Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

# --- Dark interactive HTML ---------------------------------------------------
$jsonRows = @($rows | Select-Object Setting, Change, Profile, IntunePolicy)
$json = ConvertTo-Json -InputObject $jsonRows -Depth 4
if ($jsonRows.Count -eq 1) { $json = "[$json]" }

$generated = (Get-Date).ToString('yyyy-MM-dd HH:mm')

$html = $HtmlTemplate
$html = $html.Replace('__DATA__',      $json)
$html = $html.Replace('__TITLE__',     $title)
$html = $html.Replace('__VERSION__',   [string]$target.Version)
$html = $html.Replace('__DATE__',      [string]$target.Date)
$html = $html.Replace('__COUNT__',     [string]$rows.Count)
$html = $html.Replace('__GENERATED__', $generated)

Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8

Write-Host ''
Write-Host '  Change History export complete' -ForegroundColor Cyan
Write-Host ('  Document         : {0}' -f $title)
Write-Host ('  Version exported : {0}  ({1})' -f $target.Version, $target.Date)
Write-Host ('  Total changes    : {0}' -f $rows.Count)
Write-Host ('  Previous matches : {0}' -f $resolvedFromPrevious)
Write-Host ('  PDF matches      : {0}' -f $resolvedFromPdf)
Write-Host ('  Unresolved rows  : {0}' -f $unresolvedRows.Count)
Write-Host ('  CSV              : {0}' -f $csvPath)
Write-Host ('  HTML             : {0}' -f $htmlPath)
Write-Host ''
