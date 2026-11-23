[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [Parameter()]
    [string]$OutputFile,

    [Parameter()]
    [string]$Title,

    [switch]$OpenReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    if (-not (Test-Path -Path $InputFile -PathType Leaf)) {
        throw "Input file not found: $InputFile"
    }

    $resolvedInput = (Resolve-Path -Path $InputFile).Path

    if ([string]::IsNullOrWhiteSpace($OutputFile)) {
        $OutputFile = [System.IO.Path]::ChangeExtension($resolvedInput, '.html')
    }

    if ([string]::IsNullOrWhiteSpace($Title)) {
        $Title = [System.IO.Path]::GetFileNameWithoutExtension($resolvedInput)
    }

    $markdown = Get-Content -Path $resolvedInput -Raw -Encoding UTF8
    $markdown = $markdown -replace "`0", ''

    $reportBodyHtml = $null
    $convertFromMarkdownCmd = Get-Command -Name ConvertFrom-Markdown -ErrorAction SilentlyContinue

    if ($convertFromMarkdownCmd) {
        try {
            $converted = ConvertFrom-Markdown -InputObject $markdown
            if ($converted -and $converted.Html) {
                $reportBodyHtml = $converted.Html
            }
        }
        catch {
            Write-Warning "ConvertFrom-Markdown failed. Falling back to plain text rendering. Error: $($_.Exception.Message)"
        }
    }

    if (-not $reportBodyHtml) {
        $encoded = [System.Net.WebUtility]::HtmlEncode($markdown)
        $reportBodyHtml = "<pre class='markdown-fallback'>$encoded</pre>"
    }

    $htmlTemplate = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__REPORT_TITLE__</title>
<style>
  :root {
    --bg-main: #0a1220;
    --bg-panel: #111d32;
    --bg-elevated: #182841;
    --fg-main: #e9f1ff;
    --fg-muted: #9eb3d3;
    --accent: #44b3ff;
    --accent-strong: #2d8ce6;
    --border: #2a3d5f;
    --hover: #1f3250;
    --tag-l1: #6be3a6;
    --tag-l2: #ffb576;
    --tag-yes: #6be3a6;
    --tag-no: #ff8b96;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Segoe UI Variable Text', 'Bahnschrift', 'Trebuchet MS', sans-serif;
    background:
      radial-gradient(1200px 600px at 85% -10%, #24497a 0%, rgba(36, 73, 122, 0) 55%),
      radial-gradient(950px 520px at -10% 0%, #183663 0%, rgba(24, 54, 99, 0) 58%),
      var(--bg-main);
    color: var(--fg-main);
    padding: 24px;
    line-height: 1.5;
    min-height: 100vh;
  }

  h1 {
    font-size: clamp(1.35rem, 2vw, 1.9rem);
    margin-bottom: 6px;
    color: #ffffff;
    font-weight: 700;
    letter-spacing: 0.02em;
    text-shadow: 0 2px 10px rgba(0, 0, 0, 0.35);
  }

  .subtitle {
    color: var(--fg-muted);
    font-size: 0.9rem;
    margin-bottom: 18px;
  }

  .stats {
    display: flex;
    gap: 14px;
    margin-bottom: 14px;
    flex-wrap: wrap;
  }

  .stat {
    background: linear-gradient(160deg, rgba(255, 255, 255, 0.03), rgba(255, 255, 255, 0.01));
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 12px 16px;
    min-width: 148px;
    box-shadow: 0 8px 20px rgba(0, 0, 0, 0.25);
    backdrop-filter: blur(2px);
  }

  .stat-value {
    font-size: 1.45rem;
    font-weight: 700;
    color: var(--accent);
  }

  .stat-label {
    font-size: 0.72rem;
    color: var(--fg-muted);
    text-transform: uppercase;
    letter-spacing: 0.08em;
  }

  .toolbar {
    display: flex;
    gap: 10px;
    margin-bottom: 12px;
    align-items: center;
    flex-wrap: wrap;
    padding: 10px;
    background: rgba(17, 29, 50, 0.78);
    border: 1px solid var(--border);
    border-radius: 12px;
  }

  .toolbar input[type=text] {
    padding: 9px 12px;
    border: 1px solid #385482;
    border-radius: 8px;
    font-size: 0.85rem;
    width: 280px;
    color: var(--fg-main);
    background: #0f1a2e;
    transition: border-color 0.2s ease, box-shadow 0.2s ease;
  }

  .toolbar input[type=text]:focus,
  tr.filter-row input:focus {
    outline: none;
    border-color: var(--accent);
    box-shadow: 0 0 0 2px rgba(68, 179, 255, 0.2);
  }

  .toolbar button {
    padding: 9px 14px;
    border: 1px solid #3f5d8f;
    border-radius: 8px;
    background: linear-gradient(180deg, #1e3354, #162947);
    color: var(--fg-main);
    cursor: pointer;
    font-size: 0.82rem;
    font-weight: 600;
    letter-spacing: 0.02em;
    transition: transform 0.15s ease, border-color 0.2s ease, background 0.2s ease;
  }

  .toolbar button:hover {
    border-color: #5f88ca;
    background: linear-gradient(180deg, #2a456e, #1b3358);
    transform: translateY(-1px);
  }

  .count {
    font-size: 0.84rem;
    color: var(--fg-muted);
    margin-left: auto;
    font-weight: 600;
  }

  .table-wrap {
    overflow-x: auto;
    border: 1px solid var(--border);
    border-radius: 14px;
    background: rgba(17, 29, 50, 0.85);
    box-shadow: 0 18px 36px rgba(0, 0, 0, 0.35);
  }

  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
  }

  thead {
    position: sticky;
    top: 0;
    z-index: 2;
  }

  th {
    background: linear-gradient(180deg, #2d4872, #243c62);
    color: #f4f8ff;
    padding: 11px 9px;
    text-align: left;
    font-weight: 650;
    white-space: nowrap;
    cursor: pointer;
    user-select: none;
    border-bottom: 1px solid #36547f;
  }

  th:hover {
    background: linear-gradient(180deg, #34588b, #2a4874);
  }

  th .sort-arrow { font-size: 0.65rem; margin-left: 4px; opacity: 0.6; }
  th.sorted-asc .sort-arrow::after { content: ' ▲'; opacity: 1; }
  th.sorted-desc .sort-arrow::after { content: ' ▼'; opacity: 1; }

  tr.filter-row th {
    background: #172a47;
    padding: 6px;
    border-bottom: 1px solid #264164;
  }

  tr.filter-row input {
    width: 100%;
    padding: 6px 7px;
    border: 1px solid #345a8e;
    border-radius: 6px;
    font-size: 0.78rem;
    background: #101e35;
    color: var(--fg-main);
  }

  td {
    padding: 8px 9px;
    border-bottom: 1px solid rgba(80, 117, 170, 0.25);
    vertical-align: top;
    color: #dce8ff;
  }

  tbody tr:nth-child(even) td {
    background: rgba(255, 255, 255, 0.02);
  }

  tr:hover td {
    background: var(--hover);
  }

  .tag {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 0.73rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    border: 1px solid transparent;
  }

  .tag-l1 { background: rgba(107, 227, 166, 0.15); color: var(--tag-l1); border-color: rgba(107, 227, 166, 0.35); }
  .tag-l2 { background: rgba(255, 181, 118, 0.15); color: var(--tag-l2); border-color: rgba(255, 181, 118, 0.35); }
  .tag-yes { background: rgba(107, 227, 166, 0.15); color: var(--tag-yes); border-color: rgba(107, 227, 166, 0.35); }
  .tag-no { background: rgba(255, 139, 150, 0.14); color: var(--tag-no); border-color: rgba(255, 139, 150, 0.35); }
  .tag-auto { background: rgba(116, 192, 255, 0.16); color: #8fd2ff; border-color: rgba(116, 192, 255, 0.38); }
  .tag-manual { background: rgba(255, 215, 120, 0.15); color: #ffd478; border-color: rgba(255, 215, 120, 0.35); }

  a { color: #7cc9ff; text-decoration: none; }
  a:hover { color: #b2e2ff; text-decoration: underline; }

  td:nth-child(2) { min-width: 280px; max-width: 450px; }
  #reportRoot { display: none; }

  @media (max-width: 900px) {
    body { padding: 10px; }
    .toolbar {
      padding: 10px;
      gap: 8px;
    }
    .toolbar input[type=text] {
      width: 100%;
      max-width: 100%;
    }
    .count {
      width: 100%;
      margin-left: 0;
      margin-top: 2px;
    }
    .stat {
      flex: 1 1 140px;
    }
  }
</style>
</head>
<body>
<h1>CIS Benchmark - __REPORT_TITLE__</h1>
<p class="subtitle">Generated __GENERATED_AT__ • Source: __SOURCE_NAME__</p>

<div class="stats">
  <div class="stat"><div class="stat-value" id="statRecommendations">0</div><div class="stat-label">Recommendations</div></div>
  <div class="stat"><div class="stat-value" id="statL1">0</div><div class="stat-label">Level 1 (L1)</div></div>
  <div class="stat"><div class="stat-value" id="statL2">0</div><div class="stat-label">Level 2 (L2)</div></div>
  <div class="stat"><div class="stat-value" id="statSections">0</div><div class="stat-label">Major Sections</div></div>
</div>

<div class="toolbar">
  <input type="text" id="globalSearch" placeholder="Search all columns..." oninput="applyFilters()">
  <button onclick="clearFilters()">Clear Filters</button>
  <button onclick="exportCSV()">Export CSV</button>
  <span class="count" id="rowCount"></span>
</div>

<div class="table-wrap">
<table id="cis-table">
<thead id="tableHead"></thead>
<tbody id="tableBody"></tbody>
</table>
</div>

<article id="reportRoot">__REPORT_BODY__</article>

<script>
let rows = [];
let sortCol = -1;
let sortAsc = true;

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function previousHeadingText(node) {
  let current = node;
  while (current) {
    current = current.previousElementSibling;
    if (!current) break;
    if (/^H[1-4]$/.test(current.tagName)) {
      return current.textContent.trim();
    }
  }
  return 'General';
}

function extractTableData() {
  const root = document.getElementById('reportRoot');
  const markdownTables = Array.from(root.querySelectorAll('table'));
  const allRows = [];
  let baseHeaders = [];
  const headersToRemove = new Set(['gp setting name', 'remediation app']);

  markdownTables.forEach((table) => {
    const section = previousHeadingText(table);
    const headerCells = Array.from(table.querySelectorAll('thead th'));
    let headers = headerCells.map((th) => th.textContent.trim());

    if (!headers.length) {
      const firstRow = table.querySelector('tr');
      if (firstRow) {
        headers = Array.from(firstRow.children).map((c) => c.textContent.trim());
      }
    }

    const removeIndexes = [];
    headers.forEach((header, idx) => {
      if (headersToRemove.has((header || '').trim().toLowerCase())) {
        removeIndexes.push(idx);
      }
    });
    const filteredHeaders = headers.filter((_, idx) => !removeIndexes.includes(idx));

    if (!baseHeaders.length && filteredHeaders.length) {
      baseHeaders = filteredHeaders;
    }

    const bodyRows = Array.from(table.querySelectorAll('tbody tr'));
    let dataRows = bodyRows;
    if (!dataRows.length) {
      const allTr = Array.from(table.querySelectorAll('tr'));
      dataRows = allTr.slice(1);
    }

    dataRows.forEach((tr) => {
      const cells = Array.from(tr.children)
        .map((td) => td.innerHTML.trim())
        .filter((_, idx) => !removeIndexes.includes(idx));
      if (!cells.length) return;
      if (cells.every((c) => c === '')) return;
      allRows.push([section, ...cells]);
    });
  });

  const finalHeaders = ['Section', ...baseHeaders];
  return { headers: finalHeaders, rows: allRows };
}

function renderTable(headers, dataRows) {
  const tableHead = document.getElementById('tableHead');
  const tableBody = document.getElementById('tableBody');

  const headerRow = document.createElement('tr');
  headers.forEach((header, i) => {
    const th = document.createElement('th');
    th.innerHTML = escapeHtml(header) + '<span class="sort-arrow"></span>';
    th.onclick = () => sortTable(i);
    headerRow.appendChild(th);
  });

  const filterRow = document.createElement('tr');
  filterRow.className = 'filter-row';
  headers.forEach((_, i) => {
    const th = document.createElement('th');
    th.innerHTML = '<input type="text" data-col="' + i + '" placeholder="Filter..." oninput="applyFilters()">';
    filterRow.appendChild(th);
  });

  tableHead.appendChild(headerRow);
  tableHead.appendChild(filterRow);

  dataRows.forEach((rowCells) => {
    const tr = document.createElement('tr');
    rowCells.forEach((cellHtml) => {
      const td = document.createElement('td');
      td.innerHTML = cellHtml;
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });

  rows = Array.from(tableBody.querySelectorAll('tr'));
}

function applyTags() {
  rows.forEach((row) => {
    const cells = row.querySelectorAll('td');
    if (cells.length < 2) return;

    cells.forEach((cell) => {
      const value = cell.textContent.trim();
      if (value === 'L1') cell.innerHTML = '<span class="tag tag-l1">L1</span>';
      if (value === 'L2') cell.innerHTML = '<span class="tag tag-l2">L2</span>';
      if (value === 'Yes') cell.innerHTML = '<span class="tag tag-yes">Yes</span>';
      if (value === 'No') cell.innerHTML = '<span class="tag tag-no">No</span>';
      if (value === 'Automated') cell.innerHTML = '<span class="tag tag-auto">Automated</span>';
      if (value === 'Manual') cell.innerHTML = '<span class="tag tag-manual">Manual</span>';
    });
  });
}

function updateStats() {
  const visibleRows = rows.filter((r) => r.style.display !== 'none');
  const recommendationCount = rows.length;

  let l1 = 0;
  let l2 = 0;
  const sections = new Set();

  rows.forEach((row) => {
    const cells = row.querySelectorAll('td');
    if (cells[0]) sections.add(cells[0].textContent.trim());
    cells.forEach((cell) => {
      const value = cell.textContent.trim();
      if (value === 'L1') l1 += 1;
      if (value === 'L2') l2 += 1;
    });
  });

  document.getElementById('statRecommendations').textContent = String(recommendationCount);
  document.getElementById('statL1').textContent = String(l1);
  document.getElementById('statL2').textContent = String(l2);
  document.getElementById('statSections').textContent = String(sections.size);
  document.getElementById('rowCount').textContent = visibleRows.length + ' of ' + rows.length + ' shown';
}

function applyFilters() {
  const table = document.getElementById('cis-table');
  const colFilters = table.querySelectorAll('.filter-row input');
  const globalSearch = document.getElementById('globalSearch').value.toLowerCase();
  const filters = Array.from(colFilters).map((i) => ({ col: Number(i.dataset.col), val: i.value.toLowerCase() }));

  rows.forEach((row) => {
    const cells = Array.from(row.querySelectorAll('td'));
    const texts = cells.map((c) => c.textContent.toLowerCase());
    let show = true;

    if (globalSearch && !texts.some((t) => t.includes(globalSearch))) {
      show = false;
    }

    for (const filter of filters) {
      if (filter.val && !texts[filter.col]?.includes(filter.val)) {
        show = false;
        break;
      }
    }

    row.style.display = show ? '' : 'none';
  });

  updateStats();
}

function clearFilters() {
  const table = document.getElementById('cis-table');
  document.getElementById('globalSearch').value = '';
  table.querySelectorAll('.filter-row input').forEach((i) => { i.value = ''; });
  applyFilters();
}

function sortTable(col) {
  const table = document.getElementById('cis-table');
  const tbody = table.querySelector('tbody');
  const ths = table.querySelectorAll('thead tr:first-child th');

  if (sortCol === col) {
    sortAsc = !sortAsc;
  } else {
    sortCol = col;
    sortAsc = true;
  }

  ths.forEach((th, i) => {
    th.classList.remove('sorted-asc', 'sorted-desc');
    if (i === col) {
      th.classList.add(sortAsc ? 'sorted-asc' : 'sorted-desc');
    }
  });

  rows.sort((a, b) => {
    const va = a.querySelectorAll('td')[col]?.textContent.trim() || '';
    const vb = b.querySelectorAll('td')[col]?.textContent.trim() || '';
    const na = Number.parseFloat(va);
    const nb = Number.parseFloat(vb);

    if (!Number.isNaN(na) && !Number.isNaN(nb)) {
      return sortAsc ? na - nb : nb - na;
    }

    return sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
  });

  rows.forEach((row) => tbody.appendChild(row));
  applyFilters();
}

function exportCSV() {
  const table = document.getElementById('cis-table');
  const headers = Array.from(table.querySelectorAll('thead tr:first-child th')).map((th) => th.childNodes[0].textContent.trim());
  const csvRows = [headers.map((h) => '"' + h.replace(/"/g, '""') + '"').join(',')];

  rows.forEach((row) => {
    if (row.style.display === 'none') return;
    const values = Array.from(row.querySelectorAll('td')).map((cell) => '"' + cell.textContent.replace(/"/g, '""').trim() + '"');
    csvRows.push(values.join(','));
  });

  const blob = new Blob([csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const anchor = document.createElement('a');
  anchor.href = URL.createObjectURL(blob);
  anchor.download = '__REPORT_TITLE__-CIS-Benchmark.csv';
  anchor.click();
}

const extracted = extractTableData();
renderTable(extracted.headers, extracted.rows);
applyTags();
applyFilters();
</script>
</body>
</html>
'@

    $sourceName = [System.IO.Path]::GetFileName($resolvedInput)
    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $htmlContent = $htmlTemplate.Replace('__REPORT_TITLE__', [System.Net.WebUtility]::HtmlEncode($Title))
    $htmlContent = $htmlContent.Replace('__REPORT_BODY__', $reportBodyHtml)
    $htmlContent = $htmlContent.Replace('__SOURCE_NAME__', [System.Net.WebUtility]::HtmlEncode($sourceName))
    $htmlContent = $htmlContent.Replace('__GENERATED_AT__', [System.Net.WebUtility]::HtmlEncode($generatedAt))

    $outputDirectory = Split-Path -Path $OutputFile -Parent
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -Path $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    Set-Content -Path $OutputFile -Value $htmlContent -Encoding UTF8
    Write-Output "Interactive HTML report created: $OutputFile"

    if ($OpenReport) {
        Start-Process -FilePath $OutputFile | Out-Null
    }
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
