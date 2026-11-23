#Requires -Version 5.1
<#
.SYNOPSIS
    Merges baseline Markdown and optionally generates the existing comparison report.
.EXAMPLE
    .\Invoke-BaselineMerger.ps1 -NewBaselinePath .\new.md -OldBaselinePath .\old.md -OutputPath .\merged.md -GenerateReport
.EXAMPLE
    .\Invoke-BaselineMerger.ps1 -ReportOnly -InputMarkdownPath .\merged.md -OutputHtmlPath .\report.html
.NOTES
    Dot-source to load the processing functions without performing an operation.
#>
[CmdletBinding(DefaultParameterSetName = 'Merge')]
param(
    [Parameter(ParameterSetName = 'Merge')][string]$NewBaselinePath,
    [Parameter(ParameterSetName = 'Merge')][string]$OldBaselinePath,
    [Parameter(ParameterSetName = 'Merge')][string]$OutputPath,
    [Parameter(ParameterSetName = 'Merge')][ValidateSet('All','L1','L2')][string]$ProfileFilter = 'L1',
    [Parameter(ParameterSetName = 'Merge')][ValidateRange(0.0,1.0)][double]$SimilarityThreshold = 1.0,
    [Parameter(ParameterSetName = 'Merge')][switch]$GenerateReport,
    [Parameter(ParameterSetName = 'Report')][switch]$ReportOnly,
    [Parameter(ParameterSetName = 'Report')][string]$InputMarkdownPath,
    [string]$OutputHtmlPath,
    [string]$ReportTitle,
    [switch]$OpenReport
)

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Shared\BaselineToolkit.Common.psm1') -DisableNameChecking

function Invoke-MergeBaselines {
    param(
        [string]$NewBaselinePath,
        [string]$OldBaselinePath,
        [string]$OutputPath,
        [ValidateSet('All','L1','L2')]
        [string]$ProfileFilter = 'L1',
        [ValidateRange(0.0,1.0)]
        [double]$SimilarityThreshold = 1.0
    )

    function Test-IncludedProfile {
        param([string]$Profile, [string]$Filter)
        $p = ([string]$Profile).Trim().ToUpperInvariant()
        switch ($Filter) {
            'All' { return $true }
            'L2'  { return ($p -eq 'L2' -or $p -match '(^|\s)L2$') }
            default { return ($p -eq 'L1' -or $p -match '(^|\s)L1$' -or $p -eq 'BL') }
        }
    }

    function Find-BestMatchKey {
        param(
            [string]$Target,
            [string[]]$Candidates,
            [double]$Threshold
        )

        if (-not $Target -or -not $Candidates -or $Candidates.Count -eq 0) { return $null }
        if ($Threshold -ge 1.0) {
            if ($Candidates -contains $Target) { return $Target }
            return $null
        }

        $best = $null
        $bestScore = -1.0
        foreach ($candidate in $Candidates) {
            if (-not $candidate) { continue }
            $score = Get-StringSimilarity -Left $Target -Right $candidate
            if ($score -gt $bestScore) {
                $bestScore = $score
                $best = $candidate
            }
        }

        if ($bestScore -ge $Threshold) { return $best }
        return $null
    }

    $newDoc = Read-MarkdownDocument $NewBaselinePath
    $oldDoc = Read-MarkdownDocument $OldBaselinePath
    $oldIndex  = @{}
    $twodaySections = [System.Collections.Generic.List[hashtable]]::new()
    $canonTwodayHdrs = @('Recommendation','CIS Reference','Profile','CIS Safeguard','Type','Origin','Followed','Notes','Intune Policy','References','Status')

    foreach ($sec in $oldDoc.Sections) {
        if (-not $sec['TableLines'].Count) { continue }
        $tbl = Read-MarkdownTable $sec['TableLines']
        if (-not $tbl) { continue }
        $ciI = Get-ColumnIndex $tbl.Headers 'CIS Reference'
        $foI = Get-ColumnIndex $tbl.Headers 'Followed'
        $prI = Get-ColumnIndex $tbl.Headers 'Profile'
        $noI = Get-ColumnIndex $tbl.Headers 'Notes'
        $orI = Get-ColumnIndex $tbl.Headers 'Origin'
        $reI = Get-ColumnIndex $tbl.Headers 'Recommendation'
        $ipI = Get-ColumnIndex $tbl.Headers 'Intune Policy'
        $curTwodayGroup = $null
        foreach ($row in $tbl.Rows) {
            $cisRef         = if ($ciI -ge 0) { Get-PlainText $row.Cells[$ciI] } else { '' }
            $followed       = if ($foI -ge 0) { Get-PlainText $row.Cells[$foI] } else { '' }
            $rowProfile     = if ($prI -ge 0) { Get-PlainText $row.Cells[$prI] } else { '' }
            $notes          = if ($noI -ge 0) { Get-PlainText $row.Cells[$noI] } else { '' }
            $origin         = if ($orI -ge 0) { Get-PlainText $row.Cells[$orI] } else { '' }
            $intunePolicy   = if ($ipI -ge 0) { Get-PlainText $row.Cells[$ipI] } else { '' }
            $recommendation = if ($reI -ge 0) { Get-PlainText $row.Cells[$reI] } else { '' }
            $recKey         = Get-RecKey $recommendation
            if ($recKey) { $oldIndex[$recKey] = @{ Followed = $followed; Profile = $rowProfile; Notes = $notes; CISReference = $cisRef; Origin = $origin; IntunePolicy = $intunePolicy; Section = $sec['HeadingText'] } }
            if ($origin -imatch 'twoday') {
                if (-not $curTwodayGroup) {
                    $curTwodayGroup = @{ HeadingLine = $sec['HeadingLine']; HeadingText = $sec['HeadingText']; Headers = $tbl.Headers; Rows = [System.Collections.Generic.List[hashtable]]::new() }
                    $twodaySections.Add($curTwodayGroup)
                }
                $curTwodayGroup.Rows.Add(@{ Cells = $row.Cells; RecKey = $recKey })
            }
        }
    }

    $newRecSet = @{}
    $newCisRecSet = @{}
    foreach ($sec in $newDoc.Sections) {
        if (-not $sec['TableLines'].Count) { continue }
        $tbl2 = Read-MarkdownTable $sec['TableLines']
        if (-not $tbl2) { continue }
        $prI2 = Get-ColumnIndex $tbl2.Headers 'Profile'
        $orI2 = Get-ColumnIndex $tbl2.Headers 'Origin'
        $reI2 = Get-ColumnIndex $tbl2.Headers 'Recommendation'
        foreach ($row in $tbl2.Rows) {
            $rowProf2 = if ($prI2 -ge 0) { Get-PlainText $row.Cells[$prI2] } else { '' }
            $rowOrig2 = if ($orI2 -ge 0) { Get-PlainText $row.Cells[$orI2] } else { '' }
            $rowRec2  = if ($reI2 -ge 0) { Get-PlainText $row.Cells[$reI2] } else { '' }
            $recKey2  = Get-RecKey $rowRec2
            if ($recKey2 -and (Test-IncludedProfile -Profile $rowProf2 -Filter $ProfileFilter)) { $newRecSet[$recKey2] = $true }
            if ($recKey2 -and $rowOrig2 -imatch 'cis') { $newCisRecSet[$recKey2] = $true }
        }
    }

    $mismatches = [System.Collections.Generic.List[hashtable]]::new()
    $newRecKeys = @($newRecSet.Keys)
    $newCisRecKeys = @($newCisRecSet.Keys)
    foreach ($recKey in ($oldIndex.Keys | Sort-Object)) {
        $old = $oldIndex[$recKey]
        if (-not (Test-IncludedProfile -Profile $old.Profile -Filter $ProfileFilter)) { continue }
        $isMatched = $newRecSet.ContainsKey($recKey)
        if (-not $isMatched -and $SimilarityThreshold -lt 1.0) {
            $matchedKey = Find-BestMatchKey -Target $recKey -Candidates $newRecKeys -Threshold $SimilarityThreshold
            $isMatched = [bool]$matchedKey
        }
        if (-not $isMatched) {
            $mismatches.Add(@{ CISReference = $old.CISReference; OldRec = $recKey; OldFollowed = $old.Followed; OldNotes = $old.Notes; Origin = $old.Origin; IntunePolicy = $old.IntunePolicy; Section = $old.Section })
        }
    }

    $output = [System.Collections.Generic.List[string]]::new()
    $twodayTotal = ($twodaySections | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum).Sum
    if ($null -eq $twodayTotal) { $twodayTotal = 0 }
    $stats  = @{ Overrides = 0; NotesCopied = 0; TwodayCount = $twodayTotal; L2Skipped = 0; ProfileFiltered = 0; Mismatches = $mismatches }
    foreach ($ln in $newDoc.Preamble) { $output.Add($ln) }

    foreach ($sec in $newDoc.Sections) {
        $output.Add(''); $output.Add($sec['HeadingLine']); $output.Add('')
        if (-not $sec['TableLines'].Count) { continue }
        $tbl = Read-MarkdownTable $sec['TableLines']
        if (-not $tbl) { foreach ($ln in $sec['TableLines']) { $output.Add($ln) }; continue }
        $outHeadersNoStatus = [System.Collections.Generic.List[string]]::new()
        foreach ($hdr in $tbl.Headers) {
            if ($hdr -ine 'Link' -and $hdr -ine 'twoday ID' -and $hdr -ine 'Remediation App') { $outHeadersNoStatus.Add($hdr) }
        }
        $outHeaders = [System.Collections.Generic.List[string]]::new()
        foreach ($hdr in $outHeadersNoStatus) { $outHeaders.Add($hdr) }
        $refsOutIndex = [Array]::IndexOf([string[]]$outHeadersNoStatus.ToArray(), 'References')
        if ($refsOutIndex -ge 0) {
            $outHeaders.Insert($refsOutIndex + 1, 'Status')
        } else {
            $outHeaders.Add('Status')
        }
        $output.Add('| ' + (($outHeaders.ToArray()) -join ' | ') + ' |')
        $output.Add('| ' + ((@($outHeaders | ForEach-Object { '---' })) -join ' | ') + ' |')
        $foI = Get-ColumnIndex $tbl.Headers 'Followed'
        $prI = Get-ColumnIndex $tbl.Headers 'Profile'
        $noI = Get-ColumnIndex $tbl.Headers 'Notes'
        $orI = Get-ColumnIndex $tbl.Headers 'Origin'
        $reI = Get-ColumnIndex $tbl.Headers 'Recommendation'
        $liI = Get-ColumnIndex $tbl.Headers 'Link'
        $tdI = Get-ColumnIndex $tbl.Headers 'twoday ID'
        $raI = Get-ColumnIndex $tbl.Headers 'Remediation App'
        foreach ($row in $tbl.Rows) {
            $cells      = [string[]]$row.Cells
            $rowProfile = if ($prI -ge 0) { Get-PlainText $cells[$prI] } else { '' }
            if (-not (Test-IncludedProfile -Profile $rowProfile -Filter $ProfileFilter)) {
                $stats.ProfileFiltered++
                if ($ProfileFilter -eq 'L1') { $stats.L2Skipped++ }
                continue
            }
            $newRecText = if ($reI -ge 0) { Get-RecKey $cells[$reI] } else { '' }
            $originText = if ($orI -ge 0) { Get-PlainText $cells[$orI] } else { '' }
            $matchKey = $null
            if ($newRecText) {
                $matchKey = if ($oldIndex.ContainsKey($newRecText)) { $newRecText } elseif ($SimilarityThreshold -lt 1.0) { Find-BestMatchKey -Target $newRecText -Candidates @($oldIndex.Keys) -Threshold $SimilarityThreshold } else { $null }
                if ($matchKey) {
                    $old = $oldIndex[$matchKey]
                    if ($foI -ge 0 -and $old.Followed -eq 'No') { if ((Get-PlainText $cells[$foI]) -ne 'No') { $cells[$foI] = 'No'; $stats.Overrides++ } }
                    if ($noI -ge 0 -and $old.Notes -and $old.Notes -ne '-') { $cells[$noI] = $old.Notes; $stats.NotesCopied++ }
                }
            }
            $statusValue = if ($newRecText -and -not $oldIndex.ContainsKey($newRecText) -and $originText -notmatch '(?i)twoday') { 'New' } else { '-' }

            $outRow = [System.Collections.Generic.List[string]]::new()
            for ($i = 0; $i -lt $cells.Count; $i++) {
                if ($i -ne $liI -and $i -ne $tdI -and $i -ne $raI) { $outRow.Add([string]$cells[$i]) }
            }
            if ($refsOutIndex -ge 0) {
                $insertAt = [Math]::Min($refsOutIndex + 1, $outRow.Count)
                $outRow.Insert($insertAt, $statusValue)
            } else {
                $outRow.Add($statusValue)
            }

            $output.Add((Build-PipeRow $outRow.ToArray()))
        }
    }

    if ($twodaySections.Count -gt 0) {
        $output.Add(''); $output.Add('---'); $output.Add(''); $output.Add('## Twoday Recommendations'); $output.Add('')
        $output.Add('> The following recommendations originate from Twoday and are included from the old baseline. Status shows whether the recommendation is replaced by a CIS recommendation in the new baseline.'); $output.Add('')
        $tdSep = '| ' + (($canonTwodayHdrs | ForEach-Object { '---' }) -join ' | ') + ' |'
        foreach ($tdSec in $twodaySections) {
            $output.Add($tdSec.HeadingLine); $output.Add('')
            $output.Add('| ' + ($canonTwodayHdrs -join ' | ') + ' |'); $output.Add($tdSep)
            $srcHdrs = [string[]]$tdSec.Headers
            foreach ($tr in $tdSec.Rows) {
                $srcCells = [string[]]$tr.Cells
                $remapped = @($canonTwodayHdrs | Select-Object -First ($canonTwodayHdrs.Count - 1) | ForEach-Object { $idx = [Array]::IndexOf($srcHdrs, $_); if ($idx -ge 0 -and $idx -lt $srcCells.Count) { $srcCells[$idx] } else { '' } })
                $twRecKey = if ($tr.ContainsKey('RecKey')) { [string]$tr.RecKey } else { '' }
                $replacedByCis = $false
                if ($twRecKey) {
                    $replacedByCis = $newCisRecSet.ContainsKey($twRecKey)
                    if (-not $replacedByCis -and $SimilarityThreshold -lt 1.0) {
                        $matchedCisKey = Find-BestMatchKey -Target $twRecKey -Candidates $newCisRecKeys -Threshold $SimilarityThreshold
                        $replacedByCis = [bool]$matchedCisKey
                    }
                }
                $status = if ($replacedByCis) { 'Replaced by CIS' } else { 'Included' }
                $output.Add((Build-PipeRow @($remapped + $status)))
            }
            $output.Add('')
        }
    }

    $dir = Split-Path -Path $OutputPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    $output | Set-Content -Path $OutputPath -Encoding UTF8
    return $stats
}

function Invoke-GenerateHtmlReport {
    param([string]$InputMarkdownPath, [string]$OutputHtmlPath, [string]$ReportTitle, [hashtable[]]$Mismatches = @())
    if ($null -eq $Mismatches) { $Mismatches = @() }
    if (-not (Test-Path $InputMarkdownPath -PathType Leaf)) { throw "Markdown file not found: $InputMarkdownPath" }

    # Use the same report generator as PDF to Markdown so colors/layout stay aligned.
    $converterScriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'Convert-MarkdownTextToInteractiveHtmlReport.ps1'
    if (-not (Test-Path -Path $converterScriptPath -PathType Leaf)) {
        throw "Converter script not found: $converterScriptPath"
    }

    $converterOutput = & $converterScriptPath -InputFile $InputMarkdownPath -OutputFile $OutputHtmlPath -Title $ReportTitle 2>&1
    foreach ($line in @($converterOutput)) {
        if ($null -ne $line -and "$( [string]$line )".Trim()) {
            Write-Verbose ([string]$line)
        }
    }

    if (-not (Test-Path -Path $OutputHtmlPath -PathType Leaf)) {
        throw "HTML report was not generated: $OutputHtmlPath"
    }

    if ($Mismatches.Count -gt 0) {
        $mismatchRowsHtml = foreach ($m in ($Mismatches | Sort-Object @{ Expression = { $_['Section'] } }, @{ Expression = { $_['CISReference'] } })) {
            $followedHtml = switch ($m['OldFollowed']) {
                'Yes' { '<span class="tag tag-yes">Yes</span>' }
                'No'  { '<span class="tag tag-no">No</span>' }
                default { [System.Net.WebUtility]::HtmlEncode($m['OldFollowed']) }
            }
            $section = if ($m['Section']) { [System.Net.WebUtility]::HtmlEncode($m['Section']) } else { '-' }
            $origin = if ($m['Origin']) { [System.Net.WebUtility]::HtmlEncode($m['Origin']) } else { '-' }
            $intunePolicy = if ($m['IntunePolicy']) { [System.Net.WebUtility]::HtmlEncode($m['IntunePolicy']) } else { '-' }
            "        <tr><td>$section</td><td>$([System.Net.WebUtility]::HtmlEncode($m['OldRec']))</td><td>$origin</td><td>$followedHtml</td><td>$([System.Net.WebUtility]::HtmlEncode($m['OldNotes']))</td><td>$intunePolicy</td><td><span class='tag-mismatch-missing'>Not in new baseline</span></td></tr>"
        }

        $mismatchSection = @"
<div id="mismatch-section" class="mismatch-section">
  <div class="mismatch-heading">Not in new baseline ($($Mismatches.Count))</div>
  <div class="table-wrap" style="margin-top:12px"><table><thead><tr><th>Section</th><th>Recommendation</th><th style="max-width:60px">Origin</th><th style="max-width:40px">Followed</th><th>Notes</th><th>Intune Policy</th><th style="min-width:130px">Status</th></tr><tr class='filter-row'><th><input class='col-filter' data-col='0' type='text' placeholder='Filter...' oninput='applyMismatchFilters()'></th><th><input class='col-filter' data-col='1' type='text' placeholder='Filter...' oninput='applyMismatchFilters()'></th><th><input class='col-filter' data-col='2' type='text' placeholder='Filter...' oninput='applyMismatchFilters()'></th><th><input class='col-filter' data-col='3' type='text' placeholder='Filter...' oninput='applyMismatchFilters()'></th><th><input class='col-filter' data-col='4' type='text' placeholder='Filter...' oninput='applyMismatchFilters()'></th><th><input class='col-filter' data-col='5' type='text' placeholder='Filter...' oninput='applyMismatchFilters()'></th><th><input class='col-filter' data-col='6' type='text' placeholder='Filter...' oninput='applyMismatchFilters()'></th></tr></thead><tbody id="mismatchTableBody">
$($mismatchRowsHtml -join "`n")
  </tbody></table><span id="mismatchRowCount" class="mismatch-count"></span></div>
</div>
"@

        $mismatchCss = @"
.mismatch-section{margin-top:28px;padding:16px 18px;background:rgba(17,29,50,.82);border:1px solid var(--border);border-radius:12px;box-shadow:0 14px 30px rgba(0,0,0,.28)}
.mismatch-heading{font-size:1.02rem;font-weight:700;color:var(--accent);margin-bottom:8px;letter-spacing:.02em}
.filter-row th{background:transparent}
.filter-row input{width:100%;padding:4px;border:1px solid var(--border);border-radius:3px;background:var(--surface,#f8f8f8);color:var(--fg,#1a1a1a)}
.mismatch-count{font-size:.82rem;color:#9fb4d9;margin-left:10px}
.tag-mismatch-missing{background:rgba(255,139,150,.18);color:var(--tag-no);display:inline-block;padding:1px 7px;border-radius:3px;font-size:.75rem;font-weight:600;border:1px solid rgba(255,139,150,.35);white-space:nowrap}
"@

                $mismatchScript = @"
<script>
function applyMismatchFilters(){
    const tbody = document.getElementById('mismatchTableBody');
    if(!tbody){ return; }
    const mismatchTable = tbody.closest('table');
    const filters = Array.from(mismatchTable.querySelectorAll('.filter-row .col-filter')).map(i=>({col:Number(i.dataset.col), val:(i.value||'').toLowerCase()}));
    const rows = Array.from(tbody.querySelectorAll('tr'));
    let shown = 0;
    rows.forEach(row=>{
        const cells = Array.from(row.children).map(c=>(c.textContent||'').toLowerCase());
        let visible = true;
        for(const f of filters){ if(f.val && !(cells[f.col]||'').includes(f.val)){ visible = false; break; } }
        row.style.display = visible ? '' : 'none';
        if(visible){ shown++; }
    });
    const count = document.getElementById('mismatchRowCount');
    if(count){ count.textContent = shown + ' of ' + rows.length + ' shown'; }
}
document.addEventListener('DOMContentLoaded', function(){ applyMismatchFilters(); });
</script>
"@

        $htmlContent = Get-Content -Path $OutputHtmlPath -Raw -Encoding UTF8
        if ($htmlContent -match '</style>') {
            $htmlContent = $htmlContent -replace '</style>', ($mismatchCss + "`n</style>")
        } else {
            $htmlContent = $htmlContent -replace '</head>', ("<style>`n$mismatchCss`n</style>`n</head>")
        }
        if ($htmlContent -match '</body>') {
            $htmlContent = $htmlContent -replace '</body>', ($mismatchSection + "`n" + $mismatchScript + "`n</body>")
        } else {
            $htmlContent += "`n$mismatchSection`n$mismatchScript"
        }
        Set-Content -Path $OutputHtmlPath -Value $htmlContent -Encoding UTF8
    }

    $postCss = @"
.tag-status-new{background:rgba(16,124,16,.18);color:#107c10;display:inline-block;padding:1px 7px;border-radius:3px;font-size:.75rem;font-weight:600;border:1px solid rgba(16,124,16,.35)}
.tag-status-replaced{background:rgba(255,200,10,.22);color:#a08600;display:inline-block;padding:1px 7px;border-radius:3px;font-size:.75rem;font-weight:600;border:1px solid rgba(255,200,10,.50)}
.btn-jump-mismatch{padding:7px 14px;border:1px solid var(--border,#d0d0d0);border-radius:4px;background:var(--filter-bg,#f8f8f8);cursor:pointer;font-size:.85rem;text-decoration:none;color:var(--fg,#1a1a1a);display:inline-block}
.btn-jump-mismatch:hover{background:var(--hover,#f0f6ff)}
"@

    $postScript = @"
<script>
(function(){
    function removeColumnsByHeader(table){
        const headerRow = table.querySelector('thead tr');
        if(!headerRow){ return; }
        const headersToRemove = new Set(['gp setting name','remediation app']);
        const removeIdx = [];
        Array.from(headerRow.children).forEach((th, idx)=>{
            const text = (th.textContent || '').trim().toLowerCase();
            if(headersToRemove.has(text)){ removeIdx.push(idx); }
        });
        if(!removeIdx.length){ return; }
        removeIdx.sort((a,b)=>b-a);
        const rows = Array.from(table.querySelectorAll('tr'));
        rows.forEach(tr=>{
            removeIdx.forEach(i=>{
                if(tr.children[i]){ tr.removeChild(tr.children[i]); }
            });
        });
    }

    function applyStatusBadges(table){
        const headerRow = table.querySelector('thead tr');
        if(!headerRow){ return; }
        let statusIdx = -1;
        Array.from(headerRow.children).forEach((th, idx)=>{
            const text = (th.textContent || '').trim().toLowerCase();
            if(text === 'status'){ statusIdx = idx; }
        });
        if(statusIdx < 0){ return; }

        const rows = Array.from(table.querySelectorAll('tbody tr'));
        rows.forEach(tr=>{
            const td = tr.children[statusIdx];
            if(!td){ return; }
            const raw = (td.textContent || '').trim();
            if(/^new$/i.test(raw)){
                td.innerHTML = "<span class='tag-status-new'>New</span>";
            } else if(/^replaced by cis$/i.test(raw)){
                td.innerHTML = "<span class='tag-status-replaced'>Replaced by CIS</span>";
            }
        });
    }

    function applyBaselineMergerPostProcessing(){
        const tables = Array.from(document.querySelectorAll('table'));
        tables.forEach(removeColumnsByHeader);
        tables.forEach(applyStatusBadges);

        // Add jump button to mismatch section if it exists
        const mismatchEl = document.getElementById('mismatch-section');
        if(mismatchEl){
            const toolbar = document.querySelector('.toolbar');
            if(toolbar){
                const btn = document.createElement('a');
                btn.href = '#mismatch-section';
                btn.className = 'btn-jump-mismatch';
                btn.textContent = 'Not in new baseline';
                toolbar.appendChild(btn);
            }
        }
    }

    if(document.readyState === 'loading'){
        document.addEventListener('DOMContentLoaded', applyBaselineMergerPostProcessing);
    } else {
        applyBaselineMergerPostProcessing();
    }
})();
</script>
"@

    $htmlContent = Get-Content -Path $OutputHtmlPath -Raw -Encoding UTF8
    if ($htmlContent -match '</style>') {
        $htmlContent = $htmlContent -replace '</style>', ($postCss + "`n</style>")
    } else {
        $htmlContent = $htmlContent -replace '</head>', ("<style>`n$postCss`n</style>`n</head>")
    }
    if ($htmlContent -match '</body>') {
        $htmlContent = $htmlContent -replace '</body>', ($postScript + "`n</body>")
    } else {
        $htmlContent += "`n$postScript"
    }
    Set-Content -Path $OutputHtmlPath -Value $htmlContent -Encoding UTF8

    return $OutputHtmlPath

    $markdown = Get-Content -Path $InputMarkdownPath -Raw -Encoding UTF8
    $markdown = $markdown -replace "`0", ''
    $bodyHtml = $null
    $mdCmd    = Get-Command -Name ConvertFrom-Markdown -ErrorAction SilentlyContinue
    if ($mdCmd) { try { $converted = ConvertFrom-Markdown -InputObject $markdown; if ($converted -and $converted.Html) { $bodyHtml = $converted.Html } } catch { } }
    if (-not $bodyHtml) { $encoded = [System.Net.WebUtility]::HtmlEncode($markdown); $bodyHtml = "<pre class='markdown-fallback'>$encoded</pre>" }

    $sourceName  = [System.IO.Path]::GetFileName($InputMarkdownPath)
    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $mismatchSection = ''
    if ($Mismatches.Count -gt 0) {
        $mismatchRowsHtml = foreach ($m in ($Mismatches | Sort-Object @{ Expression = { $_['Section'] } }, @{ Expression = { $_['CISReference'] } })) {
            $followedHtml = switch ($m['OldFollowed']) { 'Yes' { '<span class="tag tag-yes">Yes</span>' }; 'No' { '<span class="tag tag-no">No</span>' }; default { [System.Net.WebUtility]::HtmlEncode($m['OldFollowed']) } }
            $section = if ($m['Section']) { [System.Net.WebUtility]::HtmlEncode($m['Section']) } else { '-' }
            $origin = if ($m['Origin']) { [System.Net.WebUtility]::HtmlEncode($m['Origin']) } else { '-' }
            $intunePolicy = if ($m['IntunePolicy']) { [System.Net.WebUtility]::HtmlEncode($m['IntunePolicy']) } else { '-' }
            "        <tr><td>$section</td><td>$([System.Net.WebUtility]::HtmlEncode($m['OldRec']))</td><td>$origin</td><td>$followedHtml</td><td>$([System.Net.WebUtility]::HtmlEncode($m['OldNotes']))</td><td>$intunePolicy</td><td><span class='tag-mismatch-missing'>Not in new baseline</span></td></tr>"
        }
        $mismatchSection = @"
<div id="mismatch-section" class="mismatch-section">
  <div class="mismatch-heading">Not in new baseline ($($Mismatches.Count))</div>
  <div class="table-wrap" style="margin-top:12px"><table><thead><tr><th>Section</th><th>Recommendation</th><th style="max-width:60px">Origin</th><th style="max-width:40px">Followed</th><th>Notes</th><th>Intune Policy</th><th style="min-width:130px">Status</th></tr></thead><tbody>
$($mismatchRowsHtml -join "`n")
  </tbody></table></div>
</div>
"@
    }

    $html = @"
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>$([System.Net.WebUtility]::HtmlEncode($ReportTitle))</title>
<style>:root{--bg:#fff;--fg:#1a1a1a;--accent:#0078D4;--hover:#f0f6ff;--border:#d0d0d0;--filter-bg:#f8f8f8;--tag-l1:#107c10;--tag-l2:#d83b01;--tag-yes:#107c10;--tag-no:#a80000}*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--fg);padding:24px;line-height:1.5}h1{font-size:1.6rem;margin-bottom:4px;color:var(--accent)}.subtitle{color:#666;font-size:.9rem;margin-bottom:16px}.stats{display:flex;gap:24px;margin-bottom:16px;flex-wrap:wrap}.stat{background:var(--filter-bg);border:1px solid var(--border);border-radius:6px;padding:10px 16px}.stat-value{font-size:1.4rem;font-weight:700;color:var(--accent)}.stat-label{font-size:.75rem;color:#666;text-transform:uppercase;letter-spacing:.5px}.toolbar{display:flex;gap:12px;margin-bottom:12px;align-items:center;flex-wrap:wrap}.toolbar input[type=text]{padding:7px 12px;border:1px solid var(--border);border-radius:4px;font-size:.85rem;width:260px}.toolbar button{padding:7px 14px;border:1px solid var(--border);border-radius:4px;background:var(--filter-bg);cursor:pointer;font-size:.85rem}.toolbar button:hover{background:var(--hover)}.count{font-size:.85rem;color:#666;margin-left:auto}.table-wrap{overflow-x:auto;border:1px solid var(--border);border-radius:6px}table{width:100%;border-collapse:collapse;font-size:.82rem}thead{position:sticky;top:0;z-index:2}th{background:var(--accent);color:#fff;padding:10px 8px;text-align:left;font-weight:600;white-space:nowrap;cursor:pointer;user-select:none}th:hover{background:#005a9e}td{padding:7px 8px;border-bottom:1px solid #eee;vertical-align:top}tr:hover td{background:var(--hover)}.tag{display:inline-block;padding:1px 7px;border-radius:3px;font-size:.75rem;font-weight:600}.tag-l1{background:#dff6dd;color:var(--tag-l1)}.tag-l2{background:#fed9cc;color:var(--tag-l2)}.tag-yes{background:#dff6dd;color:var(--tag-yes)}.tag-no{background:#fde7e9;color:var(--tag-no)}.tag-auto{background:#e6f0fb;color:#0050a0}.tag-manual{background:#fff4ce;color:#7a6400}.tag-twoday{background:#f0e6fb;color:#6600cc}.mismatch-section{margin-top:32px;padding:16px 20px;background:#fff8e6;border:2px solid #f0a500;border-radius:6px}.mismatch-heading{font-size:1.15rem;font-weight:700;color:#7a4800;margin-bottom:8px}.tag-mismatch-missing{background:#fde7e9;color:#a80000;display:inline-block;padding:1px 7px;border-radius:3px;font-size:.75rem;font-weight:600}a{color:var(--accent);text-decoration:none}a:hover{text-decoration:underline}#reportRoot{display:none}</style></head><body>
<h1>CIS Benchmark &mdash; $([System.Net.WebUtility]::HtmlEncode($ReportTitle))</h1>
<p class="subtitle">Generated $generatedAt &bull; Source: $([System.Net.WebUtility]::HtmlEncode($sourceName))</p>
<div class="stats"><div class="stat"><div class="stat-value" id="statRecommendations">0</div><div class="stat-label">Recommendations</div></div><div class="stat"><div class="stat-value" id="statL1">0</div><div class="stat-label">L1</div></div><div class="stat"><div class="stat-value" id="statL2">0</div><div class="stat-label">L2</div></div><div class="stat"><div class="stat-value" id="statFollowedNo">0</div><div class="stat-label">Not Followed</div></div></div>
<div class="toolbar"><input type="text" id="globalSearch" placeholder="Search..." oninput="applyFilters()"><button onclick="clearFilters()">Clear</button><button onclick="exportCSV()">Export CSV</button><span class="count" id="rowCount"></span></div>
<div class="table-wrap"><table id="cis-table"><thead id="tableHead"></thead><tbody id="tableBody"></tbody></table></div>
<article id="reportRoot">$bodyHtml</article>
<script>let rows=[],sortCol=-1,sortAsc=!0;function escapeHtml(v){return String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}function previousHeadingText(n){let c=n;while(c){c=c.previousElementSibling;if(!c)break;if(/^H[1-4]$/.test(c.tagName))return c.textContent.trim()}return'General'}function extractTableData(){const r=document.getElementById('reportRoot'),tables=Array.from(r.querySelectorAll('table'));let allRows=[],baseHeaders=[];tables.forEach(table=>{const s=previousHeadingText(table),hc=Array.from(table.querySelectorAll('thead th'));let h=hc.map(th=>th.textContent.trim());if(!h.length){const fr=table.querySelector('tr');if(fr)h=Array.from(fr.children).map(c=>c.textContent.trim())}if(!baseHeaders.length&&h.length)baseHeaders=h;const br=Array.from(table.querySelectorAll('tbody tr')),dr=br.length?br:Array.from(table.querySelectorAll('tr')).slice(1);dr.forEach(tr=>{const cells=Array.from(tr.children).map(td=>td.innerHTML.trim());if(!cells.length||cells.every(c=>c===''))return;allRows.push([s,...cells])})});return{headers:['Section',...baseHeaders],rows:allRows}}function renderTable(h,d){const thead=document.getElementById('tableHead'),tbody=document.getElementById('tableBody'),hRow=document.createElement('tr');h.forEach((hd,i)=>{const th=document.createElement('th');th.innerHTML=escapeHtml(hd)+'<span class="sort-arrow"></span>';th.onclick=()=>sortTable(i);hRow.appendChild(th)});const fRow=document.createElement('tr');fRow.className='filter-row';h.forEach((_,i)=>{const th=document.createElement('th');th.innerHTML='<input type=text data-col='+i+' placeholder=Filter... oninput=applyFilters()>';fRow.appendChild(th)});thead.appendChild(hRow);thead.appendChild(fRow);d.forEach(cells=>{const tr=document.createElement('tr');cells.forEach(html=>{const td=document.createElement('td');td.innerHTML=html;tr.appendChild(td)});tbody.appendChild(tr)});rows=Array.from(tbody.querySelectorAll('tr'))}function applyTags(){rows.forEach(row=>{row.querySelectorAll('td').forEach(cell=>{const v=cell.textContent.trim();if(v==='L1')cell.innerHTML='<span class=tag tag-l1>L1</span>';if(v==='L2')cell.innerHTML='<span class=tag tag-l2>L2</span>';if(v==='Yes')cell.innerHTML='<span class=tag tag-yes>Yes</span>';if(v==='No')cell.innerHTML='<span class=tag tag-no>No</span>';if(v==='Automated')cell.innerHTML='<span class=tag tag-auto>Automated</span>';if(v==='Manual')cell.innerHTML='<span class=tag tag-manual>Manual</span>';if(v.toLowerCase()==='twoday')cell.innerHTML='<span class=tag tag-twoday>twoday</span>'})})}function updateStats(){let l1=0,l2=0,no=0;const sections=new Set;rows.forEach(r=>{const c=r.querySelectorAll('td');if(c[0])sections.add(c[0].textContent.trim());c.forEach(cell=>{const v=cell.textContent.trim();if(v==='L1')l1++;if(v==='L2')l2++;if(v==='No')no++})});document.getElementById('statRecommendations').textContent=String(rows.length);document.getElementById('statL1').textContent=String(l1);document.getElementById('statL2').textContent=String(l2);document.getElementById('statFollowedNo').textContent=String(no);const vis=rows.filter(r=>r.style.display!=='none');document.getElementById('rowCount').textContent=vis.length+' of '+rows.length+' shown'}function applyFilters(){const t=document.getElementById('cis-table'),cf=Array.from(t.querySelectorAll('.filter-row input')).map(i=>({col:+i.dataset.col,val:i.value.toLowerCase()})),g=document.getElementById('globalSearch').value.toLowerCase();rows.forEach(r=>{const tx=Array.from(r.querySelectorAll('td')).map(c=>c.textContent.toLowerCase());let s=!0;if(g&&!tx.some(t=>t.includes(g)))s=!1;for(const f of cf)if(f.val&&!tx[f.col]?.includes(f.val)){s=!1;break}r.style.display=s?'':'none'});updateStats()}function clearFilters(){document.getElementById('globalSearch').value='';document.getElementById('cis-table').querySelectorAll('.filter-row input').forEach(i=>{i.value=''});applyFilters()}function sortTable(col){const ths=document.querySelectorAll('#cis-table thead tr:first-child th');sortAsc=sortCol===col?!sortAsc:!0;sortCol=col;ths.forEach((th,i)=>{th.classList.remove('sorted-asc','sorted-desc');if(i===col)th.classList.add(sortAsc?'sorted-asc':'sorted-desc')});const tbody=document.querySelector('#cis-table tbody');rows.sort((a,b)=>{const va=a.querySelectorAll('td')[col]?.textContent.trim()||'',vb=b.querySelectorAll('td')[col]?.textContent.trim()||'',na=parseFloat(va),nb=parseFloat(vb);if(!isNaN(na)&&!isNaN(nb))return sortAsc?na-nb:nb-na;return sortAsc?va.localeCompare(vb):vb.localeCompare(va)});rows.forEach(r=>tbody.appendChild(r));applyFilters()}function exportCSV(){const h=Array.from(document.querySelectorAll('#cis-table thead tr:first-child th')).map(th=>th.childNodes[0].textContent.trim()),cr=[h.map(h=>'"'+h.replace(/"/g,'""')+'"').join(',')];rows.forEach(r=>{if(r.style.display==='none')return;cr.push(Array.from(r.querySelectorAll('td')).map(td=>'"'+td.textContent.replace(/"/g,'""').trim()+'"').join(','))});const b=new Blob([cr.join('\n')],{type:'text/csv;charset=utf-8;'}),a=document.createElement('a');a.href=URL.createObjectURL(b);a.download='CIS-Benchmark.csv';a.click()}const extracted=extractTableData();renderTable(extracted.headers,extracted.rows);applyTags();applyFilters();</script>
$mismatchSection
</body></html>
"@
    $dir = Split-Path -Path $OutputHtmlPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    Set-Content -Path $OutputHtmlPath -Value $html -Encoding UTF8
    return $OutputHtmlPath
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    $stats = $null
    if ($ReportOnly) {
        if ([string]::IsNullOrWhiteSpace($InputMarkdownPath)) { throw '-InputMarkdownPath is required with -ReportOnly.' }
        $reportSource = $InputMarkdownPath
    } else {
        foreach ($requiredName in @('NewBaselinePath', 'OldBaselinePath', 'OutputPath')) {
            if ([string]::IsNullOrWhiteSpace((Get-Variable -Name $requiredName -ValueOnly))) {
                throw "-$requiredName is required for a merge."
            }
        }
        $stats = Invoke-MergeBaselines -NewBaselinePath $NewBaselinePath -OldBaselinePath $OldBaselinePath -OutputPath $OutputPath -ProfileFilter $ProfileFilter -SimilarityThreshold $SimilarityThreshold
        $reportSource = $OutputPath
    }
    if ($ReportOnly -or $GenerateReport) {
        if (-not $OutputHtmlPath) { $OutputHtmlPath = [System.IO.Path]::ChangeExtension($reportSource, '.html') }
        if (-not $ReportTitle) { $ReportTitle = [System.IO.Path]::GetFileNameWithoutExtension($reportSource) }
        $reportArguments = @{ InputMarkdownPath = $reportSource; OutputHtmlPath = $OutputHtmlPath; ReportTitle = $ReportTitle }
        if ($null -ne $stats) { $reportArguments.Mismatches = $stats.Mismatches }
        Invoke-GenerateHtmlReport @reportArguments | Out-Null
        if ($OpenReport) { Start-Process -FilePath $OutputHtmlPath }
    }
    [pscustomobject]@{
        OutputPath = if ($ReportOnly) { $null } else { [System.IO.Path]::GetFullPath($OutputPath) }
        HtmlPath = if ($ReportOnly -or $GenerateReport) { [System.IO.Path]::GetFullPath($OutputHtmlPath) } else { $null }
        Statistics = $stats
    }
}
