#Requires -Version 5.1
<#
.SYNOPSIS
    Runs the toolbox PDF conversion without loading its WPF interface.
.EXAMPLE
    .\Invoke-BaselinePdfConversion.ps1 -PdfPath .\benchmark.pdf -ProductName 'Windows 11' -PrepareForIntune
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$PdfPath,
    [Parameter(Mandatory = $true)][string]$ProductName,
    [string]$OutputFolder,
    [string]$IntunePrefix,
    [ValidateSet('L1','L2','Both')][string]$ProfileFilter = 'Both',
    [switch]$PrepareForIntune,
    [scriptblock]$LogCallback
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $PdfPath -PathType Leaf)) { throw "PDF not found: $PdfPath" }
$pdfPath = (Resolve-Path -LiteralPath $PdfPath).Path
if ([string]::IsNullOrWhiteSpace($ProductName)) { throw '-ProductName cannot be blank.' }
if (-not $OutputFolder) { $OutputFolder = Split-Path $pdfPath -Parent }
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
if ($PrepareForIntune -and -not $IntunePrefix) { $IntunePrefix = 'Prod-Win-All-SecurityBaseline-' + ($ProductName -replace '\s', '') + '-' }
if (-not $PrepareForIntune) { $IntunePrefix = '' }
$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCommand) { $pythonCommand = Get-Command py -ErrorAction SilentlyContinue }
if (-not $pythonCommand) { throw 'Python 3 not found. Install Python and PyPDF2.' }
$pythonExe = $pythonCommand.Source
$expectedHtmlPath = Join-Path $OutputFolder "$ProductName.html"


        function Write-Log { param([string]$M); if ($LogCallback) { & $LogCallback $M | Out-Null } else { Write-Verbose $M } }

        try {
            Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

            Write-Log "Extracting text from PDF using Python..."

            #region Build and run Python extraction script
            $pythonScript = @"
import PyPDF2
import json
import sys
import re
import unicodedata

pdf_path = r'$pdfPath'

reader = PyPDF2.PdfReader(pdf_path)
total_pages = len(reader.pages)

summary_start = None
summary_end = None

for i in range(total_pages - 1, -1, -1):
    text = reader.pages[i].extract_text()
    if text and 'Appendix: Summary Table' in text and 'CIS Benchmark Recommendation' in text:
        summary_start = i
        break

if summary_start is None:
    for i in range(total_pages - 1, -1, -1):
        text = reader.pages[i].extract_text()
        if text and 'Summary Table' in text and 'CIS Benchmark Recommendation' in text:
            summary_start = i
            break

if summary_start is None:
    print(json.dumps({"error": "Could not find Summary Table in PDF"}))
    sys.exit(1)

for i in range(summary_start + 1, total_pages):
    text = reader.pages[i].extract_text()
    if text and ('Appendix: CIS Controls' in text or 'Appendix: Change History' in text):
        # Include this page too; parser stops when it sees non-summary Appendix lines.
        summary_end = i + 1
        break

if summary_end is None:
    summary_end = min(summary_start + 20, total_pages)

summary_text = ''
for i in range(summary_start, summary_end):
    text = reader.pages[i].extract_text()
    if text:
        summary_text += text + '\n'

detail_texts = [reader.pages[page_idx].extract_text() or '' for page_idx in range(summary_start)]
detail_reference_ids = set(re.findall(r'(?m)^\s*(\d+(?:\.\d+)+)\s+(?=\(|Ensure\b|Enable\b|Disable\b|Require\b|Block\b|Configure\b|Set\b)', '\n'.join(detail_texts)))
lines = summary_text.split('\n')

_merged = []
_j = 0
while _j < len(lines):
    _s = lines[_j].strip()
    _next = lines[_j + 1].strip() if _j + 1 < len(lines) else ''
    if re.fullmatch(r'\d+(?:\.\d+)+', _s):
        joined_reference = re.match(r'^(\d+(?:\.\d+)*)\s+', _s + _next)
        if joined_reference and joined_reference.group(1) in detail_reference_ids:
            _merged.append(_s + _next)
            _j += 2
            continue
    if re.match(r'^\d+(?:\.\d+)*\.$', _s) and re.match(r'^\d', _next):
        _merged.append(_s + _next)
        _j += 2
        continue
    if re.match(r'^\d+(?:\.\d+)*$', _s) and re.match(r'^\.\d', _next):
        _merged.append(_s + _next)
        _j += 2
        continue
    # Handle wrapped rows where profile marker and recommendation text split across lines.
    if re.match(r'^\d+(?:\.\d+)*\s+\((?:E[35]\s+Level\s*[12]|E[35]\s+L[12]|L[12]|BL)\)$', _s, re.IGNORECASE) and _next:
        _merged.append(_s + ' ' + _next)
        _j += 2
        continue
    if re.match(r'^\d+(?:\.\d+)*\s+(?:E[35]\s+Level\s*[12]|E[35]\s+L[12]|L[12]|BL)$', _s, re.IGNORECASE) and _next:
        _merged.append(_s + ' ' + _next)
        _j += 2
        continue
    _merged.append(lines[_j])
    _j += 1
lines = _merged

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
    t = re.sub(r'(?<=[A-Za-z])\s*-\s*(?=[a-z])', '-', t)
    t = re.sub(r'(?<=[A-Za-z]) -(?=[A-Za-z])', '-', t)
    t = re.sub(r'(?<=\d) -(?=[A-Za-z\d])', '-', t)
    t = re.sub(r'(?<=[A-Za-z0-9%]) \\(?=[A-Za-z0-9%_])', r'\\', t)
    t = re.sub(r"(Ensure\s+'[^']+)\s+is set to\s+'", r"\1' is set to '", t)
    return t

def is_skip_line(line):
    return (line.startswith('Page ') or
            line.startswith('CIS Benchmark Recommendation') or
            line in ('Yes No', 'Set', 'Correctly', ''))

rec_intro_pattern = r'(?:Ensure|Enable|Disable|Require|Block|Configure|Set)\b'
profile_token_pattern = r'(?:E[35]\s+Level\s*[12]|E[35]\s+L[12]|L[12]|BL)'

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
            if hdr_check and '(L' not in next_line: break
            if is_skip_line(next_line) or next_line.startswith('CIS Benchmark'):
                i += 1
                continue
            if next_line.startswith('Appendix:'): break
            rest += ' ' + next_line
            i += 1
        type_match = re.search(r'\((Automated|Manual)\)', rest)
        rec_type = type_match.group(1) if type_match else 'Automated'
        rec_text = rest[:type_match.start()].strip() if type_match else rest
        rec_text = clean_text(rec_text)
        recommendations.append({'ref': ref, 'profile': profile, 'type': rec_type, 'text': rec_text})
        ordered_items.append({'type': 'rec', 'ref': ref})
        i += 1
        continue

    rec_match3 = re.match(r'^(\d+(?:\.\d+)*)\s+(' + profile_token_pattern + r')\s+(.+)', line, re.IGNORECASE)
    if rec_match3:
        ref = rec_match3.group(1)
        profile = normalize_profile(rec_match3.group(2))
        rest = rec_match3.group(3)
        while i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if re.match(r'^(\d+(?:\.\d+)*)\s+\(' + profile_token_pattern + r'\)', next_line, re.IGNORECASE): break
            if re.match(r'^(\d+(?:\.\d+)*)\s+' + profile_token_pattern + r'\s+', next_line, re.IGNORECASE): break
            if re.match(r'^(\d+(?:\.\d+)*)\s+' + rec_intro_pattern, next_line): break
            hdr_check = re.match(r'^(\d+(?:\.\d+)*)\s+([A-Z])', next_line)
            if hdr_check and '(L' not in next_line: break
            if is_skip_line(next_line) or next_line.startswith('CIS Benchmark'):
                i += 1
                continue
            if next_line.startswith('Appendix:'): break
            rest += ' ' + next_line
            i += 1
        type_match = re.search(r'\((Automated|Manual)\)', rest)
        rec_type = type_match.group(1) if type_match else 'Automated'
        rec_text = rest[:type_match.start()].strip() if type_match else rest
        rec_text = clean_text(rec_text)
        recommendations.append({'ref': ref, 'profile': profile, 'type': rec_type, 'text': rec_text})
        ordered_items.append({'type': 'rec', 'ref': ref})
        i += 1
        continue

    rec_match2 = re.match(r'^(\d+(?:\.\d+)*)\s+(' + rec_intro_pattern + r'\s.+)', line)
    if rec_match2:
        ref = rec_match2.group(1)
        rest = rec_match2.group(2)
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
        rec_type = type_match.group(1) if type_match else 'Automated'
        rec_text = rest[:type_match.start()].strip() if type_match else rest.strip()
        rec_text = clean_text(rec_text)
        recommendations.append({'ref': ref, 'profile': 'L1', 'type': rec_type, 'text': rec_text})
        ordered_items.append({'type': 'rec', 'ref': ref})
        i += 1
        continue

    hdr_match = re.match(r'^(\d+(?:\.\d+)*)\s+([A-Z].+?)$', line)
    if hdr_match and not re.search(r'\(' + profile_token_pattern + r'\)', line, re.IGNORECASE) and not re.match(rec_intro_pattern, hdr_match.group(2)):
        hdr_num = hdr_match.group(1)
        hdr_name = clean_text(hdr_match.group(2).strip())
        section_headers[hdr_num] = hdr_name
        ordered_items.append({'type': 'header', 'num': hdr_num, 'name': hdr_name})
        i += 1
        continue
    i += 1

def recover_summary_references(recommendations, ordered_items, detail_texts):
    heading_pattern = re.compile(
        r'^\s*(\d+(?:\.\d+)*)\s+(?:\((' + profile_token_pattern + r')\)\s+|(?:' + profile_token_pattern + r')\s+)?('
        + rec_intro_pattern + r'\s+.*?)\s*\((?:Automated|Manual)\)',
        re.MULTILINE | re.DOTALL | re.IGNORECASE
    )
    detail_by_text = {}
    for heading in heading_pattern.finditer('\n'.join(detail_texts)):
        key = re.sub(r'[^a-z0-9]', '', clean_text(heading.group(3)).lower())
        detail_by_text.setdefault(key, set()).add(heading.group(1))
    for rec in recommendations:
        if '.' in rec['ref']:
            continue
        key = re.sub(r'[^a-z0-9]', '', clean_text(rec['text']).lower())
        known_refs = detail_by_text.get(key, set())
        if rec['ref'] in known_refs:
            continue
        matches = [ref for ref in known_refs if ref.endswith(rec['ref'])]
        if len(matches) != 1:
            raise ValueError('Cannot verify truncated CIS reference ' + rec['ref'] + ': ' + rec['text'])
        rec['ref'] = matches[0]
    rec_items = (item for item in ordered_items if item['type'] == 'rec')
    for rec, item in zip(recommendations, rec_items):
        item['ref'] = rec['ref']

recover_summary_references(recommendations, ordered_items, detail_texts)

# Deduplicate recommendations by ref (keep first occurrence).
# The CIS PDF summary table sometimes lists the same recommendation twice.
_seen_refs = set()
recommendations = [_r for _r in recommendations if not (_r['ref'] in _seen_refs or _seen_refs.add(_r['ref']))]

_seen_rec_refs = set()
_deduped_items = []
for _item in ordered_items:
    if _item['type'] == 'rec':
        if _item['ref'] not in _seen_rec_refs:
            _seen_rec_refs.add(_item['ref'])
            _deduped_items.append(_item)
    else:
        _deduped_items.append(_item)
ordered_items = _deduped_items

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

for rec in recommendations:
    rec['parent_header'] = rec_to_header.get(rec['ref'])

rec_refs_set = set(r['ref'] for r in recommendations)
rec_references = {}
rec_safeguards = {}
rec_values = {}
rec_policy_groups = {}
rec_gp_setting_names = {}
rec_remediation_apps = {}
rec_texts_from_detail = {}
rec_profile_levels = {}

def extract_urls_from_refs(refs_text):
    for stop_marker in ['CIS Controls:', 'Additional Information:', 'CIS Critical']:
        stop_idx = refs_text.find(stop_marker)
        if stop_idx > 0:
            refs_text = refs_text[:stop_idx]
    urls = []
    for item in re.split(r'\n\d+\.\s+', '\n' + refs_text):
        item = item.strip()
        if not item or not item.startswith('http'):
            continue
        url = item.split('\n')[0]
        for rline in item.split('\n')[1:]:
            rline = rline.strip()
            if not rline or rline.startswith('GRID') or re.match(r'^\d+\.', rline):
                break
            url += rline
        url = re.sub(r'\s+', '', url)
        url = re.sub(r'[)\s,;]+$', '', url)
        if len(url) > 20:
            urls.append(url)
    return urls

def extract_cis_controls(controls_text):
    safeguards = []
    for version in ['v8', 'v7']:
        ver_match = re.search(r'\b' + version + r'\s+(\d+\.\d+)\s+', controls_text)
        if not ver_match:
            continue
        safeguard_num = ver_match.group(1)
        if safeguard_num == '0.0':
            continue
        ver_start = ver_match.start()
        next_ver = re.search(r'\b(?:v[0-9]+)\s+\d+\.\d+', controls_text[ver_match.end():])
        if next_ver:
            ver_block = controls_text[ver_start:ver_match.end() + next_ver.start()]
        else:
            ver_block = controls_text[ver_start:]
        dot_count = ver_block.count('\u25cf')
        if dot_count >= 3:
            igs = [1, 2, 3]
        elif dot_count == 2:
            igs = [2, 3]
        elif dot_count == 1:
            igs = [3]
        else:
            igs = []
        if igs:
            safeguards.append({'version': version, 'safeguard': safeguard_num, 'igs': igs})
    return safeguards

def extract_recommended_value(combined_text):
    lines_rv = combined_text.split('\n')
    for j, line in enumerate(lines_rv):
        if 'The recommended state' not in line:
            continue
        full = line.strip()
        for k in range(j + 1, min(j + 4, len(lines_rv))):
            nl = lines_rv[k].strip()
            if not nl or nl.startswith('NOTE') or nl.startswith('Rationale') or nl.startswith('Audit') or nl.startswith('Remediation'):
                break
            full += ' ' + nl
        nums = re.findall(r'\((\d+)\)', full)
        if len(nums) == 1:
            return nums[0]
        elif len(nums) > 1:
            if re.search(r'Enabled\s+\(\d+\)\s+or\s+Disabled\s+\(\d+\)', full):
                return None
            return nums[-1]
    return None

def extract_remediation_path(combined_text):
    rem_match = re.search(r'\bRemediation\s*:', combined_text, re.IGNORECASE)
    if not rem_match:
        return None
    rem_text = combined_text[rem_match.end():]
    rem_text = re.split(
        r'(?im)^\s*(?:(?:Default\s+Value|References|CIS Controls|Additional Information|Audit|Impact|Rationale|Note)\s*:|\d+(?:\.\d+)+\s+\(L[12]\))',
        rem_text, maxsplit=1
    )[0]
    path_lines = []
    for line in rem_text.splitlines():
        stripped = line.strip()
        if re.match(r'^(?:Page\s+\d+|\d+\s*\|\s*Page|\d+)$', stripped, re.IGNORECASE):
            continue
        if not path_lines:
            root = re.match(
                r'^(?:(?:Computer|User)\s+Configuration\b|Microsoft\s+.+?\b\d{4}\b|Administrative Templates\b)',
                stripped, re.IGNORECASE
            )
            if root or ('\\' in stripped and not re.search(r'(?:HKEY_|HKLM|HKCU|https?:|:\\)', stripped, re.IGNORECASE)):
                path_lines.append(stripped)
        elif not stripped:
            break
        else:
            path_lines.append(stripped)
    if not path_lines:
        return None
    path = re.sub(r'\s+', ' ', ' '.join(path_lines)).strip()
    parts = [part.strip() for part in path.split('\\')]
    if len(parts) < 2 or not all(parts):
        return None
    return '\\'.join(parts)

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
                    path_parts.extend(sub_parts)
            else:
                break
    path_parts = [p for p in path_parts if p and len(p) > 2]
    if len(path_parts) >= 2:
        return (path_parts[-2], path_parts[-1])
    if len(path_parts) == 1:
        return (None, path_parts[0])
    # Fallback: look for M365Apps-style paths e.g. "Microsoft Excel 2016\Excel Options\Security\..."
    for line in rem_text.split('\n'):
        stripped = line.strip()
        if not stripped:
            continue
        if re.match(r'^Microsoft\s+\w.*\d{4}', stripped, re.IGNORECASE) and '\\' in stripped:
            parts = [p.strip() for p in re.split(r'\s*\\\s*', stripped) if p.strip()]
            if len(parts) >= 2:
                return (None, parts[-1], parts[0])
            if len(parts) == 1:
                return (None, parts[0], parts[0])
    return (None, None)

def extract_profile_level_from_applicability(combined_text):
    # Parse profile from the detailed recommendation section.
    # Keep E3/E5 when present and only shorten "Level X" to "LX".
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

def normalize_token_text(value):
    if value is None:
        return ''
    cleaned = re.sub(r'[^A-Za-z0-9\s]', ' ', str(value))
    cleaned = re.sub(r'\s+', ' ', cleaned).strip().lower()
    return cleaned

def build_recommendation_blocks(page_text, rec_refs):
    ref_spans = []
    seen = set()
    for ref in rec_refs:
        if ref in seen:
            continue
        seen.add(ref)
        ref_match = re.search(r'(?m)^\s*' + re.escape(ref) + r'\s+\((' + profile_token_pattern + r')\)\s+' + rec_intro_pattern, page_text, re.IGNORECASE)
        if not ref_match:
            ref_match = re.search(r'(?m)^\s*' + re.escape(ref) + r'\s+(' + profile_token_pattern + r')\s+' + rec_intro_pattern, page_text, re.IGNORECASE)
        if not ref_match:
            ref_match = re.search(r'(?m)^\s*' + re.escape(ref) + r'\s+' + rec_intro_pattern, page_text, re.IGNORECASE)
        if ref_match:
            ref_spans.append((ref, ref_match.start()))
    ref_spans.sort(key=lambda item: item[1])
    blocks = {}
    for idx, (ref, start) in enumerate(ref_spans):
        end = ref_spans[idx + 1][1] if idx + 1 < len(ref_spans) else len(page_text)
        blocks[ref] = page_text[start:end]
    return blocks

def fallback_policy_group_from_header(ref):
    # Walk up the ref's own numeric hierarchy to find the closest real section header.
    # e.g. 4.1.3.1 -> try 4.1.3, then 4.1, then 4.
    # This is immune to rec_to_header stack-order bugs (wrong "last seen" fallback).
    parts = ref.split('.')
    for length in range(len(parts) - 1, 0, -1):
        candidate = '.'.join(parts[:length])
        hdr_name = section_headers.get(candidate)
        if hdr_name:
            hdr_name = re.sub(r'^\d+(?:\.\d+)*\s*', '', hdr_name).strip()
            return hdr_name if hdr_name else None
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
    rec_blocks = build_recommendation_blocks(text, page_recs)
    refs_found = 'References:' in text
    controls_found = 'CIS Controls:' in text
    if not refs_found or not controls_found:
        for lookahead in range(1, 3):
            if page_idx + lookahead < len(detail_texts):
                next_text = detail_texts[page_idx + lookahead]
                if re.search(r'\d+(?:\.\d+)*\s+\(' + profile_token_pattern + r'\)\s+' + rec_intro_pattern, next_text, re.IGNORECASE):
                    new_rec_pos = re.search(r'\d+(?:\.\d+)*\s+\(' + profile_token_pattern + r'\)\s+' + rec_intro_pattern, next_text, re.IGNORECASE).start()
                    before_new_rec = next_text[:new_rec_pos]
                    combined_text += '\n' + before_new_rec
                    refs_found = refs_found or 'References:' in before_new_rec
                    controls_found = controls_found or 'CIS Controls:' in before_new_rec
                    break
                else:
                    combined_text += '\n' + next_text
                    refs_found = refs_found or 'References:' in next_text
                    controls_found = controls_found or 'CIS Controls:' in next_text
                    if refs_found and controls_found:
                        break

    for ref in page_recs:
        segment = rec_blocks.get(ref, combined_text)
        if ref in rec_refs_set and ref not in rec_texts_from_detail:
            m_detail = re.search(
                r'\b' + re.escape(ref) + r'\s+\(' + profile_token_pattern + r'\)\s+((?:Ensure|Enable|Disable|Require|Block|Configure|Set)\s.*?)\s*\((Automated|Manual)\)',
                segment, re.DOTALL
            )
            if not m_detail:
                m_detail = re.search(
                    r'\b' + re.escape(ref) + r'\s+((?:Ensure|Enable|Disable|Require|Block|Configure|Set)\s.*?)\s*\((Automated|Manual)\)',
                    segment, re.DOTALL
                )
            if not m_detail and segment is not text:
                m_detail = re.search(
                    r'\b' + re.escape(ref) + r'\s+\(' + profile_token_pattern + r'\)\s+((?:Ensure|Enable|Disable|Require|Block|Configure|Set)\s.*?)\s*\((Automated|Manual)\)',
                    text, re.DOTALL
                )
            if not m_detail and segment is not text:
                m_detail = re.search(
                    r'\b' + re.escape(ref) + r'\s+((?:Ensure|Enable|Disable|Require|Block|Configure|Set)\s.*?)\s*\((Automated|Manual)\)',
                    text, re.DOTALL
                )
            if m_detail:
                raw = m_detail.group(1)
                raw = re.sub(r'(\w)\s*-\n\s*(\w)', r'\1-\2', raw)
                raw = re.sub(r'\s+', ' ', raw).strip()
                rec_texts_from_detail[ref] = clean_text(raw)

    for ref in page_recs:
        if ref not in rec_refs_set:
            continue
        segment = rec_blocks.get(ref, combined_text)
        if 'The recommended state' in segment and ref not in rec_values:
            val = extract_recommended_value(segment)
            if val is None and segment is not combined_text:
                val = extract_recommended_value(combined_text)
            if val is not None:
                rec_values[ref] = val
        if refs_found and ref not in rec_references:
            refs_idx = segment.rfind('References:')
            if refs_idx >= 0:
                cleaned_urls = extract_urls_from_refs(segment[refs_idx:])
                if not cleaned_urls and segment is not combined_text:
                    refs_idx = combined_text.rfind('References:')
                    if refs_idx >= 0:
                        cleaned_urls = extract_urls_from_refs(combined_text[refs_idx:])
                if cleaned_urls:
                    rec_references[ref] = cleaned_urls
        if controls_found and ref not in rec_safeguards:
            controls_idx = segment.rfind('CIS Controls:')
            if controls_idx >= 0:
                sgs = extract_cis_controls(segment[controls_idx:])
                if not sgs and segment is not combined_text:
                    controls_idx = combined_text.rfind('CIS Controls:')
                    if controls_idx >= 0:
                        sgs = extract_cis_controls(combined_text[controls_idx:])
                if sgs:
                    rec_safeguards[ref] = sgs
        if ref not in rec_policy_groups:
            gp_result = extract_policy_group(segment)
            parsed_group = gp_result[0] if (gp_result and len(gp_result) > 0 and gp_result[0]) else None
            header_group = fallback_policy_group_from_header(ref)

            final_group = parsed_group
            if header_group and not final_group:
                final_group = header_group
            elif header_group and final_group:
                parsed_norm = normalize_token_text(final_group)
                header_norm = normalize_token_text(header_group)
                if parsed_norm and header_norm and parsed_norm != header_norm:
                    # Prefer section context when remediation parsing drifts into a nearby recommendation.
                    final_group = header_group

            if final_group:
                rec_policy_groups[ref] = final_group
            if gp_result and len(gp_result) > 2 and gp_result[2]:
                rec_remediation_apps[ref] = gp_result[2]
        if ref not in rec_gp_setting_names:
            path_segment = segment
            if ref == page_recs[-1] and combined_text.startswith(text):
                path_segment += combined_text[len(text):]
            remediation_path = extract_remediation_path(path_segment)
            if remediation_path:
                rec_gp_setting_names[ref] = remediation_path
        if ref not in rec_profile_levels:
            derived_profile = extract_profile_level_from_applicability(segment)
            if not derived_profile and segment is not combined_text:
                derived_profile = extract_profile_level_from_applicability(combined_text)
            if derived_profile:
                rec_profile_levels[ref] = derived_profile

for rec in recommendations:
    summary_profile = normalize_profile(rec.get('profile'))
    derived_profile = rec_profile_levels.get(rec['ref'])
    rec['summary_profile'] = summary_profile
    rec['profile_from_applicability'] = derived_profile
    rec['profile_conflict'] = False
    rec['profile_source'] = 'SummaryTable'
    if derived_profile:
        if summary_profile and summary_profile != derived_profile:
            rec['profile_conflict'] = True
        rec['profile'] = derived_profile
        rec['profile_source'] = 'ProfileApplicability'
    rec['references'] = rec_references.get(rec['ref'], [])
    rec['safeguards'] = rec_safeguards.get(rec['ref'], [])
    rec['rec_value'] = rec_values.get(rec['ref'], None)
    rec['policy_group'] = rec_policy_groups.get(rec['ref'], None)
    rec['gp_setting_name'] = rec_gp_setting_names.get(rec['ref'], None)
    rec['remediation_app'] = rec_remediation_apps.get(rec['ref'], None)
    if rec['ref'] in rec_texts_from_detail:
        rec['text'] = rec_texts_from_detail[rec['ref']]

output = {
    "recommendations": [r for r in recommendations],
    "section_headers": section_headers,
    "page_count": total_pages,
    "summary_start": summary_start,
    "summary_end": summary_end
}

print(json.dumps(output, ensure_ascii=True))
"@

            $tempPyFile = Join-Path $env:TEMP "cis_extract_$([guid]::NewGuid().ToString('N').Substring(0,8)).py"
            $pythonScript | Set-Content -Path $tempPyFile -Encoding UTF8 -NoNewline

            $proc = [System.Diagnostics.Process]::new()
            $proc.StartInfo.FileName               = $pythonExe
            $proc.StartInfo.Arguments              = "`"$tempPyFile`""
            $proc.StartInfo.UseShellExecute        = $false
            $proc.StartInfo.RedirectStandardOutput = $true
            $proc.StartInfo.RedirectStandardError  = $true
            $proc.StartInfo.CreateNoWindow         = $true
            $proc.StartInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $proc.StartInfo.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
            $proc.Start() | Out-Null
            $rawOutput = $proc.StandardOutput.ReadToEnd()
            $rawError  = $proc.StandardError.ReadToEnd()
            $proc.WaitForExit()
            Remove-Item $tempPyFile -Force -ErrorAction SilentlyContinue

            if ($proc.ExitCode -ne 0) {
                Write-Log "Python extraction failed (exit $($proc.ExitCode))."
                if ($rawError) { Write-Log $rawError }
                throw "Python extraction failed."
            }
            #endregion

            #region Parse JSON output
            $extracted = $rawOutput | ConvertFrom-Json
            if ($extracted.error) { throw "PDF extraction error: $($extracted.error)" }

            $recommendations = $extracted.recommendations
            $sectionHeaders  = @{}
            foreach ($prop in $extracted.section_headers.PSObject.Properties) {
                $sectionHeaders[$prop.Name] = $prop.Value
            }
            $pdfFileName = Split-Path $pdfPath -Leaf
            $isIosBenchmark = ($pdfFileName -match '(?i)(?:^|[^a-z])(ios|ipados)(?:[^a-z]|$)')
            $profileDerivedCount = @($recommendations | Where-Object { $_.profile_source -eq 'ProfileApplicability' }).Count
            $profileFallbackCount = @($recommendations | Where-Object { $_.profile_source -eq 'SummaryTable' }).Count
            $profileConflictCount = @($recommendations | Where-Object { $_.profile_conflict -eq $true }).Count
            Write-Log "Found $($recommendations.Count) recommendations across $($sectionHeaders.Count) sections"
            Write-Log "Summary table on pages $($extracted.summary_start + 1) to $($extracted.summary_end)"
            Write-Log "Profile source: applicability-derived=$profileDerivedCount, summary-fallback=$profileFallbackCount, conflicts=$profileConflictCount"
            if ($isIosBenchmark) { Write-Log "Detected iOS/iPadOS benchmark filename: enabling supervised/unsupervised Intune policy naming." }
            #endregion

            #region Helper functions
            function Get-IntunePolicySuffix {
                param([string]$SectionKey, [string]$SectionName)
                $suffix = $SectionName -replace '\([^)]*\)', '' -replace '\s+', ' '
                $suffix = $suffix -replace '[^A-Za-z0-9\s]', '' -replace '\s+', ' '
                $suffix = ($suffix.Trim().Split(' ') | ForEach-Object {
                    if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) }
                }) -join ''
                return $suffix
            }

            function Get-SectionHeading {
                param([string]$SectionKey, [string]$DefaultName)
                $name = ($DefaultName -replace '\s+', ' ' -replace '\s+-', '-' -replace '-\s+', '-').Trim()
                $name = [regex]::Replace($name, '\b\w', { param($m) $m.Value.ToUpper() })
                $words = $name.Split(' ')
                $smallWords = @('Of','Or','And','In','On','For','The','A','An','To')
                for ($idx = 1; $idx -lt $words.Count; $idx++) {
                    if ($words[$idx] -in $smallWords) { $words[$idx] = $words[$idx].ToLower() }
                }
                $name = ($words -join ' ') -replace '\bHttp\b','HTTP' -replace '\bApi\b','API' -replace '\bGpo\b','GPO' `
                    -replace '\bDns\b','DNS' -replace '\bTls\b','TLS' -replace '\bSsl\b','SSL' `
                    -replace '\bUrl\b','URL' -replace '\bUsb\b','USB' -replace '\bNtlm\b','NTLM' `
                    -replace '\bPdf\b','PDF' -replace '\bEsr\b','ESR' -replace '\bIdn\b','IDN'
                return $name
            }

            function Get-PolicyNameFromGroup {
                param([string]$PolicyGroup, [string]$IntunePrefix, [string]$Fallback)
                if ($PolicyGroup) {
                    $pg = $PolicyGroup -replace '[^A-Za-z0-9\s]','' -replace '\s+',' '
                    $pgPascal = ($pg.Trim().Split(' ') | ForEach-Object {
                        if ($_.Length -gt 0) { $_.Substring(0,1).ToUpper() + $_.Substring(1) }
                    }) -join ''
                    return "$IntunePrefix$pgPascal"
                }
                return $Fallback
            }

            function Get-ParentKey {
                param([string]$Key)
                $parts = $Key.Split('.')
                if ($parts.Count -le 1) { return $null }
                return ($parts[0..($parts.Count - 2)]) -join '.'
            }

            function Get-IosPolicyPrefix {
                param([string]$SectionKey, [hashtable]$SectionHeaderLookup)
                if (-not $isIosBenchmark) { return $null }
                $scanKey = $SectionKey
                while ($scanKey) {
                    if ($SectionHeaderLookup.ContainsKey($scanKey)) {
                        $heading = [string]$SectionHeaderLookup[$scanKey]
                        if ($heading -match '(?i)Recommendations\s+for\s+Unsupervised\s*\(BYOD\)\s*Devices') {
                            return 'Prod-iOS-Unsupervised-SecurityBaseline-'
                        }
                        if ($heading -match '(?i)Recommendations\s+for\s+Supervised\s*\(Organization\)\s*Devices') {
                            return 'Prod-iOS-Supervised-SecurityBaseline-'
                        }
                    }
                    $scanKey = Get-ParentKey -Key $scanKey
                }
                return $null
            }

            function Build-PaddedTable {
                param([array]$Rows, [string]$IntunePrefix, [string]$FallbackPolicyName, [bool]$PrepareForIntune, [bool]$UseSectionPolicyName = $false)
                $headers = @(
                    "**Recommendation**","**twoday ID**","**CIS Reference**","**Profile**",
                    "**CIS Safeguard**","**Type**","**Origin**","**Followed**","**Notes**",
                    "**Link**","**References**"
                )
                if ($PrepareForIntune) {
                    $headers = @(
                        "**Recommendation**","**twoday ID**","**CIS Reference**","**Profile**",
                        "**CIS Safeguard**","**Type**","**Origin**","**Followed**","**Notes**",
                        "**Intune Policy**","**Link**","**References**","**GP Setting Name**"
                    )
                }
                $dataRows = @()
                foreach ($rec in $Rows) {
                    $followed = if ($rec.profile -match '(^|\s)L1$' -or $rec.profile -eq 'BL') { 'Yes' } else { 'No' }
                    $safeguardValue = '-'
                    if ($rec.safeguards -and $rec.safeguards.Count -gt 0) {
                        $sgParts = @()
                        foreach ($sg in $rec.safeguards) {
                            $igStr = ($sg.igs | ForEach-Object { "IG$_" }) -join '+'
                            $sgParts += "$($sg.version): $($sg.safeguard) / $igStr"
                        }
                        $safeguardValue = $sgParts -join '<br>'
                    }
                    $refsValue = '-'
                    if ($rec.references -and $rec.references.Count -gt 0) {
                        $refLinks = @()
                        foreach ($url in $rec.references) { $refLinks += "[References]($url)" }
                        $refsValue = $refLinks -join '<br>'
                    }
                    $recText = $rec.text
                    if ($rec.rec_value) {
                        $lastQuote = $recText.LastIndexOf("'")
                        if ($lastQuote -gt 0) {
                            $recText = $recText.Substring(0, $lastQuote) + " ($($rec.rec_value))" + $recText.Substring($lastQuote)
                        } else { $recText = "$recText ($($rec.rec_value))" }
                    }
                    $rowCells = @($recText, '-', $rec.ref, $rec.profile, $safeguardValue, $rec.type, 'CIS', $followed, '-')
                    if ($PrepareForIntune) {
                        if ($UseSectionPolicyName) {
                            $policyForRow = $FallbackPolicyName
                        } else {
                            $policyForRow = Get-PolicyNameFromGroup -PolicyGroup $rec.policy_group -IntunePrefix $IntunePrefix -Fallback $FallbackPolicyName
                        }
                        $rowCells += @(
                            $policyForRow,
                            '-',
                            $refsValue,
                            $(if ($rec.gp_setting_name) { $rec.gp_setting_name.Replace('|', '&#124;') } else { '-' })
                        )
                    } else {
                        $rowCells += @('-', $refsValue)
                    }
                    $dataRows += ,$rowCells
                }
                $colCount = $headers.Count
                $colWidths = @()
                for ($c = 0; $c -lt $colCount; $c++) {
                    $maxW = $headers[$c].Length
                    foreach ($row in $dataRows) { if ($row[$c].Length -gt $maxW) { $maxW = $row[$c].Length } }
                    $colWidths += $maxW
                }
                $headerCells = @(); for ($c = 0; $c -lt $colCount; $c++) { $headerCells += " " + $headers[$c].PadRight($colWidths[$c]) + " " }
                $headerLine = "|" + ($headerCells -join "|") + "|"
                $sepCells = @(); for ($c = 0; $c -lt $colCount; $c++) { $sepCells += "-" * ($colWidths[$c] + 2) }
                $sepLine = "|" + ($sepCells -join "|") + "|"
                $tableLines = @($headerLine, $sepLine)
                foreach ($row in $dataRows) {
                    $cells = @(); for ($c = 0; $c -lt $colCount; $c++) { $cells += " " + $row[$c].PadRight($colWidths[$c]) + " " }
                    $tableLines += "|" + ($cells -join "|") + "|"
                }
                return $tableLines
            }
            #endregion

            #region Apply profile filter and group
            if ($profileFilter -ne 'Both') {
                $recommendations = @(
                    $recommendations | Where-Object {
                        $profileText = [string]$_.profile
                        switch ($profileFilter) {
                            'L2' { $profileText -match '(^|\s)L2$' }
                            default { $profileText -match '(^|\s)L1$' -or $profileText -eq 'BL' }
                        }
                    }
                )
                Write-Log "Profile filter: $profileFilter ($($recommendations.Count) retained)"
            }

            $groupedRecs       = [ordered]@{}
            $recsWithoutHeader = @()
            foreach ($rec in $recommendations) {
                $hk = $rec.parent_header
                if ($hk) {
                    if (-not $groupedRecs.Contains($hk)) { $groupedRecs[$hk] = @() }
                    $groupedRecs[$hk] += $rec
                } else { $recsWithoutHeader += $rec }
            }

            $groupingKeys    = @($groupedRecs.Keys)
            $majorHeadings   = [ordered]@{}
            $processedMajors = @{}

            foreach ($gKey in $groupingKeys) {
                $parent = Get-ParentKey -Key $gKey
                while ($parent) {
                    if ($sectionHeaders.ContainsKey($parent) -and -not $processedMajors.ContainsKey($parent)) {
                        $majorHeadings[$parent]   = $sectionHeaders[$parent]
                        $processedMajors[$parent] = $true
                    }
                    $parent = Get-ParentKey -Key $parent
                }
            }
            foreach ($key in $sectionHeaders.Keys) {
                if (($key -split '\.').Count -eq 1 -and -not $majorHeadings.Contains($key)) {
                    $majorHeadings[$key] = $sectionHeaders[$key]
                }
            }
            if ($majorHeadings.Count -eq 0) {
                foreach ($key in ($sectionHeaders.Keys | Sort-Object { ($_ -split '\.').Count })) {
                    if (($key -split '\.').Count -eq 1) { $majorHeadings[$key] = $sectionHeaders[$key] }
                }
            }
            #endregion

            #region Build Markdown
            Write-Log "Building Markdown..."
            $mdLines = @("[[_TOC_]]", "", "# 1. $productName")

            $sortedGroupKeys = $groupingKeys | Sort-Object {
                $parts = $_ -split '\.'
                $nk = ($parts | ForEach-Object { $_.PadLeft(5,'0') }) -join '.'
                if ($majorHeadings.Contains($_)) { $nk + '.zzzzz' } else { $nk }
            }

            $emittedMajors = @{}
            foreach ($gKey in $sortedGroupKeys) {
                $recs = $groupedRecs[$gKey]
                if ($recs.Count -eq 0) { continue }

                if ($majorHeadings.Contains($gKey) -and -not $emittedMajors.ContainsKey($gKey)) {
                    $majorName = Get-SectionHeading -SectionKey $gKey -DefaultName $majorHeadings[$gKey]
                    $mdLines += ""; $mdLines += "## $gKey $majorName"
                    $emittedMajors[$gKey] = $true
                }

                $pendingMajors = @()
                $parent = Get-ParentKey -Key $gKey
                while ($parent) {
                    if ($majorHeadings.Contains($parent) -and -not $emittedMajors.ContainsKey($parent)) { $pendingMajors += $parent }
                    $parent = Get-ParentKey -Key $parent
                }
                [array]::Reverse($pendingMajors)
                foreach ($mKey in $pendingMajors) {
                    $majorName = Get-SectionHeading -SectionKey $mKey -DefaultName $majorHeadings[$mKey]
                    $mdLines += ""; $mdLines += "## $mKey $majorName"
                    $emittedMajors[$mKey] = $true
                }

                $subName    = if ($sectionHeaders.ContainsKey($gKey)) { Get-SectionHeading -SectionKey $gKey -DefaultName $sectionHeaders[$gKey] } else { "Section $gKey" }
                $polSuffix  = if ($sectionHeaders.ContainsKey($gKey)) { Get-IntunePolicySuffix -SectionKey $gKey -SectionName $sectionHeaders[$gKey] } else { $gKey -replace '\.', '' }
                $iosPolicyPrefix = Get-IosPolicyPrefix -SectionKey $gKey -SectionHeaderLookup $sectionHeaders
                $effectivePrefix = if ($iosPolicyPrefix) { $iosPolicyPrefix } else { $intunePrefix }
                $policyName = "$effectivePrefix$polSuffix"
                $useSectionPolicyName = [bool]$iosPolicyPrefix

                $firstRef = $recs[0].ref; $lastRef = $recs[-1].ref
                if ($majorHeadings.Contains($gKey)) {
                    $heading = if ($firstRef -ne $lastRef) { "$firstRef - $lastRef $subName" } else { "$firstRef $subName" }
                } elseif ($sectionHeaders.ContainsKey($gKey)) {
                    $heading = "$gKey $subName"
                } else {
                    $heading = if ($firstRef -ne $lastRef) { "$firstRef - $lastRef" } else { "$firstRef" }
                }

                $mdLines += ""; $mdLines += "### $heading"
                $mdLines += Build-PaddedTable -Rows $recs -IntunePrefix $intunePrefix -FallbackPolicyName $policyName -PrepareForIntune $prepareForIntune -UseSectionPolicyName $useSectionPolicyName
            }

            if ($recsWithoutHeader.Count -gt 0) {
                $mdLines += ""; $mdLines += "### Other Recommendations"
                $mdLines += Build-PaddedTable -Rows $recsWithoutHeader -IntunePrefix $intunePrefix -FallbackPolicyName "${intunePrefix}Other" -PrepareForIntune $prepareForIntune
            }

            $mdContent = $mdLines -join "`n"
            #endregion

            #region Write Markdown/Txt files
            $mdPath  = Join-Path $outputFolder "$productName.md"
            $txtPath = Join-Path $outputFolder "$productName.txt"
            [System.IO.File]::WriteAllText($mdPath,  $mdContent, [System.Text.UTF8Encoding]::new($false))
            [System.IO.File]::WriteAllText($txtPath, $mdContent, [System.Text.UTF8Encoding]::new($false))
            Write-Log "Created: $mdPath"
            Write-Log "Created: $txtPath"
            #endregion

            #region Build HTML
            Write-Log "Building HTML report..."
            $converterScriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'Convert-MarkdownTextToInteractiveHtmlReport.ps1'
            if (-not (Test-Path -Path $converterScriptPath -PathType Leaf)) {
                throw "Converter script not found: $converterScriptPath"
            }

            $htmlPath = $expectedHtmlPath
            $converterOutput = & $converterScriptPath -InputFile $mdPath -OutputFile $htmlPath -Title $productName 2>&1
            foreach ($line in @($converterOutput)) {
                if ($null -ne $line) { Write-Log ([string]$line) }
            }
            if (-not (Test-Path -Path $htmlPath -PathType Leaf)) {
                throw 'Converter completed but output HTML was not found.'
            }
            Write-Log "Created: $htmlPath"
            #endregion

            Write-Log ""
            Write-Log "Summary:"
            Write-Log "  Product:         $productName"
            Write-Log "  Recommendations: $($recommendations.Count)"
            Write-Log "  Major sections:  $($majorHeadings.Count)"
            if ($prepareForIntune) { Write-Log "  Intune prefix:   $intunePrefix" }
            else { Write-Log "  Intune prep:     disabled" }
            if ($isIosBenchmark) { Write-Log "  iOS mode:        enabled (supervised/unsupervised section prefixes)" }
            Write-Log ""
            Write-Log "=== Conversion completed ==="

            return [pscustomobject]@{
                MarkdownPath = $mdPath
                TextPath = $txtPath
                HtmlPath = $htmlPath
                RecommendationCount = $recommendations.Count
            }
        } catch {
            Write-Log "ERROR: $($_.Exception.Message)"
            throw
        }