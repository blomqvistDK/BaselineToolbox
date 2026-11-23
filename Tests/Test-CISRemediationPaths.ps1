#Requires -Version 5.1
[CmdletBinding()]
param([string]$ScriptPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) {
    $ScriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Features\Invoke-BaselinePdfConversion.ps1'
}
Import-Module (Join-Path (Split-Path (Split-Path $ScriptPath -Parent) -Parent) 'Shared\BaselineToolkit.Common.psm1') -Force -DisableNameChecking
$source = Get-Content -LiteralPath $ScriptPath -Raw
$helper = [regex]::Match($source, '(?ms)^def extract_remediation_path\(.*?(?=^def extract_policy_group)').Value
if (-not $helper) { throw 'Embedded remediation path helper not found.' }
$embedded = [regex]::Match($source, '(?ms)^\s*\$pythonScript = @"\r?\n(.*?)^"@').Groups[1].Value
if (-not $embedded) { throw 'Embedded Python converter not found.' }
$embeddedJson = ConvertTo-Json -InputObject $embedded -Compress
$checks = @'
import ast
import unicodedata

expected = r'Microsoft PowerPoint 2016\PowerPoint Options\Security\Trust Center\Trusted Locations\Allow Trusted Locations on the network'
sample = r'''Audit:
HKEY_USERS\USER SID\Software\Policies\Microsoft\office\16.0\powerpoint\security\trusted locations:allownetworklocations
Remediation:
To establish the recommended state via configuration profiles, set the following Settings
Catalog path to Disabled.
Microsoft PowerPoint 2016\PowerPoint Options\Security\Trust Center\Trusted
Locations\Allow Trusted Locations on the network
Default Value:
Enabled'''
assert extract_remediation_path(sample) == expected
for wrapped in (
    expected.replace('Trusted Locations', 'Trusted\nLocations'),
    expected.replace('\\Trust Center', '\\\nTrust Center'),
    expected.replace('on the network', 'on the\nnetwork'),
    expected.replace('Trusted Locations', 'Trusted\nPage 42\nLocations'),
):
    assert extract_remediation_path('Remediation:\n' + wrapped + '\nDefault Value:\nEnabled') == expected
legacy = r'Computer Configuration\Policies\Administrative Templates\Windows Components\UI\Example'
assert extract_remediation_path('Remediation:\n' + legacy + '\nReferences:\nhttps://example.com') == legacy
assert extract_remediation_path('Remediation:\n' + legacy.replace('Computer', 'User')) == legacy.replace('Computer', 'User')
assert extract_remediation_path('Audit:\n' + legacy + '\nRemediation:\nConfigure manually.\nDefault Value:\nDisabled') is None
assert extract_remediation_path(r'Remediation: HKEY_LOCAL_MACHINE\Software\Policies') is None
assert extract_remediation_path('Remediation:\n' + expected + '\nNote:\nDo not include this.') == expected
assert extract_remediation_path('Remediation:\n' + expected + '\n2.6.7 (L1) Ensure next recommendation\nRemediation:\nOther\\Setting') == expected
print('PASS: wrapped Office paths, separators, plain continuation, page footer, legacy paths, missing paths, Audit and section boundaries.')

tree = ast.parse(embedded_code)
functions = [node for node in tree.body if isinstance(node, ast.FunctionDef)]
constants = [node for node in tree.body if isinstance(node, ast.Assign) and any(
    isinstance(target, ast.Name) and target.id in ('profile_token_pattern', 'rec_intro_pattern') for target in node.targets
)]
detail_loop = next(node for node in tree.body if isinstance(node, ast.For) and
    isinstance(node.target, ast.Name) and node.target.id == 'page_idx' and
    ast.unparse(node.iter) == 'range(len(detail_texts))')
runner = compile(ast.Module(body=functions + constants + [detail_loop], type_ignores=[]), '<embedded detail parser>', 'exec')

def parse_pages(pages):
    scope = dict(re=re, unicodedata=unicodedata, detail_texts=pages, section_headers={},
                 rec_refs_set={'1.1', '1.2'})
    for name in ('rec_references', 'rec_safeguards', 'rec_values', 'rec_policy_groups',
                 'rec_gp_setting_names', 'rec_remediation_apps', 'rec_texts_from_detail', 'rec_profile_levels'):
        scope[name] = {}
    exec(runner, scope)
    return scope['rec_gp_setting_names']

first = "1.1 (L1) Ensure 'Network locations' is set to 'Disabled' (Automated)\n"
second = "1.2 (L1) Ensure 'Other setting' is set to 'Disabled' (Automated)\n"
split_path = expected.index('Locations\\')
page_one = first + 'Remediation:\n' + expected[:split_path] + '\nPage 42'
page_two = expected[split_path:] + '\nDefault Value:\nEnabled\n' + second + 'Remediation:\nAdministrative Templates\\Other setting\nDefault Value:\nDisabled'
paths = parse_pages([page_one, page_two])
assert paths['1.1'] == expected, paths
assert paths['1.2'] == r'Administrative Templates\Other setting', paths
paths = parse_pages([first + 'Remediation:\nNo navigation path.\n' + second + sample])
assert '1.1' not in paths and paths['1.2'] == expected, paths
print('PASS: actual detail parser, cross-page path before next recommendation, neighboring recommendation isolation, embedded Python syntax.')

import contextlib
import io
import json
import sys
import types

def convert_pages(detail_pages, summary_lines):
    pages = detail_pages + ['Appendix: Summary Table\nCIS Benchmark Recommendation\n' + '\n'.join(summary_lines)]
    reader = types.SimpleNamespace(pages=[types.SimpleNamespace(extract_text=lambda text=text: text) for text in pages])
    previous_module = sys.modules.get('PyPDF2')
    sys.modules['PyPDF2'] = types.SimpleNamespace(PdfReader=lambda path: reader)
    output = io.StringIO()
    try:
        with contextlib.redirect_stdout(output):
            exec(compile(embedded_code, '<embedded converter>', 'exec'), {})
    finally:
        if previous_module is None:
            del sys.modules['PyPDF2']
        else:
            sys.modules['PyPDF2'] = previous_module
    return json.loads(output.getvalue())

excel_title = "Ensure 'Set default file block behavior' is set to 'Enabled: Blocked files are not opened'"
excel_path = r'Microsoft Excel 2016\Excel Options\Security\Trust Center\File Block Settings\Set Default File Block Behavior'
excel_detail = '2.2.4.7.2.2.13 (L1) ' + excel_title + ' (Automated)\nRemediation:\n' + excel_path + '\nDefault Value:\nNot configured'
summary_headers = ['2 User Configuration', '2.2 Microsoft Excel 2016', '2.2.4.7.2.2 File Block Settings']
for summary_rec in (
    ['2.2.4.7.2.2.1', '3 (L1) ' + excel_title + ' (Automated)'],
    ['2.2.4.7.2.2.', '13 (L1) ' + excel_title + ' (Automated)'],
    ['3 (L1) ' + excel_title + ' (Automated)'],
):
    converted = convert_pages([excel_detail], summary_headers + summary_rec)
    rec = converted['recommendations'][0]
    assert rec['ref'] == '2.2.4.7.2.2.13', rec
    assert rec['gp_setting_name'] == excel_path, rec
    assert rec['policy_group'] == 'File Block Settings', rec
    assert rec['parent_header'] == '2.2.4.7.2.2', rec

file_titles = (
    "Ensure 'Excel 95 workbooks' is set to 'Enabled: Open/Save Blocked, Use Open Policy'",
    "Ensure 'Excel 95-97 workbooks and templates' is set to 'Enabled: Open/Save Blocked, Use Open Policy'",
    "Ensure 'Excel 97-2003 workbooks and templates' is set to 'Enabled: Open/Save Blocked, Use Open Policy'",
    excel_title,
    "Ensure 'Web pages and Excel 2003 XML spreadsheets' is set to 'Enabled: Open/Save blocked, use open policy'",
)
file_details = []
file_summary = ['2 User Configuration', '2.2 Microsoft Excel 2016', '2.2.4.7.2.1 External Content',
                "2.2.4.7.2.1.1 (L1) Ensure 'External example' is set to 'Enabled' (Automated)",
                '2.2.4.7.2.2 File Block Settings']
for index, title in enumerate(file_titles):
    ref = '2.2.4.7.2.2.' + str(10 + index)
    file_details.append(ref + ' (L1) ' + title + ' (Automated)\nRemediation:\n' + excel_path + '\nDefault Value:\nNot configured')
    file_summary.append(str(index) + ' (L1) ' + title + ' (Automated)')
converted = convert_pages(file_details, file_summary)
file_recs = [rec for rec in converted['recommendations'] if rec['ref'].startswith('2.2.4.7.2.2.')]
assert len(file_recs) == 5, converted
assert [rec['ref'] for rec in file_recs] == ['2.2.4.7.2.2.' + str(index) for index in range(10, 15)]
assert all(rec['parent_header'] == '2.2.4.7.2.2' and rec['policy_group'] == 'File Block Settings' for rec in file_recs)
assert 'File Block Settings' not in converted['recommendations'][0]['text']

word_detail = excel_detail.replace('2.2.4.7.2.2.13', '2.11.8.7.2.1.13').replace('Microsoft Excel', 'Microsoft Word')
try:
    convert_pages([excel_detail, word_detail], summary_headers + ['3 (L1) ' + excel_title + ' (Automated)'])
    raise AssertionError('Conflicting detail references were accepted')
except ValueError as error:
    assert 'Cannot verify truncated CIS reference' in str(error)
try:
    convert_pages([], summary_headers + ['3 (L1) ' + excel_title + ' (Automated)'])
    raise AssertionError('Unverified single-digit reference was accepted')
except ValueError as error:
    assert 'Cannot verify truncated CIS reference' in str(error)
print('PASS: complete embedded conversion, digit/dot wraps, detail-verified reference recovery, GP path and grouping, conflicting/unverified rejection.')
'@
$pythonCode = "import re`nembedded_code = $embeddedJson`n$helper`n$checks"
$pythonCode | & python -
if ($LASTEXITCODE -ne 0) { throw 'Embedded Python remediation path tests failed.' }

$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join "`n") }
foreach ($statement in $ast.EndBlock.Statements) {
    if ($statement -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
        Invoke-Expression $statement.Extent.Text
    }
}
$tableFunction = $ast.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Build-PaddedTable'
}, $true)
if (-not $tableFunction) { throw 'Markdown table builder not found.' }
Invoke-Expression $tableFunction.Extent.Text
$expectedPath = 'Microsoft PowerPoint 2016\PowerPoint Options\Security\Trust Center\Trusted Locations\Allow Trusted Locations on the network'
$recommendation = [pscustomobject]@{
    text = 'Recommendation'; ref = '2.6.6.6.2.3.1'; profile = 'L1'; type = 'Automated'
    safeguards = @(); references = @(); rec_value = $null; gp_setting_name = $expectedPath
}
$table = @(Build-PaddedTable -Rows @($recommendation) -FallbackPolicyName 'Office' -PrepareForIntune $true -UseSectionPolicyName $true)
$parsed = Read-MarkdownTable -TableLines $table
$gpColumn = Get-ColumnIndex -Headers $parsed.Headers -Name 'GP Setting Name'
if ($gpColumn -lt 0 -or $parsed.Rows[0].Cells[$gpColumn] -ne $expectedPath) {
    throw 'Markdown GP Setting Name did not retain the complete path.'
}
$recommendation.gp_setting_name = 'Administrative Templates\Example | setting'
$table = @(Build-PaddedTable -Rows @($recommendation) -FallbackPolicyName 'Office' -PrepareForIntune $true -UseSectionPolicyName $true)
$parsed = Read-MarkdownTable -TableLines $table
if ($parsed.Rows[0].Cells.Count -ne $parsed.Headers.Count -or $table[-1] -notlike '*Example &#124; setting*') {
    throw 'Path delimiter corrupted the Markdown table.'
}
$recommendation.gp_setting_name = $null
$table = @(Build-PaddedTable -Rows @($recommendation) -FallbackPolicyName 'Office' -PrepareForIntune $true -UseSectionPolicyName $true)
$parsed = Read-MarkdownTable -TableLines $table
if ($parsed.Rows[0].Cells[$gpColumn] -ne '-') { throw 'Missing path should render as a dash.' }
Write-Output 'PASS: actual Markdown builder, exact full-path cell, delimiter escaping, missing-path placeholder.'
$temporaryFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('CISPathTest-' + [guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($temporaryFolder)
try {
    $markdown = @'
# Fixture
| Recommendation | CIS Reference | Origin | Intune Policy | GP Setting Name |
| --- | --- | --- | --- | --- |
| Recommendation one | 1.1 | CIS | Office | Microsoft PowerPoint 2016\Trust Center\Network locations |
| Recommendation two | 1.2 | CIS | Office | Microsoft Excel 2016\Trust Center\Network locations |
| Recommendation three | 1.3 | CIS | Office | Legacy setting |
| Recommendation four | 1.4 | CIS | Office | Microsoft Word 2016\Trust Center\Unique setting |
'@
    $markdown | Set-Content -LiteralPath (Join-Path $temporaryFolder 'fixture.md') -Encoding UTF8
    $lookup = New-BaselineCisReferenceLookup -MarkdownFolder $temporaryFolder
    foreach ($case in @(
        @('Microsoft PowerPoint 2016\Trust Center\Network locations', '1.1', 'Matched'),
        @('Unique setting', '1.4', 'Matched'),
        @('Legacy setting', '1.3', 'Matched'),
        @('Network locations', '', 'Ambiguous')
    )) {
        $row = [pscustomobject]@{ 'Baseline Policy' = 'Office'; 'Setting Name' = $case[0]; 'Expected Value' = '' }
        $result = Resolve-BaselineCisReference -Lookup $lookup -CsvRow $row
        if ($result.CISReference -ne $case[1] -or $result.Status -ne $case[2]) {
            throw "Unexpected lookup for '$($case[0])': $($result.Status) / $($result.CISReference)"
        }
    }
    Write-Output 'PASS: full-path matching, leaf aliases, legacy Markdown, duplicate-name ambiguity, PowerShell syntax.'
    foreach ($case in @(@('excel16v2', '1.2'), @('ppt16v2', '1.1'))) {
        $row = [pscustomobject]@{
            'Baseline Policy' = 'Office'; 'Setting Name' = 'Network locations'; 'Expected Value' = ''
            'Definition Id' = "user_vendor_msft_policy_config_$($case[0])~policy~l_trustcenter_l_networklocations"
        }
        $result = Resolve-BaselineCisReference -Lookup $lookup -CsvRow $row
        if ($result.CISReference -ne $case[1] -or $result.Status -ne 'Matched') {
            throw "Application disambiguation failed for '$($case[0])': $($result.Status) / $($result.CISReference)"
        }
    }
    Write-Output 'PASS: Excel and PowerPoint definition IDs disambiguate identical leaf names.'

    $settingName = 'Block macros from running in Office files from the Internet (User)'
    $applications = @(
        @('Access', 'access16v2', '2.1'), @('Excel', 'excel16v2', '2.2'),
        @('PowerPoint', 'ppt16v2', '2.3'), @('Publisher', 'pub16v3', '2.4'), @('Word', 'word16v2', '2.5')
    )
    $officeMarkdown = @(
        '# Office fixture',
        '| Recommendation | CIS Reference | Origin | Intune Policy | GP Setting Name |',
        '| --- | --- | --- | --- | --- |'
    )
    foreach ($application in $applications) {
        $officeMarkdown += "| $settingName | $($application[2]) | CIS | TrustCenter | Microsoft $($application[0]) 2016\Trust Center\$settingName |"
    }
    $fixturePath = Join-Path $temporaryFolder 'office.md'
    $officeMarkdown | Set-Content -LiteralPath $fixturePath -Encoding UTF8
    $lookup = New-BaselineCisReferenceLookup -MarkdownFolder $temporaryFolder
    foreach ($application in $applications) {
        foreach ($mode in @('Exact', 'OtherPolicy', 'ExpectedValue', 'Similarity')) {
            $row = [pscustomobject]@{
                'Baseline Policy' = 'TrustCenter'; 'Setting Name' = $settingName; 'Expected Value' = ''
                'Definition Id' = "user_vendor_msft_policy_config_$($application[1])~policy~l_microsoft$($application[0])~l_trustcenter_l_blockmacroexecutionfrominternet"
            }
            $expectedStatus = 'Matched'
            switch ($mode) {
                'OtherPolicy' { $row.'Baseline Policy' = 'Renamed TrustCenter'; $expectedStatus = 'Matched (other policy)' }
                'ExpectedValue' { $row.'Setting Name' = 'Unrelated child setting'; $row.'Expected Value' = $settingName }
                'Similarity' { $row.'Setting Name' = "$settingName extra" }
            }
            $result = Resolve-BaselineCisReference -Lookup $lookup -CsvRow $row -SimilarityThreshold 0.8
            if ($result.CISReference -ne $application[2] -or $result.Status -ne $expectedStatus) {
                throw "$mode failed for $($application[0]): $($result.Status) / $($result.CISReference)"
            }
        }
    }
    $row = [pscustomobject]@{
        'Baseline Policy' = 'TrustCenter'; 'Setting Name' = $settingName; 'Expected Value' = ''
        'Definition Id' = ''
    }
    foreach ($definitionId in @('', 'unknown', 'device_vendor_msft_policy_config_outlk16v2~policy',
        'user_vendor_msft_policy_config_excel16v2~policy~l_microsoftword~l_setting')) {
        $row.'Definition Id' = $definitionId
        $result = Resolve-BaselineCisReference -Lookup $lookup -CsvRow $row
        if ($result.Status -ne 'Ambiguous') { throw "Unsafe application match for '$definitionId'." }
    }
    foreach ($definitionId in @('user_vendor_msft_policy_config_custom~policy~l_microsoftexcel~l_setting',
        'USER_VENDOR_MSFT_POLICY_CONFIG_EXCEL160V4~POLICY~L_SETTING')) {
        $row.'Definition Id' = $definitionId
        $result = Resolve-BaselineCisReference -Lookup $lookup -CsvRow $row
        if ($result.CISReference -ne '2.2') { throw "Category/version alias failed for '$definitionId'." }
    }
    $row.'Definition Id' = 'user_vendor_msft_policy_config_excel16v2~policy~l_microsoftexcel~l_setting'
    foreach ($path in @('Microsoft Office 2016\Excel export setting', 'Microsoft Word 2016\Example',
        'Microsoft Excel 2016\Trust Center\Example', 'Legacy setting')) {
        $conflictingMarkdown = $officeMarkdown + "| $settingName | 9.9 | CIS | TrustCenter | $path |"
        $conflictingMarkdown | Set-Content -LiteralPath $fixturePath -Encoding UTF8
        $conflictingLookup = New-BaselineCisReferenceLookup -MarkdownFolder $temporaryFolder
        $result = Resolve-BaselineCisReference -Lookup $conflictingLookup -CsvRow $row
        $expectedStatus = if ($path -eq 'Microsoft Word 2016\Example') { 'Matched' } else { 'Ambiguous' }
        if ($result.Status -ne $expectedStatus) { throw "Unexpected application conflict result for '$path': $($result.Status)" }
    }
    $duplicateMarkdown = $officeMarkdown + "| $settingName | 2.2 | CIS | TrustCenter | Microsoft Excel 2016\Other category\$settingName |"
    $duplicateMarkdown | Set-Content -LiteralPath $fixturePath -Encoding UTF8
    $duplicateLookup = New-BaselineCisReferenceLookup -MarkdownFolder $temporaryFolder
    foreach ($name in @($settingName, "$settingName extra")) {
        $row.'Setting Name' = $name
        $result = Resolve-BaselineCisReference -Lookup $duplicateLookup -CsvRow $row -SimilarityThreshold 0.8
        if ($result.CISReference -ne '2.2' -or $result.Status -ne 'Matched') { throw 'Duplicate same-reference rows affected matching.' }
    }
    $result = Resolve-BaselineCisReference -Lookup $duplicateLookup -CsvRow $row -SimilarityThreshold 1.0
    if ($result.Status -ne 'No match') { throw 'Application context bypassed the similarity threshold.' }
    if ((Get-BaselineOfficeApplication -Text 'Microsoft Office 2016\Export to Microsoft Excel 2016') -ne '') {
        throw 'A setting title was mistaken for application path context.'
    }
    if ((Get-BaselineOfficeApplication -Text 'Microsoft Excel 2016\Microsoft Word 2016\Example') -ne '') {
        throw 'Conflicting GP application context was accepted.'
    }
    $twodayMarkdown = $officeMarkdown -replace '\| 2\.2 \| CIS \|', '| 2.2 | twoday |'
    $twodayMarkdown | Set-Content -LiteralPath $fixturePath -Encoding UTF8
    $twodayLookup = New-BaselineCisReferenceLookup -MarkdownFolder $temporaryFolder
    $row.'Setting Name' = $settingName
    $result = Resolve-BaselineCisReference -Lookup $twodayLookup -CsvRow $row
    if ($result.CISReference -ne 'twoday' -or $result.Status -ne 'Twoday') { throw 'Twoday status changed during disambiguation.' }
    Write-Output 'PASS: five Office apps across all four matching branches, aliases, conflicts, unknown context, duplicate references, similarity threshold and Twoday status.'
} finally {
    Remove-Item -LiteralPath $temporaryFolder -Recurse -Force
}