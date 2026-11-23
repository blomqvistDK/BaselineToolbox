# Baseline Toolbox

Run `Invoke-BaselineToolbox.ps1` from this folder for the existing GUI. Its processing uses the same scripts and helpers available to PowerShell callers. Windows PowerShell 5.1 or newer is supported. Keep the entire folder together when copying the toolbox.

```text
BaselineToolbox/
|-- Invoke-BaselineToolbox.ps1
|-- Features/       Standalone feature scripts
|-- Shared/         BaselineToolkit.Common.psm1
|-- Resources/      Customer-SecurityBaselineSettings-Overview.xlsx
|-- Tests/          Regression and GUI smoke tests
`-- README.md
```

## PowerShell Modules

Install these modules for the features that use them:

- `Microsoft.Graph.Authentication` is required for Microsoft Graph operations. The selected operation also requires the appropriate delegated Graph permissions and consent.
- `ImportExcel` is required only for Baseline Workbook export. Microsoft Excel must also be installed for that feature.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser
```

`Shared\BaselineToolkit.Common.psm1` is included in this folder and does not need to be installed separately. Python/PyPDF2 and PDF extraction prerequisites are feature-specific and documented below.

## Feature Inventory

| Feature | Script in Features |
| --- | --- |
| PDF to Markdown | [Invoke-BaselinePdfConversion.ps1](Features/Invoke-BaselinePdfConversion.ps1) |
| Change History | [Export-ChangeHistory.ps1](Features/Export-ChangeHistory.ps1) |
| Markdown to HTML | [Convert-MarkdownTextToInteractiveHtmlReport.ps1](Features/Convert-MarkdownTextToInteractiveHtmlReport.ps1) |
| Baseline Merger | [Invoke-BaselineMerger.ps1](Features/Invoke-BaselineMerger.ps1) |
| Intune Policy Creator | [Invoke-IntuneCISPolicyCreator.ps1](Features/Invoke-IntuneCISPolicyCreator.ps1) |
| Intune Policy Report | [Invoke-IntuneBaselineReport.ps1](Features/Invoke-IntuneBaselineReport.ps1) |
| Baseline Assessment | [Invoke-BaselineAssessment.ps1](Features/Invoke-BaselineAssessment.ps1) |
| Baseline Workbook | [Export-BaselineAssessmentWorkbook.ps1](Features/Export-BaselineAssessmentWorkbook.ps1) |

The launcher has moved completely from the parent Markdown folder; update existing shortcuts to `BaselineToolbox\Invoke-BaselineToolbox.ps1`. No forwarding launcher remains at the old location. Legacy tools, backups, baseline data, and existing reports remain outside this folder.

Run the following examples from this folder, or invoke the scripts by full path. Use `Get-Help .\Features\<script-name>.ps1 -Full` to inspect parameters. Script and resource lookup does not depend on the current working directory.

## PDF to Markdown

```powershell
.\Features\Invoke-BaselinePdfConversion.ps1 -PdfPath 'C:\Benchmarks\Windows.pdf' -ProductName 'Windows 11' -OutputFolder 'C:\Reports' -ProfileFilter Both -PrepareForIntune
```

Requires Python 3 and PyPDF2. Returns Markdown, text, and HTML paths and the recommendation count. `-PrepareForIntune` enables the policy-name preparation option; omit it for an ordinary conversion. `-IntunePrefix` overrides the generated prefix. This script contains the current toolbox converter; the older `Invoke-CISBenchmarkToMarkdown.ps1` remains in the parent Markdown folder for compatibility.

## Change History

```powershell
.\Features\Export-ChangeHistory.ps1 -PdfPath 'C:\Benchmarks\Windows.pdf' -PreviousMarkdownPath 'C:\Benchmarks\previous.md' -OutputDir 'C:\Reports' -SimilarityThreshold 0.80
```

Retains the existing version, page-range, diagnostic, and PDF extraction options. Requires the existing PDF extraction prerequisites described in its help. Supply `-OutputDir` to choose a report folder; the existing script-directory default now resolves to `Features` after relocation.

## Markdown to HTML

```powershell
.\Features\Convert-MarkdownTextToInteractiveHtmlReport.ps1 -InputFile 'C:\Benchmarks\Windows.md' -OutputFile 'C:\Reports\Windows.html'
```

Use `-OpenReport` to open the result explicitly.

## Baseline Merger

```powershell
.\Features\Invoke-BaselineMerger.ps1 -NewBaselinePath 'C:\Benchmarks\new.md' -OldBaselinePath 'C:\Benchmarks\old.md' -OutputPath 'C:\Reports\merged.md' -ProfileFilter L1 -SimilarityThreshold 0.90 -GenerateReport
.\Features\Invoke-BaselineMerger.ps1 -ReportOnly -InputMarkdownPath 'C:\Reports\merged.md' -OutputHtmlPath 'C:\Reports\comparison.html'
```

Omit `-GenerateReport` for merge-only. `-ReportTitle` and `-OpenReport` customize report generation. The function's existing exact-match default (`1.0`) is retained; the GUI continues to supply its configured threshold. Returns output paths and merge statistics.

## Intune Policy Creator

```powershell
.\Features\Invoke-IntuneCISPolicyCreator.ps1 -MdFilePath 'C:\Benchmarks\Windows.md' -PreviewOnly -PreviewOrigin CIS
.\Features\Invoke-IntuneCISPolicyCreator.ps1 -MdFilePath 'C:\Benchmarks\Windows.md' -ProfileLevel L1 -ExportOnly
.\Features\Invoke-IntuneCISPolicyCreator.ps1 -MdFilePath 'C:\Benchmarks\Windows.md' -ProfileLevel L1 -WhatIf
```

Retains its existing filters, profile types, platform options, and update behavior. Creating or updating policies requires appropriate Microsoft Graph write permissions. Omitting `-MdFilePath` retains this older script's own GUI; supplying it selects command-line mode.

## Intune Policy Report

```powershell
.\Features\Invoke-IntuneBaselineReport.ps1 -ListProfiles -TenantId 'tenant-guid'
.\Features\Invoke-IntuneBaselineReport.ps1 -ProfileId 'profile-guid' -OutputPath 'C:\Reports\policies.html' -CacheMode fallback
.\Features\Invoke-IntuneBaselineReport.ps1 -ProfileId 'profile-guid' -OutputPath 'C:\Reports\comparison.html' -CompareFolder 'C:\PolicyBackups' -CacheMode fallback
.\Features\Invoke-IntuneBaselineReport.ps1 -WarmCache -TenantId 'tenant-guid'
```

`-ProfileId` accepts multiple IDs. Alternatively pass the profile objects returned by `-ListProfiles` through `-Profiles`. Report results include setting rows, failed profiles, and comparison data. `-WarmCache` uses all discoverable profiles unless a selection is supplied.

Load profiles (and `-ListProfiles`) lists Settings Catalog, device-restriction, and Administrative Templates (`groupPolicyConfigurations`) policies. Administrative Template settings are reported as the enabled/disabled state of each configured setting plus each of its presentation values.

Without a Compare Folder, the report inventories selected policies and settings. With a folder, each selected policy is matched to its JSON backup by ID first, then by a unique name and policy type. Comparisons show changed values, settings missing from live, and settings added since backup. Duplicate matches, unreadable/unsupported backups, missing backups, and collection failures are reported explicitly. JSON files are read from the selected folder, not subfolders; unrelated backups are ignored. Only the currently supported Settings Catalog, device-restriction, and Administrative Templates policy formats are compared. Administrative Template backups may contain expanded `definition`/`presentation` objects or `@odata.bind` references, either as a policy with `definitionValues` or as a bare `definitionValues` array. Assignments and policy metadata are not compared. The legacy script filename and `-BaselineFolder` parameter remain supported; `-CompareFolder` is its new alias. No policies are modified.

## Baseline Assessment

```powershell
.\Features\Invoke-BaselineAssessment.ps1 -BaselineFolder 'C:\Baselines' -ProfileId 'profile-guid' -OutputPath 'C:\Reports\assessment.html' -SkipAssignmentLookups
```

Supports `-SkipUnassigned`, `-SkipAssignmentLookups`, `-Profiles`, `-TenantId`, `-CacheMode`, and `-OpenReport`. Returns the assessment data used by the GUI, including the data needed for split-policy operations. Report CSV exports remain available through the generated HTML, as in the GUI workflow. Administrative Templates appear in the profile list but are skipped (with a log entry) by the assessment; use the Intune Policy Report for them.

Creating split policies is an explicit write operation:

```powershell
.\Features\Invoke-BaselineAssessment.ps1 -BaselineFolder 'C:\Baselines' -ProfileId 'profile-guid' -OutputPath 'C:\Reports\assessment.html' -CreateSplitPolicy -SplitBaselineName 'baseline-policy-name' -WhatIf
```

Remove `-WhatIf` only when ready to create policies. Without `-CreateSplitPolicy`, assessment does not create policies. `-WhatIf` prevents policy creation, but assessment still reads tenant data and writes its local report.

Report and assessment reuse a valid Graph session matching the requested tenant, otherwise authenticate. `BASELINE_TOOLBOX_RESOLVER_TENANTID` supplies the default tenant. Read-only operations use configuration-read and organization-read scopes. Explicit split-policy creation also requires configuration-write access. Setting-definition caches remain tenant-specific under `%LOCALAPPDATA%\BaselineToolbox`; `cacheonly` skips definition-resolution requests, not the live policy requests needed for assessment.

## Baseline Workbook

```powershell
.\Features\Export-BaselineAssessmentWorkbook.ps1 -CsvPath 'C:\Reports\details.csv' -MarkdownFolder 'C:\Benchmarks' -OutputPath 'C:\Reports\assessment.xlsx' -SimilarityThreshold 0.90
```

Requires ImportExcel, Microsoft Excel, and the bundled `Resources\Customer-SecurityBaselineSettings-Overview.xlsx` template. `-TemplatePath` permits an explicit template. The template cannot be overwritten by the output. Returns the workbook path and matching/deduplication counts.

The All Settings sheet keeps its five summary counters and assignment-status data without the extra F4 counter or CIS Reference Status column. Deviations remains a manual worksheet: rows 9-11 are blank for entry, and template rows from 12 onward are removed. No deviation rows or comments are generated automatically.

## Validation

```powershell
.\Tests\Test-BaselineAssessmentHierarchy.ps1
.\Tests\Test-CISRemediationPaths.ps1
.\Tests\Test-BaselineFeatureScripts.ps1
.\Tests\Test-BaselineAssessmentHierarchy.ps1 -VerifyWorkbook
pwsh -NoProfile -STA -File .\Tests\Test-BaselineFeatureScripts.ps1 -VerifyGui
```

The GUI smoke test opens and automatically closes the toolbox without running tenant operations. The workbook check creates a temporary workbook and requires Excel. Live Graph workflows still require an authorized tenant and should be tested separately before production use.