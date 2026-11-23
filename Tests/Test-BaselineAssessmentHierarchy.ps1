#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ScriptPath,
    [switch]$VerifyWorkbook
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) {
    $ScriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Invoke-BaselineToolbox.ps1'
}
Import-Module (Join-Path (Split-Path $ScriptPath -Parent) 'Shared\BaselineToolkit.Common.psm1') -Force -DisableNameChecking
. (Join-Path (Split-Path $ScriptPath -Parent) 'Features\Export-BaselineAssessmentWorkbook.ps1')
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join "`n") }
$script:thisScriptPath = $ScriptPath

function Assert-Condition {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function New-TestRow {
    param([string]$Name, [string]$Id, [string]$Policy = 'Printers', [int]$Depth = 0)
    [pscustomobject]@{
        Classification = 'Conflicts'
        'Baseline Policy' = $Policy
        File = 'fixture.json'
        'Setting Name' = $Name
        'Setting Depth' = $Depth
        'Definition Id' = $Id
        'Expected Value' = 'Enabled'
        Status = 'Match'
        'Found In Policy' = 'Fixture'
        'Actual Value' = 'Enabled'
        'Policy Assigned' = 'Yes'
    }
}

$lookup = [pscustomobject]@{
    ExactByKey = @{}; ExactByGpKey = @{}
    ExactBySettingOnly = @{}; ExactByGpSettingOnly = @{}
    CandidatesByPolicy = @{}; HierarchicalParentsByPolicy = @{}
}
$parentId = 'device_vendor_msft_policy_config_printers_pointandprintrestrictions'
$rows = @(
    New-TestRow 'Point and Print Restrictions' $parentId
    New-TestRow 'Enter fully qualified server names separated by semicolons' "${parentId}_pointandprint_trustedservers_edit"
    New-TestRow 'Users can only point and print to machines in their forest' "${parentId}_pointandprint_trustedforest_chk"
    New-TestRow 'Users can only point and print to these servers:' "${parentId}_pointandprint_trustedservers_chk"
    New-TestRow 'When installing drivers for a new connection:' "${parentId}_pointandprint_nowarningnoelevationoninstall_enum"
    New-TestRow 'When updating drivers for an existing connection:' "${parentId}_pointandprint_nowarningnoelevationonupdate_enum"
)
$shuffled = @($rows[4], $rows[0], $rows[2], $rows[5], $rows[1], $rows[3])
$result = @(Get-BaselineAssessmentHierarchy -Rows $shuffled -Lookup $lookup)
Assert-Condition ($result.Count -eq 6 -and $result[0].Row.'Definition Id' -eq $parentId) 'Parent-first grouping failed.'
Assert-Condition (@($result | Where-Object Depth -eq 1).Count -eq 5) 'Not all five children have depth one.'
Assert-Condition (@($result.SourceIndex | Sort-Object -Unique).Count -eq 6) 'Hierarchy lost or duplicated rows.'
Assert-Condition (($result.SourceIndex -join ',') -eq '1,0,2,3,4,5') 'Sibling order changed.'

$cases = @(
    New-TestRow 'Grandchild' 'ROOT_BRANCH_LEAF'
    New-TestRow 'Unrelated' 'rooted_child' -Depth 9
    New-TestRow 'Other policy' 'root_branch' -Policy 'Other' -Depth 9
    New-TestRow 'Root' 'root' -Depth 5
    New-TestRow 'Branch' 'root_branch'
)
$result = @(Get-BaselineAssessmentHierarchy -Rows $cases -Lookup $lookup)
Assert-Condition (($result | Where-Object { $_.Row.'Setting Name' -eq 'Grandchild' }).Depth -eq 2) 'Longest-prefix nesting failed.'
Assert-Condition (($result | Where-Object { $_.Row.'Setting Name' -eq 'Grandchild' }).Path -eq 'Root: Branch: Grandchild') 'Nested path failed.'
Assert-Condition (($result | Where-Object { $_.Row.'Setting Name' -eq 'Unrelated' }).Depth -eq 0) 'Underscore boundary failed.'
Assert-Condition (($result | Where-Object { $_.Row.'Setting Name' -eq 'Other policy' }).Depth -eq 0) 'Policy isolation failed.'
$cases = @((New-TestRow 'First value' 'root'), (New-TestRow 'Second value' 'root'), (New-TestRow 'Child' 'root_child'))
$cases[1].'Expected Value' = 'Disabled'
$merged = Merge-BaselineAssessmentDetailRows -Rows $cases
$result = @(Get-BaselineAssessmentHierarchy -Rows $merged.Rows -Lookup $lookup)
Assert-Condition ($result.Count -eq 3 -and @($result | Where-Object Depth -gt 0).Count -eq 0) 'Ambiguous parent was selected arbitrarily.'
$result = @(Get-BaselineAssessmentHierarchy -Rows @((New-TestRow 'Root' ''), (New-TestRow 'Child' '' -Depth 1)) -Lookup $lookup)
Assert-Condition ($result[1].Depth -eq 1 -and $result[1].Path -eq 'Root: Child') 'Missing-ID depth fallback failed.'
Assert-Condition (@(Get-BaselineAssessmentHierarchy -Rows @() -Lookup $lookup).Count -eq 0) 'Empty input failed.'

$inheritanceHierarchy = @(
    [pscustomobject]@{ SourceIndex = 0; ParentIndex = -1; Children = @(1); DefinitionId = 'root' }
    [pscustomobject]@{ SourceIndex = 1; ParentIndex = 0; Children = @(2); DefinitionId = 'root_child' }
    [pscustomobject]@{ SourceIndex = 2; ParentIndex = 1; Children = @(); DefinitionId = 'root_child_grandchild' }
)
$inheritanceResults = [System.Collections.Generic.List[object]]::new()
foreach ($reference in @('', '4.7.10', '4.7.3')) {
    [void]$inheritanceResults.Add([pscustomobject]@{ CISReference = $reference; Status = if ($reference) { 'Matched' } else { 'No match' } })
}
Set-BaselineAssessmentCisReferenceInheritance -Hierarchy $inheritanceHierarchy -ReferenceResults $inheritanceResults
Assert-Condition (($inheritanceResults.CISReference -join ',') -eq '4.7.3,4.7.10,4.7.3') 'Numeric minimum or direct child reference preservation failed.'
Assert-Condition (($inheritanceResults.Status -join ',') -eq 'Inherited,Matched,Matched') 'Direct and inherited statuses are incorrect.'
$rpcHierarchy = @(
    [pscustomobject]@{ SourceIndex = 0; ParentIndex = -1; Children = @(1, 2); DefinitionId = 'rpc' }
    [pscustomobject]@{ SourceIndex = 1; ParentIndex = 0; Children = @(); DefinitionId = 'rpc_protocol' }
    [pscustomobject]@{ SourceIndex = 2; ParentIndex = 0; Children = @(); DefinitionId = 'rpc_authentication' }
)
$rpcResults = [System.Collections.Generic.List[object]]::new()
foreach ($reference in @('', '4.7.3', '4.7.4')) {
    [void]$rpcResults.Add([pscustomobject]@{ CISReference = $reference; Status = if ($reference) { 'Matched' } else { 'No match' } })
}
Set-BaselineAssessmentCisReferenceInheritance -Hierarchy $rpcHierarchy -ReferenceResults $rpcResults
Assert-Condition (($rpcResults.CISReference -join ',') -eq '4.7.3,4.7.3,4.7.4') 'RPC parent minimum or direct child references failed.'

$inheritanceResults = [System.Collections.Generic.List[object]]::new()
foreach ($reference in @('1.2.5.1.12', '1.2.5.1.13', '')) {
    [void]$inheritanceResults.Add([pscustomobject]@{ CISReference = $reference; Status = if ($reference) { 'Matched' } else { 'No match' } })
}
Set-BaselineAssessmentCisReferenceInheritance -Hierarchy $inheritanceHierarchy -ReferenceResults $inheritanceResults
Assert-Condition (($inheritanceResults.CISReference -join ',') -eq '1.2.5.1.12,1.2.5.1.13,1.2.5.1.13') 'Existing parent or direct child reference was not preserved.'
Assert-Condition (($inheritanceResults.Status -join ',') -eq 'Matched,Matched,Inherited') 'Parent-authoritative and inherited statuses are incorrect.'

$unresolvedHierarchy = @(
    [pscustomobject]@{ SourceIndex = 0; ParentIndex = -1; Children = @(1); DefinitionId = '' }
    [pscustomobject]@{ SourceIndex = 1; ParentIndex = 0; Children = @(); DefinitionId = '' }
)
$inheritanceResults = [System.Collections.Generic.List[object]]::new()
[void]$inheritanceResults.Add([pscustomobject]@{ CISReference = ''; Status = 'No match' })
[void]$inheritanceResults.Add([pscustomobject]@{ CISReference = 'twoday'; Status = 'Twoday' })
Set-BaselineAssessmentCisReferenceInheritance -Hierarchy $unresolvedHierarchy -ReferenceResults $inheritanceResults
Assert-Condition (($inheritanceResults.CISReference -join ',') -eq ',twoday' -and ($inheritanceResults.Status -join ',') -eq 'No match,Twoday') 'Nonnumeric-only group should retain original results.'
$standaloneResult = [System.Collections.Generic.List[object]]::new()
[void]$standaloneResult.Add([pscustomobject]@{ CISReference = '4.2.1'; Status = 'Matched' })
Set-BaselineAssessmentCisReferenceInheritance -Hierarchy @([pscustomobject]@{ SourceIndex = 0; ParentIndex = -1; Children = @(); DefinitionId = 'standalone' }) -ReferenceResults $standaloneResult
Assert-Condition ($standaloneResult[0].CISReference -eq '4.2.1' -and $standaloneResult[0].Status -eq 'Matched') 'Standalone reference was changed.'

foreach ($profile in @('Domain', 'Private', 'Public')) {
    $parentName = "Enable $profile Network Firewall"
    $policyKey = Get-BaselineExportPolicyKey 'Firewall'
    $parentKey = Get-BaselineExportSettingKey $parentName
    $lookup.HierarchicalParentsByPolicy[$policyKey] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    [void]$lookup.HierarchicalParentsByPolicy[$policyKey].Add($parentKey)
    $fullPath = "${parentName}: Log File Path"
    $settingKey = Get-BaselineExportSettingKey $fullPath
    $lookup.ExactByKey["$policyKey|$settingKey"] = @([pscustomobject]@{ OutputReference = $profile })
    $fixture = @((New-TestRow $parentName '' 'Firewall'), (New-TestRow 'Log File Path' '' 'Firewall'))
    $result = @(Get-BaselineAssessmentHierarchy -Rows $fixture -Lookup $lookup)
    $resolved = Resolve-BaselineCisReference -Lookup $lookup -CsvRow $result[1].Row -SettingPath $result[1].Path
    Assert-Condition ($result[1].Depth -eq 1 -and $resolved.CISReference -eq $profile) "Firewall matching failed: $profile"
}
Write-Output 'PASS: syntax, five children, stable ordering, nested paths, ID boundaries, policy isolation, duplicate safety, missing-ID fallback, firewall references.'

if ($VerifyWorkbook) {
    function Get-TestWorkbookSummaryColors {
        param([string]$Path)
        $excel = New-Object -ComObject Excel.Application
        $workbook = $null
        try {
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($Path, 0, $true)
            $sheet = $workbook.Worksheets.Item('All Settings')
            $colors = @{
                ActiveSheet = [string]$workbook.ActiveSheet.Name
                SelectedSheetCount = [int]$workbook.Windows.Item(1).SelectedSheets.Count
            }
            foreach ($address in @('F3', 'F4')) {
                $cell = $sheet.Range($address)
                $colors[$address] = "$($cell.Interior.Color)|$($cell.Interior.Pattern)|$($cell.Font.Color)|$($cell.DisplayFormat.Interior.Color)|$($cell.DisplayFormat.Interior.Pattern)|$($cell.MergeArea.Address())"
            }
            return $colors
        } finally {
            if ($workbook) {
                $workbook.Close($false)
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
            }
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        }
    }
    $temporaryFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('BaselineHierarchy-' + [guid]::NewGuid().ToString('N'))
    [void][System.IO.Directory]::CreateDirectory($temporaryFolder)
    $package = $null
    try {
        $csvPath = Join-Path $temporaryFolder 'details.csv'
        $outputPath = Join-Path $temporaryFolder 'assessment.xlsx'
        $markdown = @('# Printers', '| Recommendation | CIS Reference | Origin | Intune Policy |', '| --- | --- | --- | --- |')
        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
            $reference = if ($rowIndex -eq 4) { '4.7.10' } elseif ($rowIndex -eq 5) { '4.7.11' } else { '' }
            $markdown += "| $($rows[$rowIndex].'Setting Name') | $reference | CIS | Printers |"
        }
        [System.IO.File]::WriteAllLines((Join-Path $temporaryFolder 'Printers.md'), $markdown)
        $bitlockerMarkdown = @(
            '# Bitlocker'
            '| Recommendation | CIS Reference | Origin | Intune Policy |'
            '| --- | --- | --- | --- |'
            "| Ensure 'Choose how BitLocker-protected fixed drives can be recovered' is set to 'Enabled' | 4.11.7.1.1 | CIS | Bitlocker |"
            "| Ensure 'Choose how BitLocker-protected fixed drives can be recovered: Recovery Password' is set to 'Enabled: Allow 48-digit recovery password' | 4.11.7.1.3 | CIS | Bitlocker |"
        )
        [System.IO.File]::WriteAllLines((Join-Path $temporaryFolder 'Bitlocker.md'), $bitlockerMarkdown)
        $bitlockerParentId = 'device_vendor_msft_policy_config_bitlocker'
        $bitlockerRows = @(
            (New-TestRow 'Choose how BitLocker-protected fixed drives can be recovered' $bitlockerParentId 'Bitlocker')
            (New-TestRow 'Configure user storage of BitLocker recovery information:' "${bitlockerParentId}_recovery" 'Bitlocker')
            (New-TestRow 'Another recovery option' "${bitlockerParentId}_other" 'Bitlocker')
        )
        $bitlockerRows[1].'Expected Value' = 'Allow 48-digit recovery password'
        $bitlockerRows[2].'Expected Value' = 'Enabled'
        $fixture = $shuffled + @(New-TestRow 'Unrelated root' 'different') + $bitlockerRows
        $fixture | Export-Csv -Path $csvPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
        $templateColors = Get-TestWorkbookSummaryColors -Path (Join-Path (Split-Path $ScriptPath -Parent) 'Resources\Customer-SecurityBaselineSettings-Overview.xlsx')
        $export = Export-BaselineAssessmentWorkbook -CsvPath $csvPath -MarkdownFolder $temporaryFolder -OutputPath $outputPath
        $exportColors = Get-TestWorkbookSummaryColors -Path $outputPath
        Assert-Condition ($exportColors.ActiveSheet -eq 'All Settings' -and $exportColors.SelectedSheetCount -eq 1) 'The exported workbook must open with only All Settings active.'
        foreach ($address in @('F3', 'F4')) {
            Assert-Condition ($exportColors[$address] -eq $templateColors[$address]) "$address stored/displayed fill, pattern, font color or merged range differs from the resource template."
        }
        Write-Output 'PASS: F3 and F4 stored/displayed fills, patterns, font colors and merged ranges match the resource template.'
        $package = Open-ExcelPackage -Path $outputPath
        $sheet = $package.Workbook.Worksheets['All Settings']
        $headerRow = 0
        for ($rowIndex = 1; $rowIndex -le 30; $rowIndex++) {
            if ($sheet.Cells[$rowIndex, 2].Text -eq 'Setting Name') { $headerRow = $rowIndex; break }
        }
        Assert-Condition ($headerRow -gt 0 -and $export.RowCount -eq 10) 'Workbook header or row count failed.'
        $firstDataRow = $headerRow + 1
        Assert-Condition ($sheet.Cells[$firstDataRow, 2].Text -eq 'Point and Print Restrictions') 'Workbook parent order failed.'
        $rootIndent = $sheet.Cells[$firstDataRow, 2].Style.Indent
        for ($offset = 1; $offset -le 5; $offset++) {
            $cell = $sheet.Cells[($firstDataRow + $offset), 2]
            Assert-Condition ($cell.Style.Indent -eq $rootIndent + 2 -and $cell.Style.Font.Italic) 'Workbook child formatting failed.'
        }
        $referenceStatusColumn = 0
        for ($column = 1; $column -le $sheet.Dimension.End.Column; $column++) {
            if ($sheet.Cells[$headerRow, $column].Text -eq 'CIS Reference Status') { $referenceStatusColumn = $column; break }
        }
        Assert-Condition ($referenceStatusColumn -eq 0) 'CIS Reference Status column should not be exported.'
        Assert-Condition ([string]::IsNullOrEmpty($sheet.Cells['F3'].Text) -and
            [string]::IsNullOrEmpty($sheet.Cells['F4'].Text) -and
            [string]::IsNullOrEmpty($sheet.Cells['F4'].Formula)) 'The extra F4 counter or its label was retained.'
        for ($column = 1; $column -le 5; $column++) {
            Assert-Condition (-not [string]::IsNullOrEmpty(($sheet.Cells[4, $column].Formula))) 'An existing summary counter was removed.'
        }
        Assert-Condition ($sheet.Cells[$headerRow, 6].Text -eq 'Assignment status' -and
            $sheet.Cells[$firstDataRow, 6].Text -eq 'Yes' -and
            $sheet.AutoFilterAddress.End.Column -eq 6) 'Assignment status or its filter was lost.'
        $deviations = $package.Workbook.Worksheets['Deviations']
        Assert-Condition ($deviations.Dimension.End.Row -eq 11 -and $deviations.Cells['F8'].Text -eq 'Deviation comment') 'Deviations layout was not trimmed at row 12.'
        for ($rowIndex = 9; $rowIndex -le 11; $rowIndex++) {
            for ($column = 1; $column -le 9; $column++) {
                Assert-Condition ([string]::IsNullOrEmpty(($deviations.Cells[$rowIndex, $column].Text)) -and
                    [string]::IsNullOrEmpty(($deviations.Cells[$rowIndex, $column].Formula))) 'Deviations should remain blank for manual entry.'
            }
        }
        $expectedReferences = @('4.7.10', '4.7.10', '4.7.10', '4.7.11', '4.7.10', '4.7.10')
        for ($offset = 0; $offset -le 5; $offset++) {
            Assert-Condition ($sheet.Cells[($firstDataRow + $offset), 4].Text -eq $expectedReferences[$offset]) 'Direct child references were lost or missing values were not inherited.'
        }
        Assert-Condition ($export.MatchedCount -eq 9 -and $export.NoMatchCount -eq 1) 'Reference counts were not recalculated after inheritance.'
        $root = $sheet.Cells[($firstDataRow + 6), 2]
        Assert-Condition ($root.Text -eq 'Unrelated root' -and $root.Style.Indent -eq $rootIndent -and -not $root.Style.Font.Italic) 'Root formatting was not reset.'
        Assert-Condition ($root.Style.Font.Color.Rgb -eq $sheet.Cells[$firstDataRow, 2].Style.Font.Color.Rgb) 'Root font color was not reset.'
        $bitlockerStartRow = $firstDataRow + 7
        Assert-Condition ($sheet.Cells[$bitlockerStartRow, 4].Text -eq '4.11.7.1.1') 'BitLocker parent reference was incorrect.'
        Assert-Condition ($sheet.Cells[($bitlockerStartRow + 1), 4].Text -eq '4.11.7.1.3') 'Expected-value matching did not preserve the BitLocker child reference.'
        Assert-Condition ($sheet.Cells[($bitlockerStartRow + 2), 4].Text -eq '4.11.7.1.1') 'Unmatched BitLocker child did not inherit its parent reference.'
        Write-Output 'PASS: generated workbook ordering, references, counts, nesting styles, removed counter/column, and manual Deviations layout.'

        $customTemplatePath = Join-Path $temporaryFolder 'legacy-template.xlsx'
        $customOutputPath = Join-Path $temporaryFolder 'custom-assessment.xlsx'
        $sheet.InsertColumn(6, 1)
        $sheet.Cells[$headerRow, 6].Value = 'CIS Reference Status'
        $sheet.Cells[$firstDataRow, 6].Value = 'Matched'
        $sheet.Cells['F3'].Value = 'Conflict'
        $sheet.Cells['F4'].Formula = '1+1'
        $package.Workbook.Worksheets['Deviations'].Select()
        $package.SaveAs([System.IO.FileInfo]::new($customTemplatePath))
        $package.Dispose()
        $package = $null
        $customTemplateView = Get-TestWorkbookSummaryColors -Path $customTemplatePath
        Assert-Condition ($customTemplateView.ActiveSheet -eq 'Deviations') 'The supplied-template fixture must start on Deviations.'
        $customExport = Export-BaselineAssessmentWorkbook -CsvPath $csvPath -MarkdownFolder $temporaryFolder -OutputPath $customOutputPath -TemplatePath $customTemplatePath
        $customExportView = Get-TestWorkbookSummaryColors -Path $customOutputPath
        Assert-Condition ($customExportView.ActiveSheet -eq 'All Settings' -and $customExportView.SelectedSheetCount -eq 1) 'Export did not switch the supplied template from Deviations to All Settings.'
        Write-Output 'PASS: Excel opens normal and supplied-template exports on All Settings, with no grouped sheet selection.'
        $package = Open-ExcelPackage -Path $customOutputPath
        $sheet = $package.Workbook.Worksheets['All Settings']
        Assert-Condition ($sheet.Dimension.End.Column -eq 6 -and $sheet.Cells[$headerRow, 6].Text -eq 'Assignment status' -and
            $sheet.Cells[$firstDataRow, 6].Text -eq 'Yes' -and $sheet.AutoFilterAddress.End.Column -eq 6 -and
            [string]::IsNullOrEmpty($sheet.Cells['F3'].Text) -and [string]::IsNullOrEmpty($sheet.Cells['F4'].Formula) -and
            $customExport.RowCount -eq $export.RowCount -and $customExport.MatchedCount -eq $export.MatchedCount) 'Existing reference-status column was not removed safely from a supplied template.'
        Write-Output 'PASS: Excel opens the generated template; legacy column removal preserves assignment data, filters, and reference counts.'
    } finally {
        if ($package) { $package.Dispose() }
        Remove-Item -LiteralPath $temporaryFolder -Recurse -Force
    }
}