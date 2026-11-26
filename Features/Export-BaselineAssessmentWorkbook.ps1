#Requires -Version 5.1
<#
.SYNOPSIS
    Exports baseline assessment details to the existing workbook template.
.EXAMPLE
    .\Export-BaselineAssessmentWorkbook.ps1 -CsvPath .\details.csv -MarkdownFolder .\Benchmarks -OutputPath .\assessment.xlsx
.NOTES
    Requires ImportExcel and Microsoft Excel. Dot-source to load the export function only.
#>
[CmdletBinding()]
param(
    [string]$CsvPath,
    [string]$MarkdownFolder,
    [string]$OutputPath,
    [ValidateRange(0.0,1.0)]
    [double]$SimilarityThreshold = 0.9,
    [string]$TemplatePath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Resources\Customer-SecurityBaselineSettings-Overview.xlsx')
)

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Shared\BaselineToolkit.Common.psm1') -DisableNameChecking

function Export-BaselineAssessmentWorkbook {
    param(
        [string]$CsvPath,
        [string]$MarkdownFolder,
        [string]$OutputPath,
        [ValidateRange(0.0,1.0)]
        [double]$SimilarityThreshold = 0.9,
        [string]$TemplatePath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Resources\Customer-SecurityBaselineSettings-Overview.xlsx')
    )

    if ([System.IO.Path]::GetExtension($OutputPath) -ine '.xlsx') {
        throw 'The output file must use the .xlsx extension.'
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw 'Excel export requires the ImportExcel PowerShell module. Install it with: Install-Module ImportExcel -Scope CurrentUser'
    }
    Import-Module ImportExcel -ErrorAction Stop

    $importedRows = @(Import-BaselineAssessmentDetailsCsv -Path $CsvPath)
    $mergeResult = Merge-BaselineAssessmentDetailRows -Rows $importedRows
    $csvRows = @($mergeResult.Rows)
    $cisLookup = New-BaselineCisReferenceLookup -MarkdownFolder $MarkdownFolder
    $hierarchy = @(Get-BaselineAssessmentHierarchy -Rows $csvRows -Lookup $cisLookup)
    $csvRows = @($hierarchy | ForEach-Object { $_.Row })
    if (-not (Test-Path -Path $templatePath -PathType Leaf)) {
        throw "Workbook template not found: $templatePath"
    }

    $templateFullPath = [System.IO.Path]::GetFullPath($templatePath)
    $outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
    if ($templateFullPath -ieq $outputFullPath) {
        throw 'Choose a new output path; the bundled workbook template is never overwritten.'
    }
    $outputDirectory = Split-Path -Path $outputFullPath -Parent
    if (-not (Test-Path -Path $outputDirectory -PathType Container)) {
        throw "Output folder does not exist: $outputDirectory"
    }

    $referenceResults = [System.Collections.Generic.List[object]]::new()
    for ($csvIndex = 0; $csvIndex -lt $csvRows.Count; $csvIndex++) {
        $csvRow = $csvRows[$csvIndex]
        $resolved = Resolve-BaselineCisReference -Lookup $cisLookup -CsvRow $csvRow -SettingPath $hierarchy[$csvIndex].Path -SimilarityThreshold $SimilarityThreshold
        [void]$referenceResults.Add($resolved)
    }

    $workingPath = Join-Path $outputDirectory ('.{0}.{1}.xlsx' -f [System.IO.Path]::GetFileNameWithoutExtension($outputFullPath), [guid]::NewGuid().ToString('N'))
    $excel = $null
    $templateWorkbook = $null
    $package = $null
    $saveCompleted = $false
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $templateWorkbook = $excel.Workbooks.Open($templateFullPath, 0, $true)
        $templateWorkbook.SaveCopyAs($workingPath)
        $templateWorkbook.Close($false)
        $templateWorkbook = $null
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        $excel = $null

        $package = Open-ExcelPackage -Path $workingPath -ErrorAction Stop
        $settingsSheet = $package.Workbook.Worksheets['All Settings']
        if (-not $settingsSheet) { throw "Workbook template is missing the 'All Settings' sheet." }

        $templateLastRow = if ($settingsSheet.Dimension) { [int]$settingsSheet.Dimension.End.Row } else { 0 }
        $templateLastColumn = if ($settingsSheet.Dimension) { [int]$settingsSheet.Dimension.End.Column } else { 0 }
        $headerRow = 0
        $columnByName = @{}
        for ($rowIndex = 1; $rowIndex -le [Math]::Min(30, $templateLastRow); $rowIndex++) {
            $candidateColumns = @{}
            for ($columnIndex = 1; $columnIndex -le $templateLastColumn; $columnIndex++) {
                $headerText = [string]$settingsSheet.Cells[$rowIndex, $columnIndex].Text
                if ($headerText.Trim()) { $candidateColumns[$headerText.Trim()] = $columnIndex }
            }
            if (@('Baseline Policy','Setting Name','Expected Value','CIS Reference','Status','Assignment status' | Where-Object { $candidateColumns.ContainsKey($_) }).Count -eq 6) {
                $headerRow = $rowIndex
                $columnByName = $candidateColumns
                break
            }
        }
        if ($headerRow -eq 0) {
            throw "Could not find the expected settings header row on the 'All Settings' sheet."
        }

        if ($columnByName.ContainsKey('CIS Reference Status')) {
            $referenceStatusColumn = [int]$columnByName['CIS Reference Status']
            $settingsSheet.DeleteColumn($referenceStatusColumn)
            $columnByName.Remove('CIS Reference Status')
            foreach ($columnName in @($columnByName.Keys)) {
                if ($columnByName[$columnName] -gt $referenceStatusColumn) { $columnByName[$columnName]-- }
            }
            $templateLastColumn--
        }

        # Reuse curated template references for settings the Markdown files cannot resolve unambiguously.
        $templateReferences = @{}
        $getTemplateKeys = {
            param([string]$Policy, [string]$Setting, [string]$Value)
            $policyKey = Get-BaselineExportPolicyKey -PolicyName ($Policy -replace '^\s*NotDeployed-', '')
            $settingKey = Get-BaselineExportSettingKey -SettingName $Setting
            if (-not $policyKey -or -not $settingKey) { return @() }
            @(('V|{0}|{1}|{2}' -f $policyKey, $settingKey, (Normalize-CompareText -Text $Value)), ('S|{0}|{1}' -f $policyKey, $settingKey))
        }
        for ($rowIndex = $headerRow + 1; $rowIndex -le $templateLastRow; $rowIndex++) {
            $templateReference = ([string]$settingsSheet.Cells[$rowIndex, $columnByName['CIS Reference']].Text).Trim()
            if (-not $templateReference) { continue }
            $keys = & $getTemplateKeys ([string]$settingsSheet.Cells[$rowIndex, $columnByName['Baseline Policy']].Text) ([string]$settingsSheet.Cells[$rowIndex, $columnByName['Setting Name']].Text) ([string]$settingsSheet.Cells[$rowIndex, $columnByName['Expected Value']].Text)
            foreach ($key in $keys) {
                if (-not $templateReferences.ContainsKey($key)) {
                    $templateReferences[$key] = $templateReference
                } elseif ($templateReferences[$key] -ne $templateReference) {
                    $templateReferences[$key] = $null
                }
            }
        }
        for ($csvIndex = 0; $csvIndex -lt $csvRows.Count; $csvIndex++) {
            $resolved = $referenceResults[$csvIndex]
            if ($resolved.Status -notin @('No match', 'Ambiguous')) { continue }
            $csvRow = $csvRows[$csvIndex]
            foreach ($key in (& $getTemplateKeys ([string]$csvRow.'Baseline Policy') ([string]$csvRow.'Setting Name') ([string]$csvRow.'Expected Value'))) {
                if ($templateReferences.ContainsKey($key)) {
                    if ($templateReferences[$key]) {
                        $resolved.CISReference = $templateReferences[$key]
                        $resolved.Status = 'Template'
                    }
                    break
                }
            }
        }
        Set-BaselineAssessmentCisReferenceInheritance -Hierarchy $hierarchy -ReferenceResults $referenceResults
        $matchedCount = @($referenceResults | Where-Object { $_.Status -in @('Matched','Matched (other policy)','Inherited') }).Count
        $twodayCount = @($referenceResults | Where-Object { $_.Status -eq 'Twoday' }).Count
        $ambiguousCount = @($referenceResults | Where-Object { $_.Status -eq 'Ambiguous' }).Count
        $templateCount = @($referenceResults | Where-Object { $_.Status -eq 'Template' }).Count

        $firstDataRow = $headerRow + 1
        $lastDataRow = $headerRow + $csvRows.Count
        $settingNameColumn = [int]$columnByName['Setting Name']
        $baseSettingIndent = [int]$settingsSheet.Cells[$firstDataRow, $settingNameColumn].Style.Indent
        $baseSettingNameStyleId = [int]$settingsSheet.Cells[$firstDataRow, $settingNameColumn].StyleID
        $clearThroughRow = [Math]::Max($templateLastRow, $lastDataRow)
        for ($rowIndex = $firstDataRow; $rowIndex -le $clearThroughRow; $rowIndex++) {
            for ($columnIndex = 1; $columnIndex -le $templateLastColumn; $columnIndex++) {
                $settingsSheet.Cells[$rowIndex, $columnIndex].Value = $null
            }
        }

        for ($rowIndex = $firstDataRow; $rowIndex -le $lastDataRow; $rowIndex++) {
            if ($rowIndex -gt $templateLastRow) {
                for ($columnIndex = 1; $columnIndex -le $templateLastColumn; $columnIndex++) {
                    $styleId = if ($columnIndex -eq $settingNameColumn) { $baseSettingNameStyleId } else { [int]$settingsSheet.Cells[$firstDataRow, $columnIndex].StyleID }
                    $settingsSheet.Cells[$rowIndex, $columnIndex].StyleID = $styleId
                }
            }

            $csvIndex = $rowIndex - $firstDataRow
            $csvRow = $csvRows[$csvIndex]
            $resolved = $referenceResults[$csvIndex]
            $settingNameCell = $settingsSheet.Cells[$rowIndex, $settingNameColumn]
            $settingDepth = $hierarchy[$csvIndex].Depth
            $settingNameCell.StyleID = $baseSettingNameStyleId
            $settingsSheet.Cells[$rowIndex, $columnByName['Baseline Policy']].Value = [string]$csvRow.'Baseline Policy'
            $settingNameCell.Value = [string]$csvRow.'Setting Name'
            $settingNameCell.Style.Indent = [Math]::Min(250, $baseSettingIndent + ($settingDepth * 2))
            if ($settingDepth -gt 0) {
                $settingNameCell.Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
                $settingNameCell.Style.Font.Italic = $true
            }
            $settingsSheet.Cells[$rowIndex, $columnByName['Expected Value']].Value = [string]$csvRow.'Expected Value'
            $settingsSheet.Cells[$rowIndex, $columnByName['CIS Reference']].Value = [string]$resolved.CISReference
            $settingsSheet.Cells[$rowIndex, $columnByName['Status']].Value = [string]$csvRow.Status
            $settingsSheet.Cells[$rowIndex, $columnByName['Assignment status']].Value = [string]$csvRow.'Policy Assigned'
        }

        $lastSummaryRow = [Math]::Max($firstDataRow, $lastDataRow)
        $policyRange = '$A${0}:$A${1}' -f $firstDataRow, $lastSummaryRow
        $settingsRange = '$B${0}:$B${1}' -f $firstDataRow, $lastSummaryRow
        $statusRange = '$E${0}:$E${1}' -f $firstDataRow, $lastSummaryRow
        $settingsSheet.Cells['A4'].Formula = 'SUMPRODUCT((SUBTOTAL(103,OFFSET($A${0},ROW({1})-ROW($A${0}),0,1))>0)*({1}<>"")/COUNTIF({1},{1}&""))' -f $firstDataRow, $policyRange
        $settingsSheet.Cells['B4'].Formula = 'SUBTOTAL(103,{0})' -f $settingsRange
        $settingsSheet.Cells['C4'].Formula = 'SUMPRODUCT(SUBTOTAL(103,OFFSET($E${0},ROW({1})-ROW($E${0}),0,1))*({1}="NotFound"))' -f $firstDataRow, $statusRange
        $settingsSheet.Cells['D4'].Formula = 'SUMPRODUCT(SUBTOTAL(103,OFFSET($E${0},ROW({1})-ROW($E${0}),0,1))*({1}="Match"))' -f $firstDataRow, $statusRange
        $settingsSheet.Cells['E4'].Formula = 'SUMPRODUCT(SUBTOTAL(103,OFFSET($E${0},ROW({1})-ROW($E${0}),0,1))*({1}="Mismatch"))' -f $firstDataRow, $statusRange
        $settingsSheet.Cells[$headerRow, 1, $lastSummaryRow, $templateLastColumn].AutoFilter = $true

        $deviationsSheet = $package.Workbook.Worksheets['Deviations']
        if ($deviationsSheet -and $deviationsSheet.Dimension) {
            if ($deviationsSheet.Dimension.End.Row -ge 12) {
                $deviationsSheet.DeleteRow(12, $deviationsSheet.Dimension.End.Row - 11)
            }
            for ($rowIndex = 9; $rowIndex -le $deviationsSheet.Dimension.End.Row; $rowIndex++) {
                for ($columnIndex = 1; $columnIndex -le $deviationsSheet.Dimension.End.Column; $columnIndex++) {
                    $deviationsSheet.Cells[$rowIndex, $columnIndex].Value = $null
                }
            }
        }

        $package.Workbook.CalcMode = 'Automatic'
        $package.Workbook.FullCalcOnLoad = $true
        $settingsSheet.Select()
        Close-ExcelPackage -ExcelPackage $package -SaveAs $workingPath
        $package = $null
        Move-Item -Path $workingPath -Destination $outputFullPath -Force
        $saveCompleted = $true
    } finally {
        if ($templateWorkbook) {
            try { $templateWorkbook.Close($false) } catch {}
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch {}
        }
        if ($package) {
            try { Close-ExcelPackage -ExcelPackage $package -NoSave } catch {}
        }
        if (-not $saveCompleted -and (Test-Path -Path $workingPath -PathType Leaf)) {
            Remove-Item -Path $workingPath -Force -ErrorAction SilentlyContinue
        }
    }

    return [pscustomobject]@{
        OutputPath = $outputFullPath
        RowCount = $csvRows.Count
        MatchedCount = $matchedCount
        TwodayCount = $twodayCount
        NoMatchCount = @($referenceResults | Where-Object { $_.Status -eq 'No match' }).Count
        AmbiguousCount = $ambiguousCount
        TemplateCount = $templateCount
        DuplicatesRemoved = $mergeResult.DuplicatesRemoved
        ConflictCount = @($csvRows | Where-Object { [string]$_.Status -eq 'Conflict' }).Count
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    foreach ($requiredName in @('CsvPath', 'MarkdownFolder', 'OutputPath')) {
        if ([string]::IsNullOrWhiteSpace((Get-Variable -Name $requiredName -ValueOnly))) {
            throw "-$requiredName is required. See Get-Help $PSCommandPath -Examples."
        }
    }
    Export-BaselineAssessmentWorkbook -CsvPath $CsvPath -MarkdownFolder $MarkdownFolder -OutputPath $OutputPath -SimilarityThreshold $SimilarityThreshold -TemplatePath $TemplatePath
}
