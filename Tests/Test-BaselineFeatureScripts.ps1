#Requires -Version 5.1
[CmdletBinding()]
param([switch]$VerifyGui)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$initialWpfAssemblies = @([AppDomain]::CurrentDomain.GetAssemblies().GetName().Name | Where-Object { $_ -in @('PresentationFramework','PresentationCore','WindowsBase') })
$basePath = Split-Path $PSScriptRoot -Parent
$featuresPath = Join-Path $basePath 'Features'
$sharedPath = Join-Path $basePath 'Shared'
$featureFiles = @(
    'Invoke-BaselinePdfConversion.ps1'
    'Export-ChangeHistory.ps1'
    'Convert-MarkdownTextToInteractiveHtmlReport.ps1'
    'Invoke-BaselineMerger.ps1'
    'Invoke-IntuneCISPolicyCreator.ps1'
    'Invoke-IntuneBaselineReport.ps1'
    'Invoke-BaselineAssessment.ps1'
    'Export-BaselineAssessmentWorkbook.ps1'
)
foreach ($fileName in ($featureFiles + @('BaselineToolkit.Common.psm1', 'Invoke-BaselineToolbox.ps1'))) {
    $tokens = $null
    $parseErrors = $null
    $scriptDirectory = if ($fileName -eq 'BaselineToolkit.Common.psm1') { $sharedPath } elseif ($fileName -eq 'Invoke-BaselineToolbox.ps1') { $basePath } else { $featuresPath }
    $tree = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $scriptDirectory $fileName), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw ($parseErrors.Message -join "`n") }
    if ($fileName -in @('BaselineToolkit.Common.psm1', 'Invoke-BaselinePdfConversion.ps1', 'Invoke-BaselineMerger.ps1', 'Invoke-IntuneBaselineReport.ps1', 'Invoke-BaselineAssessment.ps1', 'Export-BaselineAssessmentWorkbook.ps1')) {
        $controlReferences = @($tree.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.VariableExpressionAst] -and $node.VariablePath.UserPath -ne 'txtPath' -and $node.VariablePath.UserPath -match '^(txt[A-Z]|btn[A-Z]|chk[A-Z]|cmb[A-Z]|dg[A-Z]|window$)'
        }, $true))
        if ($controlReferences.Count) { throw "GUI dependency in ${fileName}: $($controlReferences.Extent.Text -join ', ')" }
    }
}
Write-Output 'PASS: feature syntax and no toolbox-control dependencies in extracted processing.'

Import-Module (Join-Path $sharedPath 'BaselineToolkit.Common.psm1') -Force -DisableNameChecking
. (Join-Path $featuresPath 'Invoke-BaselineAssessment.ps1')
. (Join-Path $featuresPath 'Invoke-IntuneBaselineReport.ps1')

$temporaryFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('Baseline features ' + [guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($temporaryFolder)
try {
    $newPath = Join-Path $temporaryFolder 'new.md'
    $oldPath = Join-Path $temporaryFolder 'old.md'
    $outputPath = Join-Path $temporaryFolder 'merged.md'
    $header = @(
        '# Fixture'
        '## 1 Settings'
        '| Recommendation | Twoday ID | CIS Reference | Profile | Assessment | Value | Origin | Followed | Notes | Intune Policy |'
        '| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |'
    )
    $recommendation = "Ensure 'Fixture Setting' is set to 'Enabled'"
    ($header + "| $recommendation | - | 1.1 | L1 | Automated | Enabled | CIS | Yes | | Fixture |") | Set-Content -LiteralPath $newPath -Encoding UTF8
    ($header + "| $recommendation | - | 1.1 | L1 | Automated | Enabled | CIS | No | retained-note | Fixture |") | Set-Content -LiteralPath $oldPath -Encoding UTF8
    $preview = & (Join-Path $featuresPath 'Invoke-IntuneCISPolicyCreator.ps1') -MdFilePath $newPath -PreviewOnly -PreviewOrigin CIS
    if ($preview.TotalCount -ne 1 -or $preview.Recommendations.Count -ne 1 -or $preview.PolicyCount -ne 1) { throw 'Standalone policy preview did not preserve the GUI filtering rules.' }
    Write-Output 'PASS: standalone policy preview without Graph authentication.'
    Push-Location $temporaryFolder
    try {
        $result = & (Join-Path $featuresPath 'Invoke-BaselineMerger.ps1') -NewBaselinePath $newPath -OldBaselinePath $oldPath -OutputPath $outputPath -GenerateReport
        if (-not (Test-Path -LiteralPath $result.HtmlPath)) { throw 'Merge report was not created.' }
        if ((Get-Content -LiteralPath $outputPath -Raw) -notlike '*retained-note*') { throw 'Merge lost the old baseline notes.' }
        $reportOnlyPath = Join-Path $temporaryFolder 'report only.html'
        $report = & (Join-Path $featuresPath 'Invoke-BaselineMerger.ps1') -ReportOnly -InputMarkdownPath $outputPath -OutputHtmlPath $reportOnlyPath
        if (-not (Test-Path -LiteralPath $report.HtmlPath)) { throw 'Report-only did not create its output.' }
    } finally { Pop-Location }
    Write-Output 'PASS: standalone merge, retained notes, HTML and report-only from another directory with spaces.'

    $baselinePath = Join-Path $temporaryFolder 'baseline.json'
    @{
        name = 'Fixture'
        settings = @(@{
            settingInstance = @{
                '@odata.type' = '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance'
                settingDefinitionId = 'device_fixture'
                choiceSettingValue = @{ value = 'device_fixture_1'; children = @() }
            }
        })
    } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $baselinePath -Encoding UTF8
    $rows = @([pscustomobject]@{ DefinitionId = 'device_fixture'; ProfileId = 'fixture-id'; ProfileName = 'Fixture'; SettingName = 'Fixture Setting'; RawValue = 'device_fixture_1' })
    $assessment = Invoke-BaselineAssessment -BaselineFolder $temporaryFolder -SettingRows $rows
    if ($assessment.TotalBaselineFiles -ne 1 -or $assessment.Baselines[0].Classification -ne 'FullyAligned') { throw 'Offline assessment parity failed.' }
    $htmlPath = Join-Path $temporaryFolder 'assessment.html'
    New-BaselineAssessmentHtmlReport -OutputPath $htmlPath -AssessmentData $assessment -SettingRows $rows | Out-Null
    if (-not (Test-Path -LiteralPath $htmlPath)) { throw 'Assessment report was not created.' }
    Write-Output 'PASS: assessment matching and HTML generation without Graph or WPF.'

    $previousLocalAppData = $env:LOCALAPPDATA
    try {
        $env:LOCALAPPDATA = $temporaryFolder
        function Invoke-MgGraphRequest {
            param($Uri, $Method = 'Get')
            if ($Method -ine 'Get') { throw "Unexpected Graph write: $Method $Uri" }
            if ($Uri -like '*/settings*') {
                return @{ value = @(@{
                    settingInstance = @{ settingDefinitionId = 'device_fixture'; choiceSettingValue = @{ value = 'device_fixture_1'; children = @() } }
                    settingDefinitions = @(@{ id = 'device_fixture'; displayName = 'Fixture Setting'; options = @(@{ itemId = 'device_fixture_1'; displayName = 'Enabled' }) })
                }) }
            }
            if ($Uri -like '*groupPolicyConfigurations/admx-id/definitionValues*') {
                return @{ value = @(@{
                    id = 'dv-1'; enabled = $true
                    definition = @{ id = 'def-1'; displayName = 'Fixture ADMX Setting'; categoryPath = '\Windows Components\Fixture'; classType = 'machine' }
                    presentationValues = @(
                        @{ '@odata.type' = '#microsoft.graph.groupPolicyPresentationValueText'; value = 'abc'; presentation = @{ id = 'pres-1'; label = 'Fixture text' } },
                        @{ '@odata.type' = '#microsoft.graph.groupPolicyPresentationValueList'; values = @(@{ name = 'b'; value = '2' }, @{ name = 'a'; value = '1' }); presentation = @{ id = 'pres-2'; label = 'Fixture list' } }
                    )
                }) }
            }
            if ($Uri -like '*configurationPolicies?*') { return @{ value = @(@{ id = 'fixture-id'; name = 'Fixture'; platforms = 'windows10'; lastModifiedDateTime = '2026-10-30' }) } }
            if ($Uri -like '*deviceConfigurations?*') { return @{ value = @() } }
            if ($Uri -like '*groupPolicyConfigurations?*') { return @{ value = @(@{ id = 'admx-id'; displayName = 'Fixture ADMX'; lastModifiedDateTime = '2026-10-30' }) } }
            if ($Uri -like '*/organization*') { return @{ value = @(@{ displayName = 'Fixture Tenant' }) } }
            throw "Unexpected Graph query: $Uri"
        }
        $commonModule = Get-Module BaselineToolkit.Common
        & $commonModule {
            param($mockRequest)
            Set-Item Function:script:Invoke-MgGraphRequest -Value $mockRequest
        } (Get-Item Function:Invoke-MgGraphRequest).ScriptBlock
        Set-BaselineResolverTenant -TenantId 'fixture-tenant'
        $allProfiles = @((Get-BaselineIntuneProfiles).Mapped)
        $profiles = @($allProfiles | Where-Object PolicyKind -eq 'SettingsCatalog')
        if ($profiles.Count -ne 1 -or $profiles[0].Id -ne 'fixture-id') { throw 'Profile discovery failed.' }
        $admxProfiles = @($allProfiles | Where-Object PolicyKind -eq 'AdministrativeTemplate')
        if ($allProfiles.Count -ne 2 -or $admxProfiles.Count -ne 1 -or $admxProfiles[0].Id -ne 'admx-id' -or
            $admxProfiles[0].OdataType -ne '#microsoft.graph.groupPolicyConfiguration') { throw 'Administrative Template discovery failed.' }
        Write-Output 'PASS: Load profiles lists Administrative Templates.'
        $profileReportPath = Join-Path $temporaryFolder 'profiles.html'
        $profileReport = Invoke-BaselineProfileReport -TargetProfiles $profiles -OutputPath $profileReportPath -CacheMode cacheonly
        if ($profileReport.FailedProfiles.Count -or $profileReport.SettingRows.Count -ne 1 -or -not (Test-Path -LiteralPath $profileReportPath)) { throw 'Mocked profile report failed.' }
        $admxReportPath = Join-Path $temporaryFolder 'admx profiles.html'
        $admxReport = Invoke-BaselineProfileReport -TargetProfiles $admxProfiles -OutputPath $admxReportPath -CacheMode cacheonly
        $admxRows = @($admxReport.SettingRows)
        if ($admxReport.FailedProfiles.Count -or $admxRows.Count -ne 3 -or
            @($admxRows | Where-Object { $_.DefinitionId -eq 'AdminTemplate|def-1' -and $_.RawValue -eq 'Enabled' }).Count -ne 1 -or
            @($admxRows | Where-Object { $_.DefinitionId -eq 'AdminTemplate|def-1|pres-1' -and $_.RawValue -eq 'abc' }).Count -ne 1 -or
            @($admxRows | Where-Object { $_.DefinitionId -eq 'AdminTemplate|def-1|pres-2' -and $_.RawValue -eq 'a=1; b=2' }).Count -ne 1) { throw 'Administrative Template settings were not collected for the profile report.' }
        if ((Get-Content -LiteralPath $admxReportPath -Raw) -notlike '*Fixture ADMX Setting*') { throw 'Administrative Template settings missing from report HTML.' }
        $admxFolder = Join-Path $temporaryFolder 'admx-backups'
        [void](New-Item -ItemType Directory -Path $admxFolder)
        $admxBackupPath = Join-Path $admxFolder 'admx.json'
        @{
            id = 'admx-id'; displayName = 'Fixture ADMX'; '@odata.type' = '#microsoft.graph.groupPolicyConfiguration'
            definitionValues = @(@{
                enabled = $true
                'definition@odata.bind' = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyDefinitions('def-1')"
                presentationValues = @(
                    @{ '@odata.type' = '#microsoft.graph.groupPolicyPresentationValueText'; value = 'abc'; 'presentation@odata.bind' = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyDefinitions('def-1')/presentations('pres-1')" },
                    @{ '@odata.type' = '#microsoft.graph.groupPolicyPresentationValueList'; values = @(@{ name = 'a'; value = '1' }, @{ name = 'b'; value = '2' }); 'presentation@odata.bind' = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyDefinitions('def-1')/presentations('pres-2')" }
                )
            })
        } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $admxBackupPath -Encoding UTF8
        $parsedAdmx = Get-BaselineSettingsFromJson -JsonPath $admxBackupPath -PolicyReport
        if ($parsedAdmx.PolicyType -ne '#microsoft.graph.groupPolicyConfiguration' -or $parsedAdmx.Settings.Count -ne 3) { throw 'Administrative Template backup parsing failed.' }
        if (@((Get-BaselineSettingsFromJson -JsonPath $admxBackupPath).Settings).Count) { throw 'Assessment parsing should not map Administrative Template settings.' }
        $comparison = Compare-IntunePolicyBackup -CompareFolder $admxFolder -Profiles $admxProfiles -SettingRows $admxRows
        if ($comparison.Policies[0].Status -ne 'Match' -or $comparison.Policies[0].FileName -ne 'admx.json') { throw 'Administrative Template backup comparison did not match live settings.' }
        $changedAdmxRows = @($admxRows | ForEach-Object { $_.PSObject.Copy() })
        ($changedAdmxRows | Where-Object DefinitionId -eq 'AdminTemplate|def-1|pres-1').RawValue = 'xyz'
        $comparison = Compare-IntunePolicyBackup -CompareFolder $admxFolder -Profiles $admxProfiles -SettingRows $changedAdmxRows
        if ($comparison.Policies[0].Status -ne 'Changed' -or @($comparison.Policies[0].Settings | Where-Object Status -eq 'Changed').Count -ne 1) { throw 'Administrative Template backup comparison missed a changed value.' }
        $admxArrayPath = Join-Path $temporaryFolder 'admx-array.json'
        ConvertTo-Json -InputObject @(@{ enabled = $false; definition = @{ id = 'def-2'; displayName = 'Array setting' } }) -Depth 10 | Set-Content -LiteralPath $admxArrayPath -Encoding UTF8
        $parsedArray = Get-BaselineSettingsFromJson -JsonPath $admxArrayPath -PolicyReport
        if ($parsedArray.PolicyType -ne '#microsoft.graph.groupPolicyConfiguration' -or $parsedArray.Settings.Count -ne 1 -or
            $parsedArray.Settings[0].DefinitionId -ne 'AdminTemplate|def-2' -or $parsedArray.Settings[0].RawValue -ne 'Disabled') { throw 'Administrative Template definitionValues-array backup parsing failed.' }
        Remove-Item -LiteralPath $admxArrayPath
        Write-Output 'PASS: Administrative Templates are reported and compared against JSON backups.'
        $compareFolder = Join-Path $temporaryFolder 'policy-backups'
        [void](New-Item -ItemType Directory -Path $compareFolder)
        $compareProfiles = @(
            [pscustomobject]@{ Id = 'policy-a'; Name = 'Renamed A'; PolicyKind = 'SettingsCatalog'; OdataType = '' },
            [pscustomobject]@{ Id = 'policy-b'; Name = 'Policy B'; PolicyKind = 'SettingsCatalog'; OdataType = '' }
        )
        foreach ($policyId in @('policy-a', 'policy-b')) {
            @{ id = $policyId; name = $policyId; settings = @(@{ settingInstance = @{ settingDefinitionId = 'shared'; simpleSettingValue = @{ value = 'backup' } } }) } |
                ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $compareFolder "$policyId.json") -Encoding UTF8
        }
        $compareRows = @(
            [pscustomobject]@{ ProfileId = 'policy-a'; DefinitionId = 'shared'; RawValue = 'changed'; SettingName = 'Shared' },
            [pscustomobject]@{ ProfileId = 'policy-b'; DefinitionId = 'shared'; RawValue = 'backup'; SettingName = 'Shared' },
            [pscustomobject]@{ ProfileId = 'policy-a'; DefinitionId = 'added'; RawValue = '0'; SettingName = 'Added' }
        )
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles $compareProfiles -SettingRows $compareRows
        if ($comparison.Policies[0].Status -ne 'Changed' -or $comparison.Policies[1].Status -ne 'Match' -or
            @($comparison.Policies[0].Settings | Where-Object Status -eq 'Added since backup').Count -ne 1) { throw 'Policy-scoped backup comparison or added-setting detection failed.' }
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles $compareProfiles -SettingRows @()
        if ($comparison.MissingRows.Count -ne 2) { throw 'Settings missing from live were not detected.' }
        Write-Output 'PASS: backup ID matching tolerates renames, isolates policies, and detects added/missing settings.'
        $inventoryHtml = Get-Content -LiteralPath $profileReportPath -Raw
        if ($inventoryHtml -notlike '*<title>Intune Policy Report</title>*' -or $inventoryHtml -like '*id=''btnNonMatching''*') { throw 'Inventory report contains comparison controls or the wrong title.' }
        $fallbackPolicy = [pscustomobject]@{ Id = 'new-id'; Name = 'policy-a'; PolicyKind = 'SettingsCatalog'; OdataType = '' }
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($fallbackPolicy) -SettingRows @()
        if ($comparison.Policies[0].FileName -ne 'policy-a.json' -or $comparison.Policies[0].Status -ne 'Changed') { throw 'Unique name/type fallback failed.' }
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($compareProfiles[0], $fallbackPolicy) -SettingRows @()
        if ($comparison.Policies[1].Status -ne 'No backup found') { throw 'A backup reserved by ID was reused for another policy by name.' }
        $duplicateNamePolicy = [pscustomobject]@{ Id = 'another-id'; Name = 'policy-a'; PolicyKind = 'SettingsCatalog'; OdataType = '' }
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($fallbackPolicy, $duplicateNamePolicy) -SettingRows @()
        if (@($comparison.Policies | Where-Object Status -ne 'Ambiguous').Count) { throw 'Ambiguous selected policy names were guessed.' }
        Copy-Item -LiteralPath (Join-Path $compareFolder 'policy-a.json') -Destination (Join-Path $compareFolder 'duplicate.json')
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles $compareProfiles -SettingRows $compareRows
        if ($comparison.Policies[0].Status -ne 'Ambiguous') { throw 'Duplicate backups were not rejected.' }
        Remove-Item -LiteralPath (Join-Path $compareFolder 'duplicate.json')
        $failed = @([pscustomobject]@{ Id = 'policy-a'; Name = 'Renamed A'; Error = 'Read failed' })
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles $compareProfiles -SettingRows @() -FailedProfiles $failed
        if ($comparison.Policies[0].Status -ne 'Failed' -or $comparison.Policies[0].Settings.Count) { throw 'Failed collection was reported as missing settings.' }
        $unknownPolicy = [pscustomobject]@{ Id = 'unknown'; Name = 'Unknown'; PolicyKind = 'SettingsCatalog'; OdataType = '' }
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($unknownPolicy) -SettingRows @()
        if ($comparison.Policies[0].Status -ne 'No backup found' -or $comparison.IgnoredBackups -ne 2) { throw 'Unmatched policies or unrelated backup counts are incorrect.' }
        '{ invalid json' | Set-Content -LiteralPath (Join-Path $compareFolder 'invalid.json')
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($unknownPolicy) -SettingRows @()
        if ($comparison.Policies[0].Status -ne 'Unverified' -or $comparison.Diagnostics.Count -ne 1) { throw 'Unreadable backups did not produce explicit diagnostics.' }
        Remove-Item -LiteralPath (Join-Path $compareFolder 'invalid.json')
        $dcPath = Join-Path $compareFolder 'arbitrary-backup-name.json'
        @{ id = 'policy-a'; displayName = 'Device policy'; '@odata.type' = '#microsoft.graph.iosGeneralDeviceConfiguration'; passcodeBlockSimple = $false; assignments = @() } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $dcPath -Encoding UTF8
        $parsedDc = Get-BaselineSettingsFromJson -JsonPath $dcPath -PolicyReport
        if ($parsedDc.Settings.Count -ne 1 -or $parsedDc.Settings[0].RawValue -cne 'false') { throw 'Report parsing incorrectly applies baseline filename filters or includes assignments.' }
        Remove-Item -LiteralPath (Join-Path $compareFolder 'policy-a.json')
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($compareProfiles[0]) -SettingRows @()
        if ($comparison.Policies[0].Status -ne 'Type conflict') { throw 'Matching IDs with different policy types were compared.' }
        Remove-Item -LiteralPath $dcPath
        $emptyPath = Join-Path $compareFolder 'empty.json'
        @{ id = 'empty-id'; name = 'Empty'; settings = @() } | ConvertTo-Json | Set-Content -LiteralPath $emptyPath -Encoding UTF8
        $emptyPolicy = [pscustomobject]@{ Id = 'empty-id'; Name = 'Empty'; PolicyKind = 'SettingsCatalog'; OdataType = '' }
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($emptyPolicy) -SettingRows @()
        if ($comparison.Policies[0].Status -ne 'Match') { throw 'Empty matched policies should compare equal.' }
        $repeatPath = Join-Path $compareFolder 'repeated.json'
        @{ id = 'repeated-id'; name = 'Repeated'; settings = @(
            @{ settingInstance = @{ settingDefinitionId = 'repeated'; simpleSettingValue = @{ value = 'one' } } },
            @{ settingInstance = @{ settingDefinitionId = 'repeated'; simpleSettingValue = @{ value = 'two' } } }
        ) } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $repeatPath -Encoding UTF8
        $repeatedPolicy = [pscustomobject]@{ Id = 'repeated-id'; Name = 'Repeated'; PolicyKind = 'SettingsCatalog'; OdataType = '' }
        $repeatedRows = @(
            [pscustomobject]@{ ProfileId = 'repeated-id'; DefinitionId = 'repeated'; RawValue = 'two'; SettingName = '<script>alert(1)</script>' },
            [pscustomobject]@{ ProfileId = 'repeated-id'; DefinitionId = 'repeated'; RawValue = 'one'; SettingName = 'Repeated' }
        )
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($repeatedPolicy) -SettingRows $repeatedRows
        if ($comparison.Policies[0].Status -ne 'Match') { throw 'Repeated leaf values were not preserved as a collection.' }
        $repeatedRows[1].RawValue = 'two'
        $comparison = Compare-IntunePolicyBackup -CompareFolder $compareFolder -Profiles @($repeatedPolicy) -SettingRows $repeatedRows
        if ($comparison.Policies[0].Status -ne 'Changed') { throw 'Repeated leaf values were overwritten, masking a change.' }
        $escapedPath = Join-Path $temporaryFolder 'escaped.html'
        New-IntuneProfileSettingsHtmlReport -OutputPath $escapedPath -Profiles @($repeatedPolicy) -SettingRows $repeatedRows -FailedProfiles @() -BaselineComparison $comparison
        $escapedHtml = Get-Content -LiteralPath $escapedPath -Raw
        if ($escapedHtml -like '*<script>alert(1)</script>*' -or $escapedHtml -notlike '*&lt;script&gt;alert(1)&lt;/script&gt;*') { throw 'Policy comparison HTML failed to escape setting names.' }
        $integrationFolder = Join-Path $temporaryFolder 'report-integration'
        [void](New-Item -ItemType Directory -Path $integrationFolder)
        @{ id = 'fixture-id'; name = 'Fixture'; settings = @(@{ settingInstance = @{ '@odata.type' = '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance'; settingDefinitionId = 'device_fixture'; choiceSettingValue = @{ value = 'device_fixture_0'; children = @() } } }) } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $integrationFolder 'backup.json') -Encoding UTF8
        $driftPath = Join-Path $temporaryFolder 'drift.html'
        $drift = Invoke-BaselineProfileReport -TargetProfiles $profiles -OutputPath $driftPath -CompareFolder $integrationFolder -CacheMode cacheonly
        $driftHtml = Get-Content -LiteralPath $driftPath -Raw
        if ($drift.PolicyComparison.Policies[0].Status -ne 'Changed' -or $driftHtml -notlike '*Backup value*' -or
            $driftHtml -notlike '*Current value*' -or $driftHtml -notlike '*device_fixture_0*' -or
            $driftHtml -like '*Baseline Files*') { throw 'Integrated policy drift HTML or CompareFolder alias failed.' }
        $legacyDrift = Invoke-BaselineProfileReport -TargetProfiles $profiles -OutputPath $driftPath -BaselineFolder $integrationFolder -CacheMode cacheonly
        if ($legacyDrift.PolicyComparison.Policies[0].Status -ne 'Changed') { throw 'BaselineFolder compatibility was lost.' }
        $invalidFolderRejected = $false
        try { Invoke-BaselineProfileReport -TargetProfiles $profiles -OutputPath $driftPath -CompareFolder (Join-Path $temporaryFolder 'missing') -CacheMode cacheonly } catch { $invalidFolderRejected = $_.Exception.Message -like 'Compare folder not found:*' }
        if (-not $invalidFolderRejected) { throw 'Invalid compare folder silently fell back to inventory.' }
        $emptyFolder = Join-Path $temporaryFolder 'empty-compare-folder'
        [void](New-Item -ItemType Directory -Path $emptyFolder)
        $emptyFolderRejected = $false
        try { Invoke-BaselineProfileReport -TargetProfiles $profiles -OutputPath $driftPath -CompareFolder $emptyFolder -CacheMode cacheonly } catch { $emptyFolderRejected = $_.Exception.Message -like 'Compare folder contains no JSON*' }
        if (-not $emptyFolderRejected) { throw 'An empty compare folder silently fell back to inventory.' }
        $unsupportedPath = Join-Path $emptyFolder 'unsupported.json'
        @{ id = 'unknown'; displayName = 'Unsupported'; '@odata.type' = '#microsoft.graph.unsupportedPolicy' } |
            ConvertTo-Json | Set-Content -LiteralPath $unsupportedPath -Encoding UTF8
        $unsupported = Compare-IntunePolicyBackup -CompareFolder $emptyFolder -Profiles @($unknownPolicy) -SettingRows @()
        if ($unsupported.Diagnostics.Count -ne 1 -or $unsupported.Policies[0].Status -ne 'Unverified') { throw 'Unsupported backup format was presented as compared.' }
        Write-Output 'PASS: inventory HTML, name/type fallback, ambiguity, failed/empty/unmatched policies, malformed backups, device backups, comparison HTML and compatible folder parameters.'
        $workerState = [hashtable]::Synchronized(@{
            TargetProfiles = @($profiles + $admxProfiles); BaselineFolder = $temporaryFolder
            OutputPath = (Join-Path $temporaryFolder 'worker.html')
            CustomerPolicyRawSettings = @{}; CustomerPolicyMeta = @{}
            SkipUnassigned = $false; SkipAssignmentLookups = $true
            ProfileAssignments = @{}; TenantName = 'Fixture Tenant'
            ResolverTenantOverride = 'fixture-tenant'; ResolverTenantCacheKey = 'fixture-tenant'
            CacheMode = 'cacheonly'; AutoOpen = $false
            Progress = 0; IsIndeterminate = $false; StatusText = 'Starting...'
            LogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
            Done = $false; Error = $null
        })
        & {
            function Import-Module {
                param($Name)
                if ($Name -ne 'Microsoft.Graph.Authentication') { throw "Unexpected worker dependency: $Name" }
            }
            Invoke-BaselineAssessmentWorker -syncHash $workerState
        }
        if (-not $workerState.Done -or $workerState.Error) { throw "Mocked assessment worker failed: $($workerState.Error)" }
        if (-not (Test-Path -LiteralPath $workerState.OutputPath)) { throw 'Assessment worker report missing.' }
        if (-not @($workerState.LogQueue.ToArray() | Where-Object { $_ -like "Skipping Administrative Template 'Fixture ADMX'*" }).Count) { throw 'Assessment worker did not skip the Administrative Template.' }
        $admxOnlyState = [hashtable]::Synchronized(@{} + $workerState)
        $admxOnlyState.TargetProfiles = $admxProfiles; $admxOnlyState.Done = $false; $admxOnlyState.Error = $null
        $admxOnlyState.OutputPath = Join-Path $temporaryFolder 'worker-admx.html'
        $admxOnlyState.LogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        & {
            function Import-Module { param($Name) }
            Invoke-BaselineAssessmentWorker -syncHash $admxOnlyState
        }
        if (-not $admxOnlyState.Done -or $admxOnlyState.Error -ne 'NO_PROFILES') { throw 'Assessment with only Administrative Templates should end with NO_PROFILES.' }
        Write-Output 'PASS: mocked discovery, profile reporting and complete assessment worker; no Graph writes.'
    } finally {
        $env:LOCALAPPDATA = $previousLocalAppData
        Remove-Item Function:Invoke-MgGraphRequest -ErrorAction SilentlyContinue
        Import-Module (Join-Path $sharedPath 'BaselineToolkit.Common.psm1') -Force -DisableNameChecking
    }

    foreach ($case in @(
        @{ Script = 'Invoke-BaselineMerger.ps1'; Pattern = '*NewBaselinePath is required*' }
        @{ Script = 'Invoke-IntuneBaselineReport.ps1'; Pattern = '*Supply -ProfileId*' }
        @{ Script = 'Invoke-BaselineAssessment.ps1'; Pattern = '*BaselineFolder must name*' }
        @{ Script = 'Export-BaselineAssessmentWorkbook.ps1'; Pattern = '*CsvPath is required*' }
    )) {
        $rejected = $false
        try { & (Join-Path $featuresPath $case.Script) | Out-Null } catch {
            if ($_.Exception.Message -notlike $case.Pattern) { throw }
            $rejected = $true
        }
        if (-not $rejected) { throw "Missing inputs accepted: $($case.Script)" }
    }
    Write-Output 'PASS: missing CLI inputs rejected without authentication or opening a GUI.'
} finally {
    Remove-Item -LiteralPath $temporaryFolder -Recurse -Force
}

if (-not $VerifyGui) {
    $newWpfAssemblies = @([AppDomain]::CurrentDomain.GetAssemblies().GetName().Name | Where-Object { $_ -in @('PresentationFramework','PresentationCore','WindowsBase') -and $_ -notin $initialWpfAssemblies })
    if ($newWpfAssemblies.Count) { throw "Headless feature checks loaded WPF: $($newWpfAssemblies -join ', ')" }
    Write-Output 'PASS: headless feature checks did not introduce WPF assemblies.'
} else {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $script:changeHistoryLayoutChecked = $false
    $closeTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $closeTimer.Interval = [timespan]::FromSeconds(1)
    $closeTimer.Add_Tick({
        $closeTimer.Stop()
        try {
            $mainTabs.SelectedIndex = 1
            $activity = $window.FindName('changeHistoryActivityExpander')
            $log = $window.FindName('txtChangeHistoryLog')
            $log.Text = ("Fixture activity line`n" * 30)
            foreach ($size in @(
                @{ Width = $window.Width; Height = $window.Height },
                @{ Width = $window.MinWidth; Height = $window.MinHeight }
            )) {
                $window.Width = $size.Width
                $window.Height = $size.Height
                $activity.IsExpanded = $false
                $window.UpdateLayout()
                if ($activity.ActualHeight -gt 60) { throw 'Collapsed Change History activity reserves excess vertical space.' }
                $activity.IsExpanded = $true
                $window.UpdateLayout()
                $logTop = $log.TransformToAncestor($activity).Transform([System.Windows.Point]::new(0, 0)).Y
                if ($log.ActualHeight -lt 120 -or $activity.ActualHeight -gt ($log.ActualHeight + 60) -or
                    $logTop -gt 60) { throw 'Change History log is stretched or separated from its header.' }
            }
            $script:changeHistoryLayoutChecked = $true
        } finally { $window.Close() }
    })
    $closeTimer.Start()
    try {
        . (Join-Path $basePath 'Invoke-BaselineToolbox.ps1')
        if (-not $script:changeHistoryLayoutChecked) { throw 'Change History activity layout check did not run.' }
        Write-Output 'PASS: Change History activity stays compact when expanded/collapsed at normal and minimum window sizes.'
        foreach ($controlName in @('btnConvert', 'btnChangeHistoryExport', 'btnMdToHtmlConvert', 'btnMergeOnly', 'btnPolicyCreate', 'btnProfileReportGenerate', 'btnAssessmentRun', 'btnAssessmentWorkbookExport')) {
            if (-not $window.FindName($controlName)) { throw "GUI control missing: $controlName" }
        }
        $reportNavigation = $window.FindName('navIntuneBaselineReport')
        $creatorNavigation = $window.FindName('navIntunePolicyCreator')
        $navigationItems = $reportNavigation.Parent.Children
        if ($reportNavigation.Content -ne 'Intune Policy Report' -or
            $navigationItems.IndexOf($reportNavigation) -ne ($navigationItems.IndexOf($creatorNavigation) + 1)) { throw 'Policy report navigation name/order is incorrect.' }
        $mainTabs.SelectedIndex = 5
        if ($script:toolPageInfoByIndex[5].Title -ne 'Intune Policy Report') { throw 'Policy report page title is incorrect.' }
        $folderBox = $window.FindName('txtProfileReportBaselineFolder')
        $folderLabel = @($folderBox.Parent.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] })
        if ($folderLabel.Count -ne 1 -or $folderLabel[0].Text -ne 'Compare Folder:') { throw 'Compare Folder label was not updated.' }
        $reportHelp = $window.FindName('profileReportEmptyState')
        if (-not $reportHelp -or $reportHelp.Visibility -ne 'Visible') { throw 'Policy report empty-grid help should be visible before profiles are loaded.' }
        $helpText = ($reportHelp.Child.Children | ForEach-Object { $_.Text }) -join ' '
        if ($helpText -notlike '*Load profiles*' -or $helpText -notlike '*Leave Compare Folder blank for a policy report*' -or
            $helpText -notlike '*JSON backup folder to compare settings*') { throw 'Policy report help does not explain profile loading and optional comparison.' }
        $reportGrid = $window.FindName('dgProfileReportProfiles')
        $originalReportSource = $reportGrid.ItemsSource
        try {
            $reportGrid.ItemsSource = [object[]]@([pscustomobject]@{ Selected = $false; Name = 'Help fixture'; LastModifiedDateTime = '-' })
            Update-ProfileReportSelectionCount
            if ($reportHelp.Visibility -ne 'Collapsed') { throw 'Policy report help should hide after loading profiles.' }
            Apply-ProfileReportFilter -FilterText 'no matching fixture'
            $reportView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($reportGrid.ItemsSource)
            if (-not $reportView.IsEmpty -or $reportHelp.Visibility -ne 'Collapsed') { throw 'Filtering to zero matches incorrectly displays the load-profiles help.' }
            $reportGrid.ItemsSource = [object[]]@()
            Update-ProfileReportSelectionCount
            if ($reportHelp.Visibility -ne 'Visible') { throw 'Policy report help should return when the loaded list is empty.' }
        } finally {
            $reportGrid.ItemsSource = $originalReportSource
            Apply-ProfileReportFilter
        }
        $workbookNavigation = $window.FindName('navAssessmentWorkbook')
        $mainTabs.SelectedIndex = 7
        if ($workbookNavigation.Content -ne 'Baseline Workbook' -or
            $script:toolPageInfoByIndex[7].Title -ne 'Baseline Workbook' -or
            $mainTabs.Items[7].Header.Trim() -ne 'Baseline Workbook') { throw 'Baseline Workbook navigation, page title or tab header is incorrect.' }
        Write-Output 'PASS: toolbox opens and closes with all eight feature controls present.'
    } finally { $closeTimer.Stop() }
}