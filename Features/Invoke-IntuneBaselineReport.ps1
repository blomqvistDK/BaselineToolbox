#Requires -Version 5.1
<#
.SYNOPSIS
    Lists Intune policies, inventories settings, or compares them with JSON policy backups.
.EXAMPLE
    .\Invoke-IntuneBaselineReport.ps1 -ListProfiles
.EXAMPLE
    .\Invoke-IntuneBaselineReport.ps1 -ProfileId 'profile-guid' -OutputPath .\report.html -CompareFolder .\Backups
.NOTES
    Dot-source to load report functions only. A matching Graph session is reused.
#>
[CmdletBinding()]
param(
    [switch]$ListProfiles,
    [switch]$WarmCache,
    [string[]]$ProfileId,
    [object[]]$Profiles,
    [string]$OutputPath,
    [Alias('CompareFolder')][string]$BaselineFolder,
    [string]$TenantId = $env:BASELINE_TOOLBOX_RESOLVER_TENANTID,
    [ValidateSet('fallback','cacheonly','forcefresh')][string]$CacheMode = 'fallback',
    [switch]$OpenReport,
    [scriptblock]$LogCallback,
    [scriptblock]$ProgressCallback
)

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Shared\BaselineToolkit.Common.psm1') -DisableNameChecking

function Compare-IntunePolicyBackup {
    param([string]$CompareFolder, [object[]]$Profiles, [object[]]$SettingRows, [object[]]$FailedProfiles = @())
    if (-not (Test-Path -LiteralPath $CompareFolder -PathType Container)) { throw "Compare folder not found: $CompareFolder" }
    $files = @(Get-ChildItem -LiteralPath $CompareFolder -Filter '*.json' -File -ErrorAction Stop)
    if (-not $files.Count) { throw "Compare folder contains no JSON policy backups: $CompareFolder" }
    $backups = [System.Collections.Generic.List[object]]::new()
    $result = @{
        Policies = [System.Collections.Generic.List[object]]::new()
        Diagnostics = [System.Collections.Generic.List[string]]::new()
        MissingRows = [System.Collections.Generic.List[object]]::new()
        ProfileStatus = @{}
        IgnoredBackups = 0
    }
    foreach ($file in $files) {
        try {
            $parsed = Get-BaselineSettingsFromJson -JsonPath $file.FullName -PolicyReport
            $backups.Add([pscustomobject]@{ FileName = $file.Name; Policy = $parsed })
        } catch { $result.Diagnostics.Add("$($file.Name): $($_.Exception.Message)") }
    }
    $usedFiles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($policy in $Profiles) {
        $id = [string]$policy.Id
        $name = [string]$policy.Name
        $type = [string](Get-SafeObjectPropertyValue -Object $policy -Name 'OdataType')
        $kind = [string](Get-SafeObjectPropertyValue -Object $policy -Name 'PolicyKind')
        if ($kind -eq 'SettingsCatalog' -or (-not $type -and $kind -ne 'DeviceRestriction')) { $type = 'SettingsCatalog' }
        $entry = [pscustomobject]@{
            ProfileId = $id; ProfileName = $name; Status = 'No backup found'; FileName = ''; Reason = ''
            Settings = [System.Collections.Generic.List[object]]::new()
        }
        $failed = @($FailedProfiles | Where-Object {
            $failedId = [string](Get-SafeObjectPropertyValue -Object $_ -Name 'Id')
            ($failedId -and $failedId -eq $id) -or (-not $failedId -and $_.Name -eq $name)
        })
        $candidates = @($backups | Where-Object { $id -and $_.Policy.PolicyId -eq $id })
        if (-not $candidates.Count) {
            $sameLiveNames = @($Profiles | Where-Object {
                $otherType = [string](Get-SafeObjectPropertyValue -Object $_ -Name 'OdataType')
                if (-not $otherType) { $otherType = 'SettingsCatalog' }
                ([string]$_.Name).Trim() -ieq $name.Trim() -and $otherType -eq $type
            })
            $candidates = @($backups | Where-Object {
                $backupId = [string]$_.Policy.PolicyId
                ([string]$_.Policy.PolicyName).Trim() -ieq $name.Trim() -and $_.Policy.PolicyType -eq $type -and
                    (-not $backupId -or -not @($Profiles | Where-Object Id -eq $backupId).Count)
            })
            if ($candidates.Count -and $sameLiveNames.Count -gt 1) {
                $entry.Status = 'Ambiguous'; $entry.Reason = 'Multiple selected policies have this name and type.'
            }
        }
        foreach ($candidate in $candidates) { [void]$usedFiles.Add($candidate.FileName) }
        if ($failed.Count) {
            $entry.Status = 'Failed'; $entry.Reason = [string]$failed[0].Error
        } elseif ($entry.Status -eq 'Ambiguous' -or $candidates.Count -gt 1) {
            $entry.Status = 'Ambiguous'; $entry.Reason = 'Policy identity is not unique; no settings were compared.'
        } elseif ($candidates.Count -eq 1 -and $candidates[0].Policy.PolicyType -ne $type) {
            $entry.Status = 'Type conflict'; $entry.Reason = 'The backup ID matches but its policy type differs.'
        } elseif ($candidates.Count -eq 1) {
            $backup = $candidates[0]
            $entry.FileName = $backup.FileName
            $liveRows = @($SettingRows | Where-Object { $_.ProfileId -eq $id })
            $backupGroups = @{}
            $liveGroups = @{}
            foreach ($setting in $backup.Policy.Settings) {
                $key = [string]$setting.DefinitionId
                if (-not $key) { continue }
                if (-not $backupGroups.ContainsKey($key)) { $backupGroups[$key] = @() }
                $backupGroups[$key] += $setting
            }
            foreach ($setting in $liveRows) {
                $key = [string]$setting.DefinitionId
                if (-not $key -or $key -eq '-') { continue }
                if (-not $liveGroups.ContainsKey($key)) { $liveGroups[$key] = @() }
                $liveGroups[$key] += $setting
            }
            $entry.Status = 'Match'
            foreach ($key in @(@($backupGroups.Keys) + @($liveGroups.Keys) | Sort-Object -Unique)) {
                $expected = @($backupGroups[$key] | Where-Object { $null -ne $_ })
                $current = @($liveGroups[$key] | Where-Object { $null -ne $_ })
                $expectedValues = @($expected | ForEach-Object { [string]$_.RawValue } | Sort-Object -CaseSensitive)
                $currentValues = @($current | ForEach-Object { [string]$_.RawValue } | Sort-Object -CaseSensitive)
                $expectedKey = ConvertTo-Json -InputObject $expectedValues -Compress
                $currentKey = ConvertTo-Json -InputObject $currentValues -Compress
                $status = if (-not $expected.Count) { 'Added since backup' } elseif (-not $current.Count) { 'Missing from live' } elseif ($expectedKey -ceq $currentKey) { 'Match' } else { 'Changed' }
                if ($status -ne 'Match') { $entry.Status = 'Changed' }
                $settingName = if ($current.Count) { [string]$current[0].SettingName } else { $key }
                $entry.Settings.Add([pscustomobject]@{
                    DefinitionId = $key; SettingName = $settingName; Status = $status
                    BackupValue = $expectedValues -join ' | '; CurrentValue = $currentValues -join ' | '
                })
                if ($status -eq 'Missing from live') {
                    $result.MissingRows.Add([pscustomobject]@{ DefinitionId = $key; ExpectedValue = $expectedValues -join ' | '; SourceFile = $backup.FileName })
                }
            }
        } elseif ($result.Diagnostics.Count) {
            $entry.Status = 'Unverified'; $entry.Reason = 'No matching readable backup; see backup diagnostics.'
        }
        $result.ProfileStatus[$id] = $entry.Status
        $result.Policies.Add($entry)
    }
    $result.IgnoredBackups = $backups.Count - $usedFiles.Count
    return $result
}

function New-IntuneProfileSettingsHtmlReport {
    param(
        [string]$OutputPath,
        [object[]]$Profiles,
        [object[]]$SettingRows,
        [object[]]$FailedProfiles,
        [hashtable]$BaselineComparison = $null
    )

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $totalProfiles = @($Profiles).Count
    $totalSettings = @($SettingRows).Count
    $failedCount = @($FailedProfiles).Count
    $isPolicyComparison = $null -ne $BaselineComparison -and $BaselineComparison.ContainsKey('Policies')
    $hasBaseline = $null -ne $BaselineComparison -and -not $isPolicyComparison
    $hasBaselineFiles = $hasBaseline -and $BaselineComparison.ContainsKey('BaselineFiles') -and $BaselineComparison.BaselineFiles.Count -gt 0
    $baselineMatchCount    = if ($hasBaseline) { @($BaselineComparison.ProfileStatus.Values | Where-Object { $_ -eq 'Match' }).Count } else { 0 }
    $baselineMismatchCount = if ($hasBaseline) { @($BaselineComparison.ProfileStatus.Values | Where-Object { $_ -eq 'Mismatch' }).Count } else { 0 }
    $baselineMissingCount  = if ($hasBaseline) { @($BaselineComparison.MissingRows).Count } else { 0 }
    $baselineStatsHtml = if ($hasBaselineFiles) {
        $baselineTotalSettings = [int]$BaselineComparison.BaselineCount
        $totalFound = $baselineTotalSettings - $baselineMissingCount
        "  <div class='stat'><div class='v'>$($BaselineComparison.BaselineFiles.Count)</div><div class='l'>Baseline Files</div></div>`n  <div class='stat'><div class='v'>$baselineTotalSettings</div><div class='l'>Baseline Settings</div></div>`n  <div class='stat'><div class='v'>$totalFound</div><div class='l'>Found in Intune</div></div>`n  <div class='stat'><div class='v'>$baselineMissingCount</div><div class='l'>Not Found</div></div>"
    } elseif ($hasBaseline) {
        "  <div class='stat'><div class='v'>$baselineMatchCount</div><div class='l'>Profiles Matching</div></div>`n  <div class='stat'><div class='v'>$baselineMismatchCount</div><div class='l'>Profiles Mismatching</div></div>`n  <div class='stat'><div class='v'>$baselineMissingCount</div><div class='l'>Baseline Not Found</div></div>"
    } else { '' }

    $sectionBuilder = [System.Text.StringBuilder]::new()

    if ($isPolicyComparison) {
        $baselineStatsHtml = ''
        foreach ($statusGroup in @($BaselineComparison.Policies | Group-Object Status)) {
            $safeStatus = [System.Net.WebUtility]::HtmlEncode([string]$statusGroup.Name)
            $baselineStatsHtml += "<div class='stat'><div class='v'>$($statusGroup.Count)</div><div class='l'>$safeStatus</div></div>"
        }
        foreach ($diagnostic in $BaselineComparison.Diagnostics) {
            [void]$sectionBuilder.AppendLine("<p class='backup-diagnostic'>Backup not compared: $([System.Net.WebUtility]::HtmlEncode([string]$diagnostic))</p>")
        }
        [void]$sectionBuilder.AppendLine("<p class='subtitle'>$($BaselineComparison.IgnoredBackups) unrelated backup file(s) ignored.</p>")
        $displayNames = if ($BaselineComparison.ContainsKey('DisplayNames')) { $BaselineComparison.DisplayNames } else { @{} }
        foreach ($policy in $BaselineComparison.Policies) {
            $safeName = [System.Net.WebUtility]::HtmlEncode([string]$policy.ProfileName)
            $safeMeta = [System.Net.WebUtility]::HtmlEncode((@($policy.Status, $policy.FileName, $policy.Reason) | Where-Object { $_ }) -join ' | ')
            $badgeClass = if ($policy.Status -eq 'Match') { 'badge-match' } else { 'badge-mismatch' }
            [void]$sectionBuilder.AppendLine("<section class='profile-section'><div class='profile-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span><span class='badge $badgeClass'></span><span class='name'>$safeName</span><span class='meta'>$safeMeta</span></div><div class='profile-body'><table class='settings-table'><thead><tr>")
            $headers = @('Status', 'Setting Name', 'Backup value', 'Current value', 'Definition Id')
            for ($column = 0; $column -lt $headers.Count; $column++) {
                [void]$sectionBuilder.AppendLine("<th data-col='$column' onclick='sortTable(this)'>$($headers[$column])</th>")
            }
            [void]$sectionBuilder.AppendLine('</tr><tr class="filter-row">')
            for ($column = 0; $column -lt $headers.Count; $column++) {
                [void]$sectionBuilder.AppendLine("<th><input class='col-filter' data-col='$column' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            }
            [void]$sectionBuilder.AppendLine('</tr></thead><tbody>')
            foreach ($setting in $policy.Settings) {
                $rowClass = switch ($setting.Status) { 'Changed' { 'row-mismatch' }; 'Missing from live' { 'row-missing' }; 'Added since backup' { 'row-extra' }; default { '' } }
                $statusClass = switch ($setting.Status) { 'Match' { 'status-match' }; 'Changed' { 'status-mismatch' }; 'Missing from live' { 'status-missing' }; default { 'status-extra' } }
                $settingName = if ($displayNames.ContainsKey($setting.DefinitionId)) { $displayNames[$setting.DefinitionId] } else { $setting.SettingName }
                $cells = @($setting.Status, $settingName, $setting.BackupValue, $setting.CurrentValue, $setting.DefinitionId) | ForEach-Object { [System.Net.WebUtility]::HtmlEncode([string]$_) }
                [void]$sectionBuilder.AppendLine("<tr class='$rowClass'><td class='$statusClass'>$($cells[0])</td><td>$($cells[1])</td><td>$($cells[2])</td><td>$($cells[3])</td><td>$($cells[4])</td></tr>")
            }
            if (-not $policy.Settings.Count) {
                $message = if ($policy.Status -eq 'Match') { 'No settings configured in either policy.' } else { 'Not compared: ' + $safeMeta }
                [void]$sectionBuilder.AppendLine("<tr><td>$([System.Net.WebUtility]::HtmlEncode([string]$policy.Status))</td><td colspan='4'>$message</td></tr>")
            }
            [void]$sectionBuilder.AppendLine('</tbody></table></div></section>')
        }
    } elseif ($hasBaselineFiles) {
        $baselineDisplayNames = if ($BaselineComparison.ContainsKey('DisplayNames')) { $BaselineComparison['DisplayNames'] } else { @{} }
        $baselineOptionNames  = if ($BaselineComparison.ContainsKey('OptionDisplayNames')) { $BaselineComparison['OptionDisplayNames'] } else { @{} }

        foreach ($fileEntry in $BaselineComparison.BaselineFiles) {
            $safePolicyName = [System.Net.WebUtility]::HtmlEncode([string]$fileEntry.PolicyName)
            $safeFileName   = [System.Net.WebUtility]::HtmlEncode([string]$fileEntry.FileName)
            $settingsCount  = $fileEntry.Settings.Count
            $matchCount     = 0
            $mismatchCount  = 0
            $notFoundCount  = 0
            foreach ($s in $fileEntry.Settings) {
                if ($s.FoundIn.Count -eq 0) { $notFoundCount++ }
                else {
                    $hasAnyMismatch = $false
                    foreach ($f in $s.FoundIn) { if ($f.Status -eq 'Mismatch') { $hasAnyMismatch = $true; break } }
                    if ($hasAnyMismatch) { $mismatchCount++ } else { $matchCount++ }
                }
            }
            $fileBadge = if ($notFoundCount -eq 0 -and $mismatchCount -eq 0) {
                "<span class='badge badge-match' title='All settings found and matching'>&#x2705;</span>"
            } elseif ($matchCount -eq 0 -and $mismatchCount -eq 0 -and $notFoundCount -gt 0) {
                "<span class='badge badge-notfound' title='All $notFoundCount settings not found in any profile'>&#x274C;</span>"
            } elseif ($notFoundCount -gt 0 -or $mismatchCount -gt 0) {
                "<span class='badge badge-mismatch' title='$mismatchCount mismatched, $notFoundCount not found'>&#x26A0;&#xFE0F;</span>"
            } else { '' }

            [void]$sectionBuilder.AppendLine("<section class='profile-section'>")
            [void]$sectionBuilder.AppendLine("  <div class='profile-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span>$fileBadge<span class='name'>$safePolicyName</span><span class='meta'>$safeFileName &bull; $settingsCount settings &bull; $matchCount matched &bull; $mismatchCount mismatched &bull; $notFoundCount not found</span></div>")
            [void]$sectionBuilder.AppendLine("  <div class='profile-body'>")
            [void]$sectionBuilder.AppendLine("    <table class='settings-table'><thead><tr>")
            [void]$sectionBuilder.AppendLine("      <th style='width:46px'>Status</th>")
            [void]$sectionBuilder.AppendLine("      <th data-col='1' onclick='sortTable(this)'>Setting Name</th>")
            [void]$sectionBuilder.AppendLine("      <th data-col='2' onclick='sortTable(this)'>Expected Value</th>")
            [void]$sectionBuilder.AppendLine("      <th data-col='3' onclick='sortTable(this)'>Found in Policy</th>")
            [void]$sectionBuilder.AppendLine("      <th data-col='4' onclick='sortTable(this)'>Actual Value</th>")
            [void]$sectionBuilder.AppendLine("      <th data-col='5' onclick='sortTable(this)'>Definition Id</th>")
            [void]$sectionBuilder.AppendLine("    </tr><tr class='filter-row'>")
            [void]$sectionBuilder.AppendLine("      <th></th>")
            [void]$sectionBuilder.AppendLine("      <th><input class='col-filter' data-col='1' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("      <th><input class='col-filter' data-col='2' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("      <th><input class='col-filter' data-col='3' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("      <th><input class='col-filter' data-col='4' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("      <th><input class='col-filter' data-col='5' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("    </tr></thead><tbody>")

            foreach ($s in $fileEntry.Settings) {
                $defId = [System.Net.WebUtility]::HtmlEncode([string]$s.DefinitionId)
                $rawExpected = [string]$s.ExpectedValue
                $expectedDisplay = if ($rawExpected -and $baselineOptionNames.ContainsKey($rawExpected)) {
                    [System.Net.WebUtility]::HtmlEncode($baselineOptionNames[$rawExpected])
                } else {
                    [System.Net.WebUtility]::HtmlEncode($rawExpected)
                }
                $settingName = if ($s.DefinitionId -and $baselineDisplayNames.ContainsKey([string]$s.DefinitionId)) {
                    [System.Net.WebUtility]::HtmlEncode($baselineDisplayNames[[string]$s.DefinitionId])
                } elseif ($s.FoundIn.Count -gt 0 -and $s.FoundIn[0].SettingName) {
                    [System.Net.WebUtility]::HtmlEncode([string]$s.FoundIn[0].SettingName)
                } elseif ($s.DefinitionId -match '_([^_]+)$') {
                    [System.Net.WebUtility]::HtmlEncode($matches[1])
                } else {
                    $defId
                }

                if ($s.FoundIn.Count -eq 0) {
                    [void]$sectionBuilder.AppendLine("      <tr class='row-missing'><td class='status-missing'>&#x274C;</td><td>$settingName</td><td>$expectedDisplay</td><td><em>Not found in any profile</em></td><td>-</td><td>$defId</td></tr>")
                } else {
                    foreach ($hit in $s.FoundIn) {
                        $profNameHtml = [System.Net.WebUtility]::HtmlEncode([string]$hit.ProfileName)
                        $actualRaw = [string]$hit.ActualValue
                        $actualDisplay = if ($actualRaw -and $baselineOptionNames.ContainsKey($actualRaw)) {
                            [System.Net.WebUtility]::HtmlEncode($baselineOptionNames[$actualRaw])
                        } else {
                            [System.Net.WebUtility]::HtmlEncode($actualRaw)
                        }
                        if ($hit.Status -eq 'Match') {
                            [void]$sectionBuilder.AppendLine("      <tr><td class='status-match'>&#x2705;</td><td>$settingName</td><td>$expectedDisplay</td><td>$profNameHtml</td><td>$actualDisplay</td><td>$defId</td></tr>")
                        } else {
                            [void]$sectionBuilder.AppendLine("      <tr class='row-mismatch'><td class='status-mismatch'>&#x26A0;&#xFE0F;</td><td>$settingName</td><td>$expectedDisplay</td><td>$profNameHtml</td><td>$actualDisplay</td><td>$defId</td></tr>")
                        }
                    }
                }
            }

            [void]$sectionBuilder.AppendLine("    </tbody></table>")
            [void]$sectionBuilder.AppendLine("  </div>")
            [void]$sectionBuilder.AppendLine("</section>")
        }
    } else {
        # Profile-centric view (no baseline folder specified)
    foreach ($emptyPolicy in @($Profiles | Where-Object { $policyId = $_.Id; -not @($SettingRows | Where-Object ProfileId -eq $policyId).Count })) {
        $emptyName = [System.Net.WebUtility]::HtmlEncode([string]$emptyPolicy.Name)
        $emptyMessage = if (@($FailedProfiles | Where-Object { $_.Name -eq $emptyPolicy.Name }).Count) { 'Settings collection failed; see errors below.' } else { 'No settings configured.' }
        [void]$sectionBuilder.AppendLine("<section class='profile-section'><div class='profile-header' onclick='toggleSection(this)'><span class='name'>$emptyName</span></div><div class='profile-body'><table class='settings-table'><tbody><tr><td>$emptyMessage</td></tr></tbody></table></div></section>")
    }
    $grouped = $SettingRows | Group-Object ProfileId | Sort-Object { $_.Group[0].ProfileName }
    foreach ($group in $grouped) {
        $profileName = [System.Net.WebUtility]::HtmlEncode([string]$group.Group[0].ProfileName)
        $rawProfileName = [string]$group.Group[0].ProfileName
        $platforms = [System.Net.WebUtility]::HtmlEncode([string]$group.Group[0].Platforms)
        $profileId = [System.Net.WebUtility]::HtmlEncode([string]$group.Name)

        # Profile-level compliance badge
        $profileBadge = ''
        if ($hasBaseline) {
            $profStatus = if ($BaselineComparison.ProfileStatus.ContainsKey($rawProfileName)) { $BaselineComparison.ProfileStatus[$rawProfileName] } else { 'NoBaseline' }
            $profileBadge = switch ($profStatus) {
                'Match'      { "<span class='badge badge-match' title='All settings match baseline'>&#x2705;</span>" }
                'Mismatch'   { "<span class='badge badge-mismatch' title='Some settings differ from baseline'>&#x26A0;&#xFE0F;</span>" }
                'NoBaseline' { "<span class='badge badge-nobaseline' title='No matching baseline JSON found'>&#x2753;</span>" }
            }
        }

        [void]$sectionBuilder.AppendLine("<section class='profile-section' data-profile-id='$profileId'>")
        [void]$sectionBuilder.AppendLine("  <div class='profile-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span>$profileBadge<span class='name'>$profileName</span><span class='meta'>$platforms &bull; $($group.Count) settings</span></div>")
        [void]$sectionBuilder.AppendLine("  <div class='profile-body'>")
        [void]$sectionBuilder.AppendLine("    <table class='settings-table'>")
        [void]$sectionBuilder.AppendLine("      <thead>")
        [void]$sectionBuilder.AppendLine("        <tr>")
        if ($hasBaseline) {
            [void]$sectionBuilder.AppendLine("          <th data-col='0' onclick='sortTable(this)' style='width:50px'>Status</th>")
            [void]$sectionBuilder.AppendLine("          <th data-col='1' onclick='sortTable(this)'>Setting Name</th>")
            [void]$sectionBuilder.AppendLine("          <th data-col='2' onclick='sortTable(this)'>Value</th>")
            [void]$sectionBuilder.AppendLine("          <th data-col='3' onclick='sortTable(this)'>Definition Id</th>")
            [void]$sectionBuilder.AppendLine("          <th data-col='4' onclick='sortTable(this)'>Profile</th>")
        } else {
            [void]$sectionBuilder.AppendLine("          <th data-col='0' onclick='sortTable(this)'>Setting Name</th>")
            [void]$sectionBuilder.AppendLine("          <th data-col='1' onclick='sortTable(this)'>Value</th>")
            [void]$sectionBuilder.AppendLine("          <th data-col='2' onclick='sortTable(this)'>Definition Id</th>")
            [void]$sectionBuilder.AppendLine("          <th data-col='3' onclick='sortTable(this)'>Profile</th>")
        }
        [void]$sectionBuilder.AppendLine("        </tr>")
        [void]$sectionBuilder.AppendLine("        <tr class='filter-row'>")
        if ($hasBaseline) {
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='0' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='1' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='2' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='3' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='4' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
        } else {
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='0' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='1' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='2' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
            [void]$sectionBuilder.AppendLine("          <th><input class='col-filter' data-col='3' type='text' placeholder='Filter...' oninput='applyFilters()'></th>")
        }
        [void]$sectionBuilder.AppendLine("        </tr>")
        [void]$sectionBuilder.AppendLine("      </thead>")
        [void]$sectionBuilder.AppendLine("      <tbody>")

        foreach ($row in ($group.Group | Sort-Object SortGroup, Order)) {
            $settingName = [System.Net.WebUtility]::HtmlEncode([string]$row.SettingName)
            $value = [System.Net.WebUtility]::HtmlEncode([string]$row.SettingValue)
            $defId = [System.Net.WebUtility]::HtmlEncode([string]$row.DefinitionId)
            $depth = if ($row.PSObject.Properties['Depth']) { [int]$row.Depth } else { 0 }
            $indent = if ($depth -gt 0) { '<span class="sub-indicator">&#x21B3; </span>' } else { '' }
            $trClass = if ($depth -gt 0) { "sub-setting depth-$depth" } else { '' }

            $statusCell = ''
            if ($hasBaseline) {
                $compKey = New-CompareKey -ProfileName $rawProfileName -DefinitionId ([string]$row.DefinitionId)
                $compStatus = if ($BaselineComparison.SettingStatus.ContainsKey($compKey)) { $BaselineComparison.SettingStatus[$compKey] } else { $null }
                if ($compStatus) {
                    switch ($compStatus.Status) {
                        'Match'    { $statusCell = "<td class='status-match' title='Matches baseline'>&#x2705;</td>"; $trClass = if ($trClass) { "$trClass" } else { '' } }
                        'Mismatch' {
                            $rawExpectedVal = [string]$compStatus.ExpectedValue
                            $mismatchOptionNames = if ($BaselineComparison.ContainsKey('OptionDisplayNames')) { $BaselineComparison['OptionDisplayNames'] } else { @{} }
                            $friendlyExpected = if ($rawExpectedVal -and $mismatchOptionNames.ContainsKey($rawExpectedVal)) { $mismatchOptionNames[$rawExpectedVal] } else { $rawExpectedVal }
                            $expected = [System.Net.WebUtility]::HtmlEncode($friendlyExpected)
                            $statusCell = "<td class='status-mismatch' title='Expected: $expected'>&#x26A0;&#xFE0F;</td>"
                            $value = [System.Net.WebUtility]::HtmlEncode([string]$row.SettingValue) + "<br><small class='expected-value'>Expected: $expected</small>"
                            $trClass = if ($trClass) { "$trClass row-mismatch" } else { 'row-mismatch' }
                        }
                        'Extra'    {
                            $statusCell = "<td class='status-extra' title='Not in baseline (extra in Intune)'>&#x2795;</td>"
                            $trClass = if ($trClass) { "$trClass row-extra" } else { 'row-extra' }
                        }
                    }
                } else {
                    # No comparison data for this setting - show neutral indicator
                    $statusCell = "<td title='Not compared'>&#x2796;</td>"
                }
            }

            $trClassAttr = if ($trClass) { " class='$trClass'" } else { '' }
            if ($hasBaseline) {
                [void]$sectionBuilder.AppendLine("        <tr$trClassAttr>$statusCell<td>${indent}$settingName</td><td>$value</td><td>$defId</td><td>$profileName</td></tr>")
            } else {
                [void]$sectionBuilder.AppendLine("        <tr$trClassAttr><td>${indent}$settingName</td><td>$value</td><td>$defId</td><td>$profileName</td></tr>")
            }
        }

        [void]$sectionBuilder.AppendLine('      </tbody>')
        [void]$sectionBuilder.AppendLine('    </table>')
        [void]$sectionBuilder.AppendLine('  </div>')
        [void]$sectionBuilder.AppendLine('</section>')
    }
    } # end else (profile-centric view)

    # Build global section for baseline settings not found in any selected profile
    # (only needed for profile-centric view; in baseline-centric view, missing settings are shown inline per file)
    $globalMissingSection = ''
    if ($hasBaseline -and -not $hasBaselineFiles -and $BaselineComparison.MissingRows.Count -gt 0) {
        $baselineDisplayNames = if ($BaselineComparison.ContainsKey('DisplayNames')) { $BaselineComparison['DisplayNames'] } else { @{} }
        $baselineOptionNames  = if ($BaselineComparison.ContainsKey('OptionDisplayNames')) { $BaselineComparison['OptionDisplayNames'] } else { @{} }
        $globalMissingRowsHtml = [System.Text.StringBuilder]::new()
        foreach ($missing in ($BaselineComparison.MissingRows | Sort-Object DefinitionId)) {
            $mDefId = [System.Net.WebUtility]::HtmlEncode([string]$missing.DefinitionId)
            $rawExpected = [string]$missing.ExpectedValue
            $mExpected = if ($rawExpected -and $baselineOptionNames.ContainsKey($rawExpected)) {
                [System.Net.WebUtility]::HtmlEncode($baselineOptionNames[$rawExpected])
            } else {
                [System.Net.WebUtility]::HtmlEncode($rawExpected)
            }
            $mName = if ($missing.DefinitionId -and $baselineDisplayNames.ContainsKey([string]$missing.DefinitionId)) {
                [System.Net.WebUtility]::HtmlEncode($baselineDisplayNames[[string]$missing.DefinitionId])
            } elseif ($missing.DefinitionId -match '_([^_]+)$') {
                [System.Net.WebUtility]::HtmlEncode($matches[1])
            } else {
                [System.Net.WebUtility]::HtmlEncode([string]$missing.DefinitionId)
            }
            $mSource = [System.Net.WebUtility]::HtmlEncode([string]$missing.SourceFile)
            [void]$globalMissingRowsHtml.AppendLine("<tr class='row-missing'><td class='status-missing'>&#x274C;</td><td>$mName</td><td>$mExpected</td><td>$mDefId</td><td>$mSource</td></tr>")
        }
        $gMissingCount = $BaselineComparison.MissingRows.Count
        $globalMissingSection = @"
<section class='profile-section global-missing-section'>
  <div class='profile-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span><span class='name' style='color:#f87171'>&#x274C; Baseline Settings Not Found in Any Profile ($gMissingCount)</span><span class='meta'>defined in JSON file(s) but absent from all $totalProfiles selected profile(s)</span></div>
  <div class='profile-body'>
    <table class='settings-table'><thead><tr>
      <th style='width:46px'>Status</th>
      <th data-col='1' onclick='sortTable(this)'>Setting Name</th>
      <th data-col='2' onclick='sortTable(this)'>Expected Value</th>
      <th data-col='3' onclick='sortTable(this)'>Definition Id</th>
      <th data-col='4' onclick='sortTable(this)'>Source File</th>
    </tr><tr class='filter-row'>
      <th></th>
      <th><input class='col-filter' data-col='1' type='text' placeholder='Filter...' oninput='applyFilters()'></th>
      <th><input class='col-filter' data-col='2' type='text' placeholder='Filter...' oninput='applyFilters()'></th>
      <th><input class='col-filter' data-col='3' type='text' placeholder='Filter...' oninput='applyFilters()'></th>
      <th><input class='col-filter' data-col='4' type='text' placeholder='Filter...' oninput='applyFilters()'></th>
    </tr></thead><tbody>
$($globalMissingRowsHtml.ToString())
    </tbody></table>
  </div>
</section>
"@
    }

    $failedSection = ''
    if ($failedCount -gt 0) {
        $failedRows = [System.Text.StringBuilder]::new()
        foreach ($fp in $FailedProfiles) {
            $n = [System.Net.WebUtility]::HtmlEncode([string]$fp.Name)
            $e = [System.Net.WebUtility]::HtmlEncode([string]$fp.Error)
            [void]$failedRows.AppendLine("<tr><td>$n</td><td>$e</td></tr>")
        }
        $failedSection = @"
<section class='failed-section'>
  <h2>Failed Profiles ($failedCount)</h2>
  <table>
    <thead><tr><th>Profile</th><th>Error</th></tr></thead>
    <tbody>
$($failedRows.ToString())
    </tbody>
  </table>
</section>
"@
    }

    $html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Intune Policy Report</title>
<style>
:root{--bg:#f8fafc;--fg:#0f172a;--surface:#ffffff;--muted:#475569;--accent:#2563eb;--border:#cbd5e1}
body.dark{--bg:#0f172a;--fg:#e2e8f0;--surface:#1e293b;--muted:#94a3b8;--accent:#60a5fa;--border:#334155}
*{box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--fg);margin:0;padding:20px}
h1{margin:0 0 4px 0;color:var(--accent)}
.subtitle{color:var(--muted);margin:0 0 16px 0;font-size:13px}
.stats{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px}
.stat{background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:8px 12px}
.stat .v{font-size:20px;font-weight:700;color:var(--accent)}
.stat .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.toolbar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:14px}
.toolbar input{padding:7px 10px;border:1px solid var(--border);border-radius:4px;background:var(--surface);color:var(--fg)}
.toolbar button{padding:7px 10px;border:1px solid var(--border);border-radius:4px;background:var(--surface);color:var(--fg);cursor:pointer}
.profile-section{background:var(--surface);border:1px solid var(--border);border-radius:8px;margin-bottom:10px;overflow:hidden}
.profile-header{padding:10px 12px;display:flex;gap:10px;align-items:center;cursor:pointer;background:rgba(37,99,235,.08)}
.profile-header .name{font-weight:600;flex:1}
.profile-header .meta{font-size:12px;color:var(--muted)}
.profile-header .arrow{transition:transform .2s}
.profile-header.open .arrow{transform:rotate(90deg)}
.profile-body{display:none;padding:10px}
.profile-body.open{display:block}
table{width:100%;border-collapse:collapse;font-size:12px}
th,td{padding:6px 8px;border-bottom:1px solid var(--border);text-align:left;vertical-align:top}
th{background:var(--accent);color:#fff;cursor:pointer}
.filter-row th{background:transparent}
.filter-row input{width:100%;padding:4px;border:1px solid var(--border);border-radius:3px;background:var(--surface);color:var(--fg)}
tr.sub-setting td:first-child{padding-left:24px}
tr.sub-setting td{color:var(--muted);font-style:italic}
.sub-indicator{color:var(--accent);font-style:normal}
.badge{margin:0 8px;font-size:14px}
.status-match,.status-mismatch,.status-extra,.status-missing{text-align:center;font-size:14px;width:50px;padding-left:6px !important;font-style:normal !important;color:var(--fg) !important}
tr.row-mismatch{background:rgba(245,158,11,.1)}
tr.row-extra{background:rgba(59,130,246,.08)}
tr.row-missing{background:rgba(239,68,68,.1)}
tr.row-missing td{font-style:italic;color:var(--muted)}
.expected-value{color:#b45309;font-style:italic}
.failed-section{margin-top:18px;background:var(--surface);border:1px solid #f59e0b;border-radius:8px;padding:12px}
.count{margin-left:auto;color:var(--muted);font-size:12px}
</style>
</head>
<body class='dark'>
<h1>Intune Policy Report</h1>
<p class='subtitle'>Generated $generatedAt</p>

<div class='stats'>
  <div class='stat'><div class='v'>$totalProfiles</div><div class='l'>Profiles</div></div>
  <div class='stat'><div class='v'>$totalSettings</div><div class='l'>Settings</div></div>
  <div class='stat'><div class='v'>$failedCount</div><div class='l'>Failed Profiles</div></div>
  $baselineStatsHtml
</div>

<div class='toolbar'>
  <input id='globalSearch' type='text' placeholder='Search all profiles/settings...' oninput='applyFilters()'>
  <button onclick='expandAll()'>Expand All</button>
  <button onclick='collapseAll()'>Collapse All</button>
  <button onclick='clearFilters()'>Clear Filters</button>
  <button onclick='toggleTheme()'>Toggle Theme</button>
  <button onclick='exportCsv()'>Export CSV</button>
    $(if ($isPolicyComparison -or $hasBaseline) { "<button id='btnNonMatching' onclick='toggleNonMatchingOnly()'>Differences</button><button id='btnMismatchOnly' onclick='toggleMismatchOnly()'>Changed values</button><button id='btnNotFoundOnly' onclick='toggleNotFoundOnly()'>Missing from live</button><button id='btnAddedOnly' onclick='toggleAddedOnly()'>Added since backup</button>" })
  <span id='visibleCount' class='count'></span>
</div>

$($sectionBuilder.ToString())
$globalMissingSection
$failedSection

<script>
function toggleSection(header){
  const body = header.parentElement.querySelector('.profile-body');
  const open = body.classList.toggle('open');
  header.classList.toggle('open', open);
}

function expandAll(){
  document.querySelectorAll('.profile-header').forEach(h=>h.classList.add('open'));
  document.querySelectorAll('.profile-body').forEach(b=>b.classList.add('open'));
}

function collapseAll(){
  document.querySelectorAll('.profile-header').forEach(h=>h.classList.remove('open'));
  document.querySelectorAll('.profile-body').forEach(b=>b.classList.remove('open'));
}

function sortTable(th){
  const table = th.closest('table');
  const col = Number(th.dataset.col);
  const key = 'sortCol';
  const dirKey = 'sortDir';
  const currentCol = Number(table.dataset[key] || -1);
  const nextAsc = currentCol === col ? table.dataset[dirKey] !== 'asc' : true;
  table.dataset[key] = String(col);
  table.dataset[dirKey] = nextAsc ? 'asc' : 'desc';
  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));
  rows.sort((a,b)=>{
    const va = (a.children[col]?.textContent || '').trim();
    const vb = (b.children[col]?.textContent || '').trim();
    const na = Number(va), nb = Number(vb);
    if(!Number.isNaN(na) && !Number.isNaN(nb)) return nextAsc ? na-nb : nb-na;
    return nextAsc ? va.localeCompare(vb) : vb.localeCompare(va);
  });
  rows.forEach(r=>tbody.appendChild(r));
  applyFilters();
}

function applyFilters(){
  const global = (document.getElementById('globalSearch').value || '').toLowerCase();
  const nonMatchOnly = document.getElementById('btnNonMatching')?.dataset.active === '1';
  const mismatchOnly = document.getElementById('btnMismatchOnly')?.dataset.active === '1';
  const notFoundOnly = document.getElementById('btnNotFoundOnly')?.dataset.active === '1';
    const addedOnly = document.getElementById('btnAddedOnly')?.dataset.active === '1';
  let visibleRows = 0;
  document.querySelectorAll('.profile-section').forEach(section=>{
    if(nonMatchOnly && section.querySelector('.badge-match')){ section.style.display='none'; return; }
    const rows = Array.from(section.querySelectorAll('tbody tr'));
    const filters = Array.from(section.querySelectorAll('.col-filter')).map(i=>({col:Number(i.dataset.col), val:(i.value||'').toLowerCase()}));
    let sectionVisible = false;
    rows.forEach(row=>{
      const cells = Array.from(row.children).map(c=>(c.textContent||'').toLowerCase());
      let show = true;
      if(global && !cells.some(v=>v.includes(global))) show = false;
      for(const f of filters){ if(f.val && !(cells[f.col] || '').includes(f.val)){ show = false; break; } }
      // Exclusive row-level filters
      if(show && nonMatchOnly  && row.querySelector('.status-match'))   { show = false; }
      if(show && mismatchOnly  && !row.querySelector('.status-mismatch')){ show = false; }
      if(show && notFoundOnly  && !row.querySelector('.status-missing')) { show = false; }
    if(show && addedOnly && !row.querySelector('.status-extra')) { show = false; }
      row.style.display = show ? '' : 'none';
      if(show){ sectionVisible = true; visibleRows++; }
    });
    section.style.display = sectionVisible ? '' : 'none';
  });
  document.getElementById('visibleCount').textContent = visibleRows + ' visible rows';
}

function clearFilters(){
  document.getElementById('globalSearch').value = '';
  document.querySelectorAll('.col-filter').forEach(i=>i.value='');
    ['btnNonMatching','btnMismatchOnly','btnNotFoundOnly','btnAddedOnly'].forEach(id=>{
    const b = document.getElementById(id);
    if(b){ b.dataset.active='0'; b.style.fontWeight=''; b.style.outline=''; }
  });
  applyFilters();
}

function _setExclusiveFilter(activeId){
    ['btnNonMatching','btnMismatchOnly','btnNotFoundOnly','btnAddedOnly'].forEach(id=>{
    const b = document.getElementById(id);
    if(!b) return;
    if(id === activeId){ b.dataset.active='1'; b.style.fontWeight='bold'; b.style.outline='2px solid var(--accent)'; }
    else { b.dataset.active='0'; b.style.fontWeight=''; b.style.outline=''; }
  });
}

function toggleNonMatchingOnly(){
  const btn = document.getElementById('btnNonMatching');
  const wasActive = btn.dataset.active === '1';
  if(wasActive){ btn.dataset.active='0'; btn.style.fontWeight=''; btn.style.outline=''; }
  else { _setExclusiveFilter('btnNonMatching'); }
  applyFilters();
}

function toggleMismatchOnly(){
  const btn = document.getElementById('btnMismatchOnly');
  const wasActive = btn.dataset.active === '1';
  if(wasActive){ btn.dataset.active='0'; btn.style.fontWeight=''; btn.style.outline=''; }
  else { _setExclusiveFilter('btnMismatchOnly'); }
  applyFilters();
}

function toggleNotFoundOnly(){
  const btn = document.getElementById('btnNotFoundOnly');
  const wasActive = btn.dataset.active === '1';
  if(wasActive){ btn.dataset.active='0'; btn.style.fontWeight=''; btn.style.outline=''; }
  else { _setExclusiveFilter('btnNotFoundOnly'); }
  applyFilters();
}

function toggleAddedOnly(){
    const btn = document.getElementById('btnAddedOnly');
    if(btn.dataset.active === '1'){ btn.dataset.active='0'; btn.style.fontWeight=''; btn.style.outline=''; }
    else { _setExclusiveFilter('btnAddedOnly'); }
    applyFilters();
}
function toggleTheme(){ document.body.classList.toggle('dark'); }

function exportCsv(){
  const sections = document.querySelectorAll('.profile-section');
  let header = null;
  sections.forEach(section=>{
    if(!header && section.style.display!=='none'){
    const columns = Array.from(section.querySelectorAll('thead tr:first-child th')).map(th=>(th.textContent||'').trim());
    if(columns.length) header = columns;
    }
  });
  if(!header) return;
  // Find the index of 'Setting Name' to insert 'Baseline' before it; fallback to index 0
  let insertIdx = 0;
  for(let i=0;i<header.length;i++){
    if(header[i].toLowerCase()==='setting name'){insertIdx=i;break;}
  }
  const csvHeader = [...header];
    csvHeader.splice(insertIdx, 0, 'Policy');
  const lines = [csvHeader.map(v=>'"'+v.replace(/"/g,'""')+'"').join(';')];
  sections.forEach(section=>{
    if(section.style.display==='none') return;
    const nameEl = section.querySelector('.profile-header .name');
    const baseline = nameEl ? (nameEl.textContent||'').trim() : '';
    section.querySelectorAll('tbody tr').forEach(row=>{
      if(row.style.display==='none') return;
      const cells = Array.from(row.children).map(c=>{
        let t = (c.textContent||'').trim();
        if(c.classList.contains('status-match')) t='Match';
        else if(c.classList.contains('status-mismatch')) t='Changed';
        else if(c.classList.contains('status-missing')) t='Missing from live';
        else if(c.classList.contains('status-extra')) t='Added since backup';
        return t;
      });
      cells.splice(insertIdx, 0, baseline);
      lines.push(cells.map(v=>'"'+v.replace(/"/g,'""')+'"').join(';'));
    });
  });
  const blob = new Blob([lines.join('\n')],{type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href=URL.createObjectURL(blob);
    a.download='Intune-Policy-Report.csv';
  a.click();
}

collapseAll();
applyFilters();
</script>
</body>
</html>
"@

    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
}

function Invoke-BaselineProfileReport {
    param([object[]]$TargetProfiles, [string]$OutputPath, [Alias('CompareFolder')][string]$BaselineFolder, [ValidateSet('fallback','cacheonly','forcefresh')][string]$CacheMode = 'fallback', [scriptblock]$LogCallback, [scriptblock]$ProgressCallback)
    function Write-BaselineReportLog { param([string]$Message) if ($LogCallback) { & $LogCallback $Message | Out-Null } else { Write-Verbose $Message } }
    function Write-BaselineReportProgress { param([string]$Status, [int]$Percent = -1) if ($ProgressCallback) { & $ProgressCallback $Percent $Status | Out-Null } }

        if ($BaselineFolder) {
            if (-not (Test-Path -LiteralPath $BaselineFolder -PathType Container)) { throw "Compare folder not found: $BaselineFolder" }
            if (-not @(Get-ChildItem -LiteralPath $BaselineFolder -Filter '*.json' -File -ErrorAction Stop).Count) { throw "Compare folder contains no JSON policy backups: $BaselineFolder" }
        }
        Write-BaselineReportLog "Resolver mode: $(Get-CacheModeLabel -Mode $cacheMode)"

        $requiredTenant = Get-ResolverTenantOverride
        if ($requiredTenant) {
            Write-BaselineReportLog "Resolver tenant override active: $requiredTenant"
        }
        $resolverTenantCacheKey = Get-GraphTenantCacheKey
        $cacheData = Import-SettingDefinitionCache -TenantKey $resolverTenantCacheKey
        $cacheError = Get-SafeObjectPropertyValue -Object $cacheData -Name 'Error'
        if ($cacheData.Loaded) {
            Write-BaselineReportLog "Loaded resolver cache for tenant '$resolverTenantCacheKey': $($cacheData.DisplayNameCount) names, $($cacheData.OptionCount) options."
        } elseif ($cacheError) {
            Write-BaselineReportLog "Resolver cache not used: $cacheError"
        }
        Write-CacheStalenessWarning -CacheData $cacheData -LogCallback { param($message) Write-BaselineReportLog $message }

        $total = @($targetProfiles).Count
        $current = 0
        $settingRows = [System.Collections.Generic.List[object]]::new()
        $failedProfiles = [System.Collections.Generic.List[object]]::new()
        $unresolvedDefIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $allDisplayNames = $cacheData.DisplayNames
        $allOptionDisplayNames = $cacheData.OptionDisplayNames
        if ($cacheMode -eq 'forcefresh') {
            $allDisplayNames = @{}
            $allOptionDisplayNames = @{}
            Write-BaselineReportLog 'Force fresh mode selected: ignoring imported cache for this run.'
        }

        Write-BaselineReportLog "Collecting settings for $total profile(s)..."
        $policyKindCounts = @{}
        foreach ($profItem in $targetProfiles) {
            $current++
            $pct = [math]::Round(($current / [double]$total) * 100)
            Write-BaselineReportProgress -Percent $pct -Status "Processing $current/$total"
            Write-BaselineReportProgress -Status "Processing $current/$total"
            Write-BaselineReportLog "[$current/$total] $($profItem.Name)"

            # ---- Device Restriction profiles (deviceConfigurations endpoint) ----
            $profPolicyKind = Get-SafeObjectPropertyValue -Object $profItem -Name 'PolicyKind'
            if (-not $profPolicyKind) { $profPolicyKind = 'Unknown' }
            if (-not $policyKindCounts.ContainsKey($profPolicyKind)) { $policyKindCounts[$profPolicyKind] = 0 }
            $policyKindCounts[$profPolicyKind]++

            if ($profPolicyKind -eq 'DeviceRestriction') {
                Write-BaselineReportLog "  PolicyKind: $profPolicyKind | Endpoint: deviceConfigurations"
                try {
                    $dcSkipProps = [System.Collections.Generic.HashSet[string]]::new(
                        [string[]]@('id','displayName','description','createdDateTime',
                                    'lastModifiedDateTime','version','roleScopeTagIds',
                                    'supportsScopeTags','settingCount'),
                        [System.StringComparer]::OrdinalIgnoreCase
                    )
                    $dc = Invoke-MgGraphRequest -Method Get `
                        -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($profItem.Id)"
                    $profOdataType = Get-SafeObjectPropertyValue -Object $profItem -Name 'OdataType'
                    $flatDcCount = 0
                    foreach ($propKey in @($dc.Keys)) {
                        if ($propKey -like '*@odata*' -or $propKey -ieq 'assignments' -or $propKey -like 'deviceManagementApplicabilityRule*') { continue }
                        if ($dcSkipProps.Contains($propKey)) { continue }
                        $rawVal = $dc[$propKey]
                        $valStr = if ($null -eq $rawVal) { '-' }
                                  elseif ($rawVal -is [bool]) { $rawVal.ToString().ToLower() }
                                  elseif ($rawVal -is [System.Collections.IEnumerable] -and $rawVal -isnot [string]) {
                                      $items = @($rawVal | ForEach-Object { if ($null -ne $_) { [string]$_ } else { '' } })
                                      ($items | Sort-Object) -join '; '
                                  } else { [string]$rawVal }
                        # Namespace DefinitionId with @odata.type so iOS and Android property
                        # names never collide when doing baseline comparison.
                        $defId = "$profOdataType|$propKey"
                        [void]$settingRows.Add([pscustomobject]@{
                            ProfileId    = $profItem.Id
                            ProfileName  = $profItem.Name
                            Platforms    = $profItem.Platforms
                            SettingName  = $propKey
                            SettingValue = $valStr
                            RawValue     = $valStr
                            DefinitionId = $defId
                            Depth        = 0
                            SortGroup    = $propKey
                            Order        = $flatDcCount
                        })
                        $flatDcCount++
                    }
                    Write-BaselineReportLog "  -> $flatDcCount settings (Device Restriction)"
                } catch {
                    $msg = $_.Exception.Message
                    [void]$failedProfiles.Add([pscustomobject]@{ Id = $profItem.Id; Name = $profItem.Name; Error = $msg })
                    Write-BaselineReportLog "  -> FAILED: $msg"
                }
                continue
            }

            # ---- Administrative Templates (groupPolicyConfigurations endpoint) ----
            if ($profPolicyKind -eq 'AdministrativeTemplate') {
                Write-BaselineReportLog "  PolicyKind: $profPolicyKind | Endpoint: groupPolicyConfigurations"
                try {
                    $gpUri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations/$($profItem.Id)/definitionValues?`$expand=definition(`$select=id,displayName,categoryPath,classType),presentationValues(`$expand=presentation)"
                    $gpValues = [System.Collections.Generic.List[object]]::new()
                    foreach ($item in @(Invoke-GraphPagedGet -Uri $gpUri)) {
                        if ($item -is [System.Array]) {
                            foreach ($sub in $item) { [void]$gpValues.Add($sub) }
                        } else {
                            [void]$gpValues.Add($item)
                        }
                    }
                    $gpRows = @(ConvertTo-AdminTemplateSettingRows -DefinitionValues $gpValues.ToArray())
                    foreach ($gpRow in $gpRows) {
                        [void]$settingRows.Add([pscustomobject]@{
                            ProfileId    = $profItem.Id
                            ProfileName  = $profItem.Name
                            Platforms    = $profItem.Platforms
                            SettingName  = $gpRow.SettingName
                            SettingValue = $gpRow.SettingValue
                            RawValue     = $gpRow.RawValue
                            DefinitionId = $gpRow.DefinitionId
                            Depth        = $gpRow.Depth
                            SortGroup    = $gpRow.SortGroup
                            Order        = $gpRow.Order
                        })
                    }
                    Write-BaselineReportLog "  -> $($gpRows.Count) settings (Administrative Template, $($gpValues.Count) definition values)"
                } catch {
                    $msg = $_.Exception.Message
                    [void]$failedProfiles.Add([pscustomobject]@{ Id = $profItem.Id; Name = $profItem.Name; Error = $msg })
                    Write-BaselineReportLog "  -> FAILED: $msg"
                }
                continue
            }

            # ---- Settings Catalog profiles (configurationPolicies endpoint) ----
            Write-BaselineReportLog "  PolicyKind: $profPolicyKind | Endpoint: configurationPolicies"
            try {
                $settingsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($profItem.Id)/settings?`$expand=settingDefinitions"
                $settings = @(Invoke-GraphPagedGet -Uri $settingsUri)

                # Flatten settings array if Invoke-GraphPagedGet returned a nested array
                $flatSettings = [System.Collections.Generic.List[object]]::new()
                foreach ($item in $settings) {
                    if ($item -is [System.Array]) {
                        foreach ($sub in $item) { [void]$flatSettings.Add($sub) }
                    } else {
                        [void]$flatSettings.Add($item)
                    }
                }

                # Build displayName lookup from expanded settingDefinitions
                $displayNameMap = @{}
                $optionDisplayMap = @{}
                foreach ($setting in $flatSettings) {
                    $defs = Get-SafeObjectPropertyValue -Object $setting -Name 'settingDefinitions'
                    if ($defs -and $defs -is [System.Collections.IEnumerable] -and -not ($defs -is [string])) {
                        foreach ($def in $defs) {
                            $did = Get-SafeObjectPropertyValue -Object $def -Name 'id'
                            $dname = Get-SafeObjectPropertyValue -Object $def -Name 'displayName'
                            if ($did -and $dname -and -not $displayNameMap.ContainsKey([string]$did)) {
                                $displayNameMap[[string]$did] = [string]$dname
                            }
                            # Extract choice options display names (itemId -> displayName)
                            $defOptions = Get-SafeObjectPropertyValue -Object $def -Name 'options'
                            if ($defOptions -and $defOptions -is [System.Collections.IEnumerable] -and -not ($defOptions -is [string])) {
                                foreach ($opt in $defOptions) {
                                    $optId = Get-SafeObjectPropertyValue -Object $opt -Name 'itemId'
                                    $optName = Get-SafeObjectPropertyValue -Object $opt -Name 'displayName'
                                    if ($optId -and $optName -and -not $optionDisplayMap.ContainsKey([string]$optId)) {
                                        $optionDisplayMap[[string]$optId] = [string]$optName
                                    }
                                }
                            }
                        }
                    }
                }

                # Accumulate display names into combined map for baseline resolution
                foreach ($dnKey in $displayNameMap.Keys) {
                    if (-not $allDisplayNames.ContainsKey($dnKey)) { $allDisplayNames[$dnKey] = $displayNameMap[$dnKey] }
                }
                foreach ($optKey in $optionDisplayMap.Keys) {
                    if (-not $allOptionDisplayNames.ContainsKey($optKey)) { $allOptionDisplayNames[$optKey] = $optionDisplayMap[$optKey] }
                }

                # Recursive helper to extract leaf setting instances from grouped/collection settings
                # Returns [pscustomobject]@{ Instance; Depth } for each leaf
                function Get-LeafSettingInstances {
                    param($Instance, [int]$Depth = 0)
                    if (-not $Instance) { return }
                    $odataType = Get-SafeObjectPropertyValue -Object $Instance -Name '@odata.type'
                    # Group setting collection: recurse into children
                    $groupVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'groupSettingCollectionValue'
                    if ($groupVal -and $groupVal -is [System.Collections.IEnumerable]) {
                        foreach ($group in $groupVal) {
                            $children = Get-SafeObjectPropertyValue -Object $group -Name 'children'
                            if ($children -and $children -is [System.Collections.IEnumerable]) {
                                foreach ($child in $children) {
                                    Get-LeafSettingInstances -Instance $child -Depth $Depth
                                }
                            }
                        }
                        return
                    }
                    # Group setting value (non-collection): recurse into children
                    $groupSingleVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'groupSettingValue'
                    if ($groupSingleVal) {
                        $children = Get-SafeObjectPropertyValue -Object $groupSingleVal -Name 'children'
                        if ($children -and $children -is [System.Collections.IEnumerable]) {
                            foreach ($child in $children) {
                                Get-LeafSettingInstances -Instance $child -Depth $Depth
                            }
                        }
                        return
                    }
                    # Choice setting: output this instance as a leaf, then recurse into children (dependedOnBy sub-settings)
                    $choiceVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'choiceSettingValue'
                    if ($choiceVal) {
                        [pscustomobject]@{ Instance = $Instance; Depth = $Depth }
                        $choiceChildren = Get-SafeObjectPropertyValue -Object $choiceVal -Name 'children'
                        if ($choiceChildren -and $choiceChildren -is [System.Collections.IEnumerable] -and -not ($choiceChildren -is [string])) {
                            foreach ($child in $choiceChildren) {
                                Get-LeafSettingInstances -Instance $child -Depth ($Depth + 1)
                            }
                        }
                        return
                    }
                    # Simple setting: output this instance as a leaf, then recurse into children if present
                    $simpleVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'simpleSettingValue'
                    if ($simpleVal) {
                        [pscustomobject]@{ Instance = $Instance; Depth = $Depth }
                        $simpleChildren = Get-SafeObjectPropertyValue -Object $simpleVal -Name 'children'
                        if ($simpleChildren -and $simpleChildren -is [System.Collections.IEnumerable] -and -not ($simpleChildren -is [string])) {
                            foreach ($child in $simpleChildren) {
                                Get-LeafSettingInstances -Instance $child -Depth ($Depth + 1)
                            }
                        }
                        return
                    }
                    # Fallback: output as leaf
                    [pscustomobject]@{ Instance = $Instance; Depth = $Depth }
                }

                $flatSettingCount = 0

                foreach ($setting in $flatSettings) {
                    $instance = Get-SafeObjectPropertyValue -Object $setting -Name 'settingInstance'
                    if (-not $instance -and $setting -is [System.Collections.IDictionary] -and $setting.Contains('settingInstance')) {
                        $instance = $setting['settingInstance']
                    }

                    $leafInstances = @(Get-LeafSettingInstances -Instance $instance)
                    if ($leafInstances.Count -eq 0 -and $instance) {
                        $leafInstances = @([pscustomobject]@{ Instance = $instance; Depth = 0 })
                    }

                    $parentGroupName = $null
                    foreach ($leafItem in $leafInstances) {
                        if (-not $leafItem) { continue }
                        $leaf = $leafItem.Instance
                        $leafDepth = [int]$leafItem.Depth
                        if (-not $leaf) { continue }
                        $defId = Get-SafeObjectPropertyValue -Object $leaf -Name 'settingDefinitionId'
                        if (-not $defId) { $defId = '-' } else { $defId = [string]$defId }

                        # Extract value from leaf
                        $leafValue = '-'
                        $leafRawValue = '-'
                        $choiceVal = Get-SafeObjectPropertyValue -Object $leaf -Name 'choiceSettingValue'
                        $simpleVal = Get-SafeObjectPropertyValue -Object $leaf -Name 'simpleSettingValue'
                        $simpleCollVal = Get-SafeObjectPropertyValue -Object $leaf -Name 'simpleSettingCollectionValue'
                        if ($choiceVal) {
                            $v = Get-SafeObjectPropertyValue -Object $choiceVal -Name 'value'
                            if ($v) {
                                $vStr = [string]$v
                                $leafRawValue = $vStr
                                $leafValue = if ($optionDisplayMap.ContainsKey($vStr)) { $optionDisplayMap[$vStr] } else { $vStr }
                            }
                        } elseif ($simpleVal) {
                            $v = Get-SafeObjectPropertyValue -Object $simpleVal -Name 'value'
                            if ($null -ne $v) { $leafValue = [string]$v; $leafRawValue = [string]$v }
                        } elseif ($simpleCollVal -and $simpleCollVal -is [System.Collections.IEnumerable] -and -not ($simpleCollVal -is [string])) {
                            $collValues = @($simpleCollVal | ForEach-Object { $sv = Get-SafeObjectPropertyValue -Object $_ -Name 'value'; if ($null -ne $sv) { [string]$sv } })
                            $leafValue = $collValues -join '; '
                            $leafRawValue = $leafValue
                        } else {
                            $leafValue = Get-ProfileReportSettingValue -Setting $setting
                            $leafRawValue = $leafValue
                        }

                        # Use friendly displayName from expanded definitions; fall back to technical name
                        $settingName = if ($defId -ne '-' -and $displayNameMap.ContainsKey($defId)) {
                            $displayNameMap[$defId]
                        } else {
                            if ($defId -ne '-') { [void]$unresolvedDefIds.Add($defId) }
                            # Extract last segment of the ID as fallback name
                            if ($defId -match '/') { ($defId -split '/')[-1] } else { $defId }
                        }

                        # Track parent name for sort grouping
                        if ($leafDepth -eq 0) { $parentGroupName = $settingName }
                        [void]$settingRows.Add([pscustomobject]@{
                            ProfileId    = $profItem.Id
                            ProfileName  = $profItem.Name
                            Platforms    = $profItem.Platforms
                            SettingName  = $settingName
                            SettingValue = $leafValue
                            RawValue     = $leafRawValue
                            DefinitionId = $defId
                            Depth        = $leafDepth
                            SortGroup    = if ($parentGroupName) { $parentGroupName } else { $settingName }
                            Order        = $flatSettingCount
                        })
                        $flatSettingCount++
                    }
                }
                Write-BaselineReportLog "  -> $flatSettingCount settings (from $($flatSettings.Count) API objects)"
            } catch {
                $msg = $_.Exception.Message
                [void]$failedProfiles.Add([pscustomobject]@{ Id = $profItem.Id; Name = $profItem.Name; Error = $msg })
                Write-BaselineReportLog "  -> FAILED: $msg"
            }
        }

        if ($policyKindCounts.Keys.Count -gt 0) {
            $policySummary = @($policyKindCounts.Keys | Sort-Object | ForEach-Object { "$_=$($policyKindCounts[$_])" }) -join ', '
            Write-BaselineReportLog "Profile types loaded: $policySummary"
        }

        # Resolve setting definition IDs not returned by $expand=settingDefinitions via direct GET
        if ($unresolvedDefIds.Count -gt 0) {
            if ($cacheMode -eq 'cacheonly') {
                Write-BaselineReportLog "Strict cache-only mode: skipping Graph lookup for $($unresolvedDefIds.Count) unresolved setting definition(s)."
            } else {
                Write-BaselineReportLog "Resolving $($unresolvedDefIds.Count) unresolved setting definition(s) via Graph..."
                Write-BaselineReportProgress -Status 'Resolving setting names...'
                $resolveStats = Resolve-SettingDefinitionsViaGraph `
                    -DefinitionIds @($unresolvedDefIds) `
                    -DisplayNameMap $allDisplayNames `
                    -OptionDisplayMap $allOptionDisplayNames `
                    -ProgressCallback {
                        param($done, $total)
                        Write-BaselineReportProgress -Status "Resolving names $done/$total"
                    } `
                    -LogCallback {
                        param($message)
                        Write-BaselineReportLog $message
                    } `
                    -AuthChallengeCallback {
                        param($reason)
                        Write-BaselineReportLog 'Reconnecting Microsoft Graph due to authentication challenge...'
                        Connect-BaselineReadGraph -TenantId (Get-ResolverTenantOverride) -ForceReconnect
                    }

                $cachedUnresolved = $unresolvedDefIds.Count - $resolveStats.Requested
                if ($cachedUnresolved -gt 0) {
                    Write-BaselineReportLog "Cache hit(s) for unresolved definitions: $cachedUnresolved"
                }
            }

            $resolved = 0
            foreach ($row in $settingRows) {
                if ($row.DefinitionId -and $row.DefinitionId -ne '-' -and $allDisplayNames.ContainsKey($row.DefinitionId)) {
                    $row.SettingName = $allDisplayNames[$row.DefinitionId]
                    $resolved++
                }
                # Resolve choice option values to friendly names
                if ($row.SettingValue -and $allOptionDisplayNames.ContainsKey($row.SettingValue)) {
                    $row.SettingValue = $allOptionDisplayNames[$row.SettingValue]
                }
            }
            Write-BaselineReportLog "Resolved $resolved setting name(s) from available cache/Graph lookups."
        }

        $baselineComparison = $null
        
        if ($baselineFolder -and (Test-Path $baselineFolder -PathType Container)) {
            Write-BaselineReportLog "Comparing selected policies with JSON backups in: $baselineFolder"
            Write-BaselineReportProgress -Status 'Comparing policy backups...'
            $baselineComparison = Compare-IntunePolicyBackup -CompareFolder $baselineFolder -Profiles $targetProfiles -SettingRows @($settingRows) -FailedProfiles @($failedProfiles)
            $matchCount = @($baselineComparison.ProfileStatus.Values | Where-Object { $_ -eq 'Match' }).Count
            $mismatchCount = @($baselineComparison.ProfileStatus.Values | Where-Object { $_ -eq 'Changed' }).Count
            $missingSettingCount = @($baselineComparison.MissingRows).Count
            Write-BaselineReportLog "Policy backups: $matchCount unchanged | $mismatchCount changed | $missingSettingCount settings missing from live"
            foreach ($diagnostic in $baselineComparison.Diagnostics) { Write-BaselineReportLog "Backup not compared: $diagnostic" }

            # Resolve display names for missing settings not already known
            if ($missingSettingCount -gt 0) {
                $missingToResolve = @($baselineComparison.MissingRows | ForEach-Object { [string]$_.DefinitionId } | Where-Object { $_ -and -not $allDisplayNames.ContainsKey($_) })
                if ($missingToResolve.Count -gt 0) {
                    if ($cacheMode -eq 'cacheonly') {
                        Write-BaselineReportLog "Strict cache-only mode: skipping Graph lookup for $($missingToResolve.Count) missing backup setting name(s)."
                    } else {
                        Write-BaselineReportLog "Resolving $($missingToResolve.Count) missing setting name(s) via Graph..."
                        Write-BaselineReportProgress -Status 'Resolving missing setting names...'
                        $missingStats = Resolve-SettingDefinitionsViaGraph `
                            -DefinitionIds @($missingToResolve) `
                            -DisplayNameMap $allDisplayNames `
                            -OptionDisplayMap $allOptionDisplayNames `
                            -ProgressCallback {
                                param($done, $total)
                                Write-BaselineReportProgress -Status "Resolving missing names $done/$total"
                            } `
                            -LogCallback {
                                param($message)
                                Write-BaselineReportLog $message
                            } `
                            -AuthChallengeCallback {
                                param($reason)
                                Write-BaselineReportLog 'Reconnecting Microsoft Graph due to authentication challenge...'
                                Connect-BaselineReadGraph -TenantId (Get-ResolverTenantOverride) -ForceReconnect
                            }

                        Write-BaselineReportLog "Resolved $($missingStats.Resolved) missing setting name(s) in $($missingStats.GraphCalls) batch call(s)."
                    }
                }
            }
            $baselineComparison['DisplayNames'] = $allDisplayNames
            $baselineComparison['OptionDisplayNames'] = $allOptionDisplayNames
        }

        $savedCache = Export-SettingDefinitionCache -TenantKey $resolverTenantCacheKey -DisplayNameMap $allDisplayNames -OptionDisplayMap $allOptionDisplayNames -SourceMode "report-$cacheMode"
        Write-BaselineReportLog "Updated resolver cache: $($savedCache.DisplayNameCount) names, $($savedCache.OptionCount) options."

        New-IntuneProfileSettingsHtmlReport -OutputPath $outputPath -Profiles $targetProfiles -SettingRows @($settingRows) -FailedProfiles @($failedProfiles) -BaselineComparison $baselineComparison | Out-Null
    [pscustomobject]@{ OutputPath = $OutputPath; Profiles = $TargetProfiles; SettingRows = @($settingRows); FailedProfiles = @($failedProfiles); BaselineComparison = $baselineComparison; PolicyComparison = $baselineComparison }
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    if ($ListProfiles -and $WarmCache) { throw 'Use -ListProfiles or -WarmCache, not both.' }
    if (-not $ListProfiles -and -not $WarmCache) {
        if (-not $ProfileId -and -not $Profiles) { throw 'Supply -ProfileId or -Profiles, or use -ListProfiles.' }
        if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw '-OutputPath is required to generate a report.' }
        if ($BaselineFolder -and -not (Test-Path -LiteralPath $BaselineFolder -PathType Container)) { throw "Compare folder not found: $BaselineFolder" }
    }
    Initialize-BaselineGraphSession -TenantId $TenantId
    if ($ListProfiles) {
        (Get-BaselineIntuneProfiles).Mapped
    } elseif ($WarmCache) {
        $targetProfiles = @(Select-BaselineIntuneProfiles -ProfileId $ProfileId -Profiles $Profiles)
        Invoke-TenantPolicyCacheWarmup -Profiles $targetProfiles -TenantKey (Get-GraphTenantCacheKey) -CacheMode $CacheMode -LogCallback $LogCallback -ProgressCallback $ProgressCallback
    } else {
        $targetProfiles = @(Select-BaselineIntuneProfiles -ProfileId $ProfileId -Profiles $Profiles)
        if (-not $targetProfiles.Count) { throw 'No profiles selected.' }
        Invoke-BaselineProfileReport -TargetProfiles $targetProfiles -OutputPath $OutputPath -BaselineFolder $BaselineFolder -CacheMode $CacheMode -LogCallback $LogCallback -ProgressCallback $ProgressCallback
        if ($OpenReport) { Start-Process -FilePath $OutputPath }
    }
}
