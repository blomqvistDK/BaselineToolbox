#Requires -Version 5.1

$script:ResolverTenantId = $env:BASELINE_TOOLBOX_RESOLVER_TENANTID

function Set-BaselineResolverTenant {
    param([string]$TenantId)
    $script:ResolverTenantId = $TenantId
}

function Initialize-BaselineGraphSession {
    param([string]$TenantId, [switch]$WriteAccess)
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Set-BaselineResolverTenant -TenantId $TenantId
    if (-not (Test-MgSessionIsValid) -or -not (Test-MgSessionTenantMatch -TenantId $TenantId)) {
        Connect-BaselineReadGraph -TenantId $TenantId
    }
    if ($WriteAccess) {
        $context = Get-MgContext
        if ($context.AuthType -ne 'AppOnly' -and 'DeviceManagementConfiguration.ReadWrite.All' -notin $context.Scopes) {
            $connectArguments = @{ Scopes = @('DeviceManagementConfiguration.ReadWrite.All','Organization.Read.All'); NoWelcome = $true; ErrorAction = 'Stop' }
            if ($TenantId) { $connectArguments.TenantId = $TenantId }
            Connect-MgGraph @connectArguments | Out-Null
        }
    }
}

function Select-BaselineIntuneProfiles {
    param([string[]]$ProfileId, [object[]]$Profiles)
    $candidates = if ($Profiles) { @($Profiles) } else { @((Get-BaselineIntuneProfiles).Mapped) }
    if ($ProfileId) {
        $selected = @($candidates | Where-Object { $_.Id -in $ProfileId })
        $missing = @($ProfileId | Where-Object { $_ -notin @($selected | ForEach-Object { $_.Id }) })
        if ($missing.Count) { throw "Profile IDs not found: $($missing -join ', ')" }
        return $selected
    }
    return $candidates
}

function Normalize-CompareText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $normalized = [string]$Text
    $normalized = $normalized -replace '[\u200B-\u200D\uFEFF]', ''
    $normalized = $normalized.Trim()
    return $normalized.ToLowerInvariant()
}

function Get-BaselinePolicyPreview {
    param(
        [string]$MdFilePath,
        [ValidateSet('L1','L2','All')][string]$ProfileLevel = 'L1',
        [ValidateSet('All','CIS','Twoday')][string]$OriginFilter = 'All'
    )
    $parsedData = Parse-MdFile -Path $MdFilePath
    $recommendations = @(switch ($ProfileLevel) {
        'L1' { $parsedData | Where-Object { ($_.Followed -eq 'Yes' -and (([string]$_.Profile) -match '(^|\s)L1$' -or $_.Profile -eq 'BL')) -or $_.Origin -imatch 'twoday' } }
        default { $parsedData }
    })
    $recommendations = @(switch ($OriginFilter) {
        'CIS' { $recommendations | Where-Object { $_.Origin -notmatch 'twoday' } }
        'Twoday' { $recommendations | Where-Object { $_.Origin -imatch 'twoday' } }
        default { $recommendations }
    })
    [pscustomobject]@{
        TotalCount = @($parsedData).Count
        Recommendations = $recommendations
        PolicyCount = @($recommendations | ForEach-Object { $_.IntunePolicy } | Sort-Object -Unique).Count
    }
}

function Get-PlainText {
    param([string]$Cell)
    $t = $Cell.Trim()
    $t = [regex]::Replace($t, '\*\*([^*]*)\*\*', '$1')
    $t = [regex]::Replace($t, '\[([^\]]*)\]\([^)]*\)', '$1')
    $t = $t -replace '<br\s*/?>', ' '
    return $t.Trim()
}

function Get-RecKey {
    param([string]$Cell)
    $t = Get-PlainText $Cell
    $t = $t -replace [char]0x00E0, [char]0x2026
    $t = $t -replace '\.\.\.', [char]0x2026
    return $t.ToLower().Trim()
}

function Split-PipeRow {
    param([string]$Line)
    $s = $Line.Trim()
    if ($s.StartsWith('|')) { $s = $s.Substring(1) }
    if ($s.EndsWith('|'))   { $s = $s.Substring(0, $s.Length - 1) }
    return @($s -split '\|' | ForEach-Object { $_.Trim() })
}

function Test-IsSeparatorRow {
    param([string[]]$Cells)
    foreach ($c in $Cells) { if ($c -and $c -notmatch '^[\-:| ]+$') { return $false } }
    return ($Cells.Count -gt 0)
}

function Get-ColumnIndex {
    param([string[]]$Headers, [string]$Name)
    for ($i = 0; $i -lt $Headers.Count; $i++) { if ($Headers[$i] -ieq $Name) { return $i } }
    return -1
}

function Build-PipeRow {
    param([string[]]$Cells)
    return '| ' + ($Cells -join ' | ') + ' |'
}

function Get-StringSimilarity {
    param([string]$Left, [string]$Right)

    $a = if ($null -eq $Left) { '' } else { [string]$Left }
    $b = if ($null -eq $Right) { '' } else { [string]$Right }
    $a = $a.Trim().ToLowerInvariant()
    $b = $b.Trim().ToLowerInvariant()

    if ($a -eq $b) { return 1.0 }
    if (-not $a -or -not $b) { return 0.0 }

    $leftBigrams = [System.Collections.Generic.List[string]]::new()
    $rightBigrams = [System.Collections.Generic.List[string]]::new()

    if ($a.Length -lt 2) { $leftBigrams.Add($a) } else {
        for ($i = 0; $i -lt ($a.Length - 1); $i++) { $leftBigrams.Add($a.Substring($i, 2)) }
    }
    if ($b.Length -lt 2) { $rightBigrams.Add($b) } else {
        for ($i = 0; $i -lt ($b.Length - 1); $i++) { $rightBigrams.Add($b.Substring($i, 2)) }
    }

    $leftCount = [System.Collections.Generic.Dictionary[string,int]]::new([System.StringComparer]::Ordinal)
    foreach ($bg in $leftBigrams) {
        if ($leftCount.ContainsKey($bg)) { $leftCount[$bg] = $leftCount[$bg] + 1 }
        else { $leftCount[$bg] = 1 }
    }

    $intersection = 0
    foreach ($bg in $rightBigrams) {
        if ($leftCount.ContainsKey($bg) -and $leftCount[$bg] -gt 0) {
            $intersection++
            $leftCount[$bg] = $leftCount[$bg] - 1
        }
    }

    return (2.0 * $intersection) / ($leftBigrams.Count + $rightBigrams.Count)
}

function Read-MarkdownDocument {
    param([string]$FilePath)
    $raw   = Get-Content -Path $FilePath -Raw -Encoding UTF8
    $raw   = $raw -replace "`r`n", "`n" -replace "`r", "`n"
    $lines = $raw -split "`n"
    $sections = [System.Collections.Generic.List[hashtable]]::new()
    $preamble = [System.Collections.Generic.List[string]]::new()
    $cur      = $null
    $tblBuf   = [System.Collections.Generic.List[string]]::new()
    foreach ($ln in $lines) {
        if ($ln -match '^(#{1,6})\s+(.+)') {
            if ($cur) { $cur['TableLines'] = $tblBuf.ToArray(); $sections.Add($cur) }
            $cur    = @{ HeadingLine = $ln; HeadingLevel = $matches[1].Length; HeadingText = $matches[2].Trim(); TableLines = @() }
            $tblBuf = [System.Collections.Generic.List[string]]::new()
        }
        elseif ($cur -and $ln.TrimStart().StartsWith('|')) { $tblBuf.Add($ln) }
        elseif (-not $cur) { $preamble.Add($ln) }
    }
    if ($cur) { $cur['TableLines'] = $tblBuf.ToArray(); $sections.Add($cur) }
    return @{ Preamble = $preamble.ToArray(); Sections = $sections.ToArray() }
}

function Read-MarkdownTable {
    param([string[]]$TableLines)
    if ($TableLines.Count -lt 2) { return $null }
    $hdrs      = @(Split-PipeRow $TableLines[0] | ForEach-Object { Get-PlainText $_ })
    $sepCells  = Split-PipeRow $TableLines[1]
    $dataStart = if (Test-IsSeparatorRow $sepCells) { 2 } else { 1 }
    $rows = [System.Collections.Generic.List[hashtable]]::new()
    for ($i = $dataStart; $i -lt $TableLines.Count; $i++) {
        $ln = $TableLines[$i]
        if (-not $ln.TrimStart().StartsWith('|')) { continue }
        $cells = @(Split-PipeRow $ln)
        while ($cells.Count -lt $hdrs.Count) { $cells += '' }
        $rows.Add(@{ Cells = $cells })
    }
    return @{ HeaderLine = $TableLines[0]; SeparatorLine = if ($dataStart -eq 2) { $TableLines[1] } else { '| ' + (($hdrs | ForEach-Object { '---' }) -join ' | ') + ' |' }; Headers = $hdrs; Rows = $rows.ToArray() }
}

function Get-BaselineExportPolicyKey {
    param([string]$PolicyName)

    $value = if ($null -eq $PolicyName) { '' } else { [string]$PolicyName }
    $value = $value -replace '[\u200B-\u200D\uFEFF]', ''
    return [regex]::Replace($value, '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function Get-BaselineExportSettingKey {
    param([string]$SettingName)

    $value = if ($null -eq $SettingName) { '' } else { Get-PlainText ([string]$SettingName) }
    $value = $value -replace '[\u200B-\u200D\uFEFF]', ''
    $value = [regex]::Replace($value, '^\s*Ensure\s+', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $value = [regex]::Replace($value, '\s+is\s+(?:set\s+to|configured\s+to)\s+.*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $value = [regex]::Replace($value, '\s+is\s+configured\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $value = [regex]::Replace($value, '\s+is\s+(?:enabled|disabled)[.!?]*\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # Quotes must go first, otherwise the (User)/(Device) suffix is not at the end of the string.
    $value = $value.Trim().Trim([char[]]@([char]39, [char]34, [char]0x2018, [char]0x2019, [char]0x201C, [char]0x201D))
    $value = [regex]::Replace($value, '\s*\((?:user|device)\)\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    return [regex]::Replace($value, '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function Import-BaselineAssessmentDetailsCsv {
    param([string]$Path)

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "Details CSV not found: $Path"
    }

    $requiredHeaders = @(
        'Baseline Policy',
        'Setting Name',
        'Expected Value',
        'Status',
        'Policy Assigned'
    )
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
    $headerParser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path, [System.Text.Encoding]::UTF8)
    try {
        $headerParser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $headerParser.SetDelimiters(';')
        $headerParser.HasFieldsEnclosedInQuotes = $true
        $headers = @($headerParser.ReadFields())
    } finally {
        $headerParser.Close()
    }

    $missingHeaders = @($requiredHeaders | Where-Object { $_ -notin $headers })
    if ($missingHeaders.Count -gt 0) {
        throw "This is not a Baseline Assessment Details CSV. Missing column(s): $($missingHeaders -join ', '). Use Export Details CSV from the assessment report, not Export Summary CSV."
    }

    $rows = @(Import-Csv -Path $Path -Delimiter ';' -Encoding UTF8 -ErrorAction Stop)
    if ($rows.Count -eq 0) {
        throw 'The selected Details CSV contains no detail rows.'
    }

    return $rows
}

function Merge-BaselineAssessmentDetailRows {
    param([object[]]$Rows)

    $groupOrder = [System.Collections.Generic.List[string]]::new()
    $groups = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]]::new([System.StringComparer]::Ordinal)
    foreach ($row in $Rows) {
        $definitionIdProperty = $row.PSObject.Properties['Definition Id']
        $identity = if ($definitionIdProperty -and ([string]$definitionIdProperty.Value).Trim()) { [string]$definitionIdProperty.Value } else { [string]$row.'Setting Name' }
        $key = '{0}|{1}|{2}' -f (Normalize-CompareText -Text ([string]$row.'Baseline Policy')), (Normalize-CompareText -Text $identity), (Normalize-CompareText -Text ([string]$row.'Expected Value'))
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [System.Collections.Generic.List[object]]::new()
            [void]$groupOrder.Add($key)
        }
        [void]$groups[$key].Add($row)
    }

    $merged = [System.Collections.Generic.List[object]]::new()
    foreach ($key in $groupOrder) {
        $groupRows = @($groups[$key])
        if ($groupRows.Count -eq 1) {
            [void]$merged.Add($groupRows[0])
            continue
        }

        # Unassigned policies don't apply to devices, so they are ignored when an assigned one exists.
        $assignedRows = @($groupRows | Where-Object { [string]$_.'Policy Assigned' -ieq 'Yes' })
        $effectiveRows = if ($assignedRows.Count -gt 0) { $assignedRows } else { $groupRows }

        $statuses = @($effectiveRows | ForEach-Object { [string]$_.Status } | Sort-Object -Unique)
        $status = if (($statuses -contains 'Match') -and ($statuses -contains 'Mismatch')) { 'Conflict' } else { [string]$effectiveRows[0].Status }

        $assigned = if ($assignedRows.Count -gt 0) { 'Yes' } else {
            $firstAssignment = @($effectiveRows | Where-Object { ([string]$_.'Policy Assigned').Trim() } | Select-Object -First 1)
            if ($firstAssignment.Count -gt 0) { [string]$firstAssignment[0].'Policy Assigned' } else { '' }
        }

        $mergedRow = $effectiveRows[0].PSObject.Copy()
        $mergedRow.Status = $status
        $mergedRow.'Policy Assigned' = $assigned
        [void]$merged.Add($mergedRow)
    }

    return [pscustomobject]@{
        Rows = $merged.ToArray()
        DuplicatesRemoved = $Rows.Count - $merged.Count
    }
}

function New-BaselineCisReferenceLookup {
    param([string]$MarkdownFolder)

    if (-not (Test-Path -Path $MarkdownFolder -PathType Container)) {
        throw "Markdown folder not found: $MarkdownFolder"
    }

    $markdownFiles = @(Get-ChildItem -Path $MarkdownFolder -Filter '*.md' -File -ErrorAction Stop)
    if ($markdownFiles.Count -eq 0) {
        throw "No Markdown files were found in: $MarkdownFolder"
    }

    $lookup = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $gpLookup = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $settingOnlyLookup = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $gpSettingOnlyLookup = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $candidatesByPolicy = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $hierarchicalParentsByPolicy = [System.Collections.Generic.Dictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($markdownFile in $markdownFiles) {
        $document = Read-MarkdownDocument -FilePath $markdownFile.FullName
        foreach ($section in $document.Sections) {
            if (-not $section['TableLines'] -or $section['TableLines'].Count -lt 2) { continue }
            $table = Read-MarkdownTable -TableLines $section['TableLines']
            if (-not $table) { continue }

            $recommendationIndex = Get-ColumnIndex -Headers $table.Headers -Name 'Recommendation'
            $referenceIndex = Get-ColumnIndex -Headers $table.Headers -Name 'CIS Reference'
            $policyIndex = Get-ColumnIndex -Headers $table.Headers -Name 'Intune Policy'
            $originIndex = Get-ColumnIndex -Headers $table.Headers -Name 'Origin'
            $gpSettingIndex = Get-ColumnIndex -Headers $table.Headers -Name 'GP Setting Name'
            if ($recommendationIndex -lt 0 -or $referenceIndex -lt 0 -or $policyIndex -lt 0) { continue }

            foreach ($row in $table.Rows) {
                $policyKey = Get-BaselineExportPolicyKey -PolicyName (Get-PlainText $row.Cells[$policyIndex])
                $settingKey = Get-BaselineExportSettingKey -SettingName $row.Cells[$recommendationIndex]
                if (-not $policyKey -or -not $settingKey) { continue }

                $parentMatch = [regex]::Match([string]$row.Cells[$recommendationIndex], "^\s*Ensure\s+.(?<Parent>[^:]+):", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($parentMatch.Success) {
                    $parentName = $parentMatch.Groups['Parent'].Value.Trim().Trim([char[]]@([char]39, [char]34, [char]0x2018, [char]0x2019, [char]0x201C, [char]0x201D))
                    $parentKey = Get-BaselineExportSettingKey -SettingName $parentName
                    if ($parentKey) {
                        if (-not $hierarchicalParentsByPolicy.ContainsKey($policyKey)) {
                            $hierarchicalParentsByPolicy[$policyKey] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                        }
                        [void]$hierarchicalParentsByPolicy[$policyKey].Add($parentKey)
                    }
                }

                $cisReference = Get-PlainText $row.Cells[$referenceIndex]
                if ($cisReference -eq '-') { $cisReference = '' }
                $origin = if ($originIndex -ge 0) { Get-PlainText $row.Cells[$originIndex] } else { '' }
                $outputReference = if ($origin -imatch '\btwoday\b') { 'twoday' } else { $cisReference }
                $lookupKey = "$policyKey|$settingKey"
                if (-not $lookup.ContainsKey($lookupKey)) {
                    $lookup[$lookupKey] = [System.Collections.Generic.List[object]]::new()
                }
                $gpName = if ($gpSettingIndex -ge 0 -and $gpSettingIndex -lt $row.Cells.Count) {
                    Get-PlainText $row.Cells[$gpSettingIndex]
                } else { '' }
                $candidate = [pscustomobject]@{
                    CISReference = $cisReference
                    Origin = $origin
                    OutputReference = $outputReference
                    PolicyKey = $policyKey
                    SettingKey = $settingKey
                    RecommendationText = Get-PlainText $row.Cells[$recommendationIndex]
                    GpSettingName = $gpName
                    MarkdownFile = $markdownFile.Name
                }
                [void]$lookup[$lookupKey].Add($candidate)
                if (-not $settingOnlyLookup.ContainsKey($settingKey)) {
                    $settingOnlyLookup[$settingKey] = [System.Collections.Generic.List[object]]::new()
                }
                [void]$settingOnlyLookup[$settingKey].Add($candidate)
                if ($gpSettingIndex -ge 0 -and $gpSettingIndex -lt $row.Cells.Count) {
                    $gpName = Get-PlainText $row.Cells[$gpSettingIndex]
                    $gpKeys = @(
                        Get-BaselineExportSettingKey -SettingName $gpName
                        Get-BaselineExportSettingKey -SettingName (($gpName -split '\\')[-1])
                    ) | Where-Object { $_ } | Select-Object -Unique
                    foreach ($gpKey in $gpKeys) {
                        $gpLookupKey = "$policyKey|$gpKey"
                        if (-not $gpLookup.ContainsKey($gpLookupKey)) {
                            $gpLookup[$gpLookupKey] = [System.Collections.Generic.List[object]]::new()
                        }
                        [void]$gpLookup[$gpLookupKey].Add($candidate)
                        if (-not $gpSettingOnlyLookup.ContainsKey($gpKey)) {
                            $gpSettingOnlyLookup[$gpKey] = [System.Collections.Generic.List[object]]::new()
                        }
                        [void]$gpSettingOnlyLookup[$gpKey].Add($candidate)
                    }
                }
                if (-not $candidatesByPolicy.ContainsKey($policyKey)) {
                    $candidatesByPolicy[$policyKey] = [System.Collections.Generic.List[object]]::new()
                }
                [void]$candidatesByPolicy[$policyKey].Add($candidate)
            }
        }
    }

    return [pscustomobject]@{
        ExactByKey = $lookup
        ExactByGpKey = $gpLookup
        ExactBySettingOnly = $settingOnlyLookup
        ExactByGpSettingOnly = $gpSettingOnlyLookup
        CandidatesByPolicy = $candidatesByPolicy
        HierarchicalParentsByPolicy = $hierarchicalParentsByPolicy
    }
}

function Get-BaselineAssessmentSettingPaths {
    param(
        [object[]]$Rows,
        $Lookup
    )

    $paths = [System.Collections.Generic.List[string]]::new()
    $parentPaths = [System.Collections.Generic.List[string]]::new()
    $activePolicyKey = $null
    $activeParent = $null
    foreach ($row in $Rows) {
        $settingName = [string]$row.'Setting Name'
        $policyKey = Get-BaselineExportPolicyKey -PolicyName ([string]$row.'Baseline Policy')
        if ($policyKey -ine $activePolicyKey) {
            $activePolicyKey = $policyKey
            $activeParent = $null
            $parentPaths.Clear()
        }

        $depthProperty = $row.PSObject.Properties['Setting Depth']
        $hasDepth = $depthProperty -and [string]$depthProperty.Value -match '^\d+$'
        $depth = if ($hasDepth) { [int]$depthProperty.Value } else { 0 }
        if ($depth -eq 0) {
            $settingKey = Get-BaselineExportSettingKey -SettingName $settingName
            $knownParents = if ($Lookup.HierarchicalParentsByPolicy.ContainsKey($policyKey)) { $Lookup.HierarchicalParentsByPolicy[$policyKey] } else { $null }
            if ($knownParents -and $knownParents.Contains($settingKey)) {
                $activeParent = $settingName
                $parentPaths.Clear()
                [void]$parentPaths.Add($settingName)
                [void]$paths.Add($settingName)
                continue
            }

            $settingPath = $settingName
            if ($activeParent) {
                $candidatePath = '{0}: {1}' -f $activeParent, $settingName
                $candidateKey = Get-BaselineExportSettingKey -SettingName $candidatePath
                $candidateLookupKey = "$policyKey|$candidateKey"
                if ($Lookup.ExactByKey.ContainsKey($candidateLookupKey) -or $Lookup.ExactByGpKey.ContainsKey($candidateLookupKey)) {
                    $settingPath = $candidatePath
                }
            }
            $parentPaths.Clear()
            [void]$parentPaths.Add($settingPath)
            [void]$paths.Add($settingPath)
            continue
        }

        if (-not $hasDepth) {
            $parentPaths.Clear()
            [void]$paths.Add($settingName)
            continue
        }

        while ($parentPaths.Count -gt $depth) {
            $parentPaths.RemoveAt($parentPaths.Count - 1)
        }

        $settingPath = $settingName
        if ($depth -gt 0 -and $parentPaths.Count -ge $depth -and $parentPaths[$depth - 1]) {
            $settingPath = '{0}: {1}' -f $parentPaths[$depth - 1], $settingName
        }

        while ($parentPaths.Count -lt $depth) { [void]$parentPaths.Add('') }
        if ($parentPaths.Count -eq $depth) { [void]$parentPaths.Add($settingPath) }
        else { $parentPaths[$depth] = $settingPath }
        [void]$paths.Add($settingPath)
    }

    return $paths.ToArray()
}

function Get-BaselineAssessmentHierarchy {
    param(
        [object[]]$Rows,
        $Lookup
    )

    $legacyPaths = @(Get-BaselineAssessmentSettingPaths -Rows $Rows -Lookup $Lookup)
    $nodes = [System.Collections.Generic.List[object]]::new()
    $definitions = @{}
    $previousPaths = @{}
    for ($sourceIndex = 0; $sourceIndex -lt $Rows.Count; $sourceIndex++) {
        $row = $Rows[$sourceIndex]
        $policyKey = ([string]$row.'Baseline Policy').Trim().ToLowerInvariant()
        $definitionProperty = $row.PSObject.Properties['Definition Id']
        $definitionId = if ($definitionProperty) { ([string]$definitionProperty.Value).Trim().ToLowerInvariant() } else { '' }
        $node = [pscustomobject]@{
            Row = $row
            SourceIndex = $sourceIndex
            PolicyKey = $policyKey
            DefinitionId = $definitionId
            ParentIndex = -1
            Children = [System.Collections.Generic.List[int]]::new()
            Path = [string]$row.'Setting Name'
            Depth = 0
        }
        [void]$nodes.Add($node)
        if ($definitionId) {
            $definitionKey = "$policyKey|$definitionId"
            if (-not $definitions.ContainsKey($definitionKey)) {
                $definitions[$definitionKey] = [System.Collections.Generic.List[int]]::new()
            }
            [void]$definitions[$definitionKey].Add($sourceIndex)
        }
    }

    foreach ($node in $nodes) {
        if ($node.DefinitionId) {
            $prefix = $node.DefinitionId
            while ($prefix.LastIndexOf('_') -ge 0) {
                $prefix = $prefix.Substring(0, $prefix.LastIndexOf('_'))
                $parentKey = '{0}|{1}' -f $node.PolicyKey, $prefix
                if ($definitions.ContainsKey($parentKey)) {
                    if ($definitions[$parentKey].Count -eq 1) {
                        $node.ParentIndex = $definitions[$parentKey][0]
                    }
                    break
                }
            }
        } else {
            $legacyPath = $legacyPaths[$node.SourceIndex]
            $separatorIndex = $legacyPath.LastIndexOf(': ')
            if ($separatorIndex -ge 0) {
                $parentPathKey = '{0}|{1}' -f $node.PolicyKey, $legacyPath.Substring(0, $separatorIndex)
                if ($previousPaths.ContainsKey($parentPathKey)) {
                    $node.ParentIndex = $previousPaths[$parentPathKey]
                }
            }
        }
        $pathKey = '{0}|{1}' -f $node.PolicyKey, $legacyPaths[$node.SourceIndex]
        $previousPaths[$pathKey] = $node.SourceIndex
    }

    $policyOrder = [System.Collections.Generic.List[string]]::new()
    $rootsByPolicy = @{}
    foreach ($node in $nodes) {
        if (-not $rootsByPolicy.ContainsKey($node.PolicyKey)) {
            $rootsByPolicy[$node.PolicyKey] = [System.Collections.Generic.List[int]]::new()
            [void]$policyOrder.Add($node.PolicyKey)
        }
        if ($node.ParentIndex -ge 0) {
            [void]$nodes[$node.ParentIndex].Children.Add($node.SourceIndex)
        } else {
            [void]$rootsByPolicy[$node.PolicyKey].Add($node.SourceIndex)
        }
    }

    $ordered = [System.Collections.Generic.List[object]]::new()
    $pending = [System.Collections.Generic.Stack[int]]::new()
    foreach ($policyKey in $policyOrder) {
        $roots = $rootsByPolicy[$policyKey]
        for ($rootIndex = $roots.Count - 1; $rootIndex -ge 0; $rootIndex--) { $pending.Push($roots[$rootIndex]) }
        while ($pending.Count -gt 0) {
            $node = $nodes[$pending.Pop()]
            if ($node.ParentIndex -ge 0) {
                $parent = $nodes[$node.ParentIndex]
                $node.Depth = $parent.Depth + 1
                $node.Path = '{0}: {1}' -f $parent.Path, $node.Row.'Setting Name'
            } else {
                $legacyPath = $legacyPaths[$node.SourceIndex]
                $lookupPolicyKey = Get-BaselineExportPolicyKey -PolicyName ([string]$node.Row.'Baseline Policy')
                $lookupSettingKey = Get-BaselineExportSettingKey -SettingName $legacyPath
                $lookupKey = "$lookupPolicyKey|$lookupSettingKey"
                if ($Lookup.ExactByKey.ContainsKey($lookupKey) -or $Lookup.ExactByGpKey.ContainsKey($lookupKey)) {
                    $node.Path = $legacyPath
                }
            }
            [void]$ordered.Add($node)
            for ($childIndex = $node.Children.Count - 1; $childIndex -ge 0; $childIndex--) {
                $pending.Push($node.Children[$childIndex])
            }
        }
    }
    return $ordered.ToArray()
}

function Compare-BaselineCisReference {
    param(
        [Parameter(Mandatory = $true)][string]$Left,
        [Parameter(Mandatory = $true)][string]$Right
    )

    $leftParts = @($Left -split '\.')
    $rightParts = @($Right -split '\.')
    for ($partIndex = 0; $partIndex -lt [Math]::Max($leftParts.Count, $rightParts.Count); $partIndex++) {
        $leftValue = if ($partIndex -lt $leftParts.Count) { [long]$leftParts[$partIndex] } else { 0 }
        $rightValue = if ($partIndex -lt $rightParts.Count) { [long]$rightParts[$partIndex] } else { 0 }
        if ($leftValue -lt $rightValue) { return -1 }
        if ($leftValue -gt $rightValue) { return 1 }
    }
    return 0
}

function Set-BaselineAssessmentCisReferenceInheritance {
    param(
        [object[]]$Hierarchy,
        [System.Collections.Generic.List[object]]$ReferenceResults
    )

    $indexBySourceIndex = @{}
    $effectiveReferences = [string[]]::new($Hierarchy.Count)
    $minimumReferences = [string[]]::new($Hierarchy.Count)
    for ($index = 0; $index -lt $Hierarchy.Count; $index++) {
        $indexBySourceIndex[[int]$Hierarchy[$index].SourceIndex] = $index
        $reference = [string]$ReferenceResults[$index].CISReference
        $effectiveReferences[$index] = $reference
        if ($reference -match '^\d+(?:\.\d+)*$') {
            $minimumReferences[$index] = $reference
        }
    }

    for ($index = $Hierarchy.Count - 1; $index -ge 0; $index--) {
        foreach ($childSourceIndex in $Hierarchy[$index].Children) {
            $childIndex = $indexBySourceIndex[[int]$childSourceIndex]
            $candidate = $minimumReferences[$childIndex]
            if ($candidate -and (-not $minimumReferences[$index] -or (Compare-BaselineCisReference -Left $candidate -Right $minimumReferences[$index]) -lt 0)) {
                $minimumReferences[$index] = $candidate
            }
        }
    }

    for ($index = 0; $index -lt $Hierarchy.Count; $index++) {
        $originalReference = $effectiveReferences[$index]
        $parentSourceIndex = [int]$Hierarchy[$index].ParentIndex
        if ($parentSourceIndex -ge 0) {
            $parentIndex = $indexBySourceIndex[$parentSourceIndex]
            $effectiveReference = if ($originalReference) { $originalReference } else { $effectiveReferences[$parentIndex] }
        } else {
            $effectiveReference = $originalReference
            if (-not $effectiveReference -and $minimumReferences[$index]) {
                $effectiveReference = $minimumReferences[$index]
            }
        }

        if (-not $effectiveReference) { continue }
        $effectiveReferences[$index] = $effectiveReference
        $ReferenceResults[$index].CISReference = $effectiveReference
        if (-not $originalReference) {
            $ReferenceResults[$index].Status = 'Inherited'
        }
    }
}

function Get-BaselineOfficeApplication {
    param([string]$Text, [switch]$DefinitionId)

    if (-not $Text) { return '' }
    $applications = @{
        Access = 'access'; Excel = 'excel'; PowerPoint = 'ppt'
        Publisher = 'pub'; Word = 'word'
    }
    $context = $Text
    if (-not $DefinitionId) {
        $parts = $Text -split '\\'
        if ($parts.Count -lt 2) { return '' }
        $context = $parts[0..($parts.Count - 2)] -join '\'
    }
    $matches = @(
        foreach ($application in $applications.Keys) {
            $pattern = if ($DefinitionId) {
                '(?:^|_)' + $applications[$application] + '\d+v\d+(?=[~_]|$)|(?:^|[~_])l_microsoft' + $application + '(?=[~_]|$)'
            } else {
                '(?:^|\\)\s*Microsoft\s+' + $application + '(?:\s+\d{4})?\s*(?=\\|$)'
            }
            if ($context -imatch $pattern) { $application }
        }
    )
    if ($matches.Count -eq 1) { return $matches[0] }
    return ''
}

function Select-BaselineApplicationCandidates {
    param([array]$Candidates, $CsvRow)

    if ($Candidates.Count -lt 2 -or @($Candidates.OutputReference | Select-Object -Unique).Count -lt 2) {
        return $Candidates
    }
    $definitionProperty = $CsvRow.PSObject.Properties['Definition Id']
    if (-not $definitionProperty) { return $Candidates }
    $application = Get-BaselineOfficeApplication -Text ([string]$definitionProperty.Value) -DefinitionId
    if (-not $application) { return $Candidates }

    $matching = @()
    $unknown = @()
    foreach ($candidate in $Candidates) {
        $gpProperty = $candidate.PSObject.Properties['GpSettingName']
        $candidateApplication = if ($gpProperty) { Get-BaselineOfficeApplication -Text ([string]$gpProperty.Value) } else { '' }
        if (-not $candidateApplication) { $unknown += $candidate }
        elseif ($candidateApplication -eq $application) { $matching += $candidate }
    }
    $references = @($matching | ForEach-Object { $_.OutputReference } | Select-Object -Unique)
    if ($references.Count -ne 1) { return $Candidates }
    if (@($unknown | Where-Object { $_.OutputReference -ine $references[0] }).Count -gt 0) { return $Candidates }
    return $matching
}

function Resolve-BaselineCisReference {
    param(
        [Parameter(Mandatory = $true)]
        $Lookup,
        [Parameter(Mandatory = $true)]
        $CsvRow,
        [string]$SettingPath,
        [ValidateRange(0.0,1.0)]
        [double]$SimilarityThreshold = 0.9
    )

    $policyKey = Get-BaselineExportPolicyKey -PolicyName ([string]$CsvRow.'Baseline Policy')
    $settingKey = Get-BaselineExportSettingKey -SettingName ([string]$CsvRow.'Setting Name')
    $settingPathKey = if ($SettingPath) { Get-BaselineExportSettingKey -SettingName $SettingPath } else { $settingKey }
    if (-not $policyKey -or -not $settingKey) {
        return [pscustomobject]@{ CISReference = ''; Status = 'No match' }
    }

    $exactCandidates = @()
    $exactSettingKeys = @($settingPathKey, $settingKey | Select-Object -Unique)
    foreach ($exactSettingKey in $exactSettingKeys) {
        $lookupKey = "$policyKey|$exactSettingKey"
        if ($Lookup.ExactByKey.ContainsKey($lookupKey)) {
            $exactCandidates = @($Lookup.ExactByKey[$lookupKey] | Where-Object { $_.OutputReference })
        }
        # GP names are often truncated in the Markdown, so they are only trusted for exact matches.
        if ($exactCandidates.Count -eq 0 -and $Lookup.ExactByGpKey.ContainsKey($lookupKey)) {
            $exactCandidates = @($Lookup.ExactByGpKey[$lookupKey] | Where-Object { $_.OutputReference })
        }
        if ($exactCandidates.Count -gt 0) { break }
    }
    if ($exactCandidates.Count -gt 0) {
        $exactCandidates = @(Select-BaselineApplicationCandidates -Candidates $exactCandidates -CsvRow $CsvRow)
        $references = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $isTwoday = $false
        foreach ($candidate in $exactCandidates) {
            [void]$references.Add([string]$candidate.OutputReference)
            if ([string]$candidate.OutputReference -ieq 'twoday') { $isTwoday = $true }
        }
        if ($references.Count -gt 1) {
            return [pscustomobject]@{ CISReference = ''; Status = 'Ambiguous' }
        }
        return [pscustomobject]@{
            CISReference = @($references)[0]
            Status = if ($isTwoday) { 'Twoday' } else { 'Matched' }
        }
    }

    $expectedValue = [string]$CsvRow.'Expected Value'
    $expectedValueKey = [regex]::Replace((Get-PlainText $expectedValue), '[^A-Za-z0-9]', '').ToLowerInvariant()
    if ($expectedValueKey.Length -ge 20 -and $Lookup.CandidatesByPolicy.ContainsKey($policyKey)) {
        $expectedValueCandidates = @($Lookup.CandidatesByPolicy[$policyKey] | Where-Object {
            $_.OutputReference -and [regex]::Replace([string]$_.RecommendationText, '[^A-Za-z0-9]', '').ToLowerInvariant().Contains($expectedValueKey)
        })
        if ($expectedValueCandidates.Count -gt 0) {
            $expectedValueCandidates = @(Select-BaselineApplicationCandidates -Candidates $expectedValueCandidates -CsvRow $CsvRow)
            $expectedValueReferences = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($candidate in $expectedValueCandidates) { [void]$expectedValueReferences.Add([string]$candidate.OutputReference) }
            if ($expectedValueReferences.Count -gt 1) {
                return [pscustomobject]@{ CISReference = ''; Status = 'Ambiguous' }
            }
            $expectedValueReference = @($expectedValueReferences)[0]
            return [pscustomobject]@{
                CISReference = $expectedValueReference
                Status = if ($expectedValueReference -ieq 'twoday') { 'Twoday' } else { 'Matched' }
            }
        }
    }

    # Policy names can differ between the CSV and the Markdown, so fall back to the setting name alone.
    $otherPolicyCandidates = @()
    foreach ($otherSettingKey in $exactSettingKeys) {
        if ($Lookup.ExactBySettingOnly.ContainsKey($otherSettingKey)) {
            $otherPolicyCandidates = @($Lookup.ExactBySettingOnly[$otherSettingKey] | Where-Object { $_.OutputReference })
        }
        if ($otherPolicyCandidates.Count -eq 0 -and $Lookup.ExactByGpSettingOnly.ContainsKey($otherSettingKey)) {
            $otherPolicyCandidates = @($Lookup.ExactByGpSettingOnly[$otherSettingKey] | Where-Object { $_.OutputReference })
        }
        if ($otherPolicyCandidates.Count -gt 0) { break }
    }
    if ($otherPolicyCandidates.Count -gt 0) {
        $otherPolicyCandidates = @(Select-BaselineApplicationCandidates -Candidates $otherPolicyCandidates -CsvRow $CsvRow)
        $otherReferences = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($candidate in $otherPolicyCandidates) { [void]$otherReferences.Add([string]$candidate.OutputReference) }
        if ($otherReferences.Count -gt 1) {
            return [pscustomobject]@{ CISReference = ''; Status = 'Ambiguous' }
        }
        $otherReference = @($otherReferences)[0]
        return [pscustomobject]@{
            CISReference = $otherReference
            Status = if ($otherReference -ieq 'twoday') { 'Twoday' } else { 'Matched (other policy)' }
        }
    }

    if (-not $Lookup.CandidatesByPolicy.ContainsKey($policyKey)) {
        return [pscustomobject]@{ CISReference = ''; Status = 'No match' }
    }

    $bestScore = -1.0
    $bestCandidates = [System.Collections.Generic.List[object]]::new()
    $seenCandidates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($candidate in $Lookup.CandidatesByPolicy[$policyKey]) {
        if (-not $candidate.OutputReference) { continue }
        $gpProperty = $candidate.PSObject.Properties['GpSettingName']
        $candidateKey = "$($candidate.SettingKey)|$($candidate.OutputReference)|$(if ($gpProperty) { $gpProperty.Value })"
        if (-not $seenCandidates.Add($candidateKey)) { continue }

        $score = Get-StringSimilarity -Left $settingKey -Right ([string]$candidate.SettingKey)
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestCandidates.Clear()
            [void]$bestCandidates.Add($candidate)
        } elseif ([Math]::Abs($score - $bestScore) -lt 0.000000001) {
            [void]$bestCandidates.Add($candidate)
        }
    }

    if ($bestScore -lt $SimilarityThreshold -or $bestCandidates.Count -eq 0) {
        return [pscustomobject]@{ CISReference = ''; Status = 'No match' }
    }
    $bestCandidates = @(Select-BaselineApplicationCandidates -Candidates @($bestCandidates.ToArray()) -CsvRow $CsvRow |
        Group-Object -Property SettingKey, OutputReference | ForEach-Object { $_.Group[0] })
    if ($bestCandidates.Count -gt 1) {
        return [pscustomobject]@{ CISReference = ''; Status = 'Ambiguous' }
    }

    $bestCandidate = $bestCandidates[0]
    return [pscustomobject]@{
        CISReference = [string]$bestCandidate.OutputReference
        Status = if ([string]$bestCandidate.OutputReference -ieq 'twoday') { 'Twoday' } else { 'Matched' }
    }
}

function Parse-MdFile {
    param([string]$Path)
    $content = Get-Content $Path -Raw
    $lines = $content -split "`n"
    $recommendations = [System.Collections.Generic.List[PSCustomObject]]::new()
    $currentSection = ''
    $currentSubSection = ''
    foreach ($line in $lines) {
        $line = $line.TrimEnd()
        if ($line -match '^## \d+') {
            $headingText = ($line -replace '^## ', '').Trim()
            if ($headingText -imatch '\bMicrosoft\b') { $currentSection = $headingText }
        }
        if ($line -match '^### ') { $currentSubSection = ($line -replace '^### ', '').Trim() }
        if ($line -match '^\|\s*Ensure\s' -or $line -match '^\|\s*[A-Z].*\|.*\|.*\d+\.\d+' -or ($line -match '^\|' -and $line -imatch '\btwoday\b')) {
            $cells = [System.Collections.Generic.List[string]]::new()
            $cells.AddRange([string[]](Split-PipeRow $line))
            if ($cells.Count -ge 10) {
                # Skip header rows (bold markdown or literal column name)
                if ($cells[0] -match '^\*\*' -or $cells[0] -imatch '^Recommendation$') { continue }
                $recText      = $cells[0]
                $twodayId     = $cells[1]
                $cisRef       = $cells[2]
                # Use TwodayId as fallback identifier when CIS Reference is just a placeholder
                if ((-not $cisRef -or $cisRef -eq '-') -and $twodayId -and $twodayId -ne '-') {
                    $cisRef = $twodayId
                }
                $recProfile   = $cells[3]
                $origin       = if ($cells.Count -gt 6) { $cells[6] } else { '' }
                $followed     = $cells[7]
                $notes        = $cells[8]
                $intunePolicy = $cells[9]
                $settingName  = ''
                $settingValue = ''
                if ($recText -match "Ensure\s+'(.+)'\s+(?:is set to|Is|is)\s+'([^']+)'\s*$") {
                    $settingName  = $Matches[1]; $settingValue = $Matches[2]
                } elseif ($recText -match "Ensure\s+'(.+)'\s+(?:is set to|Is|is)\s+'([^']+)") {
                    $settingName  = $Matches[1]; $settingValue = $Matches[2]
                } elseif ($recText -match "Ensure\s+'(.+)'") {
                    $settingName = $Matches[1]
                }
                $recommendations.Add([PSCustomObject]@{
                    Recommendation = $recText
                    SettingName    = $settingName
                    SettingValue   = $settingValue
                    CISReference   = $cisRef
                    Profile        = $recProfile
                    Origin         = $origin
                    Followed       = $followed
                    Notes          = $notes
                    IntunePolicy   = $intunePolicy
                    Section        = $currentSection
                    SubSection     = $currentSubSection
                })
            }
        }
    }
    return ,$recommendations.ToArray()
}

function Test-MgSessionIsValid {
    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $ctx) { return $false }
        if (-not $ctx.TenantId) { return $false }

        $expiresOn = $null
        try { $expiresOn = $ctx.ExpiresOn } catch { $expiresOn = $null }
        if ($expiresOn -and (Get-Date) -ge $expiresOn) { return $false }

        return $true
    } catch {
        return $false
    }
}

function Get-ResolverTenantOverride {
    if ($script:ResolverTenantId -and $script:ResolverTenantId.Trim()) {
        return $script:ResolverTenantId.Trim()
    }
    return $null
}

function Test-MgSessionTenantMatch {
    param([string]$TenantId)

    if (-not $TenantId) { return $true }

    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $ctx -or -not $ctx.TenantId) { return $false }
        return ([string]$ctx.TenantId -ieq [string]$TenantId)
    } catch {
        return $false
    }
}

function Connect-BaselineReadGraph {
    param(
        [string]$TenantId,
        [switch]$ForceReconnect
    )

    $requiredCmdlets = @('Connect-MgGraph','Disconnect-MgGraph','Get-MgContext')
    $missingCmdlets = @($requiredCmdlets | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missingCmdlets.Count -gt 0) {
        throw "Missing required Microsoft Graph cmdlets: $($missingCmdlets -join ', '). Install/Update Microsoft.Graph modules and restart PowerShell."
    }

    $effectiveTenant = if ($TenantId -and $TenantId.Trim()) { $TenantId.Trim() } else { $null }
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if ($ForceReconnect -or ($ctx -and $ctx.TenantId)) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    $connectParams = @{
        Scopes      = @('DeviceManagementConfiguration.Read.All','Organization.Read.All')
        NoWelcome   = $true
        ErrorAction = 'Stop'
    }
    if ($effectiveTenant) {
        $connectParams['TenantId'] = $effectiveTenant
    }

    Connect-MgGraph @connectParams | Out-Null
}

function Resolve-SettingDefinitionsViaGraph {
    param(
        [string[]]$DefinitionIds,
        [hashtable]$DisplayNameMap,
        [hashtable]$OptionDisplayMap,
        [int]$BatchSize = 20,
        [scriptblock]$ProgressCallback,
        [scriptblock]$LogCallback,
        [scriptblock]$AuthChallengeCallback
    )

    if (-not $DisplayNameMap) { $DisplayNameMap = @{} }
    if (-not $OptionDisplayMap) { $OptionDisplayMap = @{} }

    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $idList = [System.Collections.Generic.List[string]]::new()
    foreach ($id in @($DefinitionIds)) {
        $idStr = [string]$id
        if (-not $idStr) { continue }
        if ($DisplayNameMap.ContainsKey($idStr)) { continue }
        if ($idStr -match '^#microsoft\.graph\.(?:iosGeneralDeviceConfiguration|androidGeneralDeviceConfiguration|androidWorkProfileGeneralDeviceConfiguration|androidDeviceOwnerGeneralDeviceConfiguration|aospDeviceOwnerGeneralDeviceConfiguration)\|(.+)$') {
            $DisplayNameMap[$idStr] = [string]$matches[1]
            continue
        }
        if ($seen.Add($idStr)) { [void]$idList.Add($idStr) }
    }

    if ($idList.Count -eq 0) {
        return [pscustomobject]@{ Requested = 0; Resolved = 0; Failed = 0; GraphCalls = 0 }
    }

    if ($BatchSize -lt 1) { $BatchSize = 1 }
    if ($BatchSize -gt 20) { $BatchSize = 20 }

    $resolved = 0
    $failed = 0
    $graphCalls = 0
    $authRetryAttempted = $false

    for ($offset = 0; $offset -lt $idList.Count; $offset += $BatchSize) {
        $chunkCount = [Math]::Min($BatchSize, $idList.Count - $offset)
        $chunk = $idList.GetRange($offset, $chunkCount)
        $reqIdToDefId = @{}
        $requests = [System.Collections.Generic.List[object]]::new()

        for ($i = 0; $i -lt $chunk.Count; $i++) {
            $reqId = [string]($i + 1)
            $defId = [string]$chunk[$i]
            $reqIdToDefId[$reqId] = $defId
            [void]$requests.Add([ordered]@{
                id     = $reqId
                method = 'GET'
                url    = "/deviceManagement/configurationSettings/$defId"
            })
        }

        if ($ProgressCallback) {
            & $ProgressCallback ($offset + $chunk.Count) $idList.Count
        }

        $chunkDone = $false
        while (-not $chunkDone) {
            try {
                $payload = @{ requests = @($requests) } | ConvertTo-Json -Depth 8 -Compress
                $batchResp = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Method POST -Body $payload -ContentType 'application/json' -ErrorAction Stop
                $graphCalls++

                $responses = Get-SafeObjectPropertyValue -Object $batchResp -Name 'responses'
                if (-not $responses) { $responses = @() }

                foreach ($respItem in @($responses)) {
                    $respId = [string](Get-SafeObjectPropertyValue -Object $respItem -Name 'id')
                    $status = [int](Get-SafeObjectPropertyValue -Object $respItem -Name 'status')
                    $defId = if ($reqIdToDefId.ContainsKey($respId)) { $reqIdToDefId[$respId] } else { $null }
                    if (-not $defId) { continue }

                    if ($status -ge 200 -and $status -lt 300) {
                        $body = Get-SafeObjectPropertyValue -Object $respItem -Name 'body'
                        $dname = Get-SafeObjectPropertyValue -Object $body -Name 'displayName'
                        if ($dname) {
                            $DisplayNameMap[$defId] = [string]$dname
                            $resolved++
                        }

                        $respOptions = Get-SafeObjectPropertyValue -Object $body -Name 'options'
                        if ($respOptions -and $respOptions -is [System.Collections.IEnumerable] -and -not ($respOptions -is [string])) {
                            foreach ($opt in $respOptions) {
                                $optId = Get-SafeObjectPropertyValue -Object $opt -Name 'itemId'
                                $optName = Get-SafeObjectPropertyValue -Object $opt -Name 'displayName'
                                if ($optId -and $optName -and -not $OptionDisplayMap.ContainsKey([string]$optId)) {
                                    $OptionDisplayMap[[string]$optId] = [string]$optName
                                }
                            }
                        }
                    } else {
                        $failed++
                        if ($LogCallback) {
                            & $LogCallback "  Could not resolve: $defId (HTTP $status)"
                        }
                    }
                }

                $chunkDone = $true
            } catch {
                $errorText = [string]$_.Exception.Message
                $isAuthChallenge = ($errorText -match 'interaction_required|login_required|consent_required|claims|MFA|AADSTS50076|AADSTS50079|AADSTS50158|AADSTS53000|Get-MgAccessToken|authentication failed')

                if ($isAuthChallenge -and -not $authRetryAttempted -and $AuthChallengeCallback) {
                    $authRetryAttempted = $true
                    if ($LogCallback) {
                        & $LogCallback '  Graph auth challenge detected during setting resolution. Re-authenticating once and retrying current batch...'
                    }
                    try {
                        & $AuthChallengeCallback $errorText
                        if ($LogCallback) {
                            & $LogCallback '  Re-authentication completed. Retrying current batch.'
                        }
                        continue
                    } catch {
                        $failed += $chunk.Count
                        if ($LogCallback) {
                            & $LogCallback "  Re-authentication failed: $($_.Exception.Message)"
                        }
                        $chunkDone = $true
                        continue
                    }
                }

                $failed += $chunk.Count
                if ($LogCallback) {
                    & $LogCallback "  Batch lookup failed for $($chunk.Count) setting definition(s): $errorText"
                }
                $chunkDone = $true
            }
        }
    }

    return [pscustomobject]@{
        Requested = $idList.Count
        Resolved  = $resolved
        Failed    = $failed
        GraphCalls = $graphCalls
    }
}

function Get-SettingDefinitionCachePath {
    param([string]$TenantKey)

    $baseDir = Join-Path $env:LOCALAPPDATA 'BaselineToolbox'
    $safeTenant = if ($TenantKey -and $TenantKey.Trim()) { $TenantKey.Trim() } else { 'default' }
    $safeTenant = ($safeTenant -replace '[^A-Za-z0-9._-]', '-')
    return (Join-Path $baseDir "setting-definition-cache-$safeTenant.json")
}

function Get-GraphTenantCacheKey {
    $tenantOverride = Get-ResolverTenantOverride
    if ($tenantOverride) {
        return $tenantOverride
    }

    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if ($ctx -and $ctx.TenantId) {
            return [string]$ctx.TenantId
        }
    } catch {}

    return 'default'
}

function Import-SettingDefinitionCache {
    param([string]$TenantKey)

    $path = Get-SettingDefinitionCachePath -TenantKey $TenantKey
    $displayNames = @{}
    $optionDisplayNames = @{}

    if (-not (Test-Path -Path $path -PathType Leaf)) {
        return [pscustomobject]@{
            TenantKey = $TenantKey
            Path = $path
            DisplayNames = $displayNames
            OptionDisplayNames = $optionDisplayNames
            DisplayNameCount = 0
            OptionCount = 0
            Loaded = $false
            UpdatedAt = $null
            AgeDays = $null
            SourceMode = $null
            Error = $null
        }
    }

    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if (-not $raw) { throw 'Cache file is empty.' }
        $cache = $raw | ConvertFrom-Json -ErrorAction Stop

        $dnNode = $cache.DisplayNames
        if ($dnNode -is [System.Collections.IDictionary]) {
            foreach ($k in $dnNode.Keys) {
                if ($k -and $dnNode[$k]) { $displayNames[[string]$k] = [string]$dnNode[$k] }
            }
        } elseif ($dnNode -and $dnNode.PSObject.Properties) {
            foreach ($p in $dnNode.PSObject.Properties) {
                if ($p.Name -and $p.Value) { $displayNames[[string]$p.Name] = [string]$p.Value }
            }
        }

        $optNode = $cache.OptionDisplayNames
        if ($optNode -is [System.Collections.IDictionary]) {
            foreach ($k in $optNode.Keys) {
                if ($k -and $optNode[$k]) { $optionDisplayNames[[string]$k] = [string]$optNode[$k] }
            }
        } elseif ($optNode -and $optNode.PSObject.Properties) {
            foreach ($p in $optNode.PSObject.Properties) {
                if ($p.Name -and $p.Value) { $optionDisplayNames[[string]$p.Name] = [string]$p.Value }
            }
        }
    } catch {
        return [pscustomobject]@{
            TenantKey = $TenantKey
            Path = $path
            DisplayNames = @{}
            OptionDisplayNames = @{}
            DisplayNameCount = 0
            OptionCount = 0
            Loaded = $false
            UpdatedAt = $null
            AgeDays = $null
            SourceMode = $null
            Error = $_.Exception.Message
        }
    }

    $updatedAt = $null
    $ageDays = $null
    $sourceMode = $null
    try {
        $updatedRaw = Get-SafeObjectPropertyValue -Object $cache -Name 'UpdatedAt'
        if ($updatedRaw) {
            $updatedAt = [datetimeoffset]::Parse([string]$updatedRaw)
            $ageDays = [math]::Floor(([datetimeoffset]::UtcNow - $updatedAt.ToUniversalTime()).TotalDays)
        }
    } catch {}
    try {
        $sourceRaw = Get-SafeObjectPropertyValue -Object $cache -Name 'SourceMode'
        if ($sourceRaw) { $sourceMode = [string]$sourceRaw }
    } catch {}

    return [pscustomobject]@{
        TenantKey = $TenantKey
        Path = $path
        DisplayNames = $displayNames
        OptionDisplayNames = $optionDisplayNames
        DisplayNameCount = $displayNames.Count
        OptionCount = $optionDisplayNames.Count
        Loaded = $true
        UpdatedAt = $updatedAt
        AgeDays = $ageDays
        SourceMode = $sourceMode
        Error = $null
    }
}

function Export-SettingDefinitionCache {
    param(
        [string]$TenantKey,
        [hashtable]$DisplayNameMap,
        [hashtable]$OptionDisplayMap,
        [string]$SourceMode = 'runtime'
    )

    $path = Get-SettingDefinitionCachePath -TenantKey $TenantKey
    $dir = Split-Path -Path $path -Parent
    if (-not (Test-Path -Path $dir -PathType Container)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }

    $payload = [ordered]@{
        SchemaVersion = 1
        TenantKey = $TenantKey
        UpdatedAt = (Get-Date).ToString('o')
        SourceMode = $SourceMode
        DisplayNames = [ordered]@{}
        OptionDisplayNames = [ordered]@{}
    }

    if ($DisplayNameMap) {
        foreach ($k in ($DisplayNameMap.Keys | Sort-Object)) {
            if ($k -and $DisplayNameMap[$k]) {
                $payload.DisplayNames[[string]$k] = [string]$DisplayNameMap[$k]
            }
        }
    }
    if ($OptionDisplayMap) {
        foreach ($k in ($OptionDisplayMap.Keys | Sort-Object)) {
            if ($k -and $OptionDisplayMap[$k]) {
                $payload.OptionDisplayNames[[string]$k] = [string]$OptionDisplayMap[$k]
            }
        }
    }

    $json = $payload | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($path, $json, [System.Text.UTF8Encoding]::new($false))

    return [pscustomobject]@{
        TenantKey = $TenantKey
        Path = $path
        DisplayNameCount = $payload.DisplayNames.Count
        OptionCount = $payload.OptionDisplayNames.Count
        UpdatedAt = $payload.UpdatedAt
        SourceMode = $payload.SourceMode
    }
}

function Get-CacheModeLabel {
    param([string]$Mode)

    $modeValue = if ($null -ne $Mode) { [string]$Mode } else { '' }
    switch ($modeValue.ToLowerInvariant()) {
        'cacheonly' { return 'Strict cache-only' }
        'forcefresh' { return 'Force fresh (ignore cache)' }
        default { return 'Cache + Graph fallback' }
    }
}

function Write-CacheStalenessWarning {
    param(
        $CacheData,
        [scriptblock]$LogCallback,
        [int]$WarnDays = 30
    )

    if (-not $CacheData -or -not $CacheData.Loaded) { return }
    if ($null -eq $CacheData.AgeDays) { return }
    if ([int]$CacheData.AgeDays -lt $WarnDays) { return }

    $updatedText = if ($CacheData.UpdatedAt) { [string]$CacheData.UpdatedAt } else { 'unknown' }
    $warn = "Warning: resolver cache is $([int]$CacheData.AgeDays) days old (updated: $updatedText)."
    if ($LogCallback) { & $LogCallback $warn }
}

function Get-SettingDefinitionIdsFromSettingInstance {
    param(
        $Instance,
        [System.Collections.Generic.HashSet[string]]$Collector
    )

    if (-not $Instance -or -not $Collector) { return }

    $defId = Get-SafeObjectPropertyValue -Object $Instance -Name 'settingDefinitionId'
    if ($defId) { [void]$Collector.Add([string]$defId) }

    $groupCollection = Get-SafeObjectPropertyValue -Object $Instance -Name 'groupSettingCollectionValue'
    if ($groupCollection -and $groupCollection -is [System.Collections.IEnumerable] -and -not ($groupCollection -is [string])) {
        foreach ($group in $groupCollection) {
            $children = Get-SafeObjectPropertyValue -Object $group -Name 'children'
            if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
                foreach ($child in $children) {
                    Get-SettingDefinitionIdsFromSettingInstance -Instance $child -Collector $Collector
                }
            }
        }
    }

    $groupSingle = Get-SafeObjectPropertyValue -Object $Instance -Name 'groupSettingValue'
    if ($groupSingle) {
        $children = Get-SafeObjectPropertyValue -Object $groupSingle -Name 'children'
        if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
            foreach ($child in $children) {
                Get-SettingDefinitionIdsFromSettingInstance -Instance $child -Collector $Collector
            }
        }
    }

    $choiceValue = Get-SafeObjectPropertyValue -Object $Instance -Name 'choiceSettingValue'
    if ($choiceValue) {
        $children = Get-SafeObjectPropertyValue -Object $choiceValue -Name 'children'
        if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
            foreach ($child in $children) {
                Get-SettingDefinitionIdsFromSettingInstance -Instance $child -Collector $Collector
            }
        }
    }

    $simpleValue = Get-SafeObjectPropertyValue -Object $Instance -Name 'simpleSettingValue'
    if ($simpleValue) {
        $children = Get-SafeObjectPropertyValue -Object $simpleValue -Name 'children'
        if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
            foreach ($child in $children) {
                Get-SettingDefinitionIdsFromSettingInstance -Instance $child -Collector $Collector
            }
        }
    }
}

function Invoke-TenantPolicyCacheWarmup {
    param(
        [object[]]$Profiles,
        [string]$TenantKey,
        [string]$CacheMode = 'fallback',
        [scriptblock]$ProgressCallback,
        [scriptblock]$LogCallback,
        [scriptblock]$AuthChallengeCallback
    )

    if (-not $Profiles -or @($Profiles).Count -eq 0) {
        throw 'No profiles available for cache warm-up.'
    }

    $mode = if ($CacheMode) { $CacheMode.ToLowerInvariant() } else { 'fallback' }
    if ($mode -notin @('fallback','cacheonly','forcefresh')) { $mode = 'fallback' }

    $cacheData = Import-SettingDefinitionCache -TenantKey $TenantKey
    $cacheError = Get-SafeObjectPropertyValue -Object $cacheData -Name 'Error'
    if ($cacheData.Loaded -and $LogCallback) {
        & $LogCallback "Loaded resolver cache for tenant '$TenantKey': $($cacheData.DisplayNameCount) names, $($cacheData.OptionCount) options."
    } elseif ($cacheError -and $LogCallback) {
        & $LogCallback "Resolver cache not used: $cacheError"
    }
    Write-CacheStalenessWarning -CacheData $cacheData -LogCallback $LogCallback

    $allDisplayNames = if ($cacheData.DisplayNames) { $cacheData.DisplayNames } else { @{} }
    $allOptionDisplayNames = if ($cacheData.OptionDisplayNames) { $cacheData.OptionDisplayNames } else { @{} }
    if ($mode -eq 'forcefresh') {
        $allDisplayNames = @{}
        $allOptionDisplayNames = @{}
        if ($LogCallback) { & $LogCallback 'Force fresh mode selected: starting warm-up with empty in-memory cache.' }
    }

    $referencedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $total = @($Profiles).Count
    $index = 0
    $settingsCatalogCount = 0

    foreach ($profItem in @($Profiles)) {
        $index++
        if ($ProgressCallback) { & $ProgressCallback $index $total }

        $policyKind = Get-SafeObjectPropertyValue -Object $profItem -Name 'PolicyKind'
        if (-not $policyKind) { $policyKind = 'SettingsCatalog' }
        if ($policyKind -in @('DeviceRestriction', 'AdministrativeTemplate')) { continue }

        $settingsCatalogCount++
        $profileId = Get-SafeObjectPropertyValue -Object $profItem -Name 'Id'
        $profileName = Get-SafeObjectPropertyValue -Object $profItem -Name 'Name'
        if (-not $profileId) { continue }

        if ($LogCallback) { & $LogCallback "[$index/$total] Warm cache scan: $profileName" }
        $settingsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$profileId/settings?`$expand=settingDefinitions"
        $settings = @(Invoke-GraphPagedGet -Uri $settingsUri)

        $flatSettings = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $settings) {
            if ($item -is [System.Array]) {
                foreach ($sub in $item) { [void]$flatSettings.Add($sub) }
            } else {
                [void]$flatSettings.Add($item)
            }
        }

        foreach ($setting in $flatSettings) {
            $defs = Get-SafeObjectPropertyValue -Object $setting -Name 'settingDefinitions'
            if ($defs -and $defs -is [System.Collections.IEnumerable] -and -not ($defs -is [string])) {
                foreach ($def in $defs) {
                    $did = Get-SafeObjectPropertyValue -Object $def -Name 'id'
                    $dname = Get-SafeObjectPropertyValue -Object $def -Name 'displayName'
                    if ($did -and $dname -and -not $allDisplayNames.ContainsKey([string]$did)) {
                        $allDisplayNames[[string]$did] = [string]$dname
                    }
                    $defOptions = Get-SafeObjectPropertyValue -Object $def -Name 'options'
                    if ($defOptions -and $defOptions -is [System.Collections.IEnumerable] -and -not ($defOptions -is [string])) {
                        foreach ($opt in $defOptions) {
                            $optId = Get-SafeObjectPropertyValue -Object $opt -Name 'itemId'
                            $optName = Get-SafeObjectPropertyValue -Object $opt -Name 'displayName'
                            if ($optId -and $optName -and -not $allOptionDisplayNames.ContainsKey([string]$optId)) {
                                $allOptionDisplayNames[[string]$optId] = [string]$optName
                            }
                        }
                    }
                }
            }

            $instance = Get-SafeObjectPropertyValue -Object $setting -Name 'settingInstance'
            if ($instance) {
                Get-SettingDefinitionIdsFromSettingInstance -Instance $instance -Collector $referencedIds
            }
        }
    }

    $missingIds = [System.Collections.Generic.List[string]]::new()
    foreach ($id in $referencedIds) {
        if (-not $allDisplayNames.ContainsKey($id)) {
            [void]$missingIds.Add([string]$id)
        }
    }

    $resolveStats = [pscustomobject]@{ Requested = 0; Resolved = 0; Failed = 0; GraphCalls = 0 }
    if ($missingIds.Count -gt 0) {
        if ($mode -eq 'cacheonly') {
            if ($LogCallback) { & $LogCallback "Strict cache-only mode: skipped Graph lookup for $($missingIds.Count) setting definition(s)." }
        } else {
            if ($LogCallback) { & $LogCallback "Resolving $($missingIds.Count) setting definition(s) via Graph..." }
            $resolveStats = Resolve-SettingDefinitionsViaGraph `
                -DefinitionIds @($missingIds) `
                -DisplayNameMap $allDisplayNames `
                -OptionDisplayMap $allOptionDisplayNames `
                -LogCallback $LogCallback `
                -AuthChallengeCallback $AuthChallengeCallback
        }
    }

    $savedCache = Export-SettingDefinitionCache `
        -TenantKey $TenantKey `
        -DisplayNameMap $allDisplayNames `
        -OptionDisplayMap $allOptionDisplayNames `
        -SourceMode "warmup-$mode"

    return [pscustomobject]@{
        TenantKey = $TenantKey
        ProfileCount = $total
        SettingsCatalogProfileCount = $settingsCatalogCount
        ReferencedDefinitionCount = $referencedIds.Count
        MissingDefinitionCount = $missingIds.Count
        ResolveRequested = $resolveStats.Requested
        ResolveResolved = $resolveStats.Resolved
        ResolveFailed = $resolveStats.Failed
        GraphCalls = $resolveStats.GraphCalls
        SavedDisplayNameCount = $savedCache.DisplayNameCount
        SavedOptionCount = $savedCache.OptionCount
        UpdatedAt = $savedCache.UpdatedAt
    }
}

function Get-SafeObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if (-not $Object) { return $null }

    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($k in $Object.Keys) {
            if ([string]$k -ieq $Name) { return $Object[$k] }
        }
        return $null
    }

    foreach ($prop in $Object.PSObject.Properties) {
        if ($prop.Name -ieq $Name) { return $prop.Value }
    }

    $additional = $Object.PSObject.Properties['AdditionalProperties']
    if ($additional -and $additional.Value -is [System.Collections.IDictionary]) {
        foreach ($k in $additional.Value.Keys) {
            if ([string]$k -ieq $Name) { return $additional.Value[$k] }
        }
    }

    return $null
}

function Invoke-GraphRequestWithTimeout {
    param(
        [string]$Uri,
        [int]$TimeoutSeconds = 20,
        [int]$RetryCount = 1
    )

    $attempt = 0
    while ($attempt -le $RetryCount) {
        try {
            # Use Graph SDK request path for compatibility across module versions.
            $response = Invoke-MgGraphRequest -Uri $Uri -Method Get -ErrorAction Stop
            return $response
        } catch {
            $attempt++
            if ($attempt -le $RetryCount) {
                Start-Sleep -Seconds 2
                continue
            }

            if ($_.Exception.Message -match 'Get-MgAccessToken') {
                throw "Microsoft Graph SDK authentication failed while calling: $Uri"
            }
            if ($_.Exception.Message -match 'timed out|timeout') {
                throw "Graph request timed out after $TimeoutSeconds seconds for: $Uri"
            }
            throw
        }
    }
}

function Invoke-GraphPagedGet {
    param(
        [string]$Uri,
        [int]$TimeoutSeconds = 20,
        [int]$RetryCount = 1
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $next = $Uri

    do {
        $response = Invoke-GraphRequestWithTimeout -Uri $next -TimeoutSeconds $TimeoutSeconds -RetryCount $RetryCount

        $value = Get-SafeObjectPropertyValue -Object $response -Name 'value'
        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            foreach ($item in $value) { [void]$results.Add($item) }
        } elseif ($response -is [System.Collections.IEnumerable] -and -not ($response -is [string]) -and -not ($response -is [System.Collections.IDictionary])) {
            foreach ($item in $response) { [void]$results.Add($item) }
        } elseif ($response) {
            [void]$results.Add($response)
        }

        $next = Get-SafeObjectPropertyValue -Object $response -Name '@odata.nextLink'
        if (-not $next) {
            $additional = Get-SafeObjectPropertyValue -Object $response -Name 'AdditionalProperties'
            if ($additional -is [System.Collections.IDictionary] -and $additional.Contains('@odata.nextLink')) {
                $next = $additional['@odata.nextLink']
            }
        }
    } while ($next)

    return ,$results.ToArray()
}

function Get-ProfileReportSettingName {
    param([object]$Setting)

    $name = $null
    if ($Setting.settingInstance -and $Setting.settingInstance.settingDefinitionId) {
        $name = [string]$Setting.settingInstance.settingDefinitionId
    } elseif ($Setting.settingDefinitionId) {
        $name = [string]$Setting.settingDefinitionId
    } elseif ($Setting.name) {
        $name = [string]$Setting.name
    } elseif ($Setting.id) {
        $name = [string]$Setting.id
    } else {
        return 'UnknownSetting'
    }

    if ($name -match '/') {
        $segments = $name -split '/'
        if ($segments.Count -gt 0 -and $segments[-1]) { return $segments[-1] }
    }
    return $name
}

function Get-ProfileReportSettingValue {
    param([object]$Setting)
    try {
        if ($Setting.settingInstance) {
            $instance = $Setting.settingInstance
            if ($instance.choiceSettingValue -and $instance.choiceSettingValue.value) {
                return [string]$instance.choiceSettingValue.value
            }
            if ($instance.simpleSettingValue -and $null -ne $instance.simpleSettingValue.value) {
                return [string]$instance.simpleSettingValue.value
            }
            if ($instance.choiceSettingCollectionValue -and $instance.choiceSettingCollectionValue.value) {
                return ($instance.choiceSettingCollectionValue.value -join '; ')
            }
            if ($instance.simpleSettingCollectionValue -and $instance.simpleSettingCollectionValue.value) {
                return ($instance.simpleSettingCollectionValue.value -join '; ')
            }
        }

        $json = $Setting | ConvertTo-Json -Depth 30 -Compress
        if ($json.Length -gt 280) { return ($json.Substring(0, 280) + '...') }
        return $json
    } catch {
        return 'Unable to parse value'
    }
}

function ConvertFrom-JsonXmlNode {
    <#
    .SYNOPSIS
    Converts an XmlNode (produced by JsonReaderWriterFactory) to a PSObject/hashtable tree.
    #>
    param([System.Xml.XmlNode]$Node)
    $type = $Node.GetAttribute('type')
    switch ($type) {
        'object' {
            $obj = [ordered]@{}
            foreach ($child in $Node.ChildNodes) {
                $propName = $child.LocalName
                # JsonReaderWriterFactory uses 'item' with a 'name' attribute for invalid XML names
                $nameAttr = $child.GetAttribute('item')
                if (-not $nameAttr) { $nameAttr = $child.GetAttribute('name') }
                if ($nameAttr) { $propName = $nameAttr }
                $obj[$propName] = ConvertFrom-JsonXmlNode $child
            }
            return [pscustomobject]$obj
        }
        'array' {
            $arr = [System.Collections.Generic.List[object]]::new()
            foreach ($child in $Node.ChildNodes) {
                [void]$arr.Add((ConvertFrom-JsonXmlNode $child))
            }
            return ,$arr.ToArray()
        }
        'number' {
            $txt = $Node.InnerText
            if ($txt -match '^\-?\d+$' -and [long]::TryParse($txt, [ref]$null)) { return [long]$txt }
            return [double]$txt
        }
        'boolean' { return ($Node.InnerText -eq 'true') }
        'null'    { return $null }
        default   { return $Node.InnerText }  # 'string' or untyped
    }
}

function ConvertFrom-JsonRobust {
    <#
    .SYNOPSIS
    Parses JSON text using ConvertFrom-Json with a fallback to .NET JsonReaderWriterFactory
    for files that contain characters (like ~) which trip up PS 5.1's JavaScriptSerializer.
    OData annotation lines are stripped first since they confuse both parsers and aren't needed.
    Also fixes malformed JSON from backup tools (e.g. bare unquoted duplicate values).
    #>
    param([string]$JsonText)

    # Strip @odata annotation lines — they contain @ . # chars that confuse both
    # JavaScriptSerializer (PS 5.1) and JsonReaderWriterFactory's XML mapper.
    $lines = $JsonText -split '\r?\n'
    $cleanedLines = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $lines) {
        if ($line -match '^\s*"[^"]*@odata') { continue }
        [void]$cleanedLines.Add($line)
    }
    $cleanedJson = $cleanedLines -join "`n"
    # Fix trailing commas that result from removed lines
    $cleanedJson = [regex]::Replace($cleanedJson, ',(\s*[}\]])', '$1')
    # Remove any BOM character
    $cleanedJson = $cleanedJson.TrimStart([char]0xFEFF)
    # Fix malformed JSON: some backup tools duplicate settingDefinitionId values as bare
    # unquoted text after the comma (e.g. "key": "value",bareUnquotedDuplicate)
    $cleanedJson = [regex]::Replace($cleanedJson, '(?<=",)\s*[a-zA-Z_][^\r\n"{}[\]]*$', '', [System.Text.RegularExpressions.RegexOptions]::Multiline)

    # Try standard ConvertFrom-Json first (fast path for JSON without ADMX ~ chars)
    try {
        return ($cleanedJson | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        # Fallback: use .NET's DataContract JSON reader which handles ~ correctly
        try {
            Add-Type -AssemblyName System.Runtime.Serialization -ErrorAction SilentlyContinue
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($cleanedJson)
            $quotas = [System.Xml.XmlDictionaryReaderQuotas]::Max
            $reader = [System.Runtime.Serialization.Json.JsonReaderWriterFactory]::CreateJsonReader($bytes, $quotas)
            $xmlDoc = [System.Xml.XmlDocument]::new()
            $xmlDoc.Load($reader)
            $reader.Close()
            return (ConvertFrom-JsonXmlNode $xmlDoc.DocumentElement)
        } catch {
            throw
        }
    }
}

function Get-DeviceConfigurationBaselineAreaInfo {
    param([string]$PolicyName)

    if ($PolicyName -notmatch '-SecurityBaseline-(.+?)(?:\.json)?$') {
        return @{ BaselineArea = $null; PropertyPattern = $null; HasAreaSuffix = $false; IsSupported = $true }
    }

    $area = [string]$matches[1]
    $areaPatterns = @{
        Applications = '^(appsAutoUpdatePolicy|appsAllowInstallFromUnknownSources|appsDefaultPermissionPolicy|appsRecommendSkippingFirstUseHints|personalProfileAppsAllowInstallFromUnknownSources|personalProfilePersonalApplications)$'
        CustomSupportInformation = '^deviceOwnerLockScreenMessage$'
        DevicePassword = '^(passwordMinimumLength|passwordBlockKeyguardFeatures|passwordRequiredType|passwordRequireUnlock|passwordExpirationDays|passwordMinutesOfInactivityBeforeScreenTimeout|passwordPreviousPasswordCountToBlock|passwordSignInFailureCountBeforeFactoryReset|passwordMinimumLetterCharacters|passwordMinimumLowerCaseCharacters|passwordMinimumNonLetterCharacters|passwordMinimumNumericCharacters|passwordMinimumSymbolCharacters|passwordMinimumUpperCaseCharacters|workProfilePassword.*)$'
        General = '^(storageBlockExternalMedia|storageBlockUsbFileTransfer|locateDeviceLostModeEnabled|systemUpdateInstallType|activationLockAllowWhenSupervised|configurationProfileBlockChanges|deviceBlockEraseContentAndSettings|deviceBlockNameModification|diagnosticDataBlockSubmission|enterpriseAppBlockTrust|privacyForceLimitAdTracking|vpnBlockCreation|dateAndTimeForceSetAutomatically|applePersonalizedAdsBlocked)$'
        PowerSettings = '^passwordMinutesOfInactivityBeforeScreenTimeout$'
        SystemSecurity = '^securityRequireVerifyApps$'
        AppStoreDocViewingGaming = '^(airDropForceUnmanagedDropTarget|documentsBlockManagedDocumentsInUnmanagedApps|managedPasteboardRequired|filesNetworkDriveAccessBlocked)$'
        'Built-inApps' = '^(safariRequireFraudWarning|siriBlockedWhenLocked)$'
        CloudandStorage = '^(iCloudBlockActivityContinuation|iCloudBlockManagedAppsSync|iCloudRequireEncryptedBackup|enterpriseBookBlockBackup|enterpriseBookBlockMetadataSync|iCloudBlockDocumentSync|keychainBlockCloudSync)$'
        ConnectedDevices = '^(airPlayForcePairingPasswordForOutgoingRequests|appleWatchForceWristDetection|autoUnlockBlocked|airPrintBlockiBeaconDiscovery|filesUsbDriveAccessBlocked)$'
        LockedScreenExperience = '^(lockScreenBlockControlCenter|lockScreenBlockPassbook|lockScreenBlockTodayView)$'
        LockScreenMessage = '^lockScreenFootnote$'
        Password = '^(autoFillForceAuthentication|passcodeBlockSimple|passcodeRequired|passwordBlockProximityRequests|passwordBlockAirDropSharing|passcodeMinimumLength|passcodeMinutesOfInactivityBeforeLock|passcodeMinutesOfInactivityBeforeScreenTimeout)$'
        Wireless = '^voiceDialingBlocked$'
    }

    if (-not $areaPatterns.ContainsKey($area)) {
        return @{ BaselineArea = $area; PropertyPattern = $null; HasAreaSuffix = $true; IsSupported = $false }
    }

    return @{ BaselineArea = $area; PropertyPattern = [string]$areaPatterns[$area]; HasAreaSuffix = $true; IsSupported = $true }
}

function Test-DeviceConfigurationSettingValue {
    param([object]$Value)

    if ($null -eq $Value) { return $false }
    if ($Value -is [bool] -and -not $Value) { return $false }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
        if ($Value -in @('notConfigured', 'deviceDefault')) { return $false }
    }
    if ($Value -is [System.Collections.ICollection] -and $Value.Count -eq 0) { return $false }
    return $true
}

function Get-AdminTemplateIdFromBind {
    param([string]$BindUrl, [string]$Segment)

    if (-not $BindUrl) { return $null }
    $m = [regex]::Match($BindUrl, "$Segment\('([^']+)'\)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($m.Success) { return $m.Groups[1].Value }
    $m = [regex]::Match($BindUrl, "$Segment/([^/\(\)\?]+)", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function ConvertTo-AdminTemplatePresentationValueString {
    param([object]$PresentationValue)

    $values = Get-SafeObjectPropertyValue -Object $PresentationValue -Name 'values'
    if ($null -ne $values -and $values -is [System.Collections.IEnumerable] -and $values -isnot [string]) {
        $items = @(foreach ($v in $values) {
            if ($null -eq $v) { continue }
            if ($v -is [string] -or $v -is [ValueType]) { [string]$v; continue }
            $n = Get-SafeObjectPropertyValue -Object $v -Name 'name'
            $val = Get-SafeObjectPropertyValue -Object $v -Name 'value'
            if ($null -ne $n -and $null -ne $val -and [string]$n -ne [string]$val) { "$n=$val" }
            elseif ($null -ne $n) { [string]$n }
            elseif ($null -ne $val) { [string]$val }
        })
        return (@($items | Sort-Object -CaseSensitive) -join '; ')
    }
    $value = Get-SafeObjectPropertyValue -Object $PresentationValue -Name 'value'
    if ($null -eq $value) { return '-' }
    if ($value -is [bool]) { return $value.ToString().ToLower() }
    return [string]$value
}

function ConvertTo-AdminTemplateSettingRows {
    <#
    .SYNOPSIS
    Converts Administrative Template (groupPolicyConfiguration) definitionValues into flat setting rows.
    Works for live Graph responses (expanded definition/presentation) and JSON backups (@odata.bind references).
    Keys: AdminTemplate|<definitionId> for the enabled state, AdminTemplate|<definitionId>|<presentationId> for each presentation value.
    #>
    param([object[]]$DefinitionValues)

    $rows = [System.Collections.Generic.List[object]]::new()
    $order = 0
    foreach ($dv in @($DefinitionValues)) {
        if ($null -eq $dv) { continue }
        $definition = Get-SafeObjectPropertyValue -Object $dv -Name 'definition'
        $defId = if ($definition) { [string](Get-SafeObjectPropertyValue -Object $definition -Name 'id') } else { $null }
        if (-not $defId) {
            $bind = Get-SafeObjectPropertyValue -Object $dv -Name 'definition@odata.bind'
            if (-not $bind) { $bind = Get-SafeObjectPropertyValue -Object $dv -Name 'definitionOdataBind' }
            $defId = Get-AdminTemplateIdFromBind -BindUrl ([string]$bind) -Segment 'groupPolicyDefinitions'
        }
        if (-not $defId) { continue }

        $displayName = if ($definition) { [string](Get-SafeObjectPropertyValue -Object $definition -Name 'displayName') } else { '' }
        $categoryPath = if ($definition) { [string](Get-SafeObjectPropertyValue -Object $definition -Name 'categoryPath') } else { '' }
        $classType = if ($definition) { [string](Get-SafeObjectPropertyValue -Object $definition -Name 'classType') } else { '' }
        $baseName = if ($displayName) { $displayName } else { $defId }
        if ($categoryPath) { $baseName = "$($categoryPath.Trim('\')) > $baseName" }
        if ($classType) { $baseName = "[$classType] $baseName" }

        $enabled = Get-SafeObjectPropertyValue -Object $dv -Name 'enabled'
        $enabledStr = if ($null -eq $enabled) { '-' } elseif ([bool]$enabled) { 'Enabled' } else { 'Disabled' }
        $key = "AdminTemplate|$defId"
        [void]$rows.Add([pscustomobject]@{
            DefinitionId = $key
            SettingName  = $baseName
            RawValue     = $enabledStr
            SettingValue = $enabledStr
            Depth        = 0
            SortGroup    = $key
            Order        = $order
        })
        $order++

        $presValues = Get-SafeObjectPropertyValue -Object $dv -Name 'presentationValues'
        if (-not $presValues -or $presValues -is [string]) { continue }
        foreach ($pv in @($presValues)) {
            if ($null -eq $pv) { continue }
            $presentation = Get-SafeObjectPropertyValue -Object $pv -Name 'presentation'
            $presId = if ($presentation) { [string](Get-SafeObjectPropertyValue -Object $presentation -Name 'id') } else { $null }
            if (-not $presId) {
                $bind = Get-SafeObjectPropertyValue -Object $pv -Name 'presentation@odata.bind'
                if (-not $bind) { $bind = Get-SafeObjectPropertyValue -Object $pv -Name 'presentationOdataBind' }
                $presId = Get-AdminTemplateIdFromBind -BindUrl ([string]$bind) -Segment 'presentations'
            }
            if (-not $presId) { $presId = [string](Get-SafeObjectPropertyValue -Object $pv -Name 'id') }
            if (-not $presId) { continue }
            $label = if ($presentation) { [string](Get-SafeObjectPropertyValue -Object $presentation -Name 'label') } else { '' }
            if (-not $label) { $label = $presId }
            $valStr = ConvertTo-AdminTemplatePresentationValueString -PresentationValue $pv
            [void]$rows.Add([pscustomobject]@{
                DefinitionId = "$key|$presId"
                SettingName  = $label
                RawValue     = $valStr
                SettingValue = $valStr
                Depth        = 1
                SortGroup    = $key
                Order        = $order
            })
            $order++
        }
    }
    return $rows.ToArray()
}

function Get-BaselineSettingsFromJson {
    <#
    .SYNOPSIS
    Parses a JSON Intune policy backup file and extracts a flat list of settingDefinitionId -> raw value pairs.
    Uses Get-SafeObjectPropertyValue to avoid PropertyNotFoundException under Set-StrictMode.
    #>
    param([string]$JsonPath, [switch]$PolicyReport)

    $jsonText = Get-Content -Path $JsonPath -Raw -Encoding UTF8
    $rootOdataTypeMatch = [regex]::Match($jsonText, '(?m)^\s*"@odata\.type"\s*:\s*"([^"]+)"')
    $rootOdataType = if ($rootOdataTypeMatch.Success) { $rootOdataTypeMatch.Groups[1].Value } else { $null }
    # Administrative Template backups reference definitions/presentations via @odata.bind, which
    # ConvertFrom-JsonRobust strips; keep them under neutral names so the IDs survive parsing.
    $jsonText = [regex]::Replace($jsonText, '"(definition|presentation)@odata\.bind"\s*:', '"$1OdataBind":')
    $json = ConvertFrom-JsonRobust -JsonText $jsonText

    $policyName = Get-SafeObjectPropertyValue -Object $json -Name 'name'
    if (-not $policyName) { $policyName = Get-SafeObjectPropertyValue -Object $json -Name 'displayName' }
    $policyId = $null
    $jsonId = Get-SafeObjectPropertyValue -Object $json -Name 'id'
    if ($jsonId) {
        $policyId = [string]$jsonId
    }
    $settings = Get-SafeObjectPropertyValue -Object $json -Name 'settings'
    $result = [System.Collections.Generic.List[object]]::new()

    # Detect Device Restriction format: no 'settings' array but has a known @odata.type
    $jsonOdataType = Get-SafeObjectPropertyValue -Object $json -Name '@odata.type'
    if (-not $jsonOdataType) { $jsonOdataType = $rootOdataType }
    $deviceRestrictionTypes = @(
        '#microsoft.graph.iosGeneralDeviceConfiguration',
        '#microsoft.graph.iosDeviceFeaturesConfiguration',
        '#microsoft.graph.androidGeneralDeviceConfiguration',
        '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration',
        '#microsoft.graph.androidDeviceOwnerGeneralDeviceConfiguration',
        '#microsoft.graph.aospDeviceOwnerGeneralDeviceConfiguration'
    )
    $hasSettingsProperty = if ($json -is [System.Collections.IDictionary]) { $json.Contains('settings') } else { $null -ne $json.PSObject.Properties['settings'] }
    $declaredPolicyType = Get-SafeObjectPropertyValue -Object $json -Name '@odata.type'
    $adminTemplateType = '#microsoft.graph.groupPolicyConfiguration'
    $adminTemplateValues = $null
    if ($json -is [System.Collections.IEnumerable] -and $json -isnot [string] -and $json -isnot [System.Collections.IDictionary]) {
        # Some backup tools store only the definitionValues array.
        $jsonItems = @($json)
        if ($jsonItems.Count -and @($jsonItems | Where-Object {
                (Get-SafeObjectPropertyValue -Object $_ -Name 'definition') -or (Get-SafeObjectPropertyValue -Object $_ -Name 'definitionOdataBind')
            }).Count -eq $jsonItems.Count) {
            $adminTemplateValues = $jsonItems
        }
    } elseif ($jsonOdataType -eq $adminTemplateType -or $null -ne (Get-SafeObjectPropertyValue -Object $json -Name 'definitionValues')) {
        $adminTemplateValues = @(Get-SafeObjectPropertyValue -Object $json -Name 'definitionValues' | Where-Object { $null -ne $_ })
    } elseif (-not $hasSettingsProperty -and ((Get-SafeObjectPropertyValue -Object $json -Name 'definition') -or (Get-SafeObjectPropertyValue -Object $json -Name 'definitionOdataBind'))) {
        # Single-item definitionValues array unrolled by the JSON parser.
        $adminTemplateValues = @($json)
    }
    $isAdminTemplate = $null -ne $adminTemplateValues
    if ($isAdminTemplate) { $jsonOdataType = $adminTemplateType }
    $policyType = if ($isAdminTemplate) {
        $adminTemplateType
    } elseif ($hasSettingsProperty -and (-not $declaredPolicyType -or $declaredPolicyType -eq '#microsoft.graph.deviceManagementConfigurationPolicy')) {
        'SettingsCatalog'
    } elseif ($jsonOdataType -in $deviceRestrictionTypes) { [string]$jsonOdataType } else { '' }
    if ($PolicyReport -and -not $policyType) { throw 'Unsupported policy backup format or missing policy type/settings.' }
    $policyPlatform = if ($isAdminTemplate) {
        'Windows'
    } elseif ($jsonOdataType -in @('#microsoft.graph.iosGeneralDeviceConfiguration', '#microsoft.graph.iosDeviceFeaturesConfiguration')) {
        'iOS'
    } elseif ($jsonOdataType -in $deviceRestrictionTypes) {
        'Android'
    } else {
        $null
    }
    $areaInfo = Get-DeviceConfigurationBaselineAreaInfo -PolicyName ([System.IO.Path]::GetFileName($JsonPath))

    if ($isAdminTemplate) {
        # Assessment baselines do not map Administrative Templates; only the policy report compares them.
        if ($PolicyReport) {
            foreach ($row in (ConvertTo-AdminTemplateSettingRows -DefinitionValues $adminTemplateValues)) {
                [void]$result.Add([pscustomobject]@{
                    DefinitionId = $row.DefinitionId
                    RawValue     = $row.RawValue
                    Depth        = $row.Depth
                })
            }
        }
    } elseif (-not $settings -and $jsonOdataType -and $jsonOdataType -in $deviceRestrictionTypes) {
        # Flatten the top-level policy properties as settings, namespaced with @odata.type
        # so comparison keys never collide across iOS/Android.
        $skipProps = [System.Collections.Generic.HashSet[string]]::new(
            [string[]]@('id','displayName','description','createdDateTime',
                        'lastModifiedDateTime','version','roleScopeTagIds',
                        'supportsScopeTags','settingCount'),
            [System.StringComparer]::OrdinalIgnoreCase
        )
        $jsonDict = $null
        if ($json -is [System.Collections.IDictionary]) {
            $jsonDict = $json
        } elseif ($json.PSObject.Properties) {
            $jsonDict = @{}
            foreach ($p in $json.PSObject.Properties) { $jsonDict[$p.Name] = $p.Value }
        }
        if ($jsonDict) {
            foreach ($propKey in @($jsonDict.Keys)) {
                if ($propKey -like '*@odata*' -or
                    $propKey -like '#microsoft.graph.*' -or
                    $propKey -ieq 'assignments' -or
                    $propKey -like 'deviceManagementApplicabilityRule*') { continue }
                if ($skipProps.Contains($propKey)) { continue }
                if (-not $PolicyReport -and -not $areaInfo.IsSupported) { continue }
                if (-not $PolicyReport -and $areaInfo.PropertyPattern -and $propKey -notmatch $areaInfo.PropertyPattern) { continue }
                $rawVal = $jsonDict[$propKey]
                if (-not $PolicyReport -and -not (Test-DeviceConfigurationSettingValue -Value $rawVal)) { continue }
                $valStr = if ($null -eq $rawVal) { '-' }
                          elseif ($rawVal -is [bool]) { $rawVal.ToString().ToLower() }
                          elseif ($rawVal -is [System.Collections.IEnumerable] -and $rawVal -isnot [string]) {
                              $items = @($rawVal | ForEach-Object { if ($null -ne $_) { [string]$_ } else { '' } })
                              ($items | Sort-Object) -join '; '
                          } else { [string]$rawVal }
                [void]$result.Add([pscustomobject]@{
                    DefinitionId = "$jsonOdataType|$propKey"
                    RawValue     = $valStr
                    Depth        = 0
                })
            }
        }
    } elseif ($settings) {
        foreach ($setting in $settings) {
            $instance = Get-SafeObjectPropertyValue -Object $setting -Name 'settingInstance'
            if (-not $instance) { continue }
            $leafItems = @(Get-BaselineJsonLeafInstances -Instance $instance)
            foreach ($leafItem in $leafItems) {
                [void]$result.Add($leafItem)
            }
        }
    }

    return [pscustomobject]@{
        PolicyName     = $policyName
        PolicyId       = $policyId
        PolicyType     = $policyType
        PolicyPlatform = $policyPlatform
        BaselineArea   = $areaInfo.BaselineArea
        AreaSupported  = $areaInfo.IsSupported
        Settings       = $result
    }
}

function Get-BaselineJsonLeafInstances {
    <#
    .SYNOPSIS
    Recursively extracts leaf settingDefinitionId -> value pairs from a JSON setting instance object.
    Uses Get-SafeObjectPropertyValue to avoid PropertyNotFoundException under Set-StrictMode.
    #>
    param($Instance, [int]$Depth = 0)
    if (-not $Instance) { return }

    # Group setting collection
    $groupCollVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'groupSettingCollectionValue'
    if ($groupCollVal) {
        foreach ($group in $groupCollVal) {
            $children = Get-SafeObjectPropertyValue -Object $group -Name 'children'
            if ($children) {
                foreach ($child in $children) {
                    Get-BaselineJsonLeafInstances -Instance $child -Depth $Depth
                }
            }
        }
        return
    }
    # Group setting value
    $groupSingleVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'groupSettingValue'
    if ($groupSingleVal) {
        $children = Get-SafeObjectPropertyValue -Object $groupSingleVal -Name 'children'
        if ($children) {
            foreach ($child in $children) {
                Get-BaselineJsonLeafInstances -Instance $child -Depth $Depth
            }
        }
        return
    }
    # Choice setting
    $choiceVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'choiceSettingValue'
    if ($choiceVal) {
        $defId = Get-SafeObjectPropertyValue -Object $Instance -Name 'settingDefinitionId'
        $rawValue = Get-SafeObjectPropertyValue -Object $choiceVal -Name 'value'
        [pscustomobject]@{ DefinitionId = $defId; RawValue = $rawValue; Depth = $Depth }
        $choiceChildren = Get-SafeObjectPropertyValue -Object $choiceVal -Name 'children'
        if ($choiceChildren) {
            foreach ($child in @($choiceChildren)) {
                Get-BaselineJsonLeafInstances -Instance $child -Depth ($Depth + 1)
            }
        }
        return
    }
    # Simple setting
    $simpleVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'simpleSettingValue'
    if ($simpleVal) {
        $defId = Get-SafeObjectPropertyValue -Object $Instance -Name 'settingDefinitionId'
        $rawValue = Get-SafeObjectPropertyValue -Object $simpleVal -Name 'value'
        [pscustomobject]@{ DefinitionId = $defId; RawValue = [string]$rawValue; Depth = $Depth }
        $simpleChildren = Get-SafeObjectPropertyValue -Object $simpleVal -Name 'children'
        if ($simpleChildren) {
            foreach ($child in @($simpleChildren)) {
                Get-BaselineJsonLeafInstances -Instance $child -Depth ($Depth + 1)
            }
        }
        return
    }
    # Simple setting collection
    $simpleCollVal = Get-SafeObjectPropertyValue -Object $Instance -Name 'simpleSettingCollectionValue'
    if ($simpleCollVal) {
        $defId = Get-SafeObjectPropertyValue -Object $Instance -Name 'settingDefinitionId'
        $values = @($simpleCollVal | ForEach-Object { $v = Get-SafeObjectPropertyValue -Object $_ -Name 'value'; [string]$v })
        $rawValue = $values -join '; '
        [pscustomobject]@{ DefinitionId = $defId; RawValue = $rawValue; Depth = $Depth }
        return
    }
    # Fallback - output definition ID with no value
    $defId = Get-SafeObjectPropertyValue -Object $Instance -Name 'settingDefinitionId'
    if ($defId) {
        [pscustomobject]@{ DefinitionId = $defId; RawValue = '-'; Depth = $Depth }
    }
}

function New-CompareKey {
    param(
        [string]$ProfileName,
        [string]$DefinitionId
    )
    $p = Normalize-CompareText -Text $ProfileName
    $d = Normalize-CompareText -Text $DefinitionId
    return "$p|$d"
}

function Compare-BaselineToLiveSettings {
    <#
    .SYNOPSIS
    Compares ALL settings from ALL baseline JSON files against ALL selected live Intune profiles.
    Returns data organized per JSON file so the report can list each baseline file as a section
    and show which live Intune profile(s) contain each setting.
    #>
    param(
        [string]$BaselineFolder,
        [object[]]$SettingRows
    )

    $result = @{
        ProfileStatus = @{}
        SettingStatus = @{}
        MissingRows   = [System.Collections.Generic.List[object]]::new()
        BaselineCount = 0
        BaselineFiles = [System.Collections.Generic.List[object]]::new()
    }

    # Build a reverse index: NormDefinitionId -> list of live profile hits
    $liveIndex = @{}
    foreach ($row in $SettingRows) {
        $defId = [string]$row.DefinitionId
        if (-not $defId -or $defId -eq '-') { continue }
        $normId = Normalize-CompareText -Text $defId
        if (-not $normId) { continue }
        if (-not $liveIndex.ContainsKey($normId)) {
            $liveIndex[$normId] = [System.Collections.Generic.List[object]]::new()
        }
        [void]$liveIndex[$normId].Add(@{
            ProfileName  = [string]$row.ProfileName
            ProfileId    = [string]$row.ProfileId
            RawValue     = [string]$row.RawValue
            SettingValue = [string]$row.SettingValue
            SettingName  = [string]$row.SettingName
        })
    }

    # Process each JSON file
    $jsonFiles = Get-ChildItem -Path $BaselineFolder -Filter '*.json' -File -ErrorAction SilentlyContinue
    $totalBaselineSettings = 0
    $profileHasMismatch = @{}
    $profileNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $SettingRows) { [void]$profileNames.Add([string]$row.ProfileName) }

    foreach ($jf in $jsonFiles) {
        try {
            $parsed = Get-BaselineSettingsFromJson -JsonPath $jf.FullName
            $policyName = if ($parsed.PolicyName) { [string]$parsed.PolicyName } else { [System.IO.Path]::GetFileNameWithoutExtension($jf.Name) }

            $fileEntry = @{
                FileName   = $jf.Name
                PolicyName = $policyName
                Settings   = [System.Collections.Generic.List[object]]::new()
            }

            foreach ($s in $parsed.Settings) {
                if (-not $s.DefinitionId) { continue }
                $totalBaselineSettings++
                $normId = Normalize-CompareText -Text ([string]$s.DefinitionId)
                $expectedRaw = [string]$s.RawValue

                $settingEntry = @{
                    DefinitionId  = [string]$s.DefinitionId
                    ExpectedValue = $expectedRaw
                    FoundIn       = [System.Collections.Generic.List[object]]::new()
                }

                if ($normId -and $liveIndex.ContainsKey($normId)) {
                    foreach ($hit in $liveIndex[$normId]) {
                        $liveRaw = [string]$hit.RawValue
                        if (-not $liveRaw -or $liveRaw -eq '-') { $liveRaw = [string]$hit.SettingValue }
                        $status = if ([string]$liveRaw -ieq [string]$expectedRaw) { 'Match' } else { 'Mismatch' }
                        [void]$settingEntry.FoundIn.Add(@{
                            ProfileName  = $hit.ProfileName
                            ProfileId    = $hit.ProfileId
                            ActualValue  = $hit.SettingValue
                            RawValue     = $hit.RawValue
                            SettingName  = $hit.SettingName
                            Status       = $status
                        })
                        if ($status -eq 'Mismatch') { $profileHasMismatch[$hit.ProfileName] = $true }
                        $key = New-CompareKey -ProfileName $hit.ProfileName -DefinitionId ([string]$s.DefinitionId)
                        $result.SettingStatus[$key] = @{ Status = $status; ExpectedValue = $expectedRaw }
                    }
                } else {
                    [void]$result.MissingRows.Add([pscustomobject]@{
                        ProfileName   = ''
                        DefinitionId  = [string]$s.DefinitionId
                        ExpectedValue = $expectedRaw
                        SourceFile    = $jf.Name
                    })
                }

                [void]$fileEntry.Settings.Add($settingEntry)
            }

            [void]$result.BaselineFiles.Add($fileEntry)
        } catch {
            Write-ProfileReportLog "  Warning: Failed to parse baseline JSON '$($jf.Name)': $($_.Exception.Message)"
        }
    }

    $result.BaselineCount = $totalBaselineSettings

    # Mark live settings not in any baseline as 'Extra'
    $baselineNormIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $result.BaselineFiles) {
        foreach ($s in $file.Settings) {
            $nid = Normalize-CompareText -Text ([string]$s.DefinitionId)
            if ($nid) { [void]$baselineNormIds.Add($nid) }
        }
    }
    foreach ($row in $SettingRows) {
        $defId = [string]$row.DefinitionId
        if (-not $defId -or $defId -eq '-') { continue }
        $normId = Normalize-CompareText -Text $defId
        if ($normId -and -not $baselineNormIds.Contains($normId)) {
            $key = New-CompareKey -ProfileName ([string]$row.ProfileName) -DefinitionId $defId
            if (-not $result.SettingStatus.ContainsKey($key)) {
                $result.SettingStatus[$key] = @{ Status = 'Extra'; ExpectedValue = '' }
            }
        }
    }

    # Profile-level status
    foreach ($profName in $profileNames) {
        $result.ProfileStatus[$profName] = if ($profileHasMismatch.ContainsKey($profName)) { 'Mismatch' } else { 'Match' }
    }

    return $result
}

function Get-BaselineIntuneProfiles {
        function Get-SafeProp {
            param($Object, [string]$Name)
            if (-not $Object) { return $null }
            if ($Object -is [System.Collections.IDictionary]) {
                foreach ($k in $Object.Keys) {
                    if ([string]$k -ieq $Name) { return $Object[$k] }
                }
                return $null
            }
            foreach ($p in $Object.PSObject.Properties) {
                if ($p.Name -ieq $Name) { return $p.Value }
            }
            $ap = $Object.PSObject.Properties['AdditionalProperties']
            if ($ap -and $ap.Value -is [System.Collections.IDictionary]) {
                foreach ($k in $ap.Value.Keys) {
                    if ([string]$k -ieq $Name) { return $ap.Value[$k] }
                }
            }
            return $null
        }

        function Get-AllPages {
            param([string]$Uri)
            $results = @()
            do {
                $response = Invoke-MgGraphRequest -Uri $Uri -Method Get -ErrorAction Stop
                $value = Get-SafeProp -Object $response -Name 'value'
                if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
                    $results += $value
                } elseif ($response) {
                    $results += $response
                }
                $Uri = Get-SafeProp -Object $response -Name '@odata.nextLink'
            } while ($Uri)
            return $results
        }

        $uri = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name,description,platforms,lastModifiedDateTime'

        $profiles = @(Get-AllPages -Uri $uri)
        $mapped = @()
        foreach ($item in $profiles) {
            $id = Get-SafeProp -Object $item -Name 'id'
            $name = Get-SafeProp -Object $item -Name 'name'
            $platforms = Get-SafeProp -Object $item -Name 'platforms'
            $lastModified = Get-SafeProp -Object $item -Name 'lastModifiedDateTime'

            if (-not $id -or -not $name) {
                continue
            }

            $mapped += [pscustomobject]@{
                Selected = $false
                Id = [string]$id
                Name = [string]$name
                Platforms = if ($platforms) { ([string[]]$platforms) -join ', ' } else { '-' }
                LastModifiedDateTime = if ($lastModified) {
                    try { (Get-Date $lastModified).ToString('yyyy-MM-dd HH:mm') } catch { [string]$lastModified }
                } else { '-' }
                PolicyKind = 'SettingsCatalog'
                OdataType  = ''
            }
        }

        # Also load Device Restriction profiles (deviceConfigurations endpoint)
        $restrictionOdataTypes = @(
            '#microsoft.graph.iosGeneralDeviceConfiguration',
            '#microsoft.graph.iosDeviceFeaturesConfiguration',
            '#microsoft.graph.androidGeneralDeviceConfiguration',
            '#microsoft.graph.androidWorkProfileGeneralDeviceConfiguration',
            '#microsoft.graph.androidDeviceOwnerGeneralDeviceConfiguration',
            '#microsoft.graph.aospDeviceOwnerGeneralDeviceConfiguration'
        )
        $dcUri = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$select=id,displayName,lastModifiedDateTime'
        try {
            $deviceConfigs = @(Get-AllPages -Uri $dcUri)
            foreach ($dcItem in $deviceConfigs) {
                $dcId      = Get-SafeProp -Object $dcItem -Name 'id'
                $dcName    = Get-SafeProp -Object $dcItem -Name 'displayName'
                $dcType    = Get-SafeProp -Object $dcItem -Name '@odata.type'
                $dcLastMod = Get-SafeProp -Object $dcItem -Name 'lastModifiedDateTime'
                if (-not $dcId -or -not $dcName) { continue }
                if (-not $dcType -or $dcType -notin $restrictionOdataTypes) { continue }
                $dcPlatform = if ($dcType -match 'ios') { 'iOS' } elseif ($dcType -match 'android') { 'Android' } else { 'Other' }
                $mapped += [pscustomobject]@{
                    Selected = $false
                    Id       = [string]$dcId
                    Name     = [string]$dcName
                    Platforms = $dcPlatform
                    LastModifiedDateTime = if ($dcLastMod) {
                        try { (Get-Date $dcLastMod).ToString('yyyy-MM-dd HH:mm') } catch { [string]$dcLastMod }
                    } else { '-' }
                    PolicyKind = 'DeviceRestriction'
                    OdataType  = [string]$dcType
                }
            }
        } catch {
            # Device Restriction load failure is non-fatal; Settings Catalog profiles are still returned.
        }

        # Also load Administrative Templates (groupPolicyConfigurations endpoint)
        $gpUri = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?$select=id,displayName,lastModifiedDateTime'
        try {
            $gpConfigs = @(Get-AllPages -Uri $gpUri)
            foreach ($gpItem in $gpConfigs) {
                $gpId      = Get-SafeProp -Object $gpItem -Name 'id'
                $gpName    = Get-SafeProp -Object $gpItem -Name 'displayName'
                $gpLastMod = Get-SafeProp -Object $gpItem -Name 'lastModifiedDateTime'
                if (-not $gpId -or -not $gpName) { continue }
                $mapped += [pscustomobject]@{
                    Selected = $false
                    Id       = [string]$gpId
                    Name     = [string]$gpName
                    Platforms = 'Windows'
                    LastModifiedDateTime = if ($gpLastMod) {
                        try { (Get-Date $gpLastMod).ToString('yyyy-MM-dd HH:mm') } catch { [string]$gpLastMod }
                    } else { '-' }
                    PolicyKind = 'AdministrativeTemplate'
                    OdataType  = '#microsoft.graph.groupPolicyConfiguration'
                }
            }
        } catch {
            # Administrative Template load failure is non-fatal; other profiles are still returned.
        }

        [pscustomobject]@{
            RawCount = $profiles.Count
            Mapped = $mapped
        }
    }

function Build-HtmlReport {
    param(
        [string]$OutputPath,
        [string]$MdFileName,
        [string]$ProfileLevel,
        [string]$Platform,
        [array]$ReportRows
    )

    $generatedDate = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $totalCreated      = @($ReportRows | Where-Object { $_.Outcome -eq 'Created' }).Count
    $totalNotFound     = @($ReportRows | Where-Object { $_.Outcome -eq 'NotFound' }).Count
    $totalMismatch     = @($ReportRows | Where-Object { $_.Outcome -eq 'ValueMismatch' }).Count
    $totalManual       = @($ReportRows | Where-Object { $_.Outcome -eq 'Manual' }).Count
    $totalCustomerInfo = @($ReportRows | Where-Object { $_.Outcome -eq 'CustomerInfoMissing' }).Count
    $totalAutoUnsupported = @($ReportRows | Where-Object { $_.Outcome -eq 'AutoCreationUnsupported' }).Count
    $policiesCreated   = @($ReportRows | Where-Object { $_.Outcome -eq 'Created' } | ForEach-Object { $_.PolicyName } | Sort-Object -Unique).Count
    $policiesSkipped   = @($ReportRows | Where-Object { $_.Outcome -in 'NotFound','Manual','ValueMismatch','CustomerInfoMissing','AutoCreationUnsupported' } | ForEach-Object { $_.PolicyName } | Sort-Object -Unique).Count

    $rowsByPolicy = $ReportRows | Group-Object PolicyName | Sort-Object Name

    $sectionsHtml = [System.Text.StringBuilder]::new()
    foreach ($grp in $rowsByPolicy) {
        $policyName = $grp.Name
        $rows = $grp.Group
        $pCreated  = @($rows | Where-Object { $_.Outcome -eq 'Created' }).Count
        $pSkipped  = @($rows | Where-Object { $_.Outcome -ne 'Created' }).Count

        $statusBadge = if ($pCreated -gt 0 -and $pSkipped -eq 0) {
            '<span class="badge badge-ok">Created</span>'
        } elseif ($pCreated -gt 0) {
            '<span class="badge badge-partial">Partial</span>'
        } elseif (@($rows | Where-Object { $_.Outcome -eq 'CustomerInfoMissing' }).Count -gt 0) {
            '<span class="badge badge-customer">Customer Info Needed</span>'
        } else {
            '<span class="badge badge-skip">Skipped</span>'
        }

        $safePolicyName = [System.Net.WebUtility]::HtmlEncode($policyName)
        [void]$sectionsHtml.Append("<section class='policy-section'>`n")
        [void]$sectionsHtml.Append("<div class='policy-header' onclick='toggleSection(this)'>`n")
        [void]$sectionsHtml.Append("  <span class='policy-arrow'>&#9654;</span>`n")
        [void]$sectionsHtml.Append("  <span class='policy-name'>$safePolicyName</span>`n")
        [void]$sectionsHtml.Append("  $statusBadge`n")
        [void]$sectionsHtml.Append("  <span class='policy-counts'>$pCreated created &bull; $pSkipped skipped</span>`n")
        [void]$sectionsHtml.Append("</div>`n")
        [void]$sectionsHtml.Append("<div class='policy-body'>`n")
        [void]$sectionsHtml.Append("<table><thead><tr><th>CIS</th><th>Setting</th><th>Value</th><th>L</th><th>Status</th><th>Detail</th></tr></thead><tbody>`n")

        foreach ($row in ($rows | Sort-Object CISRef)) {
            $outcomeCss = switch ($row.Outcome) {
                'Created'             { 'outcome-ok' }
                'NotFound'            { 'outcome-notfound' }
                'ValueMismatch'       { 'outcome-mismatch' }
                'Manual'              { 'outcome-manual' }
                'CustomerInfoMissing' { 'outcome-customerinfo' }
                'AutoCreationUnsupported' { 'outcome-autoskip' }
                default               { '' }
            }
            $outcomeLabel = switch ($row.Outcome) {
                'Created'             { '&#10003; Created' }
                'NotFound'            { '&#10008; Not found in Intune' }
                'ValueMismatch'       { '&#9888; Value mismatch' }
                'Manual'              { '&#9998; Needs manual config' }
                'CustomerInfoMissing' { '&#9888; Customer info missing' }
                'AutoCreationUnsupported' { '&#9888; Cannot be automatically created in Intune' }
                default               { $row.Outcome }
            }
            $safeDetail  = [System.Net.WebUtility]::HtmlEncode($row.Detail)
            $safeSetting = [System.Net.WebUtility]::HtmlEncode($row.SettingName)
            $safeValue   = [System.Net.WebUtility]::HtmlEncode($row.SettingValue)
            [void]$sectionsHtml.Append("<tr class='$outcomeCss'><td>$($row.CISRef)</td><td>$safeSetting</td><td>$safeValue</td><td>$($row.Profile)</td><td>$outcomeLabel</td><td>$safeDetail</td></tr>`n")
        }

        [void]$sectionsHtml.Append("</tbody></table>`n</div></section>`n")
    }

    $safeMdFileName = [System.Net.WebUtility]::HtmlEncode($MdFileName)
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Intune CIS Policy Creation Report</title>
<style>
:root{--ok:#107c10;--skip:#d83b01;--warn:#ca5010;--blue:#0078D4;--bg:#fff;--fg:#1a1a1a;--border:#ddd;--hbg:#f4f4f4}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--fg);padding:24px;line-height:1.5}
h1{font-size:1.5rem;color:var(--blue);margin-bottom:4px}
.meta{font-size:.85rem;color:#666;margin-bottom:16px}
.stats{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:20px}
.stat{background:var(--hbg);border:1px solid var(--border);border-radius:6px;padding:10px 16px;min-width:110px}
.stat-value{font-size:1.5rem;font-weight:700;color:var(--blue)}
.stat-label{font-size:.72rem;text-transform:uppercase;letter-spacing:.5px;color:#666}
.toolbar{display:flex;gap:10px;margin-bottom:12px;flex-wrap:wrap;align-items:center}
.toolbar input{padding:6px 10px;border:1px solid var(--border);border-radius:4px;font-size:.85rem;width:240px}
.toolbar button{padding:6px 12px;border:1px solid var(--border);border-radius:4px;background:var(--hbg);cursor:pointer;font-size:.85rem}
.toolbar button:hover{background:#e0e0e0}
.policy-section{border:1px solid var(--border);border-radius:6px;margin-bottom:8px;overflow:hidden}
.policy-header{display:flex;align-items:center;gap:10px;padding:10px 14px;background:var(--hbg);cursor:pointer;user-select:none}
.policy-header:hover{background:#e8e8e8}
.policy-arrow{font-size:.7rem;transition:transform .2s;display:inline-block}
.policy-arrow.open{transform:rotate(90deg)}
.policy-name{font-weight:600;flex:1;font-size:.92rem}
.policy-counts{font-size:.8rem;color:#666}
.policy-body{display:none;padding:0}
.policy-body.open{display:block}
table{width:100%;border-collapse:collapse;font-size:.82rem}
th{background:var(--blue);color:#fff;padding:7px 8px;text-align:left;white-space:nowrap}
td{padding:6px 8px;border-bottom:1px solid #eee;vertical-align:top}
.badge{display:inline-block;padding:2px 9px;border-radius:3px;font-size:.75rem;font-weight:600}
.badge-ok{background:#dff6dd;color:var(--ok)}
.badge-partial{background:#fff4ce;color:var(--warn)}
.badge-customer{background:#e6f2ff;color:#004578}
.badge-skip{background:#fde7e9;color:#a80000}
.outcome-ok td:nth-child(5){color:var(--ok);font-weight:600}
.outcome-notfound td:nth-child(5){color:#a80000;font-weight:600}
.outcome-mismatch td:nth-child(5){color:var(--warn);font-weight:600}
.outcome-manual td:nth-child(5){color:#555;font-style:italic}
.outcome-customerinfo td:nth-child(5){color:#004578;font-weight:600}
.outcome-autoskip td:nth-child(5){color:#9A3412;font-weight:600}
.hidden{display:none}
</style>
</head>
<body>
<h1>Intune CIS Policy Creation Report</h1>
<p class="meta">Generated $generatedDate &bull; Source: $safeMdFileName &bull; Profile: $ProfileLevel &bull; Platform: $Platform</p>
<div class="stats">
  <div class="stat"><div class="stat-value">$policiesCreated</div><div class="stat-label">Policies Created</div></div>
  <div class="stat"><div class="stat-value">$policiesSkipped</div><div class="stat-label">Policies Skipped</div></div>
  <div class="stat"><div class="stat-value">$totalCreated</div><div class="stat-label">Settings Created</div></div>
  <div class="stat"><div class="stat-value">$totalNotFound</div><div class="stat-label">Not in Intune</div></div>
    <div class="stat"><div class="stat-value">$totalAutoUnsupported</div><div class="stat-label">Not Auto-Creatable</div></div>
  <div class="stat"><div class="stat-value">$totalCustomerInfo</div><div class="stat-label">Customer Info Needed</div></div>
  <div class="stat"><div class="stat-value">$totalMismatch</div><div class="stat-label">Value Mismatch</div></div>
  <div class="stat"><div class="stat-value">$totalManual</div><div class="stat-label">Needs Manual</div></div>
</div>
<div class="toolbar">
  <input type="text" id="search" placeholder="Search policies or settings..." oninput="applyFilter()">
  <button onclick="expandAll()">Expand All</button>
  <button onclick="collapseAll()">Collapse All</button>
  <button onclick="showAll()">Show All</button>
  <button onclick="showSkipped()">Show Skipped Only</button>
</div>
$($sectionsHtml.ToString())
<script>
function toggleSection(h){var a=h.querySelector('.policy-arrow'),b=h.nextElementSibling,o=b.classList.toggle('open');if(o)a.classList.add('open');else a.classList.remove('open')}
function expandAll(){document.querySelectorAll('.policy-body').forEach(b=>{b.classList.add('open');b.previousElementSibling.querySelector('.policy-arrow').classList.add('open')})}
function collapseAll(){document.querySelectorAll('.policy-body').forEach(b=>{b.classList.remove('open');b.previousElementSibling.querySelector('.policy-arrow').classList.remove('open')})}
function showAll(){document.querySelectorAll('.policy-section').forEach(s=>s.classList.remove('hidden'));document.getElementById('search').value=''}
function showSkipped(){document.getElementById('search').value='';document.querySelectorAll('.policy-section').forEach(s=>{var b=s.querySelector('.badge');var skip=b&&(b.classList.contains('badge-skip')||b.classList.contains('badge-partial'));s.classList.toggle('hidden',!skip)})}
function applyFilter(){var q=document.getElementById('search').value.toLowerCase();document.querySelectorAll('.policy-section').forEach(s=>{s.classList.toggle('hidden',q&&!s.textContent.toLowerCase().includes(q))})}
collapseAll();
</script>
</body></html>
"@

    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
}
