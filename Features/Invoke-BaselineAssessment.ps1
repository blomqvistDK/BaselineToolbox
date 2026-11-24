#Requires -Version 5.1
<#
.SYNOPSIS
    Assesses baseline JSON backups against selected Intune profiles without a GUI.
.EXAMPLE
    .\Invoke-BaselineAssessment.ps1 -BaselineFolder .\Baselines -ProfileId 'profile-guid' -OutputPath .\assessment.html -SkipAssignmentLookups
.EXAMPLE
    .\Invoke-BaselineAssessment.ps1 -BaselineFolder .\Baselines -ProfileId 'profile-guid' -OutputPath .\assessment.html -CreateSplitPolicy -WhatIf
.NOTES
    A valid Microsoft Graph session is reused. Dot-source to load processing functions only.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$BaselineFolder,
    [string[]]$ProfileId,
    [object[]]$Profiles,
    [string]$OutputPath,
    [string]$TenantId = $env:BASELINE_TOOLBOX_RESOLVER_TENANTID,
    [ValidateSet('fallback','cacheonly','forcefresh')][string]$CacheMode = 'fallback',
    [switch]$SkipUnassigned,
    [switch]$SkipAssignmentLookups,
    [switch]$OpenReport,
    [switch]$CreateSplitPolicy,
    [string[]]$SplitBaselineName,
    [scriptblock]$LogCallback
)

Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'Shared\BaselineToolkit.Common.psm1') -DisableNameChecking

function Get-AlternateScopeDefinitionId {
    param([string]$DefinitionId)
    if ($DefinitionId -match '^device_') { return ($DefinitionId -replace '^device_', 'user_') }
    if ($DefinitionId -match '^user_') { return ($DefinitionId -replace '^user_', 'device_') }
    return $null
}

function Invoke-BaselineAssessment {
    <#
    .SYNOPSIS
    Compares each baseline JSON file individually against all customer policies
    and classifies each baseline into import-readiness categories.
    #>
    param(
        [string]$BaselineFolder,
        [object[]]$SettingRows,
        [hashtable]$CustomerPolicyRawSettings = @{}
    )

    # Build live index: NormDefinitionId -> list of { ProfileName, ProfileId, RawValue }
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
            ProfileName = [string]$row.ProfileName
            ProfileId   = [string]$row.ProfileId
            RawValue    = [string]$row.RawValue
            SettingName = [string]$row.SettingName
        })
    }

    # Build reverse map: for each customer policy, which baseline settings does it contain?
    $customerPolicySettings = @{}  # ProfileName -> list of all NormDefinitionIds in that policy
    foreach ($row in $SettingRows) {
        $profName = [string]$row.ProfileName
        $defId = [string]$row.DefinitionId
        if (-not $defId -or $defId -eq '-') { continue }
        $normId = Normalize-CompareText -Text $defId
        if (-not $normId) { continue }
        if (-not $customerPolicySettings.ContainsKey($profName)) {
            $customerPolicySettings[$profName] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        }
        [void]$customerPolicySettings[$profName].Add($normId)
    }

    $jsonFiles = @(Get-ChildItem -Path $BaselineFolder -Filter '*.json' -File -ErrorAction SilentlyContinue)
    $assessmentResults = [System.Collections.Generic.List[object]]::new()

    # Collect ALL baseline setting IDs across all files for the "customer-only" analysis
    $allBaselineNormIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($jf in $jsonFiles) {
        try {
            $parsed = Get-BaselineSettingsFromJson -JsonPath $jf.FullName
            $policyName = if ($parsed.PolicyName) { [string]$parsed.PolicyName } else { [System.IO.Path]::GetFileNameWithoutExtension($jf.Name) }

            $baselineSettings = $parsed.Settings
            if (-not $baselineSettings -or $baselineSettings.Count -eq 0) {
                [void]$assessmentResults.Add(@{
                    FileName       = $jf.Name
                    PolicyName     = $policyName
                    Classification = 'Error'
                    Detail         = if ($parsed.PolicyPlatform -and $parsed.BaselineArea -and -not $parsed.AreaSupported) {
                        "Unsupported Device Configuration baseline area '$($parsed.BaselineArea)' in filename."
                    } else {
                        'No configured settings matched this baseline area. Verify that the area suffix and policy settings are correct.'
                    }
                    TotalSettings  = 0
                    MatchCount     = 0
                    MismatchCount  = 0
                    NotFoundCount  = 0
                    FoundInProfiles = @{}
                    Settings       = @()
                    PolicyPlatform = $parsed.PolicyPlatform
                    BaselineArea   = $parsed.BaselineArea
                    MatchedProfileName = $null
                })
                continue
            }

            $settingDetails = [System.Collections.Generic.List[object]]::new()
            $foundInProfiles = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $totalFound = 0
            $totalMatch = 0
            $totalMismatch = 0
            $totalNotFound = 0

            foreach ($s in $baselineSettings) {
                if (-not $s.DefinitionId) { continue }
                $normId = Normalize-CompareText -Text ([string]$s.DefinitionId)
                [void]$allBaselineNormIds.Add($normId)
                $expectedRaw = [string]$s.RawValue
                $settingDepth = if ($s.PSObject.Properties['Depth']) { [int]$s.Depth } else { 0 }

                $detail = @{
                    DefinitionId  = [string]$s.DefinitionId
                    ExpectedValue = $expectedRaw
                    Depth         = $settingDepth
                    Status        = 'NotFound'
                    FoundIn       = [System.Collections.Generic.List[object]]::new()
                }

                if ($normId -and $liveIndex.ContainsKey($normId)) {
                    $totalFound++
                    foreach ($hit in $liveIndex[$normId]) {
                        $liveRaw = [string]$hit.RawValue
                        $status = if ([string]$liveRaw -ieq [string]$expectedRaw) { 'Match' } else { 'Mismatch' }
                        [void]$detail.FoundIn.Add(@{
                            ProfileName = $hit.ProfileName
                            ProfileId   = $hit.ProfileId
                            ActualValue = $hit.RawValue
                            SettingName = $hit.SettingName
                            Status      = $status
                        })

                        if (-not $foundInProfiles.ContainsKey($hit.ProfileName)) {
                            $foundInProfiles[$hit.ProfileName] = [System.Collections.Generic.List[object]]::new()
                        }
                        [void]$foundInProfiles[$hit.ProfileName].Add(@{ Status = $status; DefinitionId = [string]$s.DefinitionId })
                    }
                    # Determine aggregate status for this setting
                    $hasMatch = $false
                    $hasMismatch = $false
                    foreach ($f in $detail.FoundIn) {
                        if ($f.Status -eq 'Match') { $hasMatch = $true }
                        if ($f.Status -eq 'Mismatch') { $hasMismatch = $true }
                    }
                    if ($hasMismatch) { $detail.Status = 'Mismatch'; $totalMismatch++ }
                    else { $detail.Status = 'Match'; $totalMatch++ }
                } elseif ($normId -and ($altId = Get-AlternateScopeDefinitionId -DefinitionId ([string]$s.DefinitionId)) -and $liveIndex.ContainsKey((Normalize-CompareText -Text $altId))) {
                    $totalFound++
                    $altNormId = Normalize-CompareText -Text $altId
                    [void]$allBaselineNormIds.Add($altNormId)
                    $srcId = [string]$s.DefinitionId
                    # Choice values embed the definition id, so the expected value needs the same scope swap.
                    $altExpected = if ($expectedRaw -and $expectedRaw.StartsWith("${srcId}_", [System.StringComparison]::OrdinalIgnoreCase)) { $altId + $expectedRaw.Substring($srcId.Length) } else { $expectedRaw }
                    $altScopeLabel = if ($altId -match '^user_') { 'User' } else { 'Device' }
                    $detail.ScopeAlternate = $true
                    foreach ($hit in $liveIndex[$altNormId]) {
                        $status = if ([string]$hit.RawValue -ieq [string]$altExpected) { 'Match' } else { 'Mismatch' }
                        [void]$detail.FoundIn.Add(@{
                            ProfileName = $hit.ProfileName
                            ProfileId   = $hit.ProfileId
                            ActualValue = $hit.RawValue
                            SettingName = $hit.SettingName
                            Status      = $status
                            ScopeNote   = "Configured as $altScopeLabel scope"
                        })
                        if (-not $foundInProfiles.ContainsKey($hit.ProfileName)) {
                            $foundInProfiles[$hit.ProfileName] = [System.Collections.Generic.List[object]]::new()
                        }
                        [void]$foundInProfiles[$hit.ProfileName].Add(@{ Status = $status; DefinitionId = [string]$s.DefinitionId })
                    }
                    if (@($detail.FoundIn | Where-Object { $_.Status -eq 'Mismatch' }).Count -gt 0) { $detail.Status = 'Mismatch'; $totalMismatch++ }
                    else { $detail.Status = 'Match'; $totalMatch++ }
                } else {
                    $totalNotFound++
                }

                [void]$settingDetails.Add($detail)
            }

            # Classify this baseline JSON
            $totalSettings = $settingDetails.Count
            $classification = 'Conflicts'
            $classificationDetail = ''

            if ($totalNotFound -eq $totalSettings) {
                $classification = 'ReadyToImport'
                $classificationDetail = 'All settings are new to the customer environment'
            } elseif ($totalMatch -eq $totalFound -and $totalNotFound -eq 0) {
                # All settings found and matching
                $distinctProfiles = @($foundInProfiles.Keys)
                if ($distinctProfiles.Count -eq 1) {
                    if ($distinctProfiles[0] -ieq $policyName) {
                        $classification = 'FullyAligned'
                        $classificationDetail = "All settings match policy: $($distinctProfiles[0])"
                    } else {
                        # Check if customer policy has ONLY baseline settings (pure rename)
                        $custPolicyTotalSettings = if ($customerPolicySettings.ContainsKey($distinctProfiles[0])) {
                            $customerPolicySettings[$distinctProfiles[0]].Count
                        } else { 0 }
                        if ($custPolicyTotalSettings -le $totalSettings) {
                            $classification = 'RenameOnly'
                            $classificationDetail = "All settings match customer policy '$($distinctProfiles[0])'"
                        } else {
                            $classification = 'SplitRequired'
                            $classificationDetail = "All settings match in customer policy '$($distinctProfiles[0])' (policy has $($custPolicyTotalSettings - $totalSettings) extra settings)"
                        }
                    }
                } else {
                    $classification = 'ConsolidationNeeded'
                    $classificationDetail = "All settings match but spread across $($distinctProfiles.Count) policies"
                }
            } elseif ($totalMatch -eq $totalFound -and $totalNotFound -gt 0) {
                # Some found and matching, some not found at all
                $distinctProfiles = @($foundInProfiles.Keys)
                $classification = 'ReadyToImport'
                $classificationDetail = "$totalNotFound settings are new; $totalMatch already compliant across $($distinctProfiles.Count) policies"
            } else {
                # Has mismatches
                $classification = 'Conflicts'
                $classificationDetail = "$totalMismatch value mismatch(es), $totalNotFound not found, $totalMatch matching"
            }

            [void]$assessmentResults.Add(@{
                FileName       = $jf.Name
                PolicyName     = $policyName
                PolicyPlatform = $parsed.PolicyPlatform
                BaselineArea   = $parsed.BaselineArea
                AreaSupported  = $parsed.AreaSupported
                Classification = $classification
                Detail         = $classificationDetail
                TotalSettings  = $totalSettings
                MatchCount     = $totalMatch
                MismatchCount  = $totalMismatch
                NotFoundCount  = $totalNotFound
                FoundInProfiles = $foundInProfiles
                Settings       = $settingDetails
                MatchedProfileName = if ($classification -eq 'SplitRequired' -and $foundInProfiles.Keys.Count -gt 0) { @($foundInProfiles.Keys)[0] } elseif ($classification -eq 'RenameOnly' -and $foundInProfiles.Keys.Count -gt 0) { @($foundInProfiles.Keys)[0] } else { $null }
            })
        } catch {
            [void]$assessmentResults.Add(@{
                FileName       = $jf.Name
                PolicyName     = $jf.Name
                PolicyPlatform = $null
                BaselineArea   = $null
                AreaSupported  = $false
                Classification = 'Error'
                Detail         = $_.Exception.Message
                TotalSettings  = 0
                MatchCount     = 0
                MismatchCount  = 0
                NotFoundCount  = 0
                FoundInProfiles = @{}
                Settings       = @()
                MatchedProfileName = $null
            })
        }
    }

    # Build "Customer Policies to Split" analysis
    # For each customer policy that overlaps with any baseline, identify customer-only settings
    $customerPolicySplit = [System.Collections.Generic.List[object]]::new()
    foreach ($profName in $customerPolicySettings.Keys) {
        $profNormIds = $customerPolicySettings[$profName]
        $baselineOverlap = 0
        $customerOnly = 0
        $customerOnlyIds = [System.Collections.Generic.List[string]]::new()
        foreach ($nid in $profNormIds) {
            if ($allBaselineNormIds.Contains($nid)) {
                $baselineOverlap++
            } else {
                $customerOnly++
                [void]$customerOnlyIds.Add($nid)
            }
        }
        if ($baselineOverlap -gt 0 -and $customerOnly -gt 0) {
            # This policy has a mix — needs splitting
            [void]$customerPolicySplit.Add(@{
                ProfileName      = $profName
                TotalSettings    = $profNormIds.Count
                BaselineSettings = $baselineOverlap
                CustomerOnly     = $customerOnly
                CustomerOnlyIds  = $customerOnlyIds
            })
        }
    }

    # Reclassify SplitRequired baselines and compute how many settings in the matched
    # customer policy are TRULY customer-only (present in NO baseline JSON file). A
    # baseline stays SplitRequired only when at least one such customer-only setting
    # exists; those settings are the exact scope of the "Create Split Policy" action.
    foreach ($b in $assessmentResults) {
        if ($b.Classification -ne 'SplitRequired') { continue }
        $matchedName = [string]$b.MatchedProfileName
        if (-not $matchedName -and $b.FoundInProfiles -and $b.FoundInProfiles.Keys.Count -gt 0) {
            $matchedName = @($b.FoundInProfiles.Keys)[0]
        }
        if (-not $matchedName) { continue }

        $customerOnlyIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        if ($customerPolicySettings.ContainsKey($matchedName)) {
            foreach ($nid in $customerPolicySettings[$matchedName]) {
                if (-not $allBaselineNormIds.Contains($nid)) { [void]$customerOnlyIds.Add($nid) }
            }
        }
        $customerOnlyCount = $customerOnlyIds.Count
        $b['MatchedProfileName'] = $matchedName
        $b['CustomerOnlyCount']  = $customerOnlyCount
        $b['CustomerOnlyIds']    = $customerOnlyIds

        if ($customerOnlyCount -eq 0) {
            # Every "extra" setting belongs to other baselines - nothing customer-specific to split off.
            $b['Classification'] = 'RenameOnly'
            $b['Detail']         = "All settings match customer policy '$matchedName' (remaining settings belong to other baselines)"
        } else {
            $b['Detail'] = "All baseline settings present in '$matchedName', plus $customerOnlyCount customer-only setting(s) to split out"
        }
    }

    return @{
        Baselines              = $assessmentResults
        CustomerPolicySplit    = $customerPolicySplit
        CustomerPolicySettings = $customerPolicySettings
        CustomerPolicyRawSettings = $CustomerPolicyRawSettings
        AllBaselineNormIds     = $allBaselineNormIds
        TotalBaselineFiles     = $jsonFiles.Count
    }
}

function Get-BaselinePolicyNameWithoutSecurityBaseline {
    param([string]$PolicyName)

    if (-not $PolicyName) { return $PolicyName }
    $newName = $PolicyName -replace '-SecurityBaseline-', '-'
    $newName = $newName -replace 'SecurityBaseline-', ''
    $newName = $newName -replace '-{2,}', '-'
    return $newName.Trim('-')
}

function Remove-BaselineDefinitionsFromSettingInstance {
    param(
        [object]$Instance,
        [System.Collections.Generic.HashSet[string]]$ExcludedDefinitionIds
    )

    function Invoke-PruneSettingNode {
        param([object]$Node)

        if (-not $Node) { return $null }

        $nodeDefId = [string](Get-SafeObjectPropertyValue -Object $Node -Name 'settingDefinitionId')
        if ($nodeDefId) {
            $normId = Normalize-CompareText -Text $nodeDefId
            if ($normId -and $ExcludedDefinitionIds.Contains($normId)) { return $null }
        }

        $cloneJson = $Node | ConvertTo-Json -Depth 40
        if (-not $cloneJson) { return $null }
        $clone = $cloneJson | ConvertFrom-Json

        $groupCollection = Get-SafeObjectPropertyValue -Object $clone -Name 'groupSettingCollectionValue'
        if ($groupCollection -and $groupCollection -is [System.Collections.IEnumerable] -and -not ($groupCollection -is [string])) {
            $prunedGroups = [System.Collections.Generic.List[object]]::new()
            foreach ($group in @($groupCollection)) {
                if (-not $group) { continue }
                $children = Get-SafeObjectPropertyValue -Object $group -Name 'children'
                $prunedChildren = [System.Collections.Generic.List[object]]::new()
                foreach ($child in @($children)) {
                    $prunedChild = Invoke-PruneSettingNode -Node $child
                    if ($prunedChild) { [void]$prunedChildren.Add($prunedChild) }
                }
                if ($prunedChildren.Count -gt 0) {
                    if ($group.PSObject.Properties['children']) { $group.children = @($prunedChildren) }
                    [void]$prunedGroups.Add($group)
                }
            }
            if ($prunedGroups.Count -eq 0) { return $null }
            $clone.groupSettingCollectionValue = @($prunedGroups)
            return $clone
        }

        $groupValue = Get-SafeObjectPropertyValue -Object $clone -Name 'groupSettingValue'
        if ($groupValue) {
            $children = Get-SafeObjectPropertyValue -Object $groupValue -Name 'children'
            $prunedChildren = [System.Collections.Generic.List[object]]::new()
            foreach ($child in @($children)) {
                $prunedChild = Invoke-PruneSettingNode -Node $child
                if ($prunedChild) { [void]$prunedChildren.Add($prunedChild) }
            }
            if ($prunedChildren.Count -eq 0) { return $null }
            if ($groupValue.PSObject.Properties['children']) { $groupValue.children = @($prunedChildren) }
            $clone.groupSettingValue = $groupValue
            return $clone
        }

        $choiceValue = Get-SafeObjectPropertyValue -Object $clone -Name 'choiceSettingValue'
        if ($choiceValue) {
            $children = Get-SafeObjectPropertyValue -Object $choiceValue -Name 'children'
            if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
                $prunedChildren = [System.Collections.Generic.List[object]]::new()
                foreach ($child in @($children)) {
                    $prunedChild = Invoke-PruneSettingNode -Node $child
                    if ($prunedChild) { [void]$prunedChildren.Add($prunedChild) }
                }
                if ($choiceValue.PSObject.Properties['children']) { $choiceValue.children = @($prunedChildren) }
            }
            $clone.choiceSettingValue = $choiceValue
            return $clone
        }

        $simpleValue = Get-SafeObjectPropertyValue -Object $clone -Name 'simpleSettingValue'
        if ($simpleValue) {
            $children = Get-SafeObjectPropertyValue -Object $simpleValue -Name 'children'
            if ($children -and $children -is [System.Collections.IEnumerable] -and -not ($children -is [string])) {
                $prunedChildren = [System.Collections.Generic.List[object]]::new()
                foreach ($child in @($children)) {
                    $prunedChild = Invoke-PruneSettingNode -Node $child
                    if ($prunedChild) { [void]$prunedChildren.Add($prunedChild) }
                }
                if ($simpleValue.PSObject.Properties['children']) { $simpleValue.children = @($prunedChildren) }
            }
            $clone.simpleSettingValue = $simpleValue
            return $clone
        }

        $simpleCollection = Get-SafeObjectPropertyValue -Object $clone -Name 'simpleSettingCollectionValue'
        if ($simpleCollection -and $simpleCollection -is [System.Collections.IEnumerable] -and -not ($simpleCollection -is [string])) {
            $prunedCollection = [System.Collections.Generic.List[object]]::new()
            foreach ($entry in @($simpleCollection)) {
                $prunedEntry = Invoke-PruneSettingNode -Node $entry
                if ($prunedEntry) { [void]$prunedCollection.Add($prunedEntry) }
            }
            if ($prunedCollection.Count -eq 0) { return $null }
            $clone.simpleSettingCollectionValue = @($prunedCollection)
            return $clone
        }

        return $clone
    }

    return Invoke-PruneSettingNode -Node $Instance
}

function New-SplitRequiredPolicyFromAssessment {
    param(
        [hashtable]$AssessmentData,
        [object]$Baseline
    )

    if (-not $AssessmentData -or -not $Baseline) { return $null }

    $baselinePolicyName = [string](Get-SafeObjectPropertyValue -Object $Baseline -Name 'PolicyName')
    $matchedProfileName = [string](Get-SafeObjectPropertyValue -Object $Baseline -Name 'MatchedProfileName')
    if (-not $matchedProfileName) {
        $baselineFoundIn = Get-SafeObjectPropertyValue -Object $Baseline -Name 'FoundInProfiles'
        $foundInKeys = if ($baselineFoundIn -is [System.Collections.IDictionary]) { @($baselineFoundIn.Keys) }
                       elseif ($baselineFoundIn) { @($baselineFoundIn.PSObject.Properties.Name) }
                       else { @() }
        if ($foundInKeys.Count -gt 0) { $matchedProfileName = [string]$foundInKeys[0] }
    }
    if (-not $matchedProfileName) { throw "No matching customer policy was found for baseline '$baselinePolicyName'." }

    $rawSettingsMap = if ($AssessmentData.ContainsKey('CustomerPolicyRawSettings')) { $AssessmentData['CustomerPolicyRawSettings'] } else { @{} }
    if (-not $rawSettingsMap.ContainsKey($matchedProfileName)) {
        throw "Raw settings were not captured for customer policy '$matchedProfileName'. Run the assessment again and retry."
    }

    $sourceSettings = @($rawSettingsMap[$matchedProfileName])
    $baselineIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($setting in @(Get-SafeObjectPropertyValue -Object $Baseline -Name 'Settings')) {
        $defId = [string](Get-SafeObjectPropertyValue -Object $setting -Name 'DefinitionId')
        if ($defId) {
            $normId = Normalize-CompareText -Text $defId
            if ($normId) { [void]$baselineIds.Add($normId) }
        }
    }

    # Exclude settings that belong to ANY baseline file (not just this one), so the
    # resulting policy contains only genuinely customer-specific settings.
    $excludedIds = [System.Collections.Generic.HashSet[string]]::new($baselineIds, [System.StringComparer]::OrdinalIgnoreCase)
    if ($AssessmentData.ContainsKey('AllBaselineNormIds') -and $AssessmentData['AllBaselineNormIds']) {
        foreach ($nid in $AssessmentData['AllBaselineNormIds']) {
            if ($nid) { [void]$excludedIds.Add([string]$nid) }
        }
    }

    $extraDefinitionIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($setting in $sourceSettings) {
        $instance = Get-SafeObjectPropertyValue -Object $setting -Name 'settingInstance'
        if (-not $instance) { continue }

        $leafItems = @(Get-BaselineJsonLeafInstances -Instance $instance)
        if ($leafItems.Count -eq 0) {
            $leafItems = @([pscustomobject]@{ DefinitionId = (Get-SafeObjectPropertyValue -Object $instance -Name 'settingDefinitionId'); RawValue = '-' })
        }

        foreach ($leaf in $leafItems) {
            $leafDefId = [string]$leaf.DefinitionId
            if (-not $leafDefId) { continue }
            $normLeafId = Normalize-CompareText -Text $leafDefId
            if ($normLeafId -and -not $excludedIds.Contains($normLeafId)) {
                [void]$extraDefinitionIds.Add($normLeafId)
            }
        }
    }

    $graphSettings = [System.Collections.Generic.List[object]]::new()
    foreach ($setting in $sourceSettings) {
        $instance = Get-SafeObjectPropertyValue -Object $setting -Name 'settingInstance'
        if (-not $instance) { continue }

        $prunedInstance = Remove-BaselineDefinitionsFromSettingInstance -Instance $instance -ExcludedDefinitionIds $excludedIds
        if (-not $prunedInstance) { continue }

        [void]$graphSettings.Add(@{
            '@odata.type'   = '#microsoft.graph.deviceManagementConfigurationSetting'
            settingInstance = $prunedInstance
        })
    }

    if ($graphSettings.Count -eq 0) {
        throw 'No extra customer settings remained after pruning baseline-matching settings.'
    }

    $customerMeta = if ($AssessmentData.ContainsKey('CustomerPolicyMeta') -and $AssessmentData['CustomerPolicyMeta'].ContainsKey($matchedProfileName)) {
        $AssessmentData['CustomerPolicyMeta'][$matchedProfileName]
    } else { $null }
    $platform = if ($customerMeta -and $customerMeta.Platforms) { [string]$customerMeta.Platforms } else { 'windows10' }

    $newPolicyName = Get-BaselinePolicyNameWithoutSecurityBaseline -PolicyName $baselinePolicyName
    if (-not $newPolicyName) {
        throw "Could not derive a new policy name from baseline '$baselinePolicyName'."
    }

    $existingPolicies = @(Invoke-GraphPagedGet -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$select=id,name')
    if (@($existingPolicies | Where-Object { [string](Get-SafeObjectPropertyValue -Object $_ -Name 'name') -ieq $newPolicyName }).Count -gt 0) {
        throw "A configuration policy named '$newPolicyName' already exists."
    }

    $description = "Created from baseline '$baselinePolicyName' by splitting extra settings from '$matchedProfileName'."
    $body = @{
        name              = $newPolicyName
        description       = $description
        platforms         = $platform
        technologies      = 'mdm'
        templateReference = @{ templateId = '' }
        settings          = @($graphSettings)
    }

    $jsonBody = $body | ConvertTo-Json -Depth 30
    $created = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies' -Body $jsonBody -ContentType 'application/json' -ErrorAction Stop

    return [pscustomobject]@{
        Id                  = [string](Get-SafeObjectPropertyValue -Object $created -Name 'id')
        Name                = $newPolicyName
        SourcePolicyName    = $matchedProfileName
        BaselinePolicyName  = $baselinePolicyName
        ExtraSettingCount   = $extraDefinitionIds.Count
        CreatedPolicy       = $created
    }
}

function Get-BaselineAssessmentArea {
    param([object]$Baseline)

    $platform = [string]$Baseline.PolicyPlatform
    if ($platform -eq 'iOS' -or $platform -eq 'Android') { return $platform }
    if ([string]$Baseline.FileName -match '-SecurityBaseline-([A-Za-z0-9]+)') { return [string]$matches[1] }
    return 'Other'
}

function New-BaselineAssessmentHtmlReport {
    <#
    .SYNOPSIS
    Generates an HTML report showing per-baseline-JSON import readiness assessment.
    #>
    param(
        [string]$OutputPath,
        [hashtable]$AssessmentData,
        [hashtable]$DisplayNames = @{},
        [hashtable]$OptionDisplayNames = @{},
        [object[]]$SettingRows = @(),
        [string]$TenantName = '',
        [hashtable]$ProfileAssignments = @{}
    )

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $baselines = $AssessmentData.Baselines
    $splits = $AssessmentData.CustomerPolicySplit

    $readyCount = @($baselines | Where-Object { $_.Classification -eq 'ReadyToImport' }).Count
    $fullyAlignedCount = @($baselines | Where-Object { $_.Classification -eq 'FullyAligned' }).Count
    $consolidationCount = @($baselines | Where-Object { $_.Classification -eq 'ConsolidationNeeded' }).Count
    $splitRequiredCount = @($baselines | Where-Object { $_.Classification -eq 'SplitRequired' }).Count
    $renameCount = @($baselines | Where-Object { $_.Classification -eq 'RenameOnly' }).Count
    $conflictCount = @($baselines | Where-Object { $_.Classification -eq 'Conflicts' }).Count
    $errorCount = @($baselines | Where-Object { $_.Classification -eq 'Error' }).Count
    $splitCount = @($splits).Count

    # Group mobile baselines by platform and retain filename-based groups for other baselines.
    $areaMap = [System.Collections.Generic.Dictionary[string, hashtable]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($b in ($baselines | Where-Object { $_.Classification -ne 'Error' })) {
        $area = Get-BaselineAssessmentArea -Baseline $b
        if (-not $areaMap.ContainsKey($area)) {
            $areaMap[$area] = @{ Total = 0; Match = 0; Mismatch = 0; NotFound = 0; Files = 0 }
        }
        $areaMap[$area].Total    += [int]$b.TotalSettings
        $areaMap[$area].Match    += [int]$b.MatchCount
        $areaMap[$area].Mismatch += [int]$b.MismatchCount
        $areaMap[$area].NotFound += [int]$b.NotFoundCount
        $areaMap[$area].Files    += 1
    }
    $areaCoverageHtml = [System.Text.StringBuilder]::new()
    if ($areaMap.Count -gt 0) {
        # Compute overall totals across all areas
        $overallTotal = 0; $overallMatch = 0; $overallMismatch = 0; $overallNotFound = 0; $overallFiles = 0
        foreach ($a in $areaMap.Values) {
            $overallTotal    += [int]$a.Total
            $overallMatch    += [int]$a.Match
            $overallMismatch += [int]$a.Mismatch
            $overallNotFound += [int]$a.NotFound
            $overallFiles    += [int]$a.Files
        }
        $overallPresencePct   = if ($overallTotal -gt 0) { [math]::Round(($overallTotal - $overallNotFound) / $overallTotal * 100) } else { 0 }
        $overallCompliancePct = if ($overallTotal -gt 0) { [math]::Round($overallMatch / $overallTotal * 100) } else { 0 }
        $overallPresenceColor   = if ($overallPresencePct   -ge 80) { 'var(--green)' } elseif ($overallPresencePct   -ge 50) { 'var(--yellow)' } else { 'var(--red)' }
        $overallComplianceColor = if ($overallCompliancePct -ge 80) { 'var(--green)' } elseif ($overallCompliancePct -ge 50) { 'var(--yellow)' } else { 'var(--red)' }
        $overallPresentCount = $overallTotal - $overallNotFound

        [void]$areaCoverageHtml.AppendLine("<section id='sec-area'>")
        [void]$areaCoverageHtml.AppendLine("<h2 class='section-heading area-coverage' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x1F4CA; Area Coverage Overview ($($areaMap.Count) area(s))</h2>")
        [void]$areaCoverageHtml.AppendLine("<div class='top-section-body open'>")
        [void]$areaCoverageHtml.AppendLine("<p class='section-desc'>Presence: % of baseline settings that exist anywhere in the customer environment. Compliance: % with the correct expected value.</p>")

        # Overall summary card (full width, above per-area cards)
        [void]$areaCoverageHtml.AppendLine("<div class='area-card area-card-overall'>")
        [void]$areaCoverageHtml.AppendLine("  <div class='area-name'>Overall</div>")
        [void]$areaCoverageHtml.AppendLine("  <div class='area-meta'>$overallFiles file(s) &bull; $overallTotal settings &bull; $($areaMap.Count) area(s)</div>")
        [void]$areaCoverageHtml.AppendLine("  <div class='area-bar-row'><span class='area-bar-lbl'>Presence</span><span class='area-pct' style='color:$overallPresenceColor'>$overallPresencePct%</span></div>")
        [void]$areaCoverageHtml.AppendLine("  <div class='area-bar-wrap'><div class='area-bar' style='width:$overallPresencePct%;background:$overallPresenceColor'></div></div>")
        [void]$areaCoverageHtml.AppendLine("  <div class='area-bar-row'><span class='area-bar-lbl'>Compliance</span><span class='area-pct' style='color:$overallComplianceColor'>$overallCompliancePct%</span></div>")
        [void]$areaCoverageHtml.AppendLine("  <div class='area-bar-wrap'><div class='area-bar' style='width:$overallCompliancePct%;background:$overallComplianceColor'></div></div>")
        [void]$areaCoverageHtml.AppendLine("  <div class='area-stat-row'><span title='Present in customer env'>&#x2705; $overallPresentCount present</span><span title='Correct value'>&#x1F3AF; $overallMatch compliant</span><span title='Wrong value'>&#x26A0;&#xFE0F; $overallMismatch mismatch</span><span title='Not found'>&#x274C; $overallNotFound missing</span></div>")
        [void]$areaCoverageHtml.AppendLine("</div>")

        [void]$areaCoverageHtml.AppendLine("<div class='area-grid'>")
        foreach ($areaName in ($areaMap.Keys | Sort-Object)) {
            $a = $areaMap[$areaName]
            $presencePct   = if ($a.Total -gt 0) { [math]::Round(($a.Total - $a.NotFound) / $a.Total * 100) } else { 0 }
            $compliancePct = if ($a.Total -gt 0) { [math]::Round($a.Match / $a.Total * 100) } else { 0 }
            $presenceColor   = if ($presencePct   -ge 80) { 'var(--green)' } elseif ($presencePct   -ge 50) { 'var(--yellow)' } else { 'var(--red)' }
            $complianceColor = if ($compliancePct -ge 80) { 'var(--green)' } elseif ($compliancePct -ge 50) { 'var(--yellow)' } else { 'var(--red)' }
            $areaEnc = [System.Net.WebUtility]::HtmlEncode($areaName)
            $presentCount = $a.Total - $a.NotFound
            [void]$areaCoverageHtml.AppendLine("  <div class='area-card' data-area='$areaEnc' onclick='filterByArea(this)' title='Click to filter report to $areaEnc policies'>")            
            [void]$areaCoverageHtml.AppendLine("    <div class='area-name'>$areaEnc</div>")
            [void]$areaCoverageHtml.AppendLine("    <div class='area-meta'>$($a.Files) file(s) &bull; $($a.Total) settings</div>")
            [void]$areaCoverageHtml.AppendLine("    <div class='area-bar-row'><span class='area-bar-lbl'>Presence</span><span class='area-pct' style='color:$presenceColor'>$presencePct%</span></div>")
            [void]$areaCoverageHtml.AppendLine("    <div class='area-bar-wrap'><div class='area-bar' style='width:$presencePct%;background:$presenceColor'></div></div>")
            [void]$areaCoverageHtml.AppendLine("    <div class='area-bar-row'><span class='area-bar-lbl'>Compliance</span><span class='area-pct' style='color:$complianceColor'>$compliancePct%</span></div>")
            [void]$areaCoverageHtml.AppendLine("    <div class='area-bar-wrap'><div class='area-bar' style='width:$compliancePct%;background:$complianceColor'></div></div>")
            [void]$areaCoverageHtml.AppendLine("    <div class='area-stat-row'><span title='Present in customer env'>&#x2705; $presentCount present</span><span title='Correct value'>&#x1F3AF; $($a.Match) compliant</span><span title='Wrong value'>&#x26A0;&#xFE0F; $($a.Mismatch) mismatch</span><span title='Not found'>&#x274C; $($a.NotFound) missing</span></div>")
            [void]$areaCoverageHtml.AppendLine("  </div>")
        }
        [void]$areaCoverageHtml.AppendLine("</div>")
        [void]$areaCoverageHtml.AppendLine("<div id='areaFilterBanner'><span id='areaFilterLabel'></span><button onclick='clearAreaFilter()'>&#x2715; Clear filter</button></div>")
        [void]$areaCoverageHtml.AppendLine("</div></section>")
    }

    $sectionsHtml = [System.Text.StringBuilder]::new()

    # Section 1: Ready to Import
    if ($readyCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-ready'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading ready' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x2705; Ready to Import ($readyCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<p class='section-desc'>These baselines have NO settings in the customer environment. They can be imported directly without conflicts.</p>")
        [void]$sectionsHtml.AppendLine("<table class='summary-table'><thead><tr><th>Baseline Policy</th><th>File</th><th>Settings</th><th>Detail</th></tr></thead><tbody>")
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'ReadyToImport' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $detail = [System.Net.WebUtility]::HtmlEncode([string]$b.Detail)
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<tr data-area='$bArea'><td><strong>$name</strong></td><td>$file</td><td>$($b.TotalSettings)</td><td>$detail</td></tr>")
        }
        [void]$sectionsHtml.AppendLine("</tbody></table>")

        # Expandable detail per ReadyToImport baseline showing new and compliant settings
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'ReadyToImport' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $newSettings = @($b.Settings | Where-Object { $_.Status -eq 'NotFound' })
            $compliantSettings = @($b.Settings | Where-Object { $_.Status -eq 'Match' })
            $newCount = $newSettings.Count
            $compliantCount = $compliantSettings.Count
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<section class='conflict-section' data-area='$bArea'>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span><span class='cname'>$name</span><span class='cmeta'>$file &bull; $newCount new &bull; $compliantCount already compliant</span></div>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-body'>")
            [void]$sectionsHtml.AppendLine("    <div class='conflict-filters'><label><input type='checkbox' class='chk-new-only' checked onchange='applyReadyFilter(this)'/> Show only new settings</label></div>")
            [void]$sectionsHtml.AppendLine("    <table class='settings-table'><thead><tr><th>Status</th><th>Setting</th><th>Expected Value</th><th>Definition Id</th></tr></thead><tbody>")

            # New settings first
            foreach ($s in $newSettings) {
                $defId = [System.Net.WebUtility]::HtmlEncode([string]$s.DefinitionId)
                $expectedRaw = [string]$s.ExpectedValue
                $expectedDisplay = if ($expectedRaw -and $OptionDisplayNames.ContainsKey($expectedRaw)) {
                    [System.Net.WebUtility]::HtmlEncode($OptionDisplayNames[$expectedRaw])
                } else { [System.Net.WebUtility]::HtmlEncode($expectedRaw) }
                $settingName = if ($s.DefinitionId -and $DisplayNames.ContainsKey([string]$s.DefinitionId)) {
                    [System.Net.WebUtility]::HtmlEncode($DisplayNames[[string]$s.DefinitionId])
                } elseif ($s.DefinitionId -match '_([^_]+)$') {
                    [System.Net.WebUtility]::HtmlEncode($matches[1])
                } else { $defId }
                [void]$sectionsHtml.AppendLine("      <tr class='row-new'><td class='st-new'>&#x1F195;</td><td>$settingName</td><td>$expectedDisplay</td><td>$defId</td></tr>")
            }

            # Compliant settings
            foreach ($s in $compliantSettings) {
                $defId = [System.Net.WebUtility]::HtmlEncode([string]$s.DefinitionId)
                $expectedRaw = [string]$s.ExpectedValue
                $expectedDisplay = if ($expectedRaw -and $OptionDisplayNames.ContainsKey($expectedRaw)) {
                    [System.Net.WebUtility]::HtmlEncode($OptionDisplayNames[$expectedRaw])
                } else { [System.Net.WebUtility]::HtmlEncode($expectedRaw) }
                $settingName = if ($s.DefinitionId -and $DisplayNames.ContainsKey([string]$s.DefinitionId)) {
                    [System.Net.WebUtility]::HtmlEncode($DisplayNames[[string]$s.DefinitionId])
                } elseif ($s.FoundIn.Count -gt 0 -and $s.FoundIn[0].SettingName) {
                    [System.Net.WebUtility]::HtmlEncode([string]$s.FoundIn[0].SettingName)
                } elseif ($s.DefinitionId -match '_([^_]+)$') {
                    [System.Net.WebUtility]::HtmlEncode($matches[1])
                } else { $defId }
                $foundInProfile = if ($s.FoundIn.Count -gt 0) { [System.Net.WebUtility]::HtmlEncode([string]$s.FoundIn[0].ProfileName) } else { '' }
                $scopeBadge = if ($s.ScopeAlternate -and $s.FoundIn.Count -gt 0) { " <em>($([System.Net.WebUtility]::HtmlEncode([string]$s.FoundIn[0].ScopeNote)); counted as compliant)</em>" } else { '' }
                [void]$sectionsHtml.AppendLine("      <tr><td class='st-match'>&#x2705;</td><td>$settingName$scopeBadge</td><td>$expectedDisplay</td><td>$defId</td></tr>")
            }

            [void]$sectionsHtml.AppendLine("    </tbody></table>")
            [void]$sectionsHtml.AppendLine("  </div>")
            [void]$sectionsHtml.AppendLine("</section>")
        }

        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Section 2: Fully Aligned
    if ($fullyAlignedCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-aligned'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading aligned' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x2705; Fully Aligned ($fullyAlignedCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<p class='section-desc'>Policy name and all settings already match. No action needed.</p>")
        [void]$sectionsHtml.AppendLine("<table class='summary-table'><thead><tr><th>Baseline Policy</th><th>File</th><th>Settings</th><th>Detail</th></tr></thead><tbody>")
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'FullyAligned' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $detail = [System.Net.WebUtility]::HtmlEncode([string]$b.Detail)
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<tr data-area='$bArea'><td><strong>$name</strong></td><td>$file</td><td>$($b.TotalSettings)</td><td>$detail</td></tr>")
        }
        [void]$sectionsHtml.AppendLine("</tbody></table>")
        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Section 3: Consolidation Needed
    if ($consolidationCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-consolidation'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading consolidation' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x1F4E6; Consolidation Needed ($consolidationCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<p class='section-desc'>All baseline settings have correct values but are spread across multiple customer policies. Consider consolidating into a single policy.</p>")
        [void]$sectionsHtml.AppendLine("<table class='summary-table'><thead><tr><th>Baseline Policy</th><th>File</th><th>Settings</th><th>Spread Across</th><th>Detail</th></tr></thead><tbody>")
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'ConsolidationNeeded' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $detail = [System.Net.WebUtility]::HtmlEncode([string]$b.Detail)
            $profileCount = if ($b.FoundInProfiles) { $b.FoundInProfiles.Keys.Count } else { 0 }
            $spreadLabel = "$profileCount policies"
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<tr data-area='$bArea'><td><strong>$name</strong></td><td>$file</td><td>$($b.TotalSettings)</td><td>$spreadLabel</td><td>$detail</td></tr>")
        }
        [void]$sectionsHtml.AppendLine("</tbody></table>")
        # Expandable detail per baseline showing which policies contain which settings
        $custPolSettings = if ($AssessmentData.ContainsKey('CustomerPolicySettings')) { $AssessmentData['CustomerPolicySettings'] } else { @{} }
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'ConsolidationNeeded' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $profileCount = $b.FoundInProfiles.Keys.Count
            $baselineTotal = [int]$b.TotalSettings
            $profileListHtml = ($b.FoundInProfiles.Keys | Sort-Object | ForEach-Object {
                $pn = [System.Net.WebUtility]::HtmlEncode($_)
                $settingsInPolicy = $b.FoundInProfiles[$_].Count
                $custTotal = if ($custPolSettings.ContainsKey($_)) { $custPolSettings[$_].Count } else { 0 }
                $hasFullBaseline = ($settingsInPolicy -ge $baselineTotal)
                $isExactMatch = ($hasFullBaseline -and $custTotal -gt 0 -and $custTotal -le $baselineTotal)
                $badge = if ($isExactMatch) { " <span class='badge-exact'>Exact match</span>" }
                         elseif ($hasFullBaseline) { " <span class='badge-full'>Contains full baseline</span>" }
                         else { '' }
                $rowClass = if ($isExactMatch) { 'row-exact-match' } elseif ($hasFullBaseline) { 'row-full-baseline' } else { '' }
                $totalInfo = if ($custTotal -gt 0) { " ($custTotal total in policy)" } else { '' }
                $unassignedBadge = if ($ProfileAssignments.ContainsKey($_) -and -not $ProfileAssignments[$_]) { " <span class='badge-unassigned'>&#x26A0; Not assigned</span>" } else { '' }
                "<tr$(if ($rowClass) { " class='$rowClass'" })><td>$pn$badge$unassignedBadge</td><td>$settingsInPolicy of $baselineTotal baseline settings$totalInfo</td></tr>"
            }) -join "`n"
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<section class='conflict-section' data-area='$bArea'>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span><span class='cname'>$name</span><span class='cmeta'>$file &bull; $($b.TotalSettings) settings &bull; $profileCount policies</span></div>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-body'>")
            [void]$sectionsHtml.AppendLine("    <table class='settings-table'><thead><tr><th>Customer Policy</th><th>Settings</th></tr></thead><tbody>")
            [void]$sectionsHtml.AppendLine($profileListHtml)
            [void]$sectionsHtml.AppendLine("    </tbody></table>")
            [void]$sectionsHtml.AppendLine("  </div>")
            [void]$sectionsHtml.AppendLine("</section>")
        }
        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Section 4: Split Required
    if ($splitRequiredCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-splitreq'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading splitreq' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x2702;&#xFE0F; Compliant &#x2014; Split Required ($splitRequiredCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<p class='section-desc'>All baseline settings have correct values in a single customer policy, but that policy also contains customer-only settings (present in NO baseline file) that should be separated out before importing. The <strong>Customer-Only Settings</strong> count is the exact scope of the Create Split Policy action.</p>")
        [void]$sectionsHtml.AppendLine("<table class='summary-table'><thead><tr><th>Baseline Policy</th><th>File</th><th>Settings</th><th>Customer Policy</th><th>Customer-Only Settings</th><th>Detail</th></tr></thead><tbody>")
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'SplitRequired' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $detail = [System.Net.WebUtility]::HtmlEncode([string]$b.Detail)
            $extraCount = if ($b.ContainsKey('CustomerOnlyCount')) { [string]$b.CustomerOnlyCount } else { '0' }
            $mapsTo = if ($b.FoundInProfiles -and $b.FoundInProfiles.Keys.Count -gt 0) {
                $mapsToKey = $b.FoundInProfiles.Keys | Select-Object -First 1
                $unaBadge = if ($ProfileAssignments.ContainsKey($mapsToKey) -and -not $ProfileAssignments[$mapsToKey]) { " <span class='badge-unassigned'>&#x26A0; Not assigned</span>" } else { '' }
                [System.Net.WebUtility]::HtmlEncode($mapsToKey) + $unaBadge
            } else { '-' }
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<tr data-area='$bArea'><td><strong>$name</strong></td><td>$file</td><td>$($b.TotalSettings)</td><td>$mapsTo</td><td>$extraCount</td><td>$detail</td></tr>")
        }
        [void]$sectionsHtml.AppendLine("</tbody></table>")

        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'SplitRequired' } | Sort-Object { $_.PolicyName })) {
            if (-not ($b.FoundInProfiles -and $b.FoundInProfiles.Keys.Count -gt 0)) { continue }

            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $mapsToKey = $b.FoundInProfiles.Keys | Select-Object -First 1
            $mapsToHtml = [System.Net.WebUtility]::HtmlEncode($mapsToKey)
            $bArea = Get-BaselineAssessmentArea -Baseline $b

            # Only show TRUE customer-only settings (present in NO baseline file).
            $customerOnlyIds = if ($b.ContainsKey('CustomerOnlyIds')) { $b.CustomerOnlyIds } else { $null }
            $seenExtraIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $extraRows = @($SettingRows | Where-Object {
                [string]$_.ProfileName -ieq $mapsToKey -and
                $_.DefinitionId -and
                [string]$_.DefinitionId -ne '-' -and
                $customerOnlyIds -and
                $customerOnlyIds.Contains((Normalize-CompareText -Text ([string]$_.DefinitionId))) -and
                $seenExtraIds.Add((Normalize-CompareText -Text ([string]$_.DefinitionId)))
            })

            if ($extraRows.Count -eq 0) { continue }

            [void]$sectionsHtml.AppendLine("<section class='conflict-section' data-area='$bArea'>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span><span class='cname'>$name</span><span class='cmeta'>$file &bull; $mapsToHtml &bull; $($extraRows.Count) customer-only setting(s)</span></div>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-body'>")
            [void]$sectionsHtml.AppendLine("    <table class='settings-table'><thead><tr><th>Customer-Only Setting</th><th>Definition ID</th><th>Current Value</th></tr></thead><tbody>")
            foreach ($row in ($extraRows | Sort-Object { if ($_.SettingName) { [string]$_.SettingName } else { [string]$_.DefinitionId } })) {
                $rowDefId = [System.Net.WebUtility]::HtmlEncode([string]$row.DefinitionId)
                $rowName = if ($DisplayNames.ContainsKey([string]$row.DefinitionId)) {
                    [System.Net.WebUtility]::HtmlEncode($DisplayNames[[string]$row.DefinitionId])
                } elseif ($row.SettingName) {
                    [System.Net.WebUtility]::HtmlEncode([string]$row.SettingName)
                } else {
                    $rowDefId
                }
                $rowValueRaw = if ($null -ne $row.RawValue) { [string]$row.RawValue } else { '' }
                $rowValue = if ($rowValueRaw -and $OptionDisplayNames.ContainsKey($rowValueRaw)) {
                    [System.Net.WebUtility]::HtmlEncode($OptionDisplayNames[$rowValueRaw])
                } else {
                    [System.Net.WebUtility]::HtmlEncode($rowValueRaw)
                }
                [void]$sectionsHtml.AppendLine("      <tr><td>$rowName</td><td>$rowDefId</td><td>$rowValue</td></tr>")
            }
            [void]$sectionsHtml.AppendLine("    </tbody></table>")
            [void]$sectionsHtml.AppendLine("  </div>")
            [void]$sectionsHtml.AppendLine("</section>")
        }
        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Section 3: Rename Only
    if ($renameCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-onetone'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading onetone' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x1F504; Rename Only ($renameCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<p class='section-desc'>All baseline settings are found in exactly one customer policy with matching values, but the policy has a different name. The customer can rename their existing policy to match the baseline.</p>")
        [void]$sectionsHtml.AppendLine("<table class='summary-table'><thead><tr><th>Baseline Policy</th><th>File</th><th>Settings</th><th>Customer Policy Name</th><th>Detail</th></tr></thead><tbody>")
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'RenameOnly' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $detail = [System.Net.WebUtility]::HtmlEncode([string]$b.Detail)
            $mapsTo = if ($b.FoundInProfiles -and $b.FoundInProfiles.Keys.Count -gt 0) {
                $mapsToKey = $b.FoundInProfiles.Keys | Select-Object -First 1
                $unaBadge = if ($ProfileAssignments.ContainsKey($mapsToKey) -and -not $ProfileAssignments[$mapsToKey]) { " <span class='badge-unassigned'>&#x26A0; Not assigned</span>" } else { '' }
                [System.Net.WebUtility]::HtmlEncode($mapsToKey) + $unaBadge
            } else { '-' }
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<tr data-area='$bArea'><td><strong>$name</strong></td><td>$file</td><td>$($b.TotalSettings)</td><td>$mapsTo</td><td>$detail</td></tr>")
        }
        [void]$sectionsHtml.AppendLine("</tbody></table>")
        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Section 4: Conflicts / Discussion Needed
    if ($conflictCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-conflicts'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading conflicts' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x26A0;&#xFE0F; Conflicts / Discussion Needed ($conflictCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<p class='section-desc'>These baselines have settings with different values in the customer environment, or settings spread across multiple policies with mismatches. Requires discussion before import.</p>")

        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'Conflicts' } | Sort-Object { $_.PolicyName })) {
            $name = [System.Net.WebUtility]::HtmlEncode([string]$b.PolicyName)
            $file = [System.Net.WebUtility]::HtmlEncode([string]$b.FileName)
            $detail = [System.Net.WebUtility]::HtmlEncode([string]$b.Detail)

            # Collect unique policy names for the dropdown
            $policyOptions = [System.Collections.Generic.List[string]]::new()
            foreach ($s in $b.Settings) {
                if ($s.FoundIn -and $s.FoundIn.Count -gt 0) {
                    foreach ($hit in $s.FoundIn) {
                        $pName = [string]$hit.ProfileName
                        if ($pName -and -not $policyOptions.Contains($pName)) { [void]$policyOptions.Add($pName) }
                    }
                }
            }
            $policyOptions.Sort()
            $policyOptionsHtml = ($policyOptions | ForEach-Object { "<option value='$([System.Net.WebUtility]::HtmlEncode($_))'>$([System.Net.WebUtility]::HtmlEncode($_))</option>" }) -join ''

            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<section class='conflict-section' data-area='$bArea'>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span><span class='cname'>$name</span><span class='cmeta'>$file &bull; $($b.TotalSettings) settings &bull; $($b.MatchCount) match &bull; $($b.MismatchCount) mismatch &bull; $($b.NotFoundCount) not found</span></div>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-body'>")
            [void]$sectionsHtml.AppendLine("    <div class='conflict-filters'><label><input type='checkbox' class='chk-hide-match' checked onchange='applyConflictFilter(this)'/> Hide matching settings</label><select class='sel-policy' onchange='applyConflictFilter(this)'><option value=''>All Policies</option>$policyOptionsHtml</select></div>")
            $conflictCustPolSettings = if ($AssessmentData.ContainsKey('CustomerPolicySettings')) { $AssessmentData['CustomerPolicySettings'] } else { @{} }
            $conflictBaselineTotal = [int]$b.TotalSettings
            [void]$sectionsHtml.AppendLine("    <table class='settings-table policy-overview'><thead><tr><th>Customer Policy</th><th>Coverage</th><th>Matches</th><th>Mismatches</th><th>Status</th></tr></thead><tbody>")
            foreach ($polName in ($b.FoundInProfiles.Keys | Sort-Object)) {
                $polHits = $b.FoundInProfiles[$polName]
                $polMatchCount = @($polHits | Where-Object { $_.Status -eq 'Match' }).Count
                $polMismatchCount = @($polHits | Where-Object { $_.Status -eq 'Mismatch' }).Count
                $polSettingsCount = $polHits.Count
                $custPolTotal = if ($conflictCustPolSettings.ContainsKey($polName)) { $conflictCustPolSettings[$polName].Count } else { 0 }
                $totalInfo = if ($custPolTotal -gt 0) { " ($custPolTotal total)" } else { '' }
                $isFullyCompliant = ($polSettingsCount -ge $conflictBaselineTotal -and $polMismatchCount -eq 0)
                $conflictStatusHtml = if ($isFullyCompliant) { "<span class='badge-exact'>All settings compliant</span>" } elseif ($polMismatchCount -gt 0) { "<span class='badge-conflict-mismatch'>$polMismatchCount mismatch(es)</span>" } else { "$polMatchCount of $conflictBaselineTotal matched" }
                $conflictRowClass = if ($isFullyCompliant) { " class='row-exact-match'" } else { '' }
                $unaBadge2 = if ($ProfileAssignments.ContainsKey($polName) -and -not $ProfileAssignments[$polName]) { " <span class='badge-unassigned'>&#x26A0; Not assigned</span>" } else { '' }
                $polNameHtml = [System.Net.WebUtility]::HtmlEncode($polName) + $unaBadge2
                [void]$sectionsHtml.AppendLine("      <tr$conflictRowClass><td>$polNameHtml</td><td>$polSettingsCount of $conflictBaselineTotal$totalInfo</td><td>$polMatchCount</td><td>$polMismatchCount</td><td>$conflictStatusHtml</td></tr>")
            }
            [void]$sectionsHtml.AppendLine("    </tbody></table>")
            [void]$sectionsHtml.AppendLine("    <table class='settings-table'><thead><tr><th>Status</th><th>Setting</th><th>Expected Value</th><th>Found In Policy</th><th>Actual Value</th><th>Definition Id</th></tr></thead><tbody>")

            foreach ($s in $b.Settings) {
                $defId = [System.Net.WebUtility]::HtmlEncode([string]$s.DefinitionId)
                $expectedRaw = [string]$s.ExpectedValue
                $expectedDisplay = if ($expectedRaw -and $OptionDisplayNames.ContainsKey($expectedRaw)) {
                    [System.Net.WebUtility]::HtmlEncode($OptionDisplayNames[$expectedRaw])
                } else { [System.Net.WebUtility]::HtmlEncode($expectedRaw) }
                $settingName = if ($s.DefinitionId -and $DisplayNames.ContainsKey([string]$s.DefinitionId)) {
                    [System.Net.WebUtility]::HtmlEncode($DisplayNames[[string]$s.DefinitionId])
                } elseif ($s.FoundIn.Count -gt 0 -and $s.FoundIn[0].SettingName) {
                    [System.Net.WebUtility]::HtmlEncode([string]$s.FoundIn[0].SettingName)
                } elseif ($s.DefinitionId -match '_([^_]+)$') {
                    [System.Net.WebUtility]::HtmlEncode($matches[1])
                } else { $defId }

                if ($s.Status -eq 'NotFound') {
                    [void]$sectionsHtml.AppendLine("      <tr class='row-missing'><td class='st-missing'>&#x274C;</td><td>$settingName</td><td>$expectedDisplay</td><td><em>Not found</em></td><td>-</td><td>$defId</td></tr>")
                } else {
                    foreach ($hit in $s.FoundIn) {
                        $profHtml = [System.Net.WebUtility]::HtmlEncode([string]$hit.ProfileName)
                        $actualRaw = [string]$hit.ActualValue
                        $actualDisplay = if ($actualRaw -and $OptionDisplayNames.ContainsKey($actualRaw)) {
                            [System.Net.WebUtility]::HtmlEncode($OptionDisplayNames[$actualRaw])
                        } else { [System.Net.WebUtility]::HtmlEncode($actualRaw) }
                        $hitScopeBadge = if ($hit.ScopeNote) { " <em>($([System.Net.WebUtility]::HtmlEncode([string]$hit.ScopeNote)))</em>" } else { '' }
                        if ($hit.Status -eq 'Match') {
                            [void]$sectionsHtml.AppendLine("      <tr><td class='st-match'>&#x2705;</td><td>$settingName$hitScopeBadge</td><td>$expectedDisplay</td><td>$profHtml</td><td>$actualDisplay</td><td>$defId</td></tr>")
                        } else {
                            [void]$sectionsHtml.AppendLine("      <tr class='row-mismatch'><td class='st-mismatch'>&#x26A0;&#xFE0F;</td><td>$settingName$hitScopeBadge</td><td>$expectedDisplay</td><td>$profHtml</td><td>$actualDisplay</td><td>$defId</td></tr>")
                        }
                    }
                }
            }
            [void]$sectionsHtml.AppendLine("    </tbody></table>")
            [void]$sectionsHtml.AppendLine("  </div>")
            [void]$sectionsHtml.AppendLine("</section>")
        }
        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Section 5: Customer Policies to Split
    if ($splitCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-split'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading split' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x2702;&#xFE0F; Customer Policies to Split ($splitCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<p class='section-desc'>These customer policies contain a mix of baseline settings and customer-only settings. The customer-only settings should be separated into their own policy before importing the baseline.</p>")
        [void]$sectionsHtml.AppendLine("<table class='summary-table'><thead><tr><th>Customer Policy</th><th>Total Settings</th><th>Baseline Settings</th><th>Customer-Only Settings</th><th>Action</th></tr></thead><tbody>")

        # Build a reverse map from NormId to SettingName for display
        $normIdToName = @{}
        foreach ($row in $SettingRows) {
            $nid = Normalize-CompareText -Text ([string]$row.DefinitionId)
            if ($nid -and -not $normIdToName.ContainsKey($nid)) {
                $normIdToName[$nid] = [string]$row.SettingName
            }
        }

        foreach ($sp in ($splits | Sort-Object { $_.ProfileName })) {
            $profHtml = [System.Net.WebUtility]::HtmlEncode([string]$sp.ProfileName)
            $action = "Extract $($sp.CustomerOnly) customer-only setting(s) to a separate policy"
            [void]$sectionsHtml.AppendLine("<tr><td><strong>$profHtml</strong></td><td>$($sp.TotalSettings)</td><td>$($sp.BaselineSettings)</td><td>$($sp.CustomerOnly)</td><td>$([System.Net.WebUtility]::HtmlEncode($action))</td></tr>")
        }
        [void]$sectionsHtml.AppendLine("</tbody></table>")

        # Expandable detail per policy showing customer-only settings
        foreach ($sp in ($splits | Sort-Object { $_.ProfileName })) {
            $profHtml = [System.Net.WebUtility]::HtmlEncode([string]$sp.ProfileName)
            [void]$sectionsHtml.AppendLine("<section class='conflict-section'>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-header' onclick='toggleSection(this)'><span class='arrow'>&#x25B6;</span><span class='cname'>$profHtml</span><span class='cmeta'>$($sp.CustomerOnly) customer-only settings to extract</span></div>")
            [void]$sectionsHtml.AppendLine("  <div class='conflict-body'>")
            [void]$sectionsHtml.AppendLine("    <table class='settings-table'><thead><tr><th>Setting Name</th><th>Definition Id</th></tr></thead><tbody>")
            foreach ($nid in $sp.CustomerOnlyIds) {
                $friendlyName = if ($normIdToName.ContainsKey($nid)) { [System.Net.WebUtility]::HtmlEncode($normIdToName[$nid]) }
                                elseif ($DisplayNames.ContainsKey($nid)) { [System.Net.WebUtility]::HtmlEncode($DisplayNames[$nid]) }
                                else { [System.Net.WebUtility]::HtmlEncode($nid) }
                [void]$sectionsHtml.AppendLine("      <tr><td>$friendlyName</td><td>$([System.Net.WebUtility]::HtmlEncode($nid))</td></tr>")
            }
            [void]$sectionsHtml.AppendLine("    </tbody></table>")
            [void]$sectionsHtml.AppendLine("  </div>")
            [void]$sectionsHtml.AppendLine("</section>")
        }
        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Errors
    if ($errorCount -gt 0) {
        [void]$sectionsHtml.AppendLine("<section id='sec-error'>")
        [void]$sectionsHtml.AppendLine("<h2 class='section-heading error' onclick='toggleTopSection(this)'><span class='ts-arrow'>&#x25B6;</span>&#x274C; Errors ($errorCount)</h2>")
        [void]$sectionsHtml.AppendLine("<div class='top-section-body'>")
        [void]$sectionsHtml.AppendLine("<table class='summary-table'><thead><tr><th>File</th><th>Error</th></tr></thead><tbody>")
        foreach ($b in ($baselines | Where-Object { $_.Classification -eq 'Error' })) {
            $bArea = Get-BaselineAssessmentArea -Baseline $b
            [void]$sectionsHtml.AppendLine("<tr data-area='$bArea'><td>$([System.Net.WebUtility]::HtmlEncode($b.FileName))</td><td>$([System.Net.WebUtility]::HtmlEncode($b.Detail))</td></tr>")
        }
        [void]$sectionsHtml.AppendLine("</tbody></table>")
        [void]$sectionsHtml.AppendLine("</div></section>")
    }

    # Build JS data arrays for CSV export — ConvertTo-Json handles all escaping correctly
    $summaryArr = @($baselines | ForEach-Object {
        $sp = if ($_.FoundInProfiles) { [int]$_.FoundInProfiles.Keys.Count } else { 0 }
        $ua = if ($_.FoundInProfiles) { ($_.FoundInProfiles.Keys | Where-Object { $ProfileAssignments.ContainsKey($_) -and -not $ProfileAssignments[$_] }) -join '; ' } else { '' }
        [pscustomobject]@{ c=[string]$_.Classification; p=[string]$_.PolicyName; f=[string]$_.FileName
                           t=[int]$_.TotalSettings; m=[int]$_.MatchCount; x=[int]$_.MismatchCount
                           n=[int]$_.NotFoundCount; s=$sp; d=[string]$_.Detail; ua=[string]$ua }
    })
    $detailList = [System.Collections.Generic.List[object]]::new()
    foreach ($bx in $baselines) {
        if ($bx.Classification -eq 'Error') { continue }
        $clsx  = [string]$bx.Classification
        $bPolx = [string]$bx.PolicyName
        $bFilex= [string]$bx.FileName
        foreach ($sx in $bx.Settings) {
            $defIdx = [string]$sx.DefinitionId
            $snx  = if ($defIdx -and $DisplayNames.ContainsKey($defIdx)) { $DisplayNames[$defIdx] } else { $defIdx }
            $sdx = [int]$sx.Depth
            $evRaw= [string]$sx.ExpectedValue
            $evx  = if ($evRaw -and $OptionDisplayNames.ContainsKey($evRaw)) { $OptionDisplayNames[$evRaw] } else { $evRaw }
            if ($sx.Status -eq 'NotFound') {
                [void]$detailList.Add([pscustomobject]@{ c=$clsx; bp=$bPolx; bf=$bFilex; sn=$snx; sd=$sdx; di=$defIdx; ev=$evx; st='NotFound'; fp=''; av=''; pa='' })
            } else {
                foreach ($hitx in $sx.FoundIn) {
                    $avRaw = [string]$hitx.ActualValue
                    $avx   = if ($avRaw -and $OptionDisplayNames.ContainsKey($avRaw)) { $OptionDisplayNames[$avRaw] } else { $avRaw }
                    $pax   = if ($ProfileAssignments.ContainsKey([string]$hitx.ProfileName)) { if ($ProfileAssignments[[string]$hitx.ProfileName]) { 'Yes' } else { 'No' } } else { '' }
                    [void]$detailList.Add([pscustomobject]@{ c=$clsx; bp=$bPolx; bf=$bFilex; sn=$snx; sd=$sdx; di=$defIdx; ev=$evx; st=[string]$hitx.Status; fp=[string]$hitx.ProfileName; av=$avx; pa=$pax })
                }
            }
        }
    }
    $summaryJson = if ($summaryArr.Count -gt 0) { ConvertTo-Json -InputObject @($summaryArr) -Compress -Depth 2 } else { '[]' }
    $detailJson  = if ($detailList.Count  -gt 0) { ConvertTo-Json -InputObject @($detailList.ToArray()) -Compress -Depth 2 } else { '[]' }
    $jsDataBlock = "window._asmtSummary=$summaryJson;window._asmtDetail=$detailJson;"

    $html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Baseline Import Assessment Report</title>
<style>
:root{--bg:#0f172a;--fg:#e2e8f0;--surface:#1e293b;--muted:#94a3b8;--accent:#60a5fa;--border:#334155;--green:#22c55e;--yellow:#eab308;--red:#ef4444;--blue:#3b82f6}
*{box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--fg);margin:0;padding:20px}
h1{margin:0 0 4px 0;color:var(--accent)}
.subtitle{color:var(--muted);margin:0 0 16px 0;font-size:13px}
.stats{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.stat{background:var(--surface);border:1px solid var(--border);border-radius:6px;padding:10px 14px;min-width:100px}
.stat .v{font-size:22px;font-weight:700}
.stat .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.stat.ready .v{color:var(--green)}
.stat.onetone .v{color:var(--blue)}
.stat.aligned .v{color:#22c55e}
.stat.consolidation .v{color:#06b6d4}
.stat.splitreq .v{color:#a78bfa}
.stat.conflicts .v{color:var(--yellow)}
.stat.split .v{color:#f97316}
.section-heading{margin:24px 0 8px 0;font-size:1.1rem;padding:8px 12px;border-radius:6px;background:var(--surface);border:1px solid var(--border);cursor:pointer;user-select:none}
.section-heading:hover{opacity:.85}
.ts-arrow{font-size:.75rem;transition:transform .2s;display:inline-block;margin-right:6px;opacity:.6}
.section-heading.open .ts-arrow{transform:rotate(90deg)}
.top-section-body{display:none}
.top-section-body.open{display:block}
.stat[onclick]{cursor:pointer}
.stat[onclick]:hover{opacity:.82;transform:translateY(-2px);transition:all .14s}
.section-heading.ready{border-left:4px solid var(--green)}
.section-heading.aligned{border-left:4px solid #22c55e}
.section-heading.consolidation{border-left:4px solid #06b6d4}
.section-heading.splitreq{border-left:4px solid #a78bfa}
.section-heading.onetone{border-left:4px solid var(--blue)}
.section-heading.conflicts{border-left:4px solid var(--yellow)}
.section-heading.split{border-left:4px solid #f97316}
.section-heading.error{border-left:4px solid var(--red)}
.section-desc{color:var(--muted);font-size:13px;margin:0 0 10px 12px}
.summary-table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px}
.summary-table th{background:var(--accent);color:#fff;padding:8px;text-align:left}
.summary-table td{padding:7px 8px;border-bottom:1px solid var(--border);vertical-align:top}
.summary-table tr:hover td{background:rgba(96,165,250,.08)}
.conflict-section{background:var(--surface);border:1px solid var(--border);border-radius:8px;margin-bottom:8px;overflow:hidden}
.conflict-header{padding:10px 12px;display:flex;gap:10px;align-items:center;cursor:pointer;background:rgba(234,179,8,.08)}
.conflict-header .cname{font-weight:600;flex:1}
.conflict-header .cmeta{font-size:12px;color:var(--muted)}
.conflict-header .arrow{transition:transform .2s}
.conflict-header.open .arrow{transform:rotate(90deg)}
.conflict-body{display:none;padding:10px}
.conflict-body.open{display:block}
.conflict-filters{display:flex;gap:14px;align-items:center;margin-bottom:8px;flex-wrap:wrap}
.conflict-filters label{font-size:12px;cursor:pointer;display:flex;align-items:center;gap:4px}
.conflict-filters select{font-size:12px;padding:4px 8px;border:1px solid var(--border);border-radius:4px;background:var(--surface);color:var(--fg)}
.badge-full{display:inline-block;padding:1px 7px;border-radius:3px;font-size:.7rem;font-weight:600;background:rgba(96,165,250,.18);color:var(--blue);border:1px solid rgba(96,165,250,.35);margin-left:6px;vertical-align:middle}
.badge-exact{display:inline-block;padding:1px 7px;border-radius:3px;font-size:.7rem;font-weight:700;background:rgba(34,197,94,.18);color:var(--green);border:1px solid rgba(34,197,94,.35);margin-left:6px;vertical-align:middle}
.row-full-baseline{background:rgba(96,165,250,.06)}
.row-exact-match{background:rgba(34,197,94,.1)}
.badge-conflict-mismatch{display:inline-block;padding:1px 7px;border-radius:3px;font-size:.7rem;font-weight:600;background:rgba(234,179,8,.18);color:var(--yellow);border:1px solid rgba(234,179,8,.35);vertical-align:middle}
.badge-unassigned{display:inline-block;padding:1px 7px;border-radius:3px;font-size:.7rem;font-weight:600;background:rgba(251,146,60,.18);color:#fb923c;border:1px solid rgba(251,146,60,.4);vertical-align:middle;margin-left:4px}
.policy-overview{margin-bottom:10px}
.report-name{color:var(--fg);font-size:1rem;font-weight:400;margin:0 0 12px 0;opacity:.7}
.settings-table{width:100%;border-collapse:collapse;font-size:12px}
.settings-table th{background:var(--accent);color:#fff;padding:6px 8px;text-align:left}
.settings-table td{padding:5px 8px;border-bottom:1px solid var(--border);vertical-align:top}
.st-match,.st-mismatch,.st-missing,.st-new{text-align:center;font-size:14px;width:40px}
.row-mismatch{background:rgba(234,179,8,.1)}
.row-missing{background:rgba(239,68,68,.1)}
.row-missing td{font-style:italic;color:var(--muted)}
.row-new{background:rgba(96,165,250,.08)}
.row-new td{color:var(--fg)}
.toolbar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:14px}
.toolbar button{padding:7px 10px;border:1px solid var(--border);border-radius:4px;background:var(--surface);color:var(--fg);cursor:pointer}
.area-grid{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:16px}
.area-card{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:14px 16px;min-width:220px;flex:1 1 220px;max-width:340px}
.area-card-overall{max-width:none;flex:1 1 100%;margin-bottom:10px;border-color:var(--accent);background:rgba(96,165,250,.06)}
.area-name{font-size:1.1rem;font-weight:700;margin-bottom:2px}
.area-meta{font-size:11px;color:var(--muted);margin-bottom:10px}
.area-bar-row{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:3px}
.area-bar-lbl{font-size:11px;color:var(--muted)}
.area-pct{font-size:1rem;font-weight:700}
.area-bar-wrap{background:rgba(255,255,255,.08);border-radius:4px;height:8px;margin-bottom:8px;overflow:hidden}
.area-bar{height:8px;border-radius:4px}
.area-stat-row{display:flex;gap:8px;flex-wrap:wrap;margin-top:6px;font-size:11px;color:var(--muted)}
.section-heading.area-coverage{border-left:4px solid var(--accent)}
.area-card{cursor:pointer;transition:box-shadow .15s,border-color .15s}
.area-card:hover{box-shadow:0 0 0 2px var(--accent)}
.area-card.area-active{border-color:var(--accent);box-shadow:0 0 0 2px var(--accent)}
#areaFilterBanner{background:rgba(96,165,250,.12);border:1px solid var(--accent);border-radius:6px;padding:8px 14px;display:none;align-items:center;gap:12px;margin:8px 0 14px;font-size:13px}
#areaFilterBanner button{padding:3px 10px;border:1px solid var(--accent);border-radius:4px;background:transparent;color:var(--fg);cursor:pointer;font-size:12px}
</style>
</head>
<body>
$(if ($TenantName) { "<h1>$([System.Net.WebUtility]::HtmlEncode($TenantName))</h1><p class='report-name'>Baseline Import Assessment Report</p>" } else { "<h1>Baseline Import Assessment Report</h1>" })
<p class='subtitle'>Generated $generatedAt &bull; $($AssessmentData.TotalBaselineFiles) baseline file(s) assessed</p>

<div class='stats'>
  <div class='stat ready' onclick="jumpToSection('sec-ready')" title='Jump to Ready to Import'><div class='v'>$readyCount</div><div class='l'>Ready to Import</div></div>
  <div class='stat aligned' onclick="jumpToSection('sec-aligned')" title='Jump to Fully Aligned'><div class='v'>$fullyAlignedCount</div><div class='l'>Fully Aligned</div></div>
  <div class='stat consolidation' onclick="jumpToSection('sec-consolidation')" title='Jump to Consolidation Needed'><div class='v'>$consolidationCount</div><div class='l'>Consolidation Needed</div></div>
  <div class='stat splitreq' onclick="jumpToSection('sec-splitreq')" title='Jump to Split Required'><div class='v'>$splitRequiredCount</div><div class='l'>Split Required</div></div>
  <div class='stat onetone' onclick="jumpToSection('sec-onetone')" title='Jump to Rename Only'><div class='v'>$renameCount</div><div class='l'>Rename Only</div></div>
  <div class='stat conflicts' onclick="jumpToSection('sec-conflicts')" title='Jump to Conflicts'><div class='v'>$conflictCount</div><div class='l'>Conflicts</div></div>
  <div class='stat split' onclick="jumpToSection('sec-split')" title='Jump to Policies to Split'><div class='v'>$splitCount</div><div class='l'>Policies to Split</div></div>
</div>

<div class='toolbar'>
  <button onclick='expandAll()'>Expand All</button>
  <button onclick='collapseAll()'>Collapse All</button>
  <button onclick='exportSummaryCsv()'>Export Summary CSV</button>
  <button onclick='exportDetailCsv()'>Export Details CSV</button>
</div>

<script>$($jsDataBlock)</script>

$($areaCoverageHtml.ToString())

$($sectionsHtml.ToString())

<script>
function toggleSection(header){
  const body = header.parentElement.querySelector('.conflict-body');
  if(!body) return;
  const open = body.classList.toggle('open');
  header.classList.toggle('open', open);
}
function expandAll(){
  document.querySelectorAll('.conflict-header').forEach(h=>h.classList.add('open'));
  document.querySelectorAll('.conflict-body').forEach(b=>b.classList.add('open'));
  document.querySelectorAll('.section-heading').forEach(h=>h.classList.add('open'));
  document.querySelectorAll('.top-section-body').forEach(b=>b.classList.add('open'));
}
function collapseAll(){
  document.querySelectorAll('.conflict-header').forEach(h=>h.classList.remove('open'));
  document.querySelectorAll('.conflict-body').forEach(b=>b.classList.remove('open'));
  document.querySelectorAll('.section-heading').forEach(h=>h.classList.remove('open'));
  document.querySelectorAll('.top-section-body').forEach(b=>b.classList.remove('open'));
}
function toggleTopSection(heading){
  const body=heading.parentElement.querySelector('.top-section-body');
  if(!body)return;
  const open=body.classList.toggle('open');
  heading.classList.toggle('open',open);
}
function jumpToSection(id){
  const sec=document.getElementById(id);
  if(!sec)return;
  const body=sec.querySelector('.top-section-body');
  const heading=sec.querySelector('.section-heading');
  if(body&&!body.classList.contains('open')){body.classList.add('open');if(heading)heading.classList.add('open');}
  setTimeout(()=>sec.scrollIntoView({behavior:'smooth',block:'start'}),30);
}
function exportSummaryCsv(){
  const rows=window._asmtSummary||[];
  const h1=document.querySelector('h1');
  const t=h1?(h1.textContent||'').trim().replace(/[\\\/:*?"<>|]/g,''):'';
  const hdr='"Classification";"Baseline Policy";"File";"Total Settings";"Matches";"Mismatches";"Not Found";"Spread Across (Policies)";"Detail";"Unassigned Contributing Policies"';
  const lines=[hdr];
  rows.forEach(r=>{lines.push([r.c,r.p,r.f,r.t,r.m,r.x,r.n,r.s,r.d,r.ua].map(v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"').join(';'));});
  dlCsv(lines.join('\n'),(t?t+'_':'')+'Baseline-Assessment-Summary.csv');
}
function exportDetailCsv(){
  const rows=window._asmtDetail||[];
  const h1=document.querySelector('h1');
  const t=h1?(h1.textContent||'').trim().replace(/[\\\/:*?"<>|]/g,''):'';
    const hdr='"Classification";"Baseline Policy";"File";"Setting Name";"Setting Depth";"Definition Id";"Expected Value";"Status";"Found In Policy";"Actual Value";"Policy Assigned"';
  const lines=[hdr];
    rows.forEach(r=>{lines.push([r.c,r.bp,r.bf,r.sn,r.sd,r.di,r.ev,r.st,r.fp,r.av,r.pa].map(v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"').join(';'));});
  dlCsv(lines.join('\n'),(t?t+'_':'')+'Baseline-Assessment-Details.csv');
}
function dlCsv(content,filename){
  const blob=new Blob(['\uFEFF'+content],{type:'text/csv;charset=utf-8;'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download=filename;
  a.click();
}
function applyConflictFilter(el){
  const body = el.closest('.conflict-body');
  if(!body) return;
  const chk = body.querySelector('.chk-hide-match');
  const sel = body.querySelector('.sel-policy');
  const hideMatch = chk ? chk.checked : false;
  const policyFilter = sel ? sel.value.toLowerCase() : '';
  const rows = body.querySelectorAll('tbody tr');
  rows.forEach(row=>{
    const statusCell = row.querySelector('.st-match,.st-mismatch,.st-missing');
    const isMatch = statusCell && statusCell.classList.contains('st-match');
    const policyCell = row.children[3];
    const policyText = policyCell ? (policyCell.textContent||'').toLowerCase() : '';
    let show = true;
    if(hideMatch && isMatch) show = false;
    if(show && policyFilter && !policyText.includes(policyFilter)) show = false;
    row.style.display = show ? '' : 'none';
  });
}
function applyReadyFilter(el){
  const body = el.closest('.conflict-body');
  if(!body) return;
  const chk = body.querySelector('.chk-new-only');
  const newOnly = chk ? chk.checked : false;
  const rows = body.querySelectorAll('tbody tr');
  rows.forEach(row=>{
    const statusCell = row.querySelector('.st-new,.st-match');
    const isNew = statusCell && statusCell.classList.contains('st-new');
    row.style.display = (newOnly && !isNew) ? 'none' : '';
  });
}
// Apply default filter (hide matches) on load
document.addEventListener('DOMContentLoaded', function(){
  document.querySelectorAll('.chk-hide-match').forEach(function(chk){ applyConflictFilter(chk); });
  document.querySelectorAll('.chk-new-only').forEach(function(chk){ applyReadyFilter(chk); });
});
let _activeArea=null;
function filterByArea(card){
  const area=card.dataset.area;
  if(_activeArea===area){clearAreaFilter();return;}
  _activeArea=area;
  document.querySelectorAll('.area-card').forEach(c=>c.classList.toggle('area-active',c===card));
  document.querySelectorAll('tr[data-area],section.conflict-section[data-area]').forEach(el=>{
    el.style.display=el.dataset.area===area?'':'none';
  });
  document.querySelectorAll('section[id^="sec-"]:not(#sec-area)').forEach(sec=>{
    const any=sec.querySelector('[data-area="'+area+'"]');
    sec.style.display=any?'':'none';
    if(any){
      const h2=sec.querySelector('h2');
      if(h2){
        if(!h2.dataset.origHtml) h2.dataset.origHtml=h2.innerHTML;
        // conflict-section count avoids double-counting ConsolidationNeeded (has both tr and section rows)
        const cSec=sec.querySelectorAll('section.conflict-section[data-area="'+area+'"]').length;
        const cnt=cSec>0?cSec:sec.querySelectorAll('tr[data-area="'+area+'"]').length;
        h2.innerHTML=h2.dataset.origHtml.replace(/\(\d+\)/,'('+cnt+')');
      }
    }
  });
  document.getElementById('areaFilterLabel').textContent='Filtering: '+area+' policies \u2014 click the card again or \u00D7 to show all';
  document.getElementById('areaFilterBanner').style.display='flex';
}
function clearAreaFilter(){
  _activeArea=null;
  document.querySelectorAll('.area-card').forEach(c=>c.classList.remove('area-active'));
  document.querySelectorAll('tr[data-area],section.conflict-section[data-area]').forEach(el=>el.style.display='');
  document.querySelectorAll('section[id^="sec-"]:not(#sec-area)').forEach(s=>{
    s.style.display='';
    const h2=s.querySelector('h2[data-orig-html]');
    if(h2){h2.innerHTML=h2.dataset.origHtml;delete h2.dataset.origHtml;}
  });
  document.getElementById('areaFilterBanner').style.display='none';
}
</script>
</body>
</html>
"@

    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
}

function Invoke-BaselineAssessmentWorker {
        param([hashtable]$syncHash)

    Set-BaselineResolverTenant -TenantId $syncHash.ResolverTenantOverride

        # Rebuild all helper functions from their serialised definitions
        

        # Thread-safe log helper — enqueues for the UI-thread timer to drain
        function Log { param([string]$M); [void]$syncHash.LogQueue.Enqueue($M) }

        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

            $targetProfiles     = $syncHash.TargetProfiles
            $baselineFolder     = $syncHash.BaselineFolder
            $outputPath         = $syncHash.OutputPath
            $skipUnassigned        = $syncHash.SkipUnassigned
            $skipAssignmentLookups = $syncHash.SkipAssignmentLookups
            $resolverTenantOverride = if ($syncHash.ResolverTenantOverride) { [string]$syncHash.ResolverTenantOverride } else { $null }
            $resolverTenantCacheKey = if ($syncHash.ResolverTenantCacheKey) { [string]$syncHash.ResolverTenantCacheKey } else { 'default' }
            $cacheMode = if ($syncHash.CacheMode) { [string]$syncHash.CacheMode } else { 'fallback' }
            if ($cacheMode -notin @('fallback','cacheonly','forcefresh')) { $cacheMode = 'fallback' }

            $adminTemplateProfiles = @($targetProfiles | Where-Object { [string]$_.PolicyKind -eq 'AdministrativeTemplate' })
            if ($adminTemplateProfiles.Count -gt 0) {
                foreach ($atProfile in $adminTemplateProfiles) {
                    Log "Skipping Administrative Template '$($atProfile.Name)': not supported in baseline assessment (use the Profile Report instead)."
                }
                $targetProfiles = @($targetProfiles | Where-Object { [string]$_.PolicyKind -ne 'AdministrativeTemplate' })
                if ($targetProfiles.Count -eq 0) {
                    Log 'No supported profiles remaining after skipping Administrative Templates.'
                    $syncHash.StatusText = 'No supported profiles'
                    $syncHash.Error = 'NO_PROFILES'
                    $syncHash.Done = $true
                    return
                }
            }

            # ── Phase 0: Fetch assignment data for selected profiles ──────────
            $profileAssignments = @{}
            if ($skipAssignmentLookups) {
                Log 'Skipping assignment lookups (assignment checks disabled).'
                $syncHash.Progress   = 10
                $syncHash.StatusText = 'Skipping assignment lookups'
            } else {
                $total = $targetProfiles.Count
                Log "Fetching assignment data for $total profile(s)..."
                $skippedUnassigned = 0
                $assignedProfiles = [System.Collections.Generic.List[object]]::new()

                for ($i = 0; $i -lt $targetProfiles.Count; $i++) {
                    $profItem = $targetProfiles[$i]
                    $syncHash.Progress   = [math]::Round((($i + 1) / [double]$total) * 10)
                    $syncHash.StatusText = "Fetching assignments $($i + 1) / $total"

                    $assignmentsUri = if ($profItem.PolicyKind -eq 'DeviceRestriction') {
                        "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($profItem.Id)/assignments"
                    } else {
                        "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($profItem.Id)/assignments"
                    }

                    $hasAssignment = $true
                    try {
                        $resp  = Invoke-MgGraphRequest -Uri $assignmentsUri -Method Get -ErrorAction Stop
                        $asgns = Get-SafeObjectPropertyValue -Object $resp -Name 'value'
                        $includes = @($asgns | Where-Object {
                            $t  = Get-SafeObjectPropertyValue -Object $_ -Name 'target'
                            $tt = if ($t) { Get-SafeObjectPropertyValue -Object $t -Name '@odata.type' } else { '' }
                            $tt -and $tt -ne '#microsoft.graph.exclusionGroupAssignmentTarget'
                        })
                        if ($includes.Count -eq 0) { $hasAssignment = $false }
                    } catch {
                        # On error, assume assigned to avoid false skipping
                        $hasAssignment = $true
                    }

                    $profileAssignments[$profItem.Name] = $hasAssignment

                    if ($skipUnassigned -and -not $hasAssignment) {
                        $skippedUnassigned++
                    } else {
                        [void]$assignedProfiles.Add($profItem)
                    }
                }

                if ($skippedUnassigned -gt 0) {
                    Log "Skipped $skippedUnassigned unassigned profile(s)."
                }

                if ($skipUnassigned) {
                    $targetProfiles = $assignedProfiles.ToArray()
                    if ($targetProfiles.Count -eq 0) {
                        Log 'No assigned profiles remaining after filter.'
                        $syncHash.StatusText = 'No assigned profiles'
                        $syncHash.Error = 'NO_PROFILES'
                        $syncHash.Done = $true
                        return
                    }
                }
            }

            $syncHash.ProfileAssignments = $profileAssignments

            # ── Phase 1: Collect settings ─────────────────────────────────────
            $total      = $targetProfiles.Count
            $current    = 0
            $settingRows         = [System.Collections.Generic.List[object]]::new()
            $customerPolicyRawSettings = @{}
            $customerPolicyMeta = @{}
            $cacheData = Import-SettingDefinitionCache -TenantKey $resolverTenantCacheKey
            $cacheError = Get-SafeObjectPropertyValue -Object $cacheData -Name 'Error'
            if ($cacheData.Loaded) {
                Log "Loaded resolver cache for tenant '$resolverTenantCacheKey': $($cacheData.DisplayNameCount) names, $($cacheData.OptionCount) options."
            } elseif ($cacheError) {
                Log "Resolver cache not used: $cacheError"
            }
            Write-CacheStalenessWarning -CacheData $cacheData -LogCallback { param($message) Log $message }
            $allDisplayNames     = $cacheData.DisplayNames
            $allOptionDisplayNames = $cacheData.OptionDisplayNames
            if ($cacheMode -eq 'forcefresh') {
                $allDisplayNames = @{}
                $allOptionDisplayNames = @{}
                Log 'Force fresh mode selected: ignoring imported cache for this run.'
            }
            $policyKindCounts    = @{}

            Log "Collecting settings for $total customer profile(s)..."

            foreach ($profItem in $targetProfiles) {
                $current++
                $syncHash.Progress   = 10 + [math]::Round(($current / [double]$total) * 45)
                $syncHash.StatusText = "Collecting settings $current / $total"
                Log "[$current/$total] $($profItem.Name)"

                $customerPolicyMeta[[string]$profItem.Name] = @{
                    ProfileId  = [string]$profItem.Id
                    Platforms  = [string]$profItem.Platforms
                    PolicyKind = [string]$profItem.PolicyKind
                    OdataType  = [string]$profItem.OdataType
                }

                $policyKind = if ($profItem.PolicyKind) { [string]$profItem.PolicyKind } else { 'Unknown' }
                if (-not $policyKindCounts.ContainsKey($policyKind)) { $policyKindCounts[$policyKind] = 0 }
                $policyKindCounts[$policyKind]++

                if ($profItem.PolicyKind -eq 'DeviceRestriction') {
                    Log "  PolicyKind: $policyKind | Endpoint: deviceConfigurations"
                    try {
                        $dcSkipProps = [System.Collections.Generic.HashSet[string]]::new(
                            [string[]]@('id','displayName','description','createdDateTime',
                                        'lastModifiedDateTime','version','roleScopeTagIds',
                                        'supportsScopeTags','settingCount'),
                            [System.StringComparer]::OrdinalIgnoreCase
                        )
                        $dc = Invoke-MgGraphRequest -Method Get `
                            -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($profItem.Id)"
                        $dcAreaInfo = Get-DeviceConfigurationBaselineAreaInfo -PolicyName ([string]$profItem.Name)
                        if (-not $dcAreaInfo.IsSupported) {
                            Log "  -> Unsupported Device Configuration baseline area '$($dcAreaInfo.BaselineArea)'"
                            continue
                        }
                        $flatDcCount = 0
                        foreach ($propKey in @($dc.Keys)) {
                            if ($propKey -like '*@odata*' -or
                                $propKey -like '#microsoft.graph.*' -or
                                $propKey -ieq 'assignments' -or
                                $propKey -like 'deviceManagementApplicabilityRule*') { continue }
                            if ($dcSkipProps.Contains($propKey)) { continue }
                            if ($dcAreaInfo.PropertyPattern -and $propKey -notmatch $dcAreaInfo.PropertyPattern) { continue }
                            $rawVal = $dc[$propKey]
                            if (-not (Test-DeviceConfigurationSettingValue -Value $rawVal)) { continue }
                            $valStr = if ($null -eq $rawVal) { '-' }
                                      elseif ($rawVal -is [bool]) { $rawVal.ToString().ToLower() }
                                      elseif ($rawVal -is [System.Collections.IEnumerable] -and $rawVal -isnot [string]) {
                                          ($rawVal | ForEach-Object { if ($null -ne $_) { [string]$_ } else { '' } } | Sort-Object) -join '; '
                                      } else { [string]$rawVal }
                            $defId = "$($profItem.OdataType)|$propKey"
                            [void]$settingRows.Add([pscustomobject]@{
                                ProfileId    = $profItem.Id;   ProfileName  = $profItem.Name
                                Platforms    = $profItem.Platforms
                                SettingName  = $propKey;       SettingValue = $valStr
                                RawValue     = $valStr;        DefinitionId = $defId
                            })
                            $flatDcCount++
                        }
                        Log "  -> $flatDcCount settings (Device Restriction)"
                    } catch { Log "  -> FAILED: $($_.Exception.Message)" }
                    continue
                }

                # Settings Catalog
                Log "  PolicyKind: $policyKind | Endpoint: configurationPolicies"
                try {
                    $settingsUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($profItem.Id)/settings?`$expand=settingDefinitions"
                    $settings    = @(Invoke-GraphPagedGet -Uri $settingsUri -TimeoutSeconds 20 -RetryCount 1)

                    $flatSettings = [System.Collections.Generic.List[object]]::new()
                    foreach ($item in $settings) {
                        if ($item -is [System.Array]) { foreach ($sub in $item) { [void]$flatSettings.Add($sub) } }
                        else { [void]$flatSettings.Add($item) }
                    }

                    foreach ($setting in $flatSettings) {
                        $defs = Get-SafeObjectPropertyValue -Object $setting -Name 'settingDefinitions'
                        if ($defs -and $defs -is [System.Collections.IEnumerable] -and -not ($defs -is [string])) {
                            foreach ($def in $defs) {
                                $did   = Get-SafeObjectPropertyValue -Object $def -Name 'id'
                                $dname = Get-SafeObjectPropertyValue -Object $def -Name 'displayName'
                                if ($did -and $dname -and -not $allDisplayNames.ContainsKey([string]$did)) {
                                    $allDisplayNames[[string]$did] = [string]$dname
                                }
                                $defOptions = Get-SafeObjectPropertyValue -Object $def -Name 'options'
                                if ($defOptions -and $defOptions -is [System.Collections.IEnumerable] -and -not ($defOptions -is [string])) {
                                    foreach ($opt in $defOptions) {
                                        $optId   = Get-SafeObjectPropertyValue -Object $opt -Name 'itemId'
                                        $optName = Get-SafeObjectPropertyValue -Object $opt -Name 'displayName'
                                        if ($optId -and $optName -and -not $allOptionDisplayNames.ContainsKey([string]$optId)) {
                                            $allOptionDisplayNames[[string]$optId] = [string]$optName
                                        }
                                    }
                                }
                            }
                        }
                    }

                    $flatSettingCount = 0
                    foreach ($setting in $flatSettings) {
                        $instance = Get-SafeObjectPropertyValue -Object $setting -Name 'settingInstance'
                        if (-not $instance) { continue }
                        $leafItems = @(Get-BaselineJsonLeafInstances -Instance $instance)
                        if ($leafItems.Count -eq 0 -and $instance) {
                            $defId = Get-SafeObjectPropertyValue -Object $instance -Name 'settingDefinitionId'
                            $leafItems = @([pscustomobject]@{ DefinitionId = $defId; RawValue = '-'; Depth = 0 })
                        }
                        foreach ($leaf in $leafItems) {
                            if (-not $leaf) { continue }
                            $defId   = [string]$leaf.DefinitionId; if (-not $defId) { $defId = '-' }
                            $rawVal  = [string]$leaf.RawValue
                            $sname   = if ($defId -ne '-' -and $allDisplayNames.ContainsKey($defId)) { $allDisplayNames[$defId] }
                                       elseif ($defId -match '/') { ($defId -split '/')[-1] }
                                       else { $defId }
                            [void]$settingRows.Add([pscustomobject]@{
                                ProfileId    = $profItem.Id;   ProfileName  = $profItem.Name
                                Platforms    = $profItem.Platforms
                                SettingName  = $sname;         SettingValue = $rawVal
                                RawValue     = $rawVal;        DefinitionId = $defId
                            })
                            $flatSettingCount++
                        }
                    }
                    $customerPolicyRawSettings[[string]$profItem.Name] = @($flatSettings)
                    Log "  -> $flatSettingCount settings"
                } catch { Log "  -> FAILED: $($_.Exception.Message)" }
            }

            if ($policyKindCounts.Keys.Count -gt 0) {
                $policySummary = @($policyKindCounts.Keys | Sort-Object | ForEach-Object { "$_=$($policyKindCounts[$_])" }) -join ', '
                Log "Profile types loaded: $policySummary"
            }

            Log "Total customer settings collected: $($settingRows.Count)"
            Log "Running assessment against baseline folder..."
            $syncHash.StatusText      = 'Running assessment...'
            $syncHash.IsIndeterminate = $true

            $assessmentData = Invoke-BaselineAssessment -BaselineFolder $baselineFolder -SettingRows @($settingRows) -CustomerPolicyRawSettings $customerPolicyRawSettings
            $assessmentData['CustomerPolicyMeta'] = $customerPolicyMeta
            $assessmentData['CustomerPolicyRawSettings'] = $customerPolicyRawSettings

            $syncHash.IsIndeterminate = $false
            $syncHash.Progress        = 65
            $syncHash.AssessmentData   = $assessmentData

            $readyCount        = @($assessmentData.Baselines | Where-Object { $_.Classification -eq 'ReadyToImport' }).Count
            $fullyAlignedCount = @($assessmentData.Baselines | Where-Object { $_.Classification -eq 'FullyAligned' }).Count
            $consolidationCount= @($assessmentData.Baselines | Where-Object { $_.Classification -eq 'ConsolidationNeeded' }).Count
            $splitRequiredCount= @($assessmentData.Baselines | Where-Object { $_.Classification -eq 'SplitRequired' }).Count
            $renameCount       = @($assessmentData.Baselines | Where-Object { $_.Classification -eq 'RenameOnly' }).Count
            $conflictCount     = @($assessmentData.Baselines | Where-Object { $_.Classification -eq 'Conflicts' }).Count
            Log "Assessment complete:"
            Log "  Ready to Import: $readyCount"
            Log "  Fully Aligned: $fullyAlignedCount"
            Log "  Consolidation Needed: $consolidationCount"
            Log "  Split Required: $splitRequiredCount"
            Log "  Rename Only: $renameCount"
            Log "  Conflicts: $conflictCount"
            Log "  Customer Policies to Split: $($assessmentData.CustomerPolicySplit.Count)"

            # Resolve display names for baseline-only settings not already known from live policies
            $allBaselineDefIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($b in $assessmentData.Baselines) {
                foreach ($s in $b.Settings) {
                    $definitionId = [string]$s.DefinitionId
                    if (-not $definitionId -or $allDisplayNames.ContainsKey($definitionId)) { continue }
                    if ($definitionId -match '^#microsoft\.graph\.(?:iosGeneralDeviceConfiguration|androidGeneralDeviceConfiguration|androidWorkProfileGeneralDeviceConfiguration|androidDeviceOwnerGeneralDeviceConfiguration|aospDeviceOwnerGeneralDeviceConfiguration)\|(.+)$') {
                        $allDisplayNames[$definitionId] = [string]$matches[1]
                        continue
                    }
                    [void]$allBaselineDefIds.Add($definitionId)
                }
            }
            if ($allBaselineDefIds.Count -gt 0) {
                if ($cacheMode -eq 'cacheonly') {
                    Log "Strict cache-only mode: skipping Graph lookup for $($allBaselineDefIds.Count) setting definition(s)."
                } else {
                    Log "Resolving $($allBaselineDefIds.Count) setting name(s) via Graph..."
                    $syncHash.StatusText = 'Resolving setting names...'
                    $resolveStats = Resolve-SettingDefinitionsViaGraph `
                        -DefinitionIds @($allBaselineDefIds) `
                        -DisplayNameMap $allDisplayNames `
                        -OptionDisplayMap $allOptionDisplayNames `
                        -ProgressCallback {
                            param($done, $total)
                            $syncHash.Progress = 65 + [math]::Round(($done / [double]$total) * 20)
                        } `
                        -LogCallback {
                            param($message)
                            Log $message
                        } `
                        -AuthChallengeCallback {
                            param($reason)
                            Log 'Reconnecting Microsoft Graph due to authentication challenge...'
                            Disconnect-MgGraph -ErrorAction SilentlyContinue
                            $connectParams = @{
                                Scopes      = @('DeviceManagementConfiguration.Read.All','Organization.Read.All')
                                NoWelcome   = $true
                                ErrorAction = 'Stop'
                            }
                            if ($resolverTenantOverride) {
                                $connectParams['TenantId'] = $resolverTenantOverride
                            }
                            Connect-MgGraph @connectParams | Out-Null
                        }

                    $cachedDefs = $allBaselineDefIds.Count - $resolveStats.Requested
                    if ($cachedDefs -gt 0) {
                        Log "Cache hit(s) for setting definitions: $cachedDefs"
                    }
                    Log "Resolved $($resolveStats.Resolved) setting definition(s) in $($resolveStats.GraphCalls) batch call(s)."
                }
            }

            $savedCache = Export-SettingDefinitionCache -TenantKey $resolverTenantCacheKey -DisplayNameMap $allDisplayNames -OptionDisplayMap $allOptionDisplayNames -SourceMode "assessment-$cacheMode"
            Log "Updated resolver cache: $($savedCache.DisplayNameCount) names, $($savedCache.OptionCount) options."

            Log "Generating HTML report..."
            $syncHash.StatusText = 'Generating report...'
            $syncHash.Progress   = 88

            New-BaselineAssessmentHtmlReport `
                -OutputPath         $outputPath `
                -AssessmentData     $assessmentData `
                -DisplayNames       $allDisplayNames `
                -OptionDisplayNames $allOptionDisplayNames `
                -SettingRows        @($settingRows) `
                -TenantName         $syncHash.TenantName `
                -ProfileAssignments $profileAssignments

            Log "Report created: $outputPath"
            $syncHash.Progress   = 100
            $syncHash.StatusText = 'Completed'
            $syncHash.Done       = $true

        } catch {
            $syncHash.Error      = $_.Exception.Message
            $syncHash.StatusText = 'Error'
            $syncHash.Done       = $true
        }
    }

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    if (-not $BaselineFolder -or -not (Test-Path -LiteralPath $BaselineFolder -PathType Container)) { throw '-BaselineFolder must name an existing folder containing baseline JSON files.' }
    if (-not $ProfileId -and -not $Profiles) { throw 'Supply -ProfileId or -Profiles to select customer profiles.' }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw '-OutputPath is required.' }
    Initialize-BaselineGraphSession -TenantId $TenantId -WriteAccess:($CreateSplitPolicy -and -not $WhatIfPreference)
    $targetProfiles = @(Select-BaselineIntuneProfiles -ProfileId $ProfileId -Profiles $Profiles)
    if (-not $targetProfiles.Count) { throw 'No profiles selected.' }
    $tenantName = ''
    try {
        $organization = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization' -Method Get -ErrorAction Stop
        $tenantName = [string](Get-SafeObjectPropertyValue -Object ($organization.value | Select-Object -First 1) -Name 'displayName')
    } catch { Write-Verbose "Tenant name unavailable: $($_.Exception.Message)" }
    $workerState = [hashtable]::Synchronized(@{
        TargetProfiles = $targetProfiles
        BaselineFolder = $BaselineFolder
        OutputPath = $OutputPath
        CustomerPolicyRawSettings = @{}
        CustomerPolicyMeta = @{}
        SkipUnassigned = [bool]$SkipUnassigned
        SkipAssignmentLookups = [bool]$SkipAssignmentLookups
        ProfileAssignments = @{}
        TenantName = $tenantName
        ResolverTenantOverride = $TenantId
        ResolverTenantCacheKey = Get-GraphTenantCacheKey
        CacheMode = $CacheMode
        AutoOpen = $false
        Progress = 0
        IsIndeterminate = $false
        StatusText = 'Starting...'
        LogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
        Done = $false
        Error = $null
    })
    Invoke-BaselineAssessmentWorker -syncHash $workerState
    $message = ''
    while ($workerState.LogQueue.TryDequeue([ref]$message)) {
        if ($LogCallback) { & $LogCallback $message | Out-Null } else { Write-Verbose $message }
    }
    if ($workerState.Error) { throw $workerState.Error }
    if ($CreateSplitPolicy) {
        $splitCandidates = @($workerState.AssessmentData.Baselines | Where-Object { $_.Classification -eq 'SplitRequired' -and (-not $SplitBaselineName -or $_.PolicyName -in $SplitBaselineName) })
        foreach ($baseline in $splitCandidates) {
            if ($PSCmdlet.ShouldProcess($baseline.PolicyName, 'Create split policy containing customer-only settings')) {
                New-SplitRequiredPolicyFromAssessment -AssessmentData $workerState.AssessmentData -Baseline $baseline
            }
        }
    }
    [pscustomobject]@{ OutputPath = $OutputPath; AssessmentData = $workerState.AssessmentData; ProfileAssignments = $workerState.ProfileAssignments }
    if ($OpenReport) { Start-Process -FilePath $OutputPath }
}
