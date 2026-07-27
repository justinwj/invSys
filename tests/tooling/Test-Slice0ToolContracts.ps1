[CmdletBinding()]
param(
    [ValidateSet("Contracts", "Static", "Runtime", "All")]
    [string]$Mode = "All"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$contractRoot = Join-Path $repoRoot "tools\contracts"
$fixtureRoot = Join-Path $PSScriptRoot "fixtures"
$failures = New-Object System.Collections.Generic.List[string]
$passes = New-Object System.Collections.Generic.List[string]

function Add-Pass {
    param([string]$Name)
    $passes.Add($Name)
    Write-Host ("PASS " + $Name)
}

function Add-Failure {
    param(
        [string]$Name,
        [string]$Message
    )
    $failures.Add($Name + ": " + $Message)
    Write-Host ("FAIL " + $Name + " - " + $Message)
}

function Assert-True {
    param(
        [string]$Name,
        [bool]$Condition,
        [string]$Message
    )
    if ($Condition) {
        Add-Pass $Name
    }
    else {
        Add-Failure $Name $Message
    }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required JSON file is missing: $Path"
    }
    return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json)
}

function Get-PropertyNames {
    param($Object)
    if ($null -eq $Object) {
        return @()
    }
    return @($Object.PSObject.Properties.Name)
}

function Assert-SchemaContract {
    param(
        [string]$Name,
        [string]$Path,
        [string]$ExpectedTitle,
        [string]$ExpectedVersion = "1.0.0",
        [string[]]$RequiredProperties
    )

    try {
        $schema = Read-JsonFile $Path
        $schemaProperties = Get-PropertyNames $schema.properties
        $missingProperties = @($RequiredProperties | Where-Object { $_ -notin $schemaProperties })

        Assert-True ($Name + ".Draft") `
            ($schema.'$schema' -eq "https://json-schema.org/draft/2020-12/schema") `
            "Schema must use JSON Schema draft 2020-12."
        Assert-True ($Name + ".Version") `
            ($schema.properties.schemaVersion.const -eq $ExpectedVersion) `
            "Schema must freeze contract version $ExpectedVersion."
        Assert-True ($Name + ".Title") `
            ($schema.title -eq $ExpectedTitle) `
            "Schema title does not identify the expected evidence type."
        Assert-True ($Name + ".RequiredSections") `
            ($missingProperties.Count -eq 0) `
            ("Missing required schema sections: " + ($missingProperties -join ", "))
    }
    catch {
        Add-Failure $Name $_.Exception.Message
    }
}

function Assert-ContractFixtures {
    Assert-SchemaContract `
        -Name "Schema.ImplementationManifest" `
        -Path (Join-Path $contractRoot "implementation-manifest.schema.json") `
        -ExpectedTitle "invSys Static Implementation Manifest" `
        -RequiredProperties @(
            "schemaVersion", "reportType", "generatedAtUtc", "packages",
            "components", "procedures", "ribbons", "tables", "forms",
            "dynamicRoots", "warnings"
        )

    Assert-SchemaContract `
        -Name "Schema.MaintenanceCandidates" `
        -Path (Join-Path $contractRoot "maintenance-candidates.schema.json") `
        -ExpectedTitle "invSys VBA Maintenance Candidates" `
        -RequiredProperties @(
            "schemaVersion", "reportType", "generatedAtUtc", "baseline",
            "candidates", "ratchets", "warnings"
        )

    Assert-SchemaContract `
        -Name "Schema.RuntimeState" `
        -Path (Join-Path $contractRoot "runtime-state.schema.json") `
        -ExpectedTitle "invSys Read-Only Runtime State" `
        -ExpectedVersion "1.1.0" `
        -RequiredProperties @(
            "schemaVersion", "reportType", "capturedAtUtc", "session",
            "loadedAddins", "openWorkbooks", "runtimeResolution", "config",
            "currentUser", "domainBridges", "inboxSummary", "processor",
            "snapshotReadModels", "operatorStaging", "forms", "redaction",
            "safety", "warnings"
        )

    Assert-SchemaContract `
        -Name "Schema.DynamicRoots" `
        -Path (Join-Path $contractRoot "vba-dynamic-roots.schema.json") `
        -ExpectedTitle "invSys VBA Dynamic Root Registry" `
        -RequiredProperties @("schemaVersion", "reportType", "roots", "exceptions")

    try {
        $rootRegistry = Read-JsonFile (Join-Path $contractRoot "vba-dynamic-roots.json")
        $rootKinds = @($rootRegistry.roots | ForEach-Object { $_.rootKind })
        $requiredKinds = @(
            "RIBBON_CALLBACK", "AUTO_OPEN", "WORKBOOK_EVENT",
            "USERFORM_EVENT", "WORKSHEET_EVENT", "PROCESSOR_HANDLER",
            "STRING_DISPATCH", "CROSS_XLAM_BRIDGE", "TEST_ENTRY",
            "WINDOWS_CALLBACK", "COMPATIBILITY_SHIM"
        )
        $missingKinds = @($requiredKinds | Where-Object { $_ -notin $rootKinds })
        Assert-True "Fixture.DynamicRoots.Coverage" `
            ($missingKinds.Count -eq 0) `
            ("Dynamic-root registry fixture omits: " + ($missingKinds -join ", "))
    }
    catch {
        Add-Failure "Fixture.DynamicRoots" $_.Exception.Message
    }

    try {
        $staticSource = Get-Content -Raw -LiteralPath (
            Join-Path $fixtureRoot "static\src\Operations\Modules\modSyntheticSurface.bas"
        )
        $formSource = Get-Content -Raw -LiteralPath (
            Join-Path $fixtureRoot "static\src\Operations\Forms\frmSyntheticSurface.frm"
        )
        $ribbonSource = Get-Content -Raw -LiteralPath (
            Join-Path $fixtureRoot "static\ribbon\customUI14.xml"
        )

        Assert-True "Fixture.Static.DirectCall" `
            $staticSource.Contains("DirectWorker payload") `
            "Static fixture must contain a direct typed call."
        Assert-True "Fixture.Static.LiteralApplicationRun" `
            $staticSource.Contains("Application.Run ""'invSys.Core.xlam'!SyntheticBridge""") `
            "Static fixture must contain a literal cross-XLAM Application.Run target."
        Assert-True "Fixture.Static.DynamicApplicationRun" `
            $staticSource.Contains("Application.Run dynamicTarget") `
            "Static fixture must contain an unresolved dynamic Application.Run expression."
        Assert-True "Fixture.Static.DuplicateBodies" `
            ($staticSource.Contains("DuplicateAlpha") -and $staticSource.Contains("DuplicateBeta")) `
            "Static fixture must contain duplicate normalized procedure bodies."
        Assert-True "Fixture.Static.Unreachable" `
            $staticSource.Contains("UnreferencedCandidate") `
            "Static fixture must contain an unreachable candidate."
        Assert-True "Fixture.Static.Events" `
            ($staticSource.Contains("Auto_Open") -and $formSource.Contains("UserForm_Initialize")) `
            "Static fixture must contain convention-dispatched event procedures."
        Assert-True "Fixture.Static.Ribbon" `
            ($ribbonSource.Contains("onAction=""RibbonSyntheticOnAction""") -and
             $ribbonSource.Contains("getEnabled=""RibbonSyntheticGetEnabled""")) `
            "Ribbon fixture must identify action and state callbacks."
    }
    catch {
        Add-Failure "Fixture.Static" $_.Exception.Message
    }

    try {
        $runtimeInput = Read-JsonFile (Join-Path $fixtureRoot "runtime\session-input.json")
        $validTable = @($runtimeInput.tables | Where-Object { $_.name -eq "tblInventoryEntities" })[0]
        $legacyTable = @($runtimeInput.tables | Where-Object { $_.name -eq "tblLegacyViolation" })[0]
        $validHeaders = @($validTable.headers)
        $legacyHeaders = @($legacyTable.headers)

        Assert-True "Fixture.Runtime.SystemKey" `
            ("System_Key" -in $validHeaders) `
            "Managed inventory fixture must contain System_Key."
        Assert-True "Fixture.Runtime.CustomHeader" `
            ("Custom_Color" -in $validHeaders) `
            "Managed inventory fixture must contain an unknown custom header to preserve."
        Assert-True "Fixture.Runtime.NoRowInValidTable" `
            ("ROW" -notin $validHeaders) `
            "Valid managed inventory fixture must not contain ROW."
        Assert-True "Fixture.Runtime.RetiredRowViolation" `
            ("ROW" -in $legacyHeaders) `
            "Negative fixture must contain ROW so tools can report the retired contract."

        $runtimeText = Get-Content -Raw -LiteralPath (
            Join-Path $fixtureRoot "runtime\session-input.json"
        )
        $secretMarkers = @(
            "REDACTION_SENTINEL_ALPHA",
            "REDACTION_SENTINEL_BETA",
            "REDACTION_SENTINEL_GAMMA"
        )
        Assert-True "Fixture.Runtime.RedactionInputs" `
            (@($secretMarkers | Where-Object { -not $runtimeText.Contains($_) }).Count -eq 0) `
            "Runtime fixture must exercise every required secret redaction class."
    }
    catch {
        Add-Failure "Fixture.Runtime" $_.Exception.Message
    }

    try {
        $expectedRuntimeText = Get-Content -Raw -LiteralPath (
            Join-Path $fixtureRoot "runtime\expected\runtime-state.json"
        )
        $expectedRuntime = $expectedRuntimeText | ConvertFrom-Json
        $expectedWarningCodes = @($expectedRuntime.warnings | ForEach-Object { $_.code })
        $leakedMarkers = @(@(
            "REDACTION_SENTINEL_ALPHA",
            "REDACTION_SENTINEL_BETA",
            "REDACTION_SENTINEL_GAMMA"
        ) | Where-Object { $expectedRuntimeText.Contains($_) })

        Assert-True "Expected.Runtime.NoSecretLeak" `
            ($leakedMarkers.Count -eq 0) `
            ("Expected runtime evidence leaks markers: " + ($leakedMarkers -join ", "))
        Assert-True "Expected.Runtime.RowWarning" `
            ("RETIRED_ROW_HEADER" -in $expectedWarningCodes) `
            "Expected runtime evidence must flag ROW as a retired runtime contract."
        Assert-True "Expected.Runtime.LegacyPackageWarning" `
            ("LEGACY_ROLE_ADDINS_LOADED" -in $expectedWarningCodes) `
            "Expected runtime evidence must characterize the current legacy role package layout."
        Assert-True "Expected.Runtime.RedactionAudit" `
            ($expectedRuntime.redaction.redactedFieldCount -eq 3) `
            "Expected runtime evidence must audit all three synthetic secret fields."
        Assert-True "Expected.Runtime.SafetyProof" `
            (-not $expectedRuntime.safety.excelStartedByTool -and
             $expectedRuntime.safety.mutatingActionsInvoked -eq 0 -and
             @($expectedRuntime.safety.inspectedFiles | Where-Object {
                -not $_.unchanged -or $_.beforeSha256 -ne $_.afterSha256
             }).Count -eq 0) `
            "Expected runtime evidence must prove zero mutation and unchanged inspected files."
    }
    catch {
        Add-Failure "Expected.Runtime" $_.Exception.Message
    }

    try {
        $markdownFiles = @(
            (Join-Path $fixtureRoot "static\expected\implementation-manifest.md"),
            (Join-Path $fixtureRoot "static\expected\maintenance-candidates.md"),
            (Join-Path $fixtureRoot "runtime\expected\runtime-state.md")
        )
        foreach ($markdownFile in $markdownFiles) {
            $first = [System.IO.File]::ReadAllBytes($markdownFile)
            $second = [System.IO.File]::ReadAllBytes($markdownFile)
            $same = ([Convert]::ToBase64String($first) -eq [Convert]::ToBase64String($second))
            Assert-True ("Expected.Markdown.StableRead." + [IO.Path]::GetFileName($markdownFile)) `
                $same `
                "Expected Markdown fixture is not byte-stable."
        }
    }
    catch {
        Add-Failure "Expected.Markdown" $_.Exception.Message
    }
}

function Invoke-StaticToolContract {
    $toolPath = Join-Path $repoRoot "tools\inventory-vba-surface.ps1"
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        Add-Failure "ToolA.EntryPoint" `
            "tools/inventory-vba-surface.ps1 is absent; this is the expected Slice 0 RED."
        return
    }

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
        "invsys-slice0-static-" + [Guid]::NewGuid().ToString("N")
    )
    try {
        $firstRoot = Join-Path $tempRoot "first"
        $secondRoot = Join-Path $tempRoot "second"
        New-Item -ItemType Directory -Path $firstRoot | Out-Null
        New-Item -ItemType Directory -Path $secondRoot | Out-Null

        foreach ($outputRoot in @($firstRoot, $secondRoot)) {
            & $toolPath `
                -SourceRoot (Join-Path $fixtureRoot "static\src") `
                -BuildMapPath (Join-Path $fixtureRoot "static\build-projects.json") `
                -RibbonRoot (Join-Path $fixtureRoot "static\ribbon") `
                -TestRoot (Join-Path $fixtureRoot "static\tests") `
                -RootRegistryPath (Join-Path $contractRoot "vba-dynamic-roots.json") `
                -OutputDirectory $outputRoot `
                -ReportTimestampUtc "2026-07-27T00:00:00Z"
            if (-not $?) {
                throw "Static scanner did not complete successfully."
            }
        }

        foreach ($name in @(
            "implementation-manifest.json",
            "implementation-manifest.md",
            "maintenance-candidates.json",
            "maintenance-candidates.md"
        )) {
            $firstPath = Join-Path $firstRoot $name
            $secondPath = Join-Path $secondRoot $name
            $equal = (Test-Path -LiteralPath $firstPath) -and
                (Test-Path -LiteralPath $secondPath) -and (
                [IO.File]::ReadAllText($firstPath) -ceq [IO.File]::ReadAllText($secondPath)
            )
            Assert-True ("ToolA.Deterministic." + $name) `
                $equal `
                "Repeated static scanner runs are not byte-for-byte deterministic."
        }

        $manifest = Read-JsonFile (Join-Path $firstRoot "implementation-manifest.json")
        $candidates = Read-JsonFile (Join-Path $firstRoot "maintenance-candidates.json")
        $procedureNames = @($manifest.procedures | ForEach-Object { $_.name })
        $rootKinds = @($manifest.dynamicRoots | ForEach-Object { $_.rootKind })
        $candidateTypes = @($candidates.candidates | ForEach-Object { $_.candidateType })
        $warningCodes = @($manifest.warnings | ForEach-Object { $_.code })

        Assert-True "ToolA.Semantics.DirectProcedure" `
            ("DirectWorker" -in $procedureNames) `
            "Static manifest omitted a directly called procedure."
        Assert-True "ToolA.Semantics.DynamicRoots" `
            (("RIBBON_CALLBACK" -in $rootKinds) -and ("AUTO_OPEN" -in $rootKinds)) `
            "Static manifest omitted Ribbon or Auto_Open roots."
        Assert-True "ToolA.Semantics.UnresolvedDynamicCall" `
            ($candidates.baseline.unresolvedApplicationRunCount -eq 1) `
            "Static scanner did not report the unresolved Application.Run expression."
        Assert-True "ToolA.Semantics.DuplicateBody" `
            ("REPLACE_DUPLICATE" -in $candidateTypes) `
            "Static scanner did not report duplicate normalized procedure bodies."
        Assert-True "ToolA.Semantics.UnreachableReviewOnly" `
            (@($candidates.candidates | Where-Object {
                $_.procedureNames -contains "UnreferencedCandidate" -and $_.reviewRequired
            }).Count -eq 1) `
            "Unreachable candidate must remain review-only."
        Assert-True "ToolA.Semantics.RetiredRow" `
            ("RETIRED_ROW_HEADER" -in $warningCodes) `
            "Static scanner did not flag the retired ROW contract."
    }
    catch {
        Add-Failure "ToolA.Execution" $_.Exception.Message
    }
    finally {
        if ((Test-Path -LiteralPath $tempRoot) -and
            $tempRoot.StartsWith([IO.Path]::GetTempPath(), [StringComparison]::OrdinalIgnoreCase) -and
            ([IO.Path]::GetFileName($tempRoot) -like "invsys-slice0-static-*")) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

function Invoke-RuntimeToolContract {
    $toolPath = Join-Path $repoRoot "tools\export-invsys-runtime-state.ps1"
    if (-not (Test-Path -LiteralPath $toolPath -PathType Leaf)) {
        Add-Failure "ToolB.EntryPoint" `
            "tools/export-invsys-runtime-state.ps1 is absent; this is the expected Slice 0 RED."
        return
    }

    $toolText = Get-Content -Raw -LiteralPath $toolPath
    $forbiddenMutationPatterns = @(
        'Workbooks\s*\.\s*Open',
        '\.\s*Save(?:As|CopyAs)?\s*\(',
        '\.\s*Close\s*\(',
        '\.\s*Refresh(?:All)?\s*\(',
        'Application\s*\.\s*Run',
        'CreateObject\s*\(\s*"Excel\.Application"',
        'New-Object\s+-ComObject\s+Excel\.Application',
        '\.\s*Quit\s*\('
    )
    $foundMutationPatterns = @(
        $forbiddenMutationPatterns |
            Where-Object { [regex]::IsMatch($toolText, $_, "IgnoreCase") }
    )
    Assert-True "ToolB.Safety.NoMutationApis" `
        ($foundMutationPatterns.Count -eq 0) `
        ("Runtime extractor contains forbidden Excel mutation APIs: " +
         ($foundMutationPatterns -join ", "))
    Assert-True "ToolB.Safety.AttachOnly" `
        $toolText.Contains("GetActiveObject") `
        "Live inspection must attach to an existing Excel session and never start Excel."

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
        "invsys-slice0-runtime-" + [Guid]::NewGuid().ToString("N")
    )
    $inspectedFixturePath = Join-Path $fixtureRoot `
        "runtime\synthetic-inspected-workbook.dat"
    $fixtureHashBefore = (
        Get-FileHash -LiteralPath $inspectedFixturePath -Algorithm SHA256
    ).Hash
    try {
        $firstRoot = Join-Path $tempRoot "first"
        $secondRoot = Join-Path $tempRoot "second"
        New-Item -ItemType Directory -Path $firstRoot | Out-Null
        New-Item -ItemType Directory -Path $secondRoot | Out-Null

        foreach ($outputRoot in @($firstRoot, $secondRoot)) {
            & $toolPath `
                -FixturePath (Join-Path $fixtureRoot "runtime\session-input.json") `
                -OutputDirectory $outputRoot `
                -ReportTimestampUtc "2026-07-27T00:00:00Z"
            if (-not $?) {
                throw "Runtime extractor did not complete successfully."
            }
        }

        foreach ($name in @("runtime-state.json", "runtime-state.md")) {
            $firstPath = Join-Path $firstRoot $name
            $secondPath = Join-Path $secondRoot $name
            $equal = (Test-Path -LiteralPath $firstPath) -and
                (Test-Path -LiteralPath $secondPath) -and (
                [IO.File]::ReadAllText($firstPath) -ceq [IO.File]::ReadAllText($secondPath)
            )
            Assert-True ("ToolB.Deterministic." + $name) `
                $equal `
                "Repeated runtime extractor runs are not byte-for-byte deterministic."
        }

        $runtimeJsonPath = Join-Path $firstRoot "runtime-state.json"
        $runtimeText = [IO.File]::ReadAllText($runtimeJsonPath)
        $runtime = $runtimeText | ConvertFrom-Json
        $runtimeMarkdown = [IO.File]::ReadAllText(
            (Join-Path $firstRoot "runtime-state.md")
        )
        $warningCodes = @($runtime.warnings | ForEach-Object { $_.code })
        $secretMarkers = @(
            "REDACTION_SENTINEL_ALPHA",
            "REDACTION_SENTINEL_BETA",
            "REDACTION_SENTINEL_GAMMA"
        )
        $leaks = @($secretMarkers | Where-Object { $runtimeText.Contains($_) })

        Assert-True "ToolB.Semantics.NoSecretLeak" `
            ($leaks.Count -eq 0) `
            ("Runtime report leaked: " + ($leaks -join ", "))
        Assert-True "ToolB.Semantics.NoRowValues" `
            (-not $runtime.redaction.rowValuesIncluded) `
            "Default runtime report must not include row-level values."
        Assert-True "ToolB.Semantics.LegacyPackageWarning" `
            ("LEGACY_ROLE_ADDINS_LOADED" -in $warningCodes) `
            "Runtime report did not characterize the current legacy package layout."
        Assert-True "ToolB.Semantics.RetiredRowWarning" `
            ("RETIRED_ROW_HEADER" -in $warningCodes) `
            "Runtime report did not flag ROW as a retired managed header."
        $missingMarkdownFacts = @(
            ("- Loaded invSys add-ins: " + @($runtime.loadedAddins).Count),
            ("- Open workbooks: " + @($runtime.openWorkbooks).Count),
            ("- Mutating actions invoked: " +
             $runtime.safety.mutatingActionsInvoked)
        ) | Where-Object { -not $runtimeMarkdown.Contains($_) }
        $missingMarkdownWarnings = @(
            $warningCodes | Where-Object { -not $runtimeMarkdown.Contains($_) }
        )
        Assert-True "ToolB.Semantics.MarkdownAgrees" `
            (@($missingMarkdownFacts).Count -eq 0 -and
             $missingMarkdownWarnings.Count -eq 0) `
            "Markdown omits counts or warning codes present in canonical JSON."
        Assert-True "ToolB.Safety.ZeroMutationCounters" `
            (-not $runtime.safety.excelStartedByTool -and
             $runtime.safety.workbooksOpenedByTool -eq 0 -and
             $runtime.safety.workbooksClosedByTool -eq 0 -and
             $runtime.safety.workbooksSavedByTool -eq 0 -and
             $runtime.safety.refreshActionsInvoked -eq 0 -and
             $runtime.safety.processorActionsInvoked -eq 0 -and
             $runtime.safety.repairActionsInvoked -eq 0 -and
             $runtime.safety.mutatingActionsInvoked -eq 0) `
            "Runtime extractor reported a forbidden mutating action."
        Assert-True "ToolB.Safety.ReportedHashesUnchanged" `
            (@($runtime.safety.inspectedFiles | Where-Object {
                -not $_.unchanged -or $_.beforeSha256 -ne $_.afterSha256
            }).Count -eq 0) `
            "Runtime report did not prove identical before/after inspected-file hashes."

        $compareToolPath = Join-Path $repoRoot "tools\compare-invsys-reports.ps1"
        if (-not (Test-Path -LiteralPath $compareToolPath -PathType Leaf)) {
            Add-Failure "ToolB.Comparison.EntryPoint" `
                "tools/compare-invsys-reports.ps1 is absent; reports cannot be compared offline."
        }
        else {
            $comparisonPath = Join-Path $tempRoot "comparison.json"
            & $compareToolPath `
                -BeforePath (Join-Path $firstRoot "runtime-state.json") `
                -AfterPath (Join-Path $secondRoot "runtime-state.json") `
                -OutputPath $comparisonPath
            if (-not $?) {
                throw "Offline report comparison did not complete successfully."
            }
            $comparison = Read-JsonFile $comparisonPath
            Assert-True "ToolB.Comparison.NoExcelDependency" `
                (-not (
                    (Get-Content -Raw -LiteralPath $compareToolPath) -match
                    "(?i)Excel\.Application|GetActiveObject|Workbooks"
                )) `
                "Comparison command must not depend on reopening Excel."
            Assert-True "ToolB.Comparison.IdenticalReports" `
                ($comparison.identical -and @($comparison.differences).Count -eq 0) `
                "Comparison command did not identify byte-equivalent semantic reports."

            $changedReportPath = Join-Path $tempRoot "runtime-state-changed.json"
            $changedReport = Read-JsonFile (Join-Path $secondRoot "runtime-state.json")
            $changedReport.capturedAtUtc = "2026-07-27T00:00:01Z"
            [IO.File]::WriteAllText(
                $changedReportPath,
                (($changedReport | ConvertTo-Json -Depth 100).TrimEnd() +
                 [Environment]::NewLine),
                (New-Object Text.UTF8Encoding($false))
            )
            $changedComparisonPath = Join-Path $tempRoot "comparison-changed.json"
            & $compareToolPath `
                -BeforePath (Join-Path $firstRoot "runtime-state.json") `
                -AfterPath $changedReportPath `
                -OutputPath $changedComparisonPath
            if (-not $?) {
                throw "Offline changed-report comparison did not complete successfully."
            }
            $changedComparison = Read-JsonFile $changedComparisonPath
            Assert-True "ToolB.Comparison.DetectsChange" `
                (-not $changedComparison.identical -and
                 $changedComparison.differenceCount -eq 1 -and
                 @($changedComparison.differences)[0].path -eq '$.capturedAtUtc') `
                "Comparison command did not report the expected capturedAtUtc change."
        }
    }
    catch {
        Add-Failure "ToolB.Execution" $_.Exception.Message
    }
    finally {
        $fixtureHashAfter = (
            Get-FileHash -LiteralPath $inspectedFixturePath -Algorithm SHA256
        ).Hash
        Assert-True "ToolB.Safety.IndependentFixtureHash" `
            ($fixtureHashBefore -eq $fixtureHashAfter) `
            "Independent hash proof shows the inspected fixture changed."
        if ((Test-Path -LiteralPath $tempRoot) -and
            $tempRoot.StartsWith([IO.Path]::GetTempPath(), [StringComparison]::OrdinalIgnoreCase) -and
            ([IO.Path]::GetFileName($tempRoot) -like "invsys-slice0-runtime-*")) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

if ($Mode -in @("Contracts", "All")) {
    Assert-ContractFixtures
}
if ($Mode -in @("Static", "All")) {
    Invoke-StaticToolContract
}
if ($Mode -in @("Runtime", "All")) {
    Invoke-RuntimeToolContract
}

Write-Host ("RESULT passed=" + $passes.Count + " failed=" + $failures.Count)
if ($failures.Count -gt 0) {
    foreach ($failure in $failures) {
        Write-Host ("  " + $failure)
    }
    exit 1
}

exit 0
