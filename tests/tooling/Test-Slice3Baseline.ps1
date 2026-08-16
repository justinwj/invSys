[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$generatorPath = Join-Path $repoRoot "tools\create-maintenance-baseline.ps1"
$baselineRoot = Join-Path $repoRoot "reports\static-baseline"
$backlogSchemaPath = Join-Path $repoRoot `
    "tools\contracts\reviewed-cleanup-backlog.schema.json"
$failures = New-Object System.Collections.Generic.List[string]
$passes = New-Object System.Collections.Generic.List[string]

function Add-Check {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$FailureMessage
    )
    if ($Passed) {
        $passes.Add($Name)
        Write-Host ("PASS " + $Name)
    }
    else {
        $failures.Add($Name + ": " + $FailureMessage)
        Write-Host ("FAIL " + $Name + " - " + $FailureMessage)
    }
}

function Read-JsonFile {
    param([string]$Path)
    return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json)
}

$requiredArtifacts = @(
    "implementation-manifest.json",
    "implementation-manifest.md",
    "maintenance-candidates.json",
    "maintenance-candidates.md",
    "reviewed-cleanup-backlog.json",
    "reviewed-cleanup-backlog.md"
)

if (-not (Test-Path -LiteralPath $generatorPath -PathType Leaf)) {
    Add-Check "Slice3.EntryPoint" $false `
        "tools/create-maintenance-baseline.ps1 is absent; this is the expected Slice 3 RED."
    Write-Host ("RESULT passed=" + $passes.Count + " failed=" + $failures.Count)
    exit 1
}

Add-Check "Slice3.Schema.Exists" `
    (Test-Path -LiteralPath $backlogSchemaPath -PathType Leaf) `
    "Reviewed backlog schema is missing."
foreach ($artifact in $requiredArtifacts) {
    Add-Check ("Slice3.Baseline.Exists." + $artifact) `
        (Test-Path -LiteralPath (Join-Path $baselineRoot $artifact) -PathType Leaf) `
        ("Committed baseline artifact is missing: " + $artifact)
}

if ($failures.Count -eq 0) {
    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
        "invsys-slice3-baseline-" + [Guid]::NewGuid().ToString("N")
    )
    try {
        New-Item -ItemType Directory -Path $tempRoot | Out-Null
        & $generatorPath `
            -OutputDirectory $tempRoot `
            -ReportTimestampUtc "2026-08-16T20:00:00Z"
        if (-not $?) {
            throw "Baseline generator failed."
        }

        foreach ($artifact in $requiredArtifacts) {
            $committedPath = Join-Path $baselineRoot $artifact
            $regeneratedPath = Join-Path $tempRoot $artifact
            Add-Check ("Slice3.Deterministic." + $artifact) `
                ([IO.File]::ReadAllText($committedPath) -ceq
                 [IO.File]::ReadAllText($regeneratedPath)) `
                ("Regenerated artifact differs from the committed baseline: " + $artifact)
        }

        $validatorPath = Join-Path $repoRoot "tools\validate-json-contract.ps1"
        & $validatorPath `
            -JsonPath (Join-Path $tempRoot "reviewed-cleanup-backlog.json") `
            -SchemaPath $backlogSchemaPath
        if (-not $?) {
            throw "Reviewed cleanup backlog failed schema validation."
        }

        $manifest = Read-JsonFile (Join-Path $tempRoot "implementation-manifest.json")
        $backlog = Read-JsonFile (Join-Path $tempRoot "reviewed-cleanup-backlog.json")
        $workstreamNames = @($backlog.workstreams | ForEach-Object { $_.name })
        $requiredWorkstreams = @(
            "RECEIVING", "PRODUCTION", "SHIPPING", "SHARED_OPERATIONS"
        )
        Add-Check "Slice3.Backlog.RoleSeparation" `
            (@($requiredWorkstreams | Where-Object {
                $_ -notin $workstreamNames
            }).Count -eq 0) `
            "Backlog does not separately identify every role/shared Operations workstream."
        Add-Check "Slice3.Backlog.NoAutomaticDeletion" `
            (-not $backlog.policy.automaticDeletionAllowed -and
             @($backlog.candidates | Where-Object {
                 $_.deletionApproved
             }).Count -eq 0) `
            "Baseline authorizes an automatic or unreviewed deletion."

        $deletionDispositions = @(
            "REMOVE", "REPLACE_DUPLICATE", "ISOLATE_LEGACY_IMPORT"
        )
        $unsafeHigh = @($backlog.candidates | Where-Object {
            $_.reviewedConfidence -eq "HIGH" -and
            $_.disposition -in $deletionDispositions -and (
                [string]::IsNullOrWhiteSpace([string]$_.reason) -or
                @($_.protectingTests).Count -eq 0
            )
        })
        Add-Check "Slice3.Backlog.HighDeletionProtected" `
            ($unsafeHigh.Count -eq 0) `
            "A HIGH-confidence deletion candidate lacks a reason or protecting test."

        $missingRatchets = @($manifest.components | Where-Object {
            $_.lineCount -gt $backlog.ratchets.maxNewModuleLines -and
            $_.sourcePath -notin @(
                $backlog.ratchets.oversizedModules |
                    ForEach-Object { $_.sourcePath }
            )
        })
        Add-Check "Slice3.Ratchets.OversizedModulesRecorded" `
            ($missingRatchets.Count -eq 0 -and
             @($backlog.ratchets.oversizedModules).Count -gt 0) `
            "One or more oversized runtime modules lack a growth baseline."
        Add-Check "Slice3.Ratchets.NoGrowthDefaults" `
            (-not $backlog.ratchets.allowSameProjectApplicationRunGrowth -and
             -not $backlog.ratchets.allowUnresolvedDynamicCallGrowth -and
             -not $backlog.ratchets.allowDuplicateBodyGrowth) `
            "Dynamic-call or duplicate-body ratchets allow unreviewed growth."

        $registry = Read-JsonFile (
            Join-Path $repoRoot "tools\contracts\vba-dynamic-roots.json"
        )
        Add-Check "Slice3.RootRegistry.ClassEvents" `
            ("CLASS_EVENT" -in @($registry.roots | ForEach-Object {
                $_.rootKind
            })) `
            "Class and WithEvents callback procedures are not registered roots."
    }
    catch {
        $failures.Add("Slice3.Execution: " + $_.Exception.Message)
        Write-Host ("FAIL Slice3.Execution - " + $_.Exception.Message)
    }
    finally {
        if ((Test-Path -LiteralPath $tempRoot) -and
            $tempRoot.StartsWith([IO.Path]::GetTempPath(), [StringComparison]::OrdinalIgnoreCase) -and
            ([IO.Path]::GetFileName($tempRoot) -like "invsys-slice3-baseline-*")) {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force
        }
    }
}

Write-Host ("RESULT passed=" + $passes.Count + " failed=" + $failures.Count)
if ($failures.Count -gt 0) {
    foreach ($failure in $failures) {
        Write-Host ("  " + $failure)
    }
    exit 1
}

exit 0
