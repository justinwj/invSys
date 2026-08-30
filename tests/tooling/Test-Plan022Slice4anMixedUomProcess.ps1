[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) {
    Get-Content -Raw -LiteralPath $Path
}

$worksheet = Read-Text (Join-Path $repo "src\Production\Modules\modProductionProcessWorksheet.bas")
$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")
$historical = Read-Text (Join-Path $repo "tests\tooling\Test-Plan022Slice4yProcessWorksheet.ps1")
$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")

$checks = @(
    [pscustomobject]@{ Name = "Docs.MixedUomGroupingContract"; Pass =
        $spec -match 'partitioned by normalized UOM' -and
        $spec -match 'Every populated UOM group must\s+total 100\.0% independently' -and
        $plan -match 'Slice 4an -- Mixed-UOM Process assembly retrieval' -and
        $controls -match 'Slice 4an mixed-UOM Process assembly' },
    [pscustomobject]@{ Name = "Worksheet.OperatorInstructionAllowsGroups"; Pass =
        $worksheet -match 'INPUT quantities are grouped by UOM' -and
        $worksheet -notmatch 'Enter INPUT quantities in one compatible UOM' },
    [pscustomobject]@{ Name = "Worksheet.FormulasUsePerUomBasis"; Pass =
        $worksheet -match 'SUMIFS\(\[Qty\],\[Record Type\],""INPUT"",\[UOM\],\[@UOM\]\)' -and
        $worksheet -match '\[@Qty\]/\[@\[Basis Qty\]\]\*100' -and
        $worksheet -notmatch 'If inputUoms\.Count > 1 Then' },
    [pscustomobject]@{ Name = "Worksheet.ValidatesEveryUomGroup"; Pass =
        $worksheet -match 'inputPercentTotals' -and
        $worksheet -match 'For Each inputUom' -and
        $worksheet -match 'INPUT formula percentages for' -and
        $worksheet -match 'UOM must total 100\.0%' },
    [pscustomobject]@{ Name = "Form.RealRetrieveAcceptsMixedUom"; Pass =
        $form -match 'mBtnProcessWorksheetRetrieve_Click' -and
        $form -match 'MixedUomAccepted=True' -and
        $form -match 'MixedUomRowsPreserved=True' -and
        $form -notmatch 'MixedUomRejected=True' },
    [pscustomobject]@{ Name = "Packaged.RequiresMixedUomAcceptance"; Pass =
        $validator -match 'MixedUomAccepted=True' -and
        $validator -match 'MixedUomRowsPreserved=True' -and
        $validator -notmatch 'MixedUomRejected=True' },
    [pscustomobject]@{ Name = "HistoricalContractSuperseded"; Pass =
        $historical -match 'per-UOM groups' -and
        $historical -notmatch 'one compatible UOM' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "FAIL" }), $check.Name
}

"PLAN022_SLICE4AN_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) {
    throw "Plan 022 Slice 4an source contract RED: $($failed.Name -join ', ')"
}
