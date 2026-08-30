[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) {
    Get-Content -Raw -LiteralPath $Path
}

$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$controller = Read-Text (Join-Path $repo "src\Production\Modules\mProduction.bas")
$worksheet = Read-Text (Join-Path $repo "src\Production\Modules\modProductionProcessWorksheet.bas")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")
$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")

# Slice 4an supersedes Slice 4y's single-UOM rejection with per-UOM groups.
$checks = @(
    [pscustomobject]@{ Name = "Docs.NormativeContract"; Pass =
        $spec -match 'exactly\s+three uppercase Base-36\s+characters' -and
        $spec -match 'Create Process Table' -and
        $spec -match 'Batch basis' -and
        $plan -match 'Slice 4y -- Process formulation worksheet round-trip' -and
        $controls -match 'btnProcessWorksheetCreate' },
    [pscustomobject]@{ Name = "Form.GeneratedLockedIds"; Pass =
        $form -match 'mTxtProcessId\.Locked = True' -and
        $form -match 'mTxtReusableRecipeId\.Locked = True' -and
        $form -match 'mTxtRequirementId\.Locked = True' -and
        $form -match 'mTxtProcessOutputId\.Locked = True' -and
        $form -match 'NextListBase36Id' -and
        $controller -match 'Public Function NextBase36Identifier' },
    [pscustomobject]@{ Name = "Form.OperatorHandler"; Pass =
        $form -match 'Private Sub mBtnProcessWorksheetCreate_Click\(\)' -and
        $form -match 'Private Sub mBtnProcessWorksheetRetrieve_Click\(\)' -and
        $form -match 'SendProcessDraftToWorksheet' -and
        $form -match 'ReadProcessDraftFromWorksheet' -and
        $form -match 'DeleteProcessWorksheetTable' -and
        $validator -match 'RunProcessWorksheetRoundTripContractTest' },
    [pscustomobject]@{ Name = "Worksheet.CapturedBinding"; Pass =
        $worksheet -match 'ByVal wb As Workbook' -and
        $worksheet -match 'FindOutstandingProcessWorksheetTable' -and
        $worksheet -notmatch '(?i)ActiveWorkbook' -and
        $worksheet -match 'wb\.Save' },
    [pscustomobject]@{ Name = "Worksheet.FormulasAndValidation"; Pass =
        $worksheet -match 'SUMIFS\(\[Qty\]' -and
        $worksheet -match '\[@Qty\]/\[@\[Basis Qty\]\]\*100' -and
        $worksheet -match 'grouped by UOM' -and
        $worksheet -match 'must total 100\.0%' -and
        $worksheet -match 'Every Process must declare at least one OUTPUT' },
    [pscustomobject]@{ Name = "Worksheet.SafeLifecycle"; Pass =
        $worksheet -match 'lo\.Delete' -and
        $form -match 'If Not ValidateProcessDraft\(validationReport\)' -and
        $form -match 'retrieval failed' -and
        $form -match 'Retrieve Selected Process' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "FAIL" }), $check.Name
}
if ($failed.Count -gt 0) {
    throw "Plan 022 Slice 4y source contract failed: $($failed.Name -join ', ')"
}

"PLAN022_SLICE4Y_SOURCE_GREEN $($checks.Count)/$($checks.Count)"
