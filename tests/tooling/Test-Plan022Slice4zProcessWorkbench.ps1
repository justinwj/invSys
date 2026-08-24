[CmdletBinding()]
param([string]$RepoRoot = ".")

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) {
    Get-Content -Raw -LiteralPath $Path
}

$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")
$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$worksheet = Read-Text (Join-Path $repo "src\Production\Modules\modProductionProcessWorksheet.bas")
$events = Read-Text (Join-Path $repo "src\Production\ClassModules\cProductionAppEvents.cls")
$picker = Read-Text (Join-Path $repo "src\Core\ClassModules\cDynItemSearch.cls")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")

$checks = @(
    [pscustomobject]@{ Name = "Docs.MultiTableContract"; Pass =
        $spec -match 'Create Process Table' -and
        $spec -match 'Any number of invSys Process tables may coexist' -and
        $plan -match 'Slice 4z -- multi-table Process import workbench' -and
        $controls -match 'btnProcessWorksheetRetrieve' },
    [pscustomobject]@{ Name = "Form.SeparateHandlers"; Pass =
        $form -match 'mBtnProcessWorksheetCreate' -and
        $form -match 'mBtnProcessWorksheetRetrieve' -and
        $form -match 'Private Sub mBtnProcessWorksheetCreate_Click\(\)' -and
        $form -match 'Private Sub mBtnProcessWorksheetRetrieve_Click\(\)' },
    [pscustomobject]@{ Name = "Worksheet.MultipleSelectedTables"; Pass =
        $worksheet -match 'FindSelectedProcessWorksheetTable' -and
        $worksheet -match 'NextProcessTableTopRow' -and
        $worksheet -notmatch 'Retrieve or discard the outstanding Process worksheet table first' },
    [pscustomobject]@{ Name = "Worksheet.ValidationAndCalculatedPercent"; Pass =
        $worksheet -match '\.Validation' -and
        $worksheet -match '\.Add Type:=xlValidateList' -and
        $worksheet -match 'INPUT,OUTPUT,INSTRUCTION,ALTERNATIVE' -and
        $worksheet -match 'ListColumns\("Percent"\)\.DataBodyRange\.Formula' -and
        $worksheet -match 'ListColumns\("Basis Qty"\)\.DataBodyRange\.Formula' },
    [pscustomobject]@{ Name = "Worksheet.GeneratedDesignNoItemCode"; Pass =
        $worksheet -match 'GeneratedOutputDesignId' -and
        $worksheet -notmatch 'headers = Array\([^\r\n]*"Item Code"' -and
        $form -match 'mTxtProcessOutputItemCode\.Visible = False' -and
        $form -match 'mTxtProcessOutputDesignId\.Locked = True' -and
        $form -match 'mTxtProcessOutputDesignVersion\.Locked = True' },
    [pscustomobject]@{ Name = "Worksheet.AssignmentItemSearch"; Pass =
        $worksheet -match 'Acceptable Managed Item' -and
        $worksheet -match 'Accepted SKU' -and
        $events -match 'IsProcessWorksheetItemSearchTarget' -and
        $picker -match 'invsys_process_' },
    [pscustomobject]@{ Name = "Packaged.SameHandlers"; Pass =
        $validator -match 'RunProcessWorksheetWorkbenchContractTest' -and
        $validator -match 'MultipleTables=True' -and
        $validator -match 'SelectedOnly=True' -and
        $validator -match 'ItemSearch=True' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "RED" }), $check.Name
}
"PLAN022_SLICE4Z_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) { exit 1 }
