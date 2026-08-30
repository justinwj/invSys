[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$form = Get-Content -LiteralPath (Join-Path $repo "src\Receiving\Forms\frmReceiving.frm") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw

$checks = [ordered]@{}
$checks["Form.TenVisibleColumns"] =
    $form.Contains('AddListBox("lstReceiveItems", 18, 150, 964, 116, 10') -and
    $form.Contains('Array("Code", "Item", "UOM", "Available", "Location", "Capacity (coming later)", "Lot", "Condition", "Description", "Vendor")') -and
    $form.Contains("mLstReceiveItems.ColumnCount = 10") -and
    -not $form.Contains("mLstReceiveItems.ColumnCount = 11") -and
    -not $form.Contains("mLstReceiveItems.List(rowIndex, 10)")
$checks["Form.HiddenSystemKeyMap"] =
    $form.Contains("Private mReceiveItemSystemKeys As Collection") -and
    $form.Contains("Private Function SelectedReceiveItemSystemKey") -and
    $form.Contains("sourceSystemKey = SelectedReceiveItemSystemKey(idx)")
$checks["Form.CapacityMappingWithinBounds"] =
    $form.Contains("mLstReceiveItems.List(rowIndex, 5) = vbNullString") -and
    $form.Contains("mLstReceiveItems.List(rowIndex, 9) = NzText(values(r, 10))")
$checks["Form.SearchHandlerEvidence"] =
    $form.Contains("mTxtItemSearch_Change") -and
    $form.Contains("|SearchRowsLoaded=") -and
    $form.Contains("|HiddenSystemKeyMap=") -and
    $form.Contains("|TenColumnItemResults=")
$checks["Form.NonWrappingHeaders"] =
    $form.Contains("headerLabel.WordWrap = False") -and
    $form.Contains("headerLabel.AutoSize = False") -and
    $form.Contains("|HeadersSingleLine=")
$checks["Packaged.RequiresSlice4auEvidence"] =
    $validator.Contains("SearchRowsLoaded=True") -and
    $validator.Contains("HiddenSystemKeyMap=True") -and
    $validator.Contains("TenColumnItemResults=True") -and
    $validator.Contains("HeadersSingleLine=True")
$checks["Docs.Slice4auContract"] =
    $plan.Contains("Slice 4au -- Receiving 10-column results and non-wrapping headers") -and
    $controls.Contains("Slice 4au Receiving result projection")

$passed = 0
$red = 0
foreach ($entry in $checks.GetEnumerator()) {
    if ($entry.Value) {
        Write-Host ("PASS " + $entry.Key)
        $passed++
    }
    else {
        Write-Host ("RED  " + $entry.Key)
        $red++
    }
}

Write-Host ("PLAN022_SLICE4AU_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
