[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Forms\frmAddInventoryItem.frm") -Raw
$admin = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Modules\modAdmin.bas") -Raw
$apply = Get-Content -LiteralPath (Join-Path $repo "src\InventoryDomain\Modules\modInventoryApply.bas") -Raw
$schema = Get-Content -LiteralPath (Join-Path $repo "src\InventoryDomain\Modules\modInventorySchema.bas") -Raw
$queries = Get-Content -LiteralPath (Join-Path $repo "src\InventoryDomain\Modules\modInventoryQueries.bas") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_phase6_packaged_xlams.ps1") -Raw

$checks = [ordered]@{}
$checks["Docs.Slice4alContract"] =
    $spec.Contains("single-item **Add/Edit Inventory Items** form exposes **Delete Item**") -and
    $plan.Contains("Slice 4al -- Add/Edit managed-inventory deletion") -and
    $controls.Contains('`btnDeleteItem`')
$checks["Form.RealDeleteHandler"] =
    $form.Contains("Private WithEvents mBtnDeleteItem As MSForms.CommandButton") -and
    $form.Contains('AddButton("btnDeleteItem"') -and
    $form.Contains("Private Sub mBtnDeleteItem_Click()") -and
    $form -match 'TestInventoryDeleteActionContract(?s).*?mBtnDeleteItem_Click'
$checks["Admin.ExactSkuRetirementService"] =
    $admin.Contains("Public Function RetireInventoryItemForWarehouse") -and
    $admin.Contains('item("System_Key") = systemKey') -and
    $admin.Contains('item("InventoryState") = "RETIRED"') -and
    $admin.Contains("QueueAdminInventoryAdjustEvent")
$checks["Domain.RetiredProjectionState"] =
    $schema.Contains('"InventoryState", "AttributesJson", "Note"') -and
    $apply.Contains('lineItem("InventoryState") = inventoryState') -and
    $apply.Contains('entity("Retired") = True') -and
    $apply.Contains('IIf(entity.Exists("Retired") And CBool(entity("Retired")), "RETIRED"')
$checks["Domain.ZeroDeltaNonCountedRetirement"] =
    $apply.Contains("PayloadLineIsRetirementApply") -and
    $apply -match 'EVENT_TYPE_ADMIN_INVENTORY_ADJUST(?s).*?qty = 0(?s).*?PayloadLineIsRetirementApply'
$checks["ActiveQueriesAndCatalogOmitRetired"] =
    $queries.Contains('StrComp(inventoryState, "ACTIVE"') -and
    $admin.Contains('item("CATALOG_STATE")') -and
    $admin.Contains('CatalogCellAdmin(lo, rowIndex, "CATALOG_STATE")') -and
    $admin.Contains('MarkInventoryCatalogRetiredAdmin')
$checks["Packaged.RealHandlerEvidence"] =
    $admin.Contains("InventoryDeleteContractForAutomation") -and
    $validator.Contains("Admin.InventoryDeleteItem") -and
    $validator.Contains("InventoryDomain.InventoryRetirement") -and
    $validator.Contains("DeleteHandler=True") -and
    $validator.Contains("ExactKey=True") -and
    $validator.Contains("UtilityZeroDelta=True")

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

Write-Host ("PLAN022_SLICE4AL_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
