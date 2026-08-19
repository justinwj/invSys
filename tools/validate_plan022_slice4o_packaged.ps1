Param(
    [string]$RepoRoot = ".",
    [string]$PackageRoot = "deploy/current"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject([object]$Object) {
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Find-ListObject([object]$Workbook, [string]$TableName) {
    foreach ($sheet in @($Workbook.Worksheets)) {
        try {
            $table = $sheet.ListObjects.Item($TableName)
            if ($null -ne $table) { return $table }
        } catch {}
    }
    return $null
}

function Set-TableValue([object]$Table, [int]$RowIndex, [string]$ColumnName, [object]$Value) {
    $column = $Table.ListColumns.Item($ColumnName)
    $column.DataBodyRange.Cells.Item($RowIndex, 1).Value2 = [string]$Value
    Release-ComObject $column
}

$repo = (Resolve-Path $RepoRoot).Path
$packages = Join-Path $repo $PackageRoot
$resultPath = Join-Path $repo "tests/integration/plan022_slice4q_packaged_results.md"
$excel = $null
$core = $null
$operations = $null
$admin = $null
$operator = $null
$checks = New-Object System.Collections.Generic.List[object]

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $core = $excel.Workbooks.Open((Join-Path $packages "invSys.Core.xlam"), 0, $false)
    $operations = $excel.Workbooks.Open((Join-Path $packages "invSys.Operations.xlam"), 0, $false)
    $admin = $excel.Workbooks.Open((Join-Path $packages "invSys.Admin.xlam"), 0, $false)
    $core.IsAddin = $true
    $operations.IsAddin = $true
    $admin.IsAddin = $true
    $operator = $excel.Workbooks.Add()

    $surfaceOk = [bool]$excel.Run("'" + $core.Name + "'!modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface", $operator)
    $checks.Add([pscustomobject]@{ Name = "Packaged.ReceivingSurface"; Passed = $surfaceOk; Detail = "Generated operator surface accepts the expanded Receiving schema." })

    $inventory = Find-ListObject $operator "invSys"
    if ($null -eq $inventory) { throw "Packaged Receiving surface did not create invSys." }
    if ($inventory.ListRows.Count -eq 0) { [void]$inventory.ListRows.Add() }
    Set-TableValue $inventory 1 "System_Key" "SYS-PACKAGED-RETURN"
    Set-TableValue $inventory 1 "ITEM_CODE" "SKU-PACKAGED-RETURN"
    Set-TableValue $inventory 1 "ITEM" "Packaged Return Item"
    Set-TableValue $inventory 1 "UOM" "EA"
    Set-TableValue $inventory 1 "LOCATION" "RETURNS"
    Set-TableValue $inventory 1 "QtyAvailable" 100
    Set-TableValue $inventory 1 "TOTAL INV" 100
    Set-TableValue $inventory 1 "Condition" "DAMAGED"

    $staging = Find-ListObject $operator "ReceivedTally"
    while ($staging.ListRows.Count -lt 2) { [void]$staging.ListRows.Add() }

    $contract = [string]$excel.Run("'" + $operations.Name + "'!modTS_Received.RunReceivingReturnsTabContractForTest", $operator)
    $checks.Add([pscustomobject]@{
        Name = "Packaged.ReturnsTabContract"
        Passed = $contract.StartsWith("OK|") -and $contract.Contains("Selected=Returns") -and
            $contract.Contains("AddCaption=Add Disposition") -and $contract.Contains("ReceiptEventType=RETURN") -and
            $contract.Contains("DispositionVisible=True") -and $contract.Contains("DispositionDefault=RETURN") -and
            $contract.Contains("DispositionOptions=RETURN,DUMP") -and
            $contract.Contains("HistoryTitle=Return Entries History") -and
            $contract.Contains("TallyTitle=Return Tally") -and
            $contract.Contains("AggregateTitle=Aggregate Returns") -and
            $contract.Contains("ItemConditionColumn=True")
        Detail = $contract
    })

    $returnAction = [string]$excel.Run("'" + $operations.Name + "'!modTS_Received.RunReceivingInboundReturnFormActionForTest", $operator)
    $checks.Add([pscustomobject]@{
        Name = "Packaged.OutboundDispositionFormAction"
        Passed = $returnAction.StartsWith("OK|") -and $returnAction.Contains("ReceiptType=RETURN") -and $returnAction.Contains("Condition=DAMAGED") -and $returnAction.Contains("Reason=TEST RETURN")
        Detail = $returnAction
    })

    $aggregate = Find-ListObject $operator "AggregateReceived"
    $secondRow = $staging.ListRows.Add()
    Set-TableValue $staging 2 "REF_NUMBER" "RETURN-SECOND"
    Set-TableValue $staging 2 "RECEIPT_TYPE" "RETURN"
    Set-TableValue $staging 2 "ITEMS" "Packaged Return Item"
    Set-TableValue $staging 2 "QUANTITY" 2
    Set-TableValue $staging 2 "UOM" "EA"
    Set-TableValue $staging 2 "VENDOR" ""
    Set-TableValue $staging 2 "LOCATION" "RETURNS"
    Set-TableValue $staging 2 "LOT_NUMBER" ""
    Set-TableValue $staging 2 "Condition" "DAMAGED"
    Set-TableValue $staging 2 "RETURN_REASON" "TEST RETURN"
    Set-TableValue $staging 2 "System_Key" "SYS-PACKAGED-RETURN"
    Set-TableValue $staging 2 "ITEM_CODE" "SKU-PACKAGED-RETURN"
    Set-TableValue $staging 2 "Source_System_Key" "SYS-PACKAGED-RETURN"
    Set-TableValue $staging 2 "EventId" "EVT-PACKAGED-RETURN-SECOND"
    Set-TableValue $staging 2 "WorkflowState" "STAGED"
    $aggregateRebuilt = [bool]$excel.Run("'" + $operations.Name + "'!modTS_Received.RebuildAggregationForWorkbook", $operator)
    $stagingOk = ($null -ne $staging) -and ($staging.ListRows.Count -eq 2) -and
        ([string]$staging.ListColumns.Item("RECEIPT_TYPE").DataBodyRange.Cells.Item(1, 1).Value2 -eq "RETURN") -and
        ([string]$staging.ListColumns.Item("Condition").DataBodyRange.Cells.Item(1, 1).Value2 -eq "DAMAGED") -and
        ([string]$staging.ListColumns.Item("System_Key").DataBodyRange.Cells.Item(1, 1).Value2 -eq "SYS-PACKAGED-RETURN") -and
        ([string]$staging.ListColumns.Item("Source_System_Key").DataBodyRange.Cells.Item(1, 1).Value2 -eq "SYS-PACKAGED-RETURN")
    $aggregateOk = ($null -ne $aggregate) -and ($aggregate.ListRows.Count -eq 1) -and
        ([string]$aggregate.ListColumns.Item("RETURN_REASON").DataBodyRange.Cells.Item(1, 1).Value2 -eq "TEST RETURN") -and
        ([double]$aggregate.ListColumns.Item("QUANTITY").DataBodyRange.Cells.Item(1, 1).Value2 -eq 3) -and
        ([string]$aggregate.ListColumns.Item("REF_NUMBER").DataBodyRange.Cells.Item(1, 1).Value2 -eq "RETURN-TEST, RETURN-SECOND") -and
        ([string]$aggregate.ListColumns.Item("Condition").DataBodyRange.Cells.Item(1, 1).Value2 -eq "DAMAGED") -and
        ([string]$aggregate.ListColumns.Item("System_Key").DataBodyRange.Cells.Item(1, 1).Value2 -eq "SYS-PACKAGED-RETURN")
    $aggregateDetail = "Rebuilt=$aggregateRebuilt; Rows=$($aggregate.ListRows.Count); Ref=$([string]$aggregate.ListColumns.Item('REF_NUMBER').DataBodyRange.Cells.Item(1, 1).Value2); Qty=$([string]$aggregate.ListColumns.Item('QUANTITY').DataBodyRange.Cells.Item(1, 1).Value2); Condition=$([string]$aggregate.ListColumns.Item('Condition').DataBodyRange.Cells.Item(1, 1).Value2); Reason=$([string]$aggregate.ListColumns.Item('RETURN_REASON').DataBodyRange.Cells.Item(1, 1).Value2)"
    $checks.Add([pscustomobject]@{ Name = "Packaged.ReturnProjection"; Passed = $stagingOk -and $aggregateRebuilt -and $aggregateOk; Detail = $aggregateDetail })

    $adminContract = [string]$excel.Run("'" + $admin.Name + "'!modAdmin.DemoInventoryFormContractForAutomation")
    $checks.Add([pscustomobject]@{
        Name = "Packaged.DemoInventorySilentClose"
        Passed = $adminContract.StartsWith("OK|") -and $adminContract.Contains("Cancel=False") -and $adminContract.Contains("CloseIsSilent=True")
        Detail = $adminContract
    })
}
finally {
    if ($null -ne $operator) { try { $operator.Close($false) } catch {} }
    if ($null -ne $admin) { try { $admin.Close($false) } catch {} }
    if ($null -ne $operations) { try { $operations.Close($false) } catch {} }
    if ($null -ne $core) { try { $core.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $operator
    Release-ComObject $admin
    Release-ComObject $operations
    Release-ComObject $core
    Release-ComObject $excel
}

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$lines = @(
    "# Plan 022 Slice 4q Packaged Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Detail |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $detail = ([string]$check.Detail).Replace("|", "\|")
    $lines += "| $($check.Name) | $result | $detail |"
}
[IO.File]::WriteAllLines($resultPath, $lines)
$lines -join [Environment]::NewLine
if ($failed -gt 0) { exit 1 }
