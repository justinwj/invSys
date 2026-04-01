Param(
    [string]$OutputDir = "tests/fixtures"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    Param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Add-Table {
    Param(
        [object]$Worksheet,
        [string]$TableName,
        [object[]]$Headers,
        [object[][]]$Rows
    )

    $colCount = $Headers.Count
    $rowCount = [Math]::Max($Rows.Count, 1)

    $Worksheet.Range("A1").Resize(1, $colCount).Value = ,$Headers
    if ($Rows.Count -gt 0) {
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            $Worksheet.Range("A" + ($r + 2)).Resize(1, $colCount).Value = ,$Rows[$r]
        }
    } else {
        $Worksheet.Range("A2").Resize(1, $colCount).Value = ,([object[]]::new($colCount))
    }

    $endCell = $Worksheet.Cells($rowCount + 1, $colCount)
    $range = $Worksheet.Range("A1", $endCell)
    $listObject = $Worksheet.ListObjects.Add(1, $range, $null, 1)
    $listObject.Name = $TableName
    return $listObject
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $inventoryPath = Join-Path (Resolve-Path $OutputDir) "WH1.invSys.Data.Inventory.xlsb"
    $receivePath = Join-Path (Resolve-Path $OutputDir) "invSys.Inbox.Receiving.S1.xlsb"
    $shipPath = Join-Path (Resolve-Path $OutputDir) "invSys.Inbox.Shipping.S1.xlsb"
    $prodPath = Join-Path (Resolve-Path $OutputDir) "invSys.Inbox.Production.S1.xlsb"

    foreach ($path in @($inventoryPath, $receivePath, $shipPath, $prodPath)) {
        if (Test-Path $path) { Remove-Item $path -Force }
    }

    $wbInventory = $excel.Workbooks.Add()
    try {
        $wsInventoryLog = $wbInventory.Worksheets(1)
        $wsInventoryLog.Name = "InventoryLog"
        Add-Table -Worksheet $wsInventoryLog -TableName "tblInventoryLog" -Headers @(
            "EventID", "UndoOfEventId", "AppliedSeq", "EventType", "OccurredAtUTC", "AppliedAtUTC",
            "WarehouseId", "StationId", "UserId", "SKU", "QtyDelta", "Location", "Note"
        ) -Rows @()

        $wsApplied = $wbInventory.Worksheets.Add()
        $wsApplied.Name = "AppliedEvents"
        Add-Table -Worksheet $wsApplied -TableName "tblAppliedEvents" -Headers @(
            "EventID", "UndoOfEventId", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status"
        ) -Rows @()

        $wsLocks = $wbInventory.Worksheets.Add()
        $wsLocks.Name = "Locks"
        Add-Table -Worksheet $wsLocks -TableName "tblLocks" -Headers @(
            "LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", "ExpiresAtUTC", "HeartbeatAtUTC", "Status"
        ) -Rows @(
            @("INVENTORY", "", "", "", "", "", "", "EXPIRED")
        )

        $wsSku = $wbInventory.Worksheets.Add()
        $wsSku.Name = "SkuCatalog"
        Add-Table -Worksheet $wsSku -TableName "tblSkuCatalog" -Headers @("SKU") -Rows @(
            @("SKU-001"),
            @("SKU-002"),
            @("SKU-COMP"),
            @("SKU-FG")
        )

        $wbInventory.SaveAs($inventoryPath, 50)
    }
    finally {
        $wbInventory.Close($true)
        Release-ComObject $wbInventory
    }

    $inboxHeaders = @(
        "EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", "WarehouseId", "StationId",
        "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", "ErrorCode", "ErrorMessage", "FailedAtUTC"
    )

    $wbInbox = $excel.Workbooks.Add()
    try {
        $wsInbox = $wbInbox.Worksheets(1)
        $wsInbox.Name = "InboxReceive"
        Add-Table -Worksheet $wsInbox -TableName "tblInboxReceive" -Headers $inboxHeaders -Rows @(
            @("EVT-FIXTURE-001", "", "", "RECEIVE", (Get-Date).ToUniversalTime(), "WH1", "S1", "user1", "SKU-001", 5, "A1", "fixture row", "", "NEW", 0, "", "", "")
        )
        $wbInbox.SaveAs($receivePath, 50)
    }
    finally {
        $wbInbox.Close($true)
        Release-ComObject $wbInbox
    }

    $wbShip = $excel.Workbooks.Add()
    try {
        $wsShip = $wbShip.Worksheets(1)
        $wsShip.Name = "InboxShip"
        Add-Table -Worksheet $wsShip -TableName "tblInboxShip" -Headers $inboxHeaders -Rows @(
            @("EVT-SHIP-FIXTURE-001", "", "", "SHIP", (Get-Date).ToUniversalTime(), "WH1", "S1", "user1", "", "", "", "fixture ship row", '[{"Row":101,"SKU":"SKU-001","Qty":2,"Location":"DOCK","Note":"fixture ship"}]', "NEW", 0, "", "", "")
        )
        $wbShip.SaveAs($shipPath, 50)
    }
    finally {
        $wbShip.Close($true)
        Release-ComObject $wbShip
    }

    $wbProd = $excel.Workbooks.Add()
    try {
        $wsProd = $wbProd.Worksheets(1)
        $wsProd.Name = "InboxProd"
        Add-Table -Worksheet $wsProd -TableName "tblInboxProd" -Headers $inboxHeaders -Rows @(
            @("EVT-PROD-FIXTURE-001", "", "", "PROD_COMPLETE", (Get-Date).ToUniversalTime(), "WH1", "S1", "user1", "", "", "", "fixture prod row", '[{"Row":201,"SKU":"SKU-FG","Qty":1,"Location":"FG","Note":"fixture prod","IoType":"COMPLETE"}]', "NEW", 0, "", "", "")
        )
        $wbProd.SaveAs($prodPath, 50)
    }
    finally {
        $wbProd.Close($true)
        Release-ComObject $wbProd
    }

    Write-Output "PHASE2_FIXTURES_OK"
    Write-Output "INVENTORY_XLSB=$inventoryPath"
    Write-Output "RECEIVE_XLSB=$receivePath"
    Write-Output "SHIP_XLSB=$shipPath"
    Write-Output "PROD_XLSB=$prodPath"
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}
