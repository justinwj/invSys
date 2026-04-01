Param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$harnessPath = Join-Path $fixtures "Phase2_TestHarness.xlsm"
$cfgXlsb = Join-Path $fixtures "WH1.invSys.Config.xlsb"
$authXlsb = Join-Path $fixtures "WH1.invSys.Auth.xlsb"
$inventoryXlsb = Join-Path $fixtures "WH1.invSys.Data.Inventory.xlsb"
$inboxXlsb = Join-Path $fixtures "invSys.Inbox.Receiving.S1.xlsb"

& (Join-Path $repo "tools/run_phase2_excel_validation.ps1") -RepoRoot $repo | Out-Host

function Release-ComObject {
    Param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Run-TestFunction {
    Param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$FunctionName
    )

    $fullMacro = "'$WorkbookName'!$FunctionName"
    return $Excel.Run($fullMacro)
}

$excel = $null
$harness = $null
$opened = @()

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $true
    $excel.EnableEvents = $false

    foreach ($path in @($cfgXlsb, $authXlsb, $inventoryXlsb, $inboxXlsb, $harnessPath)) {
        if (Test-Path $path) {
            $wb = $excel.Workbooks.Open($path)
            $opened += $wb
            if ($path -eq $harnessPath) {
                $harness = $wb
            }
        }
    }

    if ($null -eq $harness) {
        throw "Harness workbook not found at $harnessPath"
    }

    $tests = @(
        "Run_TestCoreLockManager_TestAcquireReleaseLock_Lifecycle",
        "Run_TestInventoryApply_TestApplyReceive_ValidEvent",
        "Run_TestCoreProcessor_TestRunBatch_ProcessesInboxRow",
        "Run_TestCoreProcessor_TestRunBatch_ProcessesShipRow",
        "Run_TestCoreProcessor_TestRunBatch_ProcessesProdCompleteRow"
    )

    foreach ($testName in $tests) {
        Write-Host "RUNNING_VISIBLE_TEST=$testName"
        [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName $testName)
        Start-Sleep -Seconds 2
    }

    $excel.EnableEvents = $true
    Write-Host "VISIBLE_PHASE2_RUN_COMPLETE"
    Write-Host "Excel remains open for inspection. Close it manually when finished."
}
catch {
    Write-Host ("VISIBLE_PHASE2_RUN_FAILED=" + $_.Exception.Message)
    Write-Host "Excel remains open for inspection if it was launched."
}
finally {
    foreach ($wb in $opened) {
        Release-ComObject $wb
    }
    Release-ComObject $harness
    Release-ComObject $excel
}
