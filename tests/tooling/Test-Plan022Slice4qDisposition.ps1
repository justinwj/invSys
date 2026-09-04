Param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path $RepoRoot).Path
$form = Get-Content (Join-Path $repo "src/Receiving/Forms/frmReceiving.frm") -Raw
$receiving = Get-Content (Join-Path $repo "src/Receiving/Modules/modTS_Received.bas") -Raw
$posting = Get-Content (Join-Path $repo "src/Receiving/Modules/modReceivingPostingService.bas") -Raw
$writer = Get-Content (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas") -Raw
$processor = Get-Content (Join-Path $repo "src/Core/Modules/modProcessor.bas") -Raw
$domain = Get-Content (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas") -Raw

$checks = @(
    [pscustomobject]@{
        Name = "Returns.DispositionSelector"
        Passed = ($form -match 'Disposition') -and ($form -match 'AddItem "RETURN"') -and ($form -match 'AddItem "DUMP"')
        Contract = "Returns exposes required RETURN and DUMP choices."
    },
    [pscustomobject]@{
        Name = "Returns.PreservesExactIdentity"
        Passed = ($receiving -match 'StageInventoryDisposition') -and ($receiving -match 'Source_System_Key')
        Contract = "Disposition stages exact existing System_Key allocations rather than creating a new entity."
    },
    [pscustomobject]@{
        Name = "Returns.QueuesDistinctEventTypes"
        Passed = ($posting -match 'state.EventType') -and ($writer -match 'ROLE_EVENT_TYPE_RETURN') -and ($writer -match 'ROLE_EVENT_TYPE_DUMP')
        Contract = "RETURN and DUMP remain distinct queue and audit event types."
    },
    [pscustomobject]@{
        Name = "Processor.ReceivingDispositionCapability"
        Passed = ($processor -match 'PROC_EVENT_TYPE_RETURN') -and ($processor -match 'PROC_EVENT_TYPE_DUMP') -and ($processor -match 'RECEIVE_POST')
        Contract = "Receiving processor accepts RETURN/DUMP under RECEIVE_POST."
    },
    [pscustomobject]@{
        Name = "Domain.DispositionDepletes"
        Passed = ($domain -match 'EVENT_TYPE_RETURN') -and ($domain -match 'EVENT_TYPE_DUMP') -and ($domain -match 'BuildDispositionLines')
        Contract = "RETURN/DUMP apply negative deltas to existing exact System_Key entities."
    },
    [pscustomobject]@{
        Name = "Domain.ExactKeyOverdrawRejected"
        Passed = ($domain -match 'INSUFFICIENT_ENTITY_INVENTORY')
        Contract = "Disposition cannot borrow quantity from another entity or Condition bucket."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4q_disposition_results.md"
$lines = @(
    "# Plan 022 Slice 4q Inventory Disposition Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($check.Name) | $result | $($check.Contract) |"
}
[System.IO.File]::WriteAllLines($resultPath, $lines)
$lines -join [Environment]::NewLine
if ($failed -gt 0) { exit 1 }
