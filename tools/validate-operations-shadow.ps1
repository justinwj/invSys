[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$ShadowRoot = "",
    [string]$ReportTimestampUtc = "2026-07-27T23:30:00Z"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path $RepoRoot).Path
$createdShadowRoot = $false
$rows = New-Object System.Collections.Generic.List[object]
$excel = $null
$openedWorkbooks = New-Object System.Collections.Generic.List[object]
$operatorWorkbooks = New-Object System.Collections.Generic.List[object]

function Release-ComObject {
    param([object]$Object)

    if ($null -ne $Object) {
        try {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
        }
        catch {}
    }
}

function Add-Result {
    param(
        [string]$Check,
        [bool]$Passed,
        [string]$Detail
    )

    $safeDetail = [regex]::Replace(
        [string]$Detail,
        '(?i)[A-Z]:\\[^|;\r\n]+',
        '<redacted-path>'
    )
    $rows.Add([pscustomobject]@{
        Check = $Check
        Passed = $Passed
        Detail = $safeDetail.Replace("|", "/")
    })
}

function Run-WorkbookMacro {
    param(
        [object]$Excel,
        [object]$Workbook,
        [string]$MacroName,
        [object[]]$Arguments = @()
    )

    $qualified = "'" + $Workbook.Name.Replace("'", "''") +
                 "'!" + $MacroName
    switch ($Arguments.Count) {
        0 { return $Excel.Run($qualified) }
        1 { return $Excel.Run($qualified, $Arguments[0]) }
        2 { return $Excel.Run($qualified, $Arguments[0], $Arguments[1]) }
        default {
            throw "Shadow validator supports at most two macro arguments."
        }
    }
}

function Test-ComponentPresent {
    param(
        [object]$Workbook,
        [string]$ComponentName
    )

    try {
        $component = $Workbook.VBProject.VBComponents.Item($ComponentName)
        $present = $null -ne $component
        Release-ComObject $component
        return $present
    }
    catch {
        return $false
    }
}

function Test-NoBrokenReferences {
    param([object]$Workbook)

    foreach ($reference in $Workbook.VBProject.References) {
        if ($reference.IsBroken) {
            return $false
        }
    }
    return $true
}

try {
    if ([string]::IsNullOrWhiteSpace($ShadowRoot)) {
        $ShadowRoot = Join-Path ([IO.Path]::GetTempPath()) (
            "invsys-operations-shadow-" + [Guid]::NewGuid().ToString("N")
        )
        $createdShadowRoot = $true
        & (Join-Path $repo "tools\build-operations-shadow.ps1") `
            -RepoRoot $repo `
            -OutputDirectory $ShadowRoot `
            -ReportTimestampUtc $ReportTimestampUtc
    }
    elseif (-not [IO.Path]::IsPathRooted($ShadowRoot)) {
        $ShadowRoot = Join-Path $repo $ShadowRoot
    }
    $ShadowRoot = [IO.Path]::GetFullPath($ShadowRoot)

    $expectedFiles = @(
        "invSys.Core.xlam",
        "invSys.Inventory.Domain.xlam",
        "invSys.Designs.Domain.xlam",
        "invSys.Operations.xlam"
    )
    $missingFiles = @($expectedFiles | Where-Object {
        -not (Test-Path -LiteralPath (
            Join-Path $ShadowRoot $_
        ) -PathType Leaf)
    })
    $buildOutputDetail = if ($missingFiles.Count -eq 0) {
        "Core, both Domain packages, and Operations are present."
    }
    else {
        "Missing=" + ($missingFiles -join ",")
    }
    Add-Result "Shadow.BuildOutputs" `
        ($missingFiles.Count -eq 0) `
        $buildOutputDetail
    if ($missingFiles.Count -gt 0) {
        throw "Required shadow packages are missing."
    }

    $collisionReportPath = Join-Path $repo `
        "reports\operations-shadow\collision-report.json"
    $collisionReport = Get-Content -Raw -LiteralPath $collisionReportPath |
        ConvertFrom-Json
    $collisionGreen = (
        [int]$collisionReport.summary.componentCollisionCount -eq 0 -and
        [int]$collisionReport.summary.publicProcedureCollisionCount -eq 0 -and
        [int]$collisionReport.summary.ribbonCallbackCollisionCount -eq 0 -and
        [int]$collisionReport.summary.unresolvedCollisionCount -eq 0
    )
    Add-Result "Shadow.CollisionReport" $collisionGreen (
        "Components=$($collisionReport.summary.componentCollisionCount);" +
        "PublicProcedures=$($collisionReport.summary.publicProcedureCollisionCount);" +
        "RibbonCallbacks=$($collisionReport.summary.ribbonCallbackCollisionCount);" +
        "Unresolved=$($collisionReport.summary.unresolvedCollisionCount)"
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $opsArchive = [IO.Compression.ZipFile]::OpenRead(
        (Join-Path $ShadowRoot "invSys.Operations.xlam")
    )
    try {
        $ribbonParts = @($opsArchive.Entries | Where-Object {
            $_.FullName -like "customUI/*"
        })
        Add-Result "Shadow.NoRibbonRegistration" `
            ($ribbonParts.Count -eq 0) `
            "Disposable shadow has no RibbonX part."
    }
    finally {
        $opsArchive.Dispose()
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    $workbookByName = @{}
    foreach ($fileName in $expectedFiles) {
        $workbook = $excel.Workbooks.Open(
            (Join-Path $ShadowRoot $fileName),
            0,
            $true
        )
        $openedWorkbooks.Add($workbook) | Out-Null
        $workbook.IsAddin = $true
        $workbookByName[$fileName] = $workbook
    }
    $opsWorkbook = $workbookByName["invSys.Operations.xlam"]
    Add-Result "Shadow.LoadOrder" `
        ($workbookByName.Count -eq 4) `
        "Loaded Core, Inventory Domain, Designs Domain, then Operations."
    Add-Result "Shadow.References" `
        (Test-NoBrokenReferences $opsWorkbook) `
        "Operations contains no broken VBA references."

    $requiredComponents = @(
        "modOperationsInit",
        "modReceivingInit",
        "modTS_Received",
        "frmReceiving",
        "modProductionInit",
        "mProduction",
        "frmProduction",
        "modShippingInit",
        "modTS_Shipments",
        "frmShipmentsTally",
        "cReceivingAppEvents",
        "cProductionAppEvents",
        "cShippingAppEvents"
    )
    $missingComponents = @($requiredComponents | Where-Object {
        -not (Test-ComponentPresent $opsWorkbook $_)
    })
    Add-Result "Shadow.RoleComponents" `
        ($missingComponents.Count -eq 0) `
        $(if ($missingComponents.Count -eq 0) {
            "All three role module/form sets are present."
        }
        else {
            "Missing=" + ($missingComponents -join ",")
        })

    $forbiddenComponents = @(
        "modReceivingAutoOpen",
        "modProductionAutoOpen",
        "modShippingAutoOpen",
        "ufDynItemSearchTemplate",
        "modRibbonGenerated"
    )
    $presentForbidden = @($forbiddenComponents | Where-Object {
        Test-ComponentPresent $opsWorkbook $_
    })
    Add-Result "Shadow.ExcludedComponents" `
        ($presentForbidden.Count -eq 0) `
        $(if ($presentForbidden.Count -eq 0) {
            "Standalone startup wrappers, template form, and Ribbon callbacks are absent."
        }
        else {
            "Unexpected=" + ($presentForbidden -join ",")
        })

    $startupReport = [string](Run-WorkbookMacro `
        -Excel $excel `
        -Workbook $opsWorkbook `
        -MacroName "modOperationsInit.OperationsShadowStartupForTest")
    $startupGreen = $startupReport.StartsWith(
        "OK|",
        [StringComparison]::Ordinal
    )
    Add-Result "Shadow.Startup" `
        $startupGreen `
        $startupReport

    $targetSpecs = @(
        @{
            Name = "Receiving"
            File = "WH1.S1.Receiving.Operator.xlsx"
            Macro = "modTS_Received.ReceivingFormInitializeSmokeForWorkbook"
        },
        @{
            Name = "Production"
            File = "WH1.S1.Production.Operator.xlsx"
            Macro = "mProduction.ProductionFormInitializeSmokeForWorkbook"
        },
        @{
            Name = "Shipping"
            File = "WH1.S1.Shipping.Operator.xlsx"
            Macro = "modTS_Shipments.ShippingFormInitializeSmokeForWorkbook"
        }
    )
    $formInitializeGreen = $true
    foreach ($spec in $targetSpecs) {
        $targetWorkbook = $excel.Workbooks.Add()
        $operatorWorkbooks.Add($targetWorkbook) | Out-Null
        $targetWorkbook.SaveAs(
            (Join-Path $ShadowRoot $spec.File),
            51
        )
        $targetWorkbook.Activate()
        $formReport = [string](Run-WorkbookMacro `
            -Excel $excel `
            -Workbook $opsWorkbook `
            -MacroName $spec.Macro `
            -Arguments @($targetWorkbook))
        $formGreen = $formReport.StartsWith(
            "OK|",
            [StringComparison]::Ordinal
        )
        if (-not $formGreen) {
            $formInitializeGreen = $false
        }
        Add-Result "Shadow.$($spec.Name)FormInitialize" `
            $formGreen `
            $formReport
    }
    Add-Result "Shadow.Compile" `
        ($startupGreen -and $formInitializeGreen) `
        "Excel compiled and executed unified startup plus all three role-form initialization paths."

    $loadedNames = @(
        $excel.Workbooks |
            ForEach-Object { [string]$_.Name }
    )
    $legacyLoaded = @(
        "invSys.Receiving.xlam",
        "invSys.Production.xlam",
        "invSys.Shipping.xlam"
    ) | Where-Object { $_ -in $loadedNames }
    Add-Result "Shadow.LegacyNotLoadedBesideOperations" `
        (@($legacyLoaded).Count -eq 0) `
        $(if (@($legacyLoaded).Count -eq 0) {
            "No standalone role XLAM is loaded in the shadow session."
        }
        else {
            "Loaded=" + (@($legacyLoaded) -join ",")
        })
}
catch {
    Add-Result "Shadow.ValidatorException" $false $_.Exception.Message
}
finally {
    if ($null -ne $excel) {
        for ($index = $operatorWorkbooks.Count - 1; $index -ge 0; $index--) {
            $workbook = $operatorWorkbooks[$index]
            try { $workbook.Close($false) } catch {}
            Release-ComObject $workbook
        }
        for ($index = $openedWorkbooks.Count - 1; $index -ge 0; $index--) {
            $workbook = $openedWorkbooks[$index]
            try { $workbook.Close($false) } catch {}
            Release-ComObject $workbook
        }
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
        $excel = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($createdShadowRoot -and
        -not [string]::IsNullOrWhiteSpace($ShadowRoot)) {
        $resolvedShadowRoot = [IO.Path]::GetFullPath($ShadowRoot)
        $tempPrefix = [IO.Path]::GetFullPath(
            [IO.Path]::GetTempPath()
        ).TrimEnd("\") + "\"
        $leafName = [IO.Path]::GetFileName(
            $resolvedShadowRoot.TrimEnd("\")
        )
        if ($resolvedShadowRoot.StartsWith(
            $tempPrefix,
            [StringComparison]::OrdinalIgnoreCase
        ) -and $leafName -like "invsys-operations-shadow-*") {
            Remove-Item -LiteralPath $resolvedShadowRoot -Recurse -Force `
                -ErrorAction SilentlyContinue
        }
    }
}

$passed = @($rows | Where-Object Passed).Count
$failed = $rows.Count - $passed
$resultPath = Join-Path $repo `
    "tests\unit\slice6_shadow_validation_results.md"
$lines = @(
    "# Slice 6 Operations Shadow Validation Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Detail |",
    "|---|---|---|"
)
foreach ($row in $rows) {
    $result = if ($row.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($row.Check) | $result | $($row.Detail) |"
}
[IO.File]::WriteAllText(
    $resultPath,
    (($lines -join "`n") + "`n"),
    (New-Object Text.UTF8Encoding($false))
)

if ($failed -gt 0) {
    Write-Host "OPERATIONS_SHADOW_VALIDATION_FAILED"
}
else {
    Write-Host "OPERATIONS_SHADOW_VALIDATION_OK"
}
Write-Host "RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($rows.Count)"

if ($failed -gt 0) {
    throw "Operations shadow validation failed."
}
