[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OutputDirectory = "",
    [string]$ResultPath = ""
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repo "tests/integration/slice9-layout"
}
if ([string]::IsNullOrWhiteSpace($ResultPath)) {
    $ResultPath = Join-Path $repo "tests/integration/slice9_layout_results.md"
}

$corePath = Join-Path $repo "deploy/current/invSys.Core.xlam"
$productionPath = Join-Path $repo "deploy/current/invSys.Operations.xlam"
foreach ($required in @($corePath, $productionPath)) {
    if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
        throw "Required packaged add-in is missing: $required"
    }
}

$preexistingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue)
if ($preexistingExcel.Count -gt 0) {
    throw "Close Excel before running the Slice 9 packaged layout validator."
}

if (-not ("Slice9NativeWindow" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class Slice9NativeWindow {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr parameter);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr FindWindow(string className, string windowName);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")]
    public static extern bool IsZoomed(IntPtr hwnd);
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maximumCount);
    public static IntPtr FindVisibleOwnedWindow(uint expectedProcessId, string titlePart) {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hwnd, IntPtr parameter) {
            uint ownerProcessId;
            GetWindowThreadProcessId(hwnd, out ownerProcessId);
            if (ownerProcessId != expectedProcessId || !IsWindowVisible(hwnd)) return true;
            StringBuilder title = new StringBuilder(512);
            GetWindowText(hwnd, title, title.Capacity);
            if (title.ToString().IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0) {
                found = hwnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
"@
}

Add-Type -AssemblyName System.Drawing

function Release-ComObject {
    param([object]$Value)
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Wait-ForLayoutWindow {
    param([int]$ExpectedProcessId)
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        $handle = [Slice9NativeWindow]::FindVisibleOwnedWindow(
            [uint32]$ExpectedProcessId, "Production Layout Validation")
        if ($handle -ne [IntPtr]::Zero) { return $handle }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Production layout validation window did not appear."
}

function Save-WindowScreenshot {
    param([IntPtr]$Handle, [string]$Path)
    $rect = New-Object Slice9NativeWindow+RECT
    if (-not [Slice9NativeWindow]::GetWindowRect($Handle, [ref]$rect)) {
        throw "Could not read the Production form window bounds."
    }
    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top
    if ($width -le 0 -or $height -le 0) {
        throw "Production form window bounds were invalid: ${width}x${height}."
    }

    [void][Slice9NativeWindow]::ShowWindow($Handle, 9)
    [void][Slice9NativeWindow]::SetForegroundWindow($Handle)
    Start-Sleep -Milliseconds 150
    $bitmap = New-Object Drawing.Bitmap($width, $height)
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen($rect.Left, $rect.Top, 0, 0, $bitmap.Size)
    }
    finally {
        $graphics.Dispose()
    }
    try {
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $bitmap.Dispose()
    }
}

function Assert-HealthyGeometryReport {
    param([string]$Name, [string]$Report)
    if (-not $Report.StartsWith("OK|", [StringComparison]::OrdinalIgnoreCase) -or
        $Report -notmatch '\|OutOfBounds=0\|' -or
        $Report -notmatch '\|Overlap=0\|' -or
        $Report -notmatch 'Resizable=True' -or
        $Report -notmatch 'Minimize=True' -or
        $Report -notmatch 'Maximize=True') {
        throw "$Name geometry failed: $Report"
    }
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$cases = @(
    [pscustomobject]@{ Name = "minimum"; Width = 900.0; Height = 600.0; Screenshot = "production-minimum.png" },
    [pscustomobject]@{ Name = "default"; Width = 1110.0; Height = 690.0; Screenshot = "production-default.png" },
    [pscustomobject]@{ Name = "expanded"; Width = 1350.0; Height = 750.0; Screenshot = "production-expanded.png" }
)

$excel = $null
$core = $null
$production = $null
$excelProcess = $null
$geometryRows = New-Object System.Collections.Generic.List[object]
$nativeRows = New-Object System.Collections.Generic.List[object]

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    [uint32]$excelProcessId = 0
    [void][Slice9NativeWindow]::GetWindowThreadProcessId(
        [IntPtr]([int64]$excel.Hwnd), [ref]$excelProcessId)
    $excelProcess = Get-Process -Id ([int]$excelProcessId)

    $core = $excel.Workbooks.Open($corePath, 0, $true)
    $production = $excel.Workbooks.Open($productionPath, 0, $true)
    $showMacro = "'" + $production.Name + "'!mProduction.ShowProductionLayoutForValidation"
    $currentMacro = "'" + $production.Name + "'!mProduction.CurrentProductionLayoutValidationReport"
    $closeMacro = "'" + $production.Name + "'!mProduction.CloseProductionLayoutValidation"

    foreach ($case in $cases) {
        Write-Host ("Geometry case: " + $case.Name)
        $pageReports = New-Object System.Collections.Generic.List[string]
        for ($pageIndex = 0; $pageIndex -lt 4; $pageIndex++) {
            Write-Host ("  Page " + $pageIndex)
            $report = [string]$excel.Run(
                $showMacro, $case.Width, $case.Height, [int]$pageIndex)
            Assert-HealthyGeometryReport -Name "$($case.Name)/page-$pageIndex" -Report $report
            $pageReports.Add($report)
        }

        $report = [string]$excel.Run($showMacro, $case.Width, $case.Height, 2)
        Assert-HealthyGeometryReport -Name "$($case.Name)/screenshot" -Report $report
        $handle = Wait-ForLayoutWindow -ExpectedProcessId ([int]$excelProcessId)
        $screenshotPath = Join-Path $OutputDirectory $case.Screenshot
        Write-Host ("  Capturing " + $screenshotPath)
        Save-WindowScreenshot -Handle $handle -Path $screenshotPath
        Write-Host "  Capture complete"
        $geometryRows.Add([pscustomobject]@{
            Case = $case.Name
            Requested = "$($case.Width)x$($case.Height)"
            Pages = 4
            Report = $report
            Screenshot = "slice9-layout/$($case.Screenshot)"
        })
    }

    Write-Host "Native window transitions"
    $handle = Wait-ForLayoutWindow -ExpectedProcessId ([int]$excelProcessId)
    Write-Host "  Minimize"
    [void][Slice9NativeWindow]::ShowWindow($handle, 6)
    Start-Sleep -Milliseconds 300
    $minimized = [Slice9NativeWindow]::IsIconic($handle)
    Write-Host ("  Minimized=" + $minimized)
    $nativeRows.Add([pscustomobject]@{ Action = "Minimize"; Passed = $minimized })
    if (-not $minimized) { throw "Production form did not enter the minimized state." }

    Write-Host "  Restore"
    [void][Slice9NativeWindow]::ShowWindow($handle, 9)
    Start-Sleep -Milliseconds 300
    $restored = -not [Slice9NativeWindow]::IsIconic($handle)
    Write-Host ("  Restored=" + $restored)
    $nativeRows.Add([pscustomobject]@{ Action = "Restore"; Passed = $restored })
    if (-not $restored) { throw "Production form did not restore from minimized state." }

    Write-Host "  Maximize"
    [void][Slice9NativeWindow]::ShowWindow($handle, 3)
    Start-Sleep -Milliseconds 500
    $maximized = [Slice9NativeWindow]::IsZoomed($handle)
    Write-Host ("  Maximized=" + $maximized)
    $nativeRows.Add([pscustomobject]@{ Action = "Maximize"; Passed = $maximized })
    if (-not $maximized) { throw "Production form did not enter the maximized state." }

    for ($pageIndex = 0; $pageIndex -lt 4; $pageIndex++) {
        Write-Host ("  Maximized page " + $pageIndex)
        $maximizedReport = [string]$excel.Run($currentMacro, [int]$pageIndex)
        Assert-HealthyGeometryReport -Name "maximized/page-$pageIndex" -Report $maximizedReport
    }

    Write-Host "  Restore after maximize"
    [void][Slice9NativeWindow]::ShowWindow($handle, 9)
    Start-Sleep -Milliseconds 300
    $restoredAgain = -not [Slice9NativeWindow]::IsIconic($handle) -and
                     -not [Slice9NativeWindow]::IsZoomed($handle)
    Write-Host ("  RestoredAfterMaximize=" + $restoredAgain)
    $nativeRows.Add([pscustomobject]@{ Action = "RestoreAfterMaximize"; Passed = $restoredAgain })
    if (-not $restoredAgain) { throw "Production form did not restore from maximized state." }

    [void]$excel.Run($closeMacro)
}
finally {
    if ($null -ne $production) { try { $production.Close($false) } catch {} }
    if ($null -ne $core) { try { $core.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Release-ComObject $production
    Release-ComObject $core
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($null -ne $excelProcess) {
        try { $excelProcess.WaitForExit(5000) | Out-Null } catch {}
    }
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("# Slice 9 Production Layout Runtime Results")
$lines.Add("")
$lines.Add("- Packaged geometry: PASS (3 sizes x 4 pages; zero out-of-bounds and zero interactive-control overlaps)")
$lines.Add("- Native window behavior: PASS (minimize, restore, maximize, restore)")
$lines.Add("- Screenshots: PASS (minimum/default/expanded production-run page)")
$lines.Add("")
$lines.Add("| Case | Requested points | Pages | Screenshot |")
$lines.Add("|---|---:|---:|---|")
foreach ($row in $geometryRows) {
    $lines.Add("| $($row.Case) | $($row.Requested) | $($row.Pages) | $($row.Screenshot) |")
}
$lines.Add("")
$lines.Add("| Native action | Result |")
$lines.Add("|---|---|")
foreach ($row in $nativeRows) {
    $lines.Add("| $($row.Action) | $(if ($row.Passed) { 'PASS' } else { 'FAIL' }) |")
}
$lines.Add("")
$lines.Add("Representative packaged reports:")
$lines.Add("")
foreach ($row in $geometryRows) {
    $lines.Add("- $($row.Case): ``$($row.Report)``")
}

[IO.File]::WriteAllLines(
    $ResultPath,
    $lines,
    (New-Object Text.UTF8Encoding($false))
)

Write-Host "Slice 9 packaged Production layout validation passed."
Write-Host "Result=$ResultPath"
