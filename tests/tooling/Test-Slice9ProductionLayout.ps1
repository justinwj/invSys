[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$formPath = Join-Path $repo "src/Production/Forms/frmProduction.frm"
$windowPath = Join-Path $repo "src/Production/Modules/modProductionFormWindow.bas"
$buildPath = Join-Path $repo "tools/build-xlam.ps1"
$scannerPath = Join-Path $repo "tools/inventory-vba-surface.ps1"
$schemaPath = Join-Path $repo "tools/contracts/implementation-manifest.schema.json"
$manifestPath = Join-Path $repo "reports/static-baseline/implementation-manifest.json"
$anchorItemPath = Join-Path $repo "src/Operations/ClassModules/cOperationsAnchorItem.cls"
$anchorManagerPath = Join-Path $repo "src/Operations/ClassModules/cOperationsAnchorManager.cls"
$layoutPath = Join-Path $repo "src/Operations/Modules/modOperationsLayout.bas"

$formText = Get-Content -Raw -LiteralPath $formPath
$windowText = Get-Content -Raw -LiteralPath $windowPath
$buildText = Get-Content -Raw -LiteralPath $buildPath
$scannerText = Get-Content -Raw -LiteralPath $scannerPath
$schemaText = Get-Content -Raw -LiteralPath $schemaPath

$results = New-Object System.Collections.Generic.List[object]
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Detail)
    $results.Add([pscustomobject]@{ Check = $Name; Passed = $Passed; Detail = $Detail })
}

Add-Check "Production.Layout.OperationsAnchorTypes" `
    ((Test-Path -LiteralPath $anchorItemPath -PathType Leaf) -and
     (Test-Path -LiteralPath $anchorManagerPath -PathType Leaf) -and
     (Test-Path -LiteralPath $layoutPath -PathType Leaf)) `
    "Production layout support must live inside the Operations VBA project, not cross the Core XLAM boundary."

Add-Check "Production.Layout.DeclarativeRegistration" `
    (($formText -match '(?i)Private\s+mLayout\s+As\s+cOperationsAnchorManager') -and
     ($formText -match '(?i)ConfigureProductionAnchors') -and
     ($formText -match '(?i)mLayout\.ApplyAnchoredLayout')) `
    "The form must register declarative anchors once and apply them from resize events."

$resizeBody = ""
$resizeMatch = [regex]::Match(
    $formText,
    '(?is)Private\s+Sub\s+ResizeProductionLayout\s*\(\s*\)(?<body>.*?)End\s+Sub'
)
if ($resizeMatch.Success) { $resizeBody = $resizeMatch.Groups["body"].Value }
Add-Check "Production.Layout.NoResizeCoordinateArithmetic" `
    (($resizeBody -notmatch '(?i)\.Move\b') -and
     ($resizeBody -notmatch '(?i)\.(Left|Top|Width|Height)\s*=') -and
     ($formText -notmatch '(?i)Private\s+Sub\s+ResizeProductionPages\b')) `
    "Resize callbacks may apply anchors but must not contain one-off per-control coordinate arithmetic."

Add-Check "Production.Layout.GeometryContract" `
    (($formText -match '(?i)PRODUCTION_MIN_WIDTH\s+As\s+Double\s*=\s*1110') -and
     ($formText -match '(?i)PRODUCTION_MIN_HEIGHT\s+As\s+Double\s*=\s*690') -and
     ($formText -match '(?i)PRODUCTION_LAYOUT_TEST_MAX_WIDTH\s+As\s+Double\s*=\s*1350') -and
     ($formText -match '(?i)PRODUCTION_LAYOUT_TEST_MAX_HEIGHT\s+As\s+Double\s*=\s*750')) `
    "Minimum/default/expanded acceptance sizes must be explicit and stable."

Add-Check "Production.Layout.NativeWindowBehavior" `
    (($formText -match '(?i)modProductionFormWindow\.EnableResizable\s+Me\s*,\s*True\s*,\s*True') -and
     ($windowText -match '(?i)WS_THICKFRAME') -and
     ($windowText -match '(?i)WS_MINIMIZEBOX') -and
     ($windowText -match '(?i)WS_MAXIMIZEBOX') -and
     ($windowText -match '(?i)ApplyDpiLayoutZoom') -and
     ($windowText -match '(?i)GetDpiForWindow') -and
     ($windowText -match '(?i)DiagnoseWindowStyle')) `
    "The operator form must retain Windows resize, minimize, maximize, DPI-normalized layout, and diagnostic support."

Add-Check "Production.Layout.LegacyAndShadowComposition" `
    (($buildText -match '(?is)Key\s*=\s*"Production".*?SourceDirs\s*=\s*@\(.*?src/Operations.*?src/Production') -and
     ($buildText -match '(?is)Key\s*=\s*"Production".*?ExcludeFiles\s*=\s*@\(.*?modOperationsInit\.bas')) `
    "Until Slice 13 cutover, both the legacy Production package and Operations shadow must compile the shared Operations-local layout types."

Add-Check "Production.Layout.StaticManifestContract" `
    (($scannerText -match '(?i)FormLayout') -and
     ($schemaText -match '(?i)\"layout\"') -and
     ($formText -match '(?i)@FormLayout\s+Strategy=WINDOWS_API_ANCHORS')) `
    "The static manifest must expose declared form layout strategy and acceptance geometry."

$manifestHasLayout = $false
if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
    $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
    $productionForm = @($manifest.forms | Where-Object { $_.name -eq "frmProduction" }) | Select-Object -First 1
    if ($null -ne $productionForm -and $null -ne $productionForm.layout) {
        $manifestHasLayout =
            ([string]$productionForm.layout.strategy -eq "WINDOWS_API_ANCHORS") -and
            ([double]$productionForm.layout.minimum.width -eq 1110) -and
            ([double]$productionForm.layout.minimum.height -eq 690) -and
            ([double]$productionForm.layout.expandedTest.width -eq 1350) -and
            ([double]$productionForm.layout.expandedTest.height -eq 750)
    }
}
Add-Check "Production.Layout.StaticManifestEvidence" $manifestHasLayout `
    "The regenerated baseline must record Production's anchor strategy and minimum/expanded geometry."

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.Passed })
Write-Host ("Slice 9 Production layout contract: {0} passed, {1} failed" -f
    ($results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
