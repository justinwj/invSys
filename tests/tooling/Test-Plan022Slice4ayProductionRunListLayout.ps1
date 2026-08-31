[CmdletBinding()]
param(
    [string]$RepoRoot = "."
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$formPath = Join-Path $repo "src/Production/Forms/frmProduction.frm"
$modulePath = Join-Path $repo "src/Production/Modules/mProduction.bas"
$anchorManagerPath = Join-Path $repo "src/Operations/ClassModules/cOperationsAnchorManager.cls"
$anchorItemPath = Join-Path $repo "src/Operations/ClassModules/cOperationsAnchorItem.cls"

$formText = Get-Content -Raw -LiteralPath $formPath
$moduleText = Get-Content -Raw -LiteralPath $modulePath
$anchorManagerText = Get-Content -Raw -LiteralPath $anchorManagerPath
$anchorItemText = Get-Content -Raw -LiteralPath $anchorItemPath

$results = New-Object System.Collections.Generic.List[object]
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Contract)
    $results.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Contract = $Contract
    })
}

Add-Check "Production.RunList.ReadableBaselineRows" `
    (($formText -match '(?i)RUN_LIST_CHECK_MIN_HEIGHT\s+As\s+Single\s*=\s*102') -and
     ($formText -match '(?i)RUN_LIST_INSTRUCTIONS_MIN_HEIGHT\s+As\s+Single\s*=\s*52') -and
     ($formText -match '(?i)RUN_LIST_CHECK_READABLE_HEIGHT\s+As\s+Single\s*=\s*96') -and
     ($formText -match '(?i)RUN_LIST_INSTRUCTIONS_READABLE_HEIGHT\s+As\s+Single\s*=\s*48') -and
     ($formText -match '(?is)AddList\(pg,\s*"lstManagerCheck".*?RUN_LIST_CHECK_MIN_HEIGHT') -and
     ($formText -match '(?is)AddList\(pg,\s*"lstRunInstructions".*?RUN_LIST_INSTRUCTIONS_MIN_HEIGHT')) `
    "Production Run - List must show at least eight Inventory Check rows and four Selected Process Instruction rows at its readable baseline."

Add-Check "Production.RunList.ProportionalVerticalBands" `
    (($anchorManagerText -match '(?i)RegisterProportionalVerticalControl') -and
     ($anchorItemText -match '(?i)CaptureProportionalVerticalAnchor') -and
     ($formText -match '(?i)ConfigureRunListProportionalLayout') -and
     ($formText -match '(?is)ConfigureRunListProportionalLayout.*?mLstLoaderRecipes.*?mLstLoaderLines.*?mLstRunPalette.*?mLstManagerCheck.*?mLstRunInstructions.*?mLstManagerOutput')) `
    "Every visible Production Run - List list must receive a declarative proportional vertical resize registration, rather than leaving growth solely to Production Output."

Add-Check "Production.RunList.InstructionsLeftAnchored" `
    ($formText -match '(?is)ConfigureRunListProportionalLayout.*?mLstRunInstructions\s*,\s*leftRightTop') `
    "Selected Process Instructions must remain fixed to the left edge while its right edge follows form width."

Add-Check "Production.RunList.HeadersFollowBands" `
    (($formText -match '(?i)RegisterRunListHeaderBand') -and
     ($formText -match '(?is)ConfigureRunListProportionalLayout.*?RegisterRunListHeaderBand\s+"ManagerCheck".*?RegisterRunListHeaderBand\s+"RunInstructions".*?RegisterRunListHeaderBand\s+"ManagerOutput"')) `
    "Inventory Check, Selected Process Instructions, and Production Output headings and column labels must move with their associated list bands."

Add-Check "Production.RunList.PackagedPublicLauncherSeam" `
    (($formText -match '(?i)Public\s+Function\s+TestRunListResponsiveLayoutReportForSize') -and
     ($moduleText -match '(?i)Public\s+Function\s+RunProductionRunListResponsiveLayoutTest') -and
     ($moduleText -match '(?is)RunProductionRunListResponsiveLayoutTest.*?BtnOpenProductionForm')) `
    "The packaged geometry check must enter the same public Production launcher used by operators and report the actual Run List geometry."

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.Passed })
Write-Host ("Plan 022 Slice 4ay Production Run List layout: {0} passed, {1} failed" -f `
    ($results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
