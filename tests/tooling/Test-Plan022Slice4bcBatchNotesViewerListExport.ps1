[CmdletBinding()]
param(
    [string]$RepoRoot = ""
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($RepoRoot)) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$productionForm = Get-Content -Raw -LiteralPath (Join-Path $repo 'src\Production\Forms\frmProduction.frm')
$productionRun = Get-Content -Raw -LiteralPath (Join-Path $repo 'src\Production\Modules\modProductionReusableRun.bas')
$productionPublic = Get-Content -Raw -LiteralPath (Join-Path $repo 'src\Production\Modules\mProduction.bas')
$viewerForm = Get-Content -Raw -LiteralPath (Join-Path $repo 'src\Operations\Forms\frmInventoryViewer.frm')
$viewerModule = Get-Content -Raw -LiteralPath (Join-Path $repo 'src\Operations\Modules\modInventoryViewer.bas')
$ribbon = Get-Content -Raw -LiteralPath (Join-Path $repo 'tools\build-xlam.ps1')

$checks = @(
    @{ Name = 'Native-first Production palette'; Passed = $productionForm -match 'Stock / Requirement UOM' -and $productionForm -match 'Native / Requirement Available' -and $productionForm -match '145 pt;230 pt' -and $productionRun -match 'result\(outRow, 8\).*bucket\(2\)' -and $productionRun -match '" / " & RunRecordText\(requirement, "UOM"\)' }
    @{ Name = 'Batch Note public handler contract'; Passed = $productionForm -match 'txtRunBatchNote' -and $productionForm -match 'mBtnManagerCheckIn_Click' -and $productionRun -match 'SetReusableRunBatchNote' -and $productionRun -match 'BatchNote=' -and $productionPublic -match 'RunProductionBatchNoteHandlerContractTest' }
    @{ Name = 'Viewer name and Events geometry'; Passed = $viewerForm -match 'Caption\s*=\s*"Viewer"' -and $viewerForm -match 'ConfigureViewerHeaderGeometry' -and $viewerForm -match 'ViewerEventHeadersAlignedForTest' }
    @{ Name = 'ListBox to Table public action'; Passed = $viewerForm -match 'ListBox->Table' -and $viewerForm -match 'Export ListBox to Table' -and $viewerModule -match 'RunInventoryViewerListBoxTableActionForTest' -and $viewerModule -match 'ExportDeclaredListBoxToTable' -and $viewerModule -match 'ADMIN_MAINT' -and $viewerModule -match 'DeclaredListBoxHeaders' }
    @{ Name = 'Ribbon placement and Viewer label'; Passed = $ribbon -match 'Label\s+=\s+"Viewer"' -and $ribbon -match 'PostStatusMenuButtons' -and $ribbon -match 'btnOperationsCurrentUser' }
)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { Write-Output ("{0}: {1}" -f $_.Name, $(if ($_.Passed) { 'PASS' } else { 'FAIL' })) }
if ($failed.Count -gt 0) { exit 1 }
