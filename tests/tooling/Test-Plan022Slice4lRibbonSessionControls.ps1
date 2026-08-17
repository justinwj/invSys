[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
function Read-Source([string]$relativePath) {
    Get-Content -Raw -LiteralPath (Join-Path $repo $relativePath)
}

$build = Read-Source "tools\build-xlam.ps1"
$role = Read-Source "src\Core\Modules\modRoleEventWriter.bas"
$nas = Read-Source "src\Core\Modules\modNasConnection.bas"
$status = Read-Source "src\Core\Modules\modRibbonRuntimeStatus.bas"
$failures = [System.Collections.Generic.List[string]]::new()
$passes = [System.Collections.Generic.List[string]]::new()

function Check([string]$name, [bool]$passed, [string]$contract) {
    if ($passed) {
        $passes.Add($name)
        Write-Host "PASS $name"
    } else {
        $failures.Add("${name}: ${contract}")
        Write-Host "FAIL $name - $contract"
    }
}

Check "Ribbon.ServerToggleWiring" (
    $build.Contains('GetLabel = "RibbonServerSessionGetLabel"') -and
    $build.Contains('DirectAction = "modRoleEventWriter.ToggleServerSessionForCapability') -and
    $build.Contains('returnedVal = modRibbonRuntimeStatus.GetServerSessionActionLabel()')
) "Operations/Admin must expose one dynamic Server Sign In / Server Sign Out action."

Check "Ribbon.InvSysToggleWiring" (
    $build.Contains('DirectAction = "modRoleEventWriter.ToggleCurrentInvSysUserForCapability') -and
    $build.Contains('returnedVal = modRibbonRuntimeStatus.GetCurrentUserActionLabel()') -and
    $status.Contains('"invSys Sign In"') -and
    $status.Contains('"invSys Sign Out"')
) "The invSys user action must toggle with explicit invSys wording."

Check "Ribbon.NoGenericSessionLabels" (
    $build -notmatch 'Label = "Connect Server"' -and
    $build -notmatch 'Label = "Sign In"' -and
    $build -notmatch 'Label = "Sign Out"'
) "Generic server/user session labels must not remain on Operations or Admin."

Check "Ribbon.TargetSelectionInvalidatesOperations" (
    $status.Contains('ribbon.InvalidateControl "ddOperationsWarehouseTarget"') -and
    $status.Contains('ribbon.InvalidateControl "lblOperationsServerStatus"') -and
    $status.Contains('ribbon.InvalidateControl "lblOperationsAccessStatus"')
) "Selecting Send To must invalidate the actual Operations dropdown and status controls."

Check "Ribbon.TargetSelectionFullInvalidate" (
    [regex]::Match($status, '(?ms)^Private Sub InvalidateWarehouseTargetRibbonsStatus\(\).*?^End Sub').Value -match '(?m)^\s*ribbon\.Invalidate\s*$'
) "Warehouse selection must force Excel to requery the live ribbon immediately."

Check "Session.ServerSignOutDisconnects" (
    $role.Contains('Public Sub SignOutServerSession') -and
    $role.Contains('ApplyServerSignOutRole(True, True)') -and
    $role.Contains('modNasConnection.DisconnectCurrentNasSession(disconnectWindowsSession)') -and
    $nas.Contains('Public Function DisconnectCurrentNasSession') -and
    $nas.Contains('WNetCancelConnection2')
) "Server Sign Out must clear invSys state, the target, and the Windows SMB session."

Check "Session.InvSysSignOutRetainsServer" (
    $role.Contains('Public Sub ToggleCurrentInvSysUserForCapability') -and
    $role.Contains('SignOutCurrentUser') -and
    $role.Contains('Warehouse storage remains connected')
) "invSys Sign Out must retain the separately authenticated server session."

Check "Session.DisconnectedInvSysSignInFailsClosed" (
    [regex]::Match($role, '(?ms)^Public Sub PromptSetCurrentUserForCapability.*?^End Sub').Value.Contains('InvSysSignInPrerequisiteRole') -and
    [regex]::Match($role, '(?ms)^Private Function InvSysSignInPrerequisiteRole.*?^End Function').Value.Contains('If Not modNasConnection.HasConnectedUncRoot() Then') -and
    $role.Contains('Use Server Sign In before invSys Sign In')
) "invSys Sign In must not revive a remembered warehouse while the server session is disconnected."

Check "Session.OperationsAccessStatusFailsClosed" (
    [regex]::Match($status, '(?ms)^Public Function GetAccessStatusLabel.*?^End Function').Value.Contains('Access: Server Sign In required') -and
    [regex]::Match($status, '(?ms)^Public Function GetAccessStatusLabel.*?^End Function').Value.Contains('Access: invSys Sign In required')
) "The Operations access label must not say Ready while either session layer is signed out."

Write-Host "RESULT passed=$($passes.Count) failed=$($failures.Count)"
if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "  $_" }
    exit 1
}
