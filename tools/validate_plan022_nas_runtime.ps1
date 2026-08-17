[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$DeployRoot = "deploy/current",
    [Parameter(Mandatory = $true)]
    [string]$NasTestLeaf,
    [string]$ServerConfigPath = "C:\Users\Justin\OneDrive\Documents\invsys-scv.txt",
    [string]$OutputDirectory = "reports/runtime/plan022-nas",
    [switch]$PrepareUserAcceptancePin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Import-FunctionDefinitions {
    param([string]$ScriptPath)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $ScriptPath,
        [ref]$tokens,
        [ref]$errors
    )
    if ($errors.Count -gt 0) {
        throw "Unable to parse helper source."
    }
    $definitions = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)
    foreach ($definition in $definitions) {
        $scriptScopedDefinition = $definition.Extent.Text -replace (
            '^(?i)function\s+' + [regex]::Escape($definition.Name)
        ), ('function script:' + $definition.Name)
        . ([scriptblock]::Create($scriptScopedDefinition))
    }
}

function Add-Check {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Check,
        [bool]$Passed,
        [string]$Detail
    )

    $Rows.Add([pscustomobject]@{
        Check = $Check
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

function Get-WorkbookByName {
    param(
        [object]$Excel,
        [string]$Name
    )

    try { return $Excel.Workbooks.Item($Name) } catch { return $null }
}

function Get-PackageMap {
    param(
        [object]$Excel,
        [string]$DeployPath,
        [System.Collections.Generic.List[object]]$OwnedBooks
    )

    $map = @{}
    foreach ($name in @(
        "invSys.Core.xlam",
        "invSys.Inventory.Domain.xlam",
        "invSys.Designs.Domain.xlam",
        "invSys.Operations.xlam",
        "invSys.Admin.xlam"
    )) {
        $expectedPath = [IO.Path]::GetFullPath((Join-Path $DeployPath $name))
        $wb = Get-WorkbookByName -Excel $Excel -Name $name
        if ($null -eq $wb) {
            $wb = $Excel.Workbooks.Open($expectedPath)
            $OwnedBooks.Add($wb) | Out-Null
        }
        if ([string]::IsNullOrWhiteSpace([string]$wb.FullName) -or
            -not [string]::Equals(
                [IO.Path]::GetFullPath([string]$wb.FullName),
                $expectedPath,
                [StringComparison]::OrdinalIgnoreCase
            )) {
            throw "A loaded package did not come from the approved deployment root: $name"
        }
        $map[$name] = $wb
    }
    return $map
}

function Get-NasFileHashMap {
    param([string]$RootPath)

    $result = @{}
    foreach ($file in Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction Stop) {
        $relative = $file.FullName.Substring($RootPath.Length).TrimStart("\")
        $hash = Get-SharedNasFileSha256 -Path $file.FullName
        if ([string]::IsNullOrWhiteSpace($hash)) {
            throw "A dedicated NAS test file could not be hashed with shared-read access."
        }
        $result[$relative] = $hash
    }
    return $result
}

function Get-SharedNasFileSha256 {
    param([string]$Path)

    $stream = $null
    $sha = $null
    try {
        $stream = New-Object IO.FileStream(
            $Path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
        )
        $sha = [Security.Cryptography.SHA256]::Create()
        return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace("-", "")
    }
    finally {
        if ($null -ne $sha) { $sha.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-ChangedRelativeFiles {
    param(
        [hashtable]$Before,
        [hashtable]$After
    )

    $keys = @($Before.Keys) + @($After.Keys) | Sort-Object -Unique
    return @($keys | Where-Object {
        -not $Before.ContainsKey($_) -or
        -not $After.ContainsKey($_) -or
        $Before[$_] -ne $After[$_]
    })
}

function Get-ChangedNasFileLabels {
    param([string[]]$RelativePaths)

    return @($RelativePaths | ForEach-Object {
        $name = [IO.Path]::GetFileName([string]$_)
        if ([string]::IsNullOrWhiteSpace($name)) { "<unknown>" } else { $name }
    } | Sort-Object -Unique)
}

function Get-ExpectedRoleFormWindowState {
    param([int]$ExcelProcessId)

    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
        $ExcelProcessId
    )
    $classCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty,
        "ThunderDFrame"
    )
    $condition = New-Object System.Windows.Automation.AndCondition(
        $processCondition,
        $classCondition
    )
    $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $condition
    )
    $counts = @{
        Receiving = 0
        Production = 0
        Shipping = 0
        Other = 0
    }
    foreach ($window in $windows) {
        $caption = [string]$window.Current.Name
        switch ($caption) {
            "Receiving" { $counts.Receiving += 1 }
            "Production" { $counts.Production += 1 }
            "Shipping Shipments" { $counts.Shipping += 1 }
            default { $counts.Other += 1 }
        }
    }
    return [pscustomobject]$counts
}

function Get-PathFingerprint {
    param([string]$Path)

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Path)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").Substring(0, 12)
    }
    finally {
        $sha.Dispose()
    }
}

function Connect-ServerConfigSession {
    param([string]$ConfigPath)

    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        throw "The server configuration file is unavailable."
    }
    $lines = @(
        [IO.File]::ReadAllLines($ConfigPath) |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
    $urlMatch = [regex]::Match(($lines -join " "), 'https?://[^\s]+')
    if (-not $urlMatch.Success) {
        throw "The server configuration does not contain a server URL."
    }
    $serverUri = [Uri]$urlMatch.Value
    $userLine = $lines | Where-Object { $_ -match '^\s*Username\s*[:=]' } |
        Select-Object -First 1
    $passwordLine = $lines | Where-Object { $_ -match '^\s*Password\s*[:=]' } |
        Select-Object -First 1
    $userName = $userLine -replace '^\s*Username\s*[:=]\s*', ''
    $password = $passwordLine -replace '^\s*Password\s*[:=]\s*', ''
    if ([string]::IsNullOrWhiteSpace($userName) -or [string]::IsNullOrEmpty($password)) {
        throw "The server configuration is missing a username or password."
    }

    $ipcRoot = "\\" + $serverUri.Host + "\IPC$"
    $resource = New-Object Plan022NasRuntime+NETRESOURCE
    $resource.dwType = 1
    $resource.lpRemoteName = $ipcRoot
    $resultCode = [Plan022NasRuntime]::WNetAddConnection2(
        [ref]$resource,
        $password,
        $userName,
        0
    )
    if ($resultCode -notin @(0, 85, 1219)) {
        throw "The authenticated NAS session could not be established. Windows status=$resultCode"
    }

    $viewOutput = @(& net.exe view ("\\" + $serverUri.Host) 2>&1)
    $divider = $viewOutput | Where-Object { $_ -match '^-{3,}' } | Select-Object -First 1
    $dividerIndex = [array]::IndexOf($viewOutput, $divider)
    $shares = @()
    for ($i = $dividerIndex + 1; $i -lt $viewOutput.Count; $i++) {
        if ($viewOutput[$i] -match 'The command completed|System error') { break }
        if ($viewOutput[$i] -match '^([^\s]+)\s+Disk\s*') {
            $shares += $matches[1]
        }
    }
    $hubShares = @()
    foreach ($share in @($shares | Where-Object { $_ -match '(?i)inv.*sys' })) {
        $sharePath = "\\" + $serverUri.Host + "\" + $share
        try {
            $configProbe = @(
                Get-ChildItem -LiteralPath $sharePath -Filter "*.invSys.Config.xlsb" `
                    -File -Recurse -Depth 2 -ErrorAction Stop |
                    Select-Object -First 1
            )
            if ($configProbe.Count -eq 1) {
                $hubShares += $sharePath
            }
        }
        catch {}
    }
    if ($hubShares.Count -ne 1) {
        throw "The server session did not resolve exactly one invSys hub share."
    }

    return [pscustomobject]@{
        HubShare = $hubShares[0]
        UserName = $userName
        Password = $password
    }
}

function Set-TestUserPinHash {
    param(
        [object]$Excel,
        [object]$CoreWorkbook,
        [string]$AuthPath,
        [string]$UserId,
        [string]$SecretText
    )

    $pinHash = [string](Run-WorkbookMacro -Excel $Excel `
        -WorkbookName ([string]$CoreWorkbook.Name) `
        -MacroName "modAuth.HashUserCredential" -Arguments @($SecretText))
    $authWb = Get-WorkbookByName -Excel $Excel -Name ([IO.Path]::GetFileName($AuthPath))
    $openedHere = $false
    if ($null -eq $authWb) {
        $authWb = $Excel.Workbooks.Open($AuthPath)
        $openedHere = $true
    }
    try {
        $users = $authWb.Worksheets("Users").ListObjects("tblUsers")
        $userIndex = [int]$users.ListColumns("UserId").Index
        $pinIndex = [int]$users.ListColumns("PinHash").Index
        $matched = $false
        for ($row = 1; $row -le [int]$users.ListRows.Count; $row++) {
            if ([string]::Equals(
                    [string]$users.DataBodyRange.Cells.Item($row, $userIndex).Value2,
                    $UserId,
                    [StringComparison]::OrdinalIgnoreCase
                )) {
                $users.DataBodyRange.Cells.Item($row, $pinIndex).Value2 = $pinHash
                $matched = $true
                break
            }
        }
        if (-not $matched) {
            throw "The dedicated test user was not found in the auth workbook."
        }
        $authWb.Save()
    }
    finally {
        if ($openedHere) {
            $authWb.Close($false)
            Release-ComObject $authWb
        }
    }
}

function Read-ConfirmedUserAcceptancePin {
    $pin = Read-Host "Enter a temporary six-digit PIN for Plan 022 UAT" -AsSecureString
    $confirm = Read-Host "Confirm the temporary six-digit PIN" -AsSecureString
    $pinPtr = [IntPtr]::Zero
    $confirmPtr = [IntPtr]::Zero
    try {
        $pinPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($pin)
        $confirmPtr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($confirm)
        $pinText = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pinPtr)
        $confirmText = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($confirmPtr)
        if ($pinText -notmatch '^\d{6}$') {
            throw "The UAT PIN must contain exactly six digits."
        }
        if (-not [string]::Equals(
                $pinText,
                $confirmText,
                [StringComparison]::Ordinal
            )) {
            throw "The two UAT PIN entries did not match."
        }
        return $pinText
    }
    finally {
        if ($pinPtr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pinPtr)
        }
        if ($confirmPtr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($confirmPtr)
        }
        Remove-Variable pinText, confirmText -ErrorAction SilentlyContinue
    }
}

function Test-InventoryHasManagedRows {
    param(
        [object]$Excel,
        [string]$InventoryPath
    )

    $name = [IO.Path]::GetFileName($InventoryPath)
    $wb = Get-WorkbookByName -Excel $Excel -Name $name
    $openedHere = $false
    if ($null -eq $wb) {
        $wb = $Excel.Workbooks.Open(
            $InventoryPath,
            0,
            $true
        )
        $openedHere = $true
    }
    try {
        foreach ($worksheet in $wb.Worksheets) {
            foreach ($table in $worksheet.ListObjects) {
                if ([string]::Equals(
                        [string]$table.Name,
                        "tblInventoryEntities",
                        [StringComparison]::OrdinalIgnoreCase
                    )) {
                    return ([int]$table.ListRows.Count -gt 0)
                }
            }
        }
        return $false
    }
    finally {
        if ($openedHere -and $null -ne $wb) {
            $wb.Close($false)
            Release-ComObject $wb
        }
    }
}

function Remove-TemporaryOperatorRuntime {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or
        -not (Test-Path -LiteralPath $Path -PathType Container)) {
        return
    }
    $resolvedPath = [IO.Path]::GetFullPath($Path)
    $temporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if (-not $resolvedPath.StartsWith(
            $temporaryRoot,
            [StringComparison]::OrdinalIgnoreCase
        ) -or
        [IO.Path]::GetFileName($resolvedPath) -notmatch
            '^invsys-plan022-nas-operator-[A-Fa-f0-9]{32}$') {
        throw "Refusing to remove an operator runtime outside the dedicated temporary root."
    }
    Remove-Item -LiteralPath $resolvedPath -Recurse -Force
}

function Invoke-RuntimeExtraction {
    param(
        [string]$Repo,
        [string]$OutputPath,
        [hashtable]$ManifestHashes
    )

    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    $output = @(
        & powershell -NoProfile -ExecutionPolicy Bypass -File (
            Join-Path $Repo "tools/export-invsys-runtime-state.ps1"
        ) -OutputDirectory $OutputPath 2>&1
    )
    if ($LASTEXITCODE -ne 0) {
        throw "The read-only runtime extractor failed."
    }
    $runtime = Get-Content -Raw -LiteralPath (
        Join-Path $OutputPath "runtime-state.json"
    ) | ConvertFrom-Json
    $names = @($runtime.loadedAddins | ForEach-Object { [string]$_.name } | Sort-Object)
    $expected = @(
        "invSys.Admin.xlam",
        "invSys.Core.xlam",
        "invSys.Designs.Domain.xlam",
        "invSys.Inventory.Domain.xlam",
        "invSys.Operations.xlam"
    )
    $hashesMatch = $true
    foreach ($addin in @($runtime.loadedAddins)) {
        $name = [string]$addin.name
        if (-not $ManifestHashes.ContainsKey($name) -or
            -not [string]::Equals(
                [string]$addin.sha256,
                [string]$ManifestHashes[$name],
                [StringComparison]::OrdinalIgnoreCase
            )) {
            $hashesMatch = $false
        }
    }
    $safe = [int]$runtime.safety.mutatingActionsInvoked -eq 0 -and
        @($runtime.safety.inspectedFiles | Where-Object { -not $_.unchanged }).Count -eq 0
    return [pscustomobject]@{
        FivePackages = (($names -join "|") -eq ($expected -join "|"))
        HashesMatch = $hashesMatch
        ReadOnlySafe = $safe
        PackageNames = $names
    }
}

function Stop-OwnedExcel {
    param(
        [object]$Excel,
        [int]$ProcessId,
        [System.Collections.Generic.List[object]]$OwnedBooks
    )

    foreach ($wb in $OwnedBooks) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $Excel) {
        try { $Excel.Quit() } catch {}
        Release-ComObject $Excel
    }
    Start-Sleep -Milliseconds 500
    $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($null -ne $process -and $process.ProcessName -eq "EXCEL") {
        Stop-Process -Id $ProcessId -Force
    }
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Plan022LauncherWindow
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
public static class Plan022NasRuntime
{
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    public struct NETRESOURCE
    {
        public int dwScope;
        public int dwType;
        public int dwDisplayType;
        public int dwUsage;
        public string lpLocalName;
        public string lpRemoteName;
        public string lpComment;
        public string lpProvider;
    }
    [DllImport("mpr.dll", CharSet=CharSet.Unicode)]
    public static extern int WNetAddConnection2(
        ref NETRESOURCE resource,
        string password,
        string username,
        int flags);
}
"@

if ($NasTestLeaf -notmatch '^Plan022-R1-UAT-\d{8}-[A-F0-9]{8}$') {
    throw "The NAS test leaf does not match the dedicated Plan 022 naming contract."
}

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deployPath = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$outputPath = Join-Path $repo $OutputDirectory
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
Import-FunctionDefinitions -ScriptPath (
    Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
)
Import-FunctionDefinitions -ScriptPath (
    Join-Path $repo "tools/validate_plan022_packaged_launchers.ps1"
)

$manifest = Get-Content -Raw -LiteralPath (
    Join-Path $deployPath "addins-manifest.json"
) | ConvertFrom-Json
$manifestHashes = @{}
foreach ($package in @($manifest.packages)) {
    $manifestHashes[[string]$package.name] = [string]$package.sha256
}

$server = Connect-ServerConfigSession -ConfigPath $ServerConfigPath
$nasRoot = Join-Path $server.HubShare $NasTestLeaf
if (-not (Test-Path -LiteralPath $nasRoot -PathType Container)) {
    throw "The dedicated NAS test root is unavailable."
}
$rootFingerprint = Get-PathFingerprint -Path $nasRoot
$warehouseId = "WHT" + (($NasTestLeaf -split "-")[-1]).Substring(0, 6)
$stationId = if ([string]::IsNullOrWhiteSpace($env:COMPUTERNAME)) {
    "LOCAL-COMPUTER"
} else {
    $env:COMPUTERNAME.Trim()
}
$uatUser = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) {
    "plan022-user"
} else {
    $env:USERNAME
}
$testUser = "plan022-auto-" + (($NasTestLeaf -split "-")[-1]).ToLowerInvariant()
$testSecret = [guid]::NewGuid().ToString("N")
$operatorRuntimeBase = Join-Path ([IO.Path]::GetTempPath()) (
    "invsys-plan022-nas-operator-" + [guid]::NewGuid().ToString("N")
)
$configPath = Join-Path $nasRoot ($warehouseId + ".invSys.Config.xlsb")
$authPath = Join-Path $nasRoot ($warehouseId + ".invSys.Auth.xlsb")
$inventoryPath = Join-Path $nasRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
$results = New-Object System.Collections.Generic.List[object]

if ($PrepareUserAcceptancePin) {
    $excel = $null
    $excelProcessId = 0
    $ownedBooks = New-Object System.Collections.Generic.List[object]
    $uatPin = ""
    try {
        $uatPin = Read-ConfirmedUserAcceptancePin
        $excel = New-Object -ComObject Excel.Application
        $excelProcessId = Get-ExcelProcessId -Excel $excel
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 1
        $packages = Get-PackageMap -Excel $excel -DeployPath $deployPath `
            -OwnedBooks $ownedBooks
        foreach ($role in @("ADMIN", "RECEIVE", "PRODUCTION", "SHIPPING")) {
            $authResult = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName ([string]$packages["invSys.Core.xlam"].Name) `
                -MacroName "modAuth.EnsureStationRoleAuthForAutomation" `
                -Arguments @(
                    $warehouseId,
                    $stationId,
                    $uatUser,
                    $uatUser,
                    $role,
                    $authPath,
                    "svc_processor"
                ))
            if (-not $authResult.StartsWith("OK|")) {
                throw "The UAT user capability set could not be provisioned."
            }
        }
        Set-TestUserPinHash -Excel $excel `
            -CoreWorkbook $packages["invSys.Core.xlam"] `
            -AuthPath $authPath -UserId $uatUser -SecretText $uatPin
        Write-Output (
            "Plan 022 UAT PIN hash prepared for user " + $uatUser +
            " in warehouse " + $warehouseId + "."
        )
    }
    finally {
        $uatPin = ""
        Stop-OwnedExcel -Excel $excel -ProcessId $excelProcessId `
            -OwnedBooks $ownedBooks
    }
    exit 0
}

for ($sessionNumber = 1; $sessionNumber -le 2; $sessionNumber++) {
    $excel = $null
    $excelProcessId = 0
    $ownedBooks = New-Object System.Collections.Generic.List[object]
    $sessionOperatorRoot = Join-Path $operatorRuntimeBase ("session-" + $sessionNumber)
    $sessionStep = "startup"
    try {
        New-Item -ItemType Directory -Path $sessionOperatorRoot -Force | Out-Null
        $excel = New-Object -ComObject Excel.Application
        $excelProcessId = Get-ExcelProcessId -Excel $excel
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $true
        $excel.AutomationSecurity = 1
        Start-Sleep -Milliseconds 1200

        $sessionStep = "load approved package set"
        $packages = Get-PackageMap -Excel $excel -DeployPath $deployPath `
            -OwnedBooks $ownedBooks
        $core = $packages["invSys.Core.xlam"]
        $operations = $packages["invSys.Operations.xlam"]
        $admin = $packages["invSys.Admin.xlam"]

        if ($sessionNumber -eq 1 -and
            -not (Test-Path -LiteralPath $configPath -PathType Leaf)) {
            $sessionStep = "bootstrap dedicated NAS warehouse"
            [void](Run-WorkbookMacro -Excel $excel `
                -WorkbookName ([string]$core.Name) `
                -MacroName "modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride" `
                -Arguments @((Join-Path $deployPath "templates")))
            $bootstrap = [bool](Run-WorkbookMacro -Excel $excel `
                -WorkbookName ([string]$admin.Name) `
                -MacroName "modAdminConsole.BootstrapWarehouseLocalAdmin" `
                -Arguments @(
                    $warehouseId,
                    "Plan 022 Release 1 UAT",
                    $stationId,
                    $testUser,
                    $nasRoot,
                    (Join-Path $nasRoot "Published")
                ))
            [void](Run-WorkbookMacro -Excel $excel `
                -WorkbookName ([string]$core.Name) `
                -MacroName "modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride")
            if (-not $bootstrap) {
                throw "Packaged Admin did not bootstrap the dedicated NAS warehouse."
            }
        }

        if ($sessionNumber -eq 1) {
            foreach ($role in @("ADMIN", "RECEIVE", "PRODUCTION", "SHIPPING")) {
                $authResult = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName ([string]$core.Name) `
                    -MacroName "modAuth.EnsureStationRoleAuthForAutomation" `
                    -Arguments @(
                        $warehouseId,
                        $stationId,
                        $testUser,
                        $testUser,
                        $role,
                        $authPath,
                        "svc_processor"
                    ))
                if (-not $authResult.StartsWith("OK|")) {
                    throw "The dedicated test capability set could not be provisioned."
                }
            }
            Set-TestUserPinHash -Excel $excel -CoreWorkbook $core `
                -AuthPath $authPath -UserId $testUser -SecretText $testSecret
        }

        $sessionStep = "connect and select dedicated NAS target"
        $connectCode = [int](Run-WorkbookMacro -Excel $excel `
            -WorkbookName ([string]$core.Name) `
            -MacroName "modNasConnection.ConnectNasRootWithCredentials" `
            -Arguments @($nasRoot, $server.UserName, $server.Password))
        if ($connectCode -eq 3 -and
            (Test-Path -LiteralPath $nasRoot -PathType Container)) {
            $connectCode = [int](Run-WorkbookMacro -Excel $excel `
                -WorkbookName ([string]$core.Name) `
                -MacroName "modNasConnection.TryRevalidateRememberedRoot" `
                -Arguments @($nasRoot))
        }
        if ($connectCode -ne 0) {
            throw "Core did not establish or revalidate the NAS target session. Status=$connectCode"
        }
        $targetResult = [string](Run-WorkbookMacro -Excel $excel `
            -WorkbookName ([string]$core.Name) `
            -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" `
            -Arguments @($nasRoot, $nasRoot, $stationId, $false))
        if (-not $targetResult.StartsWith("OK|")) {
            throw "Core did not select the dedicated NAS warehouse."
        }
        $configLoaded = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName ([string]$core.Name) `
            -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
        $authLoaded = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName ([string]$core.Name) `
            -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
        $signIn = [string](Run-WorkbookMacro -Excel $excel `
            -WorkbookName ([string]$core.Name) `
            -MacroName "modAuth.SignInCurrentTargetForAutomation" `
            -Arguments @($testUser, $testSecret, "RECEIVE_POST"))
        if (-not $configLoaded -or -not $authLoaded -or
            -not $signIn.StartsWith("OK|")) {
            throw (
                "The dedicated invSys test user could not sign in. " +
                "ConfigLoaded=$configLoaded; AuthLoaded=$authLoaded; Result=$signIn"
            )
        }
        $operatorRootOverrideSet = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName ([string]$core.Name) `
            -MacroName "modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation" `
            -Arguments @($sessionOperatorRoot))
        if (-not $operatorRootOverrideSet) {
            throw "Core did not accept the isolated station-local operator root."
        }

        if ($sessionNumber -eq 1 -and
            -not (Test-InventoryHasManagedRows -Excel $excel `
                -InventoryPath $inventoryPath)) {
            $sessionStep = "seed dedicated test inventory"
            $seed = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName ([string]$admin.Name) `
                -MacroName "modAdminConsole.SeedDemoInventoryForAutomation" `
                -Arguments @($warehouseId, $stationId, $testUser))
            if (-not $seed.StartsWith("OK|")) {
                throw "Packaged Admin did not seed the dedicated test inventory."
            }
        }

        $nasHashesBefore = Get-NasFileHashMap -RootPath $nasRoot
        $sessionStep = "self-provision station-local operator workbooks"
        $receivingCapture = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName ([string]$operations.Name) `
            -MacroName "modTS_Received.ShowReceivingForm"
        $productionCapture = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName ([string]$operations.Name) `
            -MacroName "mProduction.BtnOpenProductionForm"
        $shippingCapture = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName ([string]$operations.Name) `
            -MacroName "modTS_Shipments.BtnOpenShipmentsForm"
        $receivingName = $warehouseId + ".Receiving.Operator.xlsm"
        $productionName = $warehouseId + ".Production.Operator.xlsm"
        $shippingName = $warehouseId + ".Shipping.Operator.xlsm"
        $receivingWb = Get-WorkbookByName -Excel $excel -Name $receivingName
        $productionWb = Get-WorkbookByName -Excel $excel -Name $productionName
        $shippingWb = Get-WorkbookByName -Excel $excel -Name $shippingName
        if ($null -eq $receivingWb -or
            $null -eq $productionWb -or
            $null -eq $shippingWb) {
            throw "A role launcher did not self-provision its station-local operator workbook."
        }
        $ownedBooks.Add($receivingWb) | Out-Null
        $ownedBooks.Add($productionWb) | Out-Null
        $ownedBooks.Add($shippingWb) | Out-Null
        $operatorPaths = @(
            [string]$receivingWb.FullName,
            [string]$productionWb.FullName,
            [string]$shippingWb.FullName
        )
        $operatorRootFullPath = [IO.Path]::GetFullPath($sessionOperatorRoot)
        $operatorPathsLocal = @($operatorPaths | Where-Object {
            [IO.Path]::GetFullPath($_).StartsWith(
                $operatorRootFullPath,
                [StringComparison]::OrdinalIgnoreCase
            )
        }).Count -eq 3
        $operatorFileCount = @(
            Get-ChildItem -LiteralPath $sessionOperatorRoot `
                -Filter "*.Operator.xlsm" -File -Recurse -ErrorAction Stop
        ).Count
        if (-not $operatorPathsLocal -or $operatorFileCount -ne 3) {
            throw "Role launchers did not create exactly three isolated station-local operator workbooks."
        }
        Add-Check -Rows $results -Check "Session$sessionNumber.OperatorSelfProvision" `
            -Passed $true `
            -Detail "All three role callbacks created exactly one isolated station-local operator workbook."

        $sessionStep = "repeat all packaged launcher callbacks"
        $productionWb.Activate()
        $productionSecond = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName ([string]$operations.Name) `
            -MacroName "mProduction.BtnOpenProductionForm"
        $shippingWb.Activate()
        $shippingSecond = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName ([string]$operations.Name) `
            -MacroName "modTS_Shipments.BtnOpenShipmentsForm"
        $receivingWb.Activate()
        $receivingSecond = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName ([string]$operations.Name) `
            -MacroName "modTS_Received.ShowReceivingForm"
        $callbackText = @(
            $receivingCapture.WindowText,
            $productionCapture.WindowText,
            $shippingCapture.WindowText,
            $receivingSecond.WindowText,
            $productionSecond.WindowText,
            $shippingSecond.WindowText
        ) -join " "
        $callbackErrors = @(
            $receivingCapture.Error,
            $productionCapture.Error,
            $shippingCapture.Error,
            $receivingSecond.Error,
            $productionSecond.Error,
            $shippingSecond.Error
        ) -join " "
        $formState = Get-ExpectedRoleFormWindowState -ExcelProcessId $excelProcessId
        $nasHashesAfter = Get-NasFileHashMap -RootPath $nasRoot
        $changedNasFiles = @(
            Get-ChangedRelativeFiles -Before $nasHashesBefore -After $nasHashesAfter
        )
        $changedNasLabels = @(Get-ChangedNasFileLabels -RelativePaths $changedNasFiles)
        $launchersOk = [string]::IsNullOrWhiteSpace($callbackErrors) -and
            $callbackText -notmatch '(?i)(Type mismatch|failed|operator workbook)' -and
            $formState.Receiving -eq 1 -and
            $formState.Production -eq 1 -and
            $formState.Shipping -eq 1
        Add-Check -Rows $results -Check "Session$sessionNumber.Launchers" `
            -Passed $launchersOk `
            -Detail (
                "Callbacks completed twice; modeless forms: Receiving=" +
                $formState.Receiving + ", Production=" + $formState.Production +
                ", Shipping=" + $formState.Shipping + ", other=" + $formState.Other + "."
            )
        Add-Check -Rows $results -Check "Session$sessionNumber.NasCanonicalHashes" `
            -Passed ($changedNasFiles.Count -eq 0) `
            -Detail (
                "Canonical files changed merely from launcher use=" +
                $changedNasFiles.Count + "; files=" +
                $(if ($changedNasLabels.Count -eq 0) {
                    "<none>"
                } else {
                    $changedNasLabels -join ","
                })
            )

        $sessionStep = "extract deployed runtime state read-only"
        $runtimeEvidence = Invoke-RuntimeExtraction -Repo $repo `
            -OutputPath (Join-Path $outputPath ("runtime-session-" + $sessionNumber)) `
            -ManifestHashes $manifestHashes
        Add-Check -Rows $results -Check "Session$sessionNumber.RuntimeFivePackages" `
            -Passed (
                $runtimeEvidence.FivePackages -and
                $runtimeEvidence.HashesMatch
            ) `
            -Detail "Read-only extraction observed the exact five approved manifest hashes."
        Add-Check -Rows $results -Check "Session$sessionNumber.RuntimeReadOnlySafety" `
            -Passed $runtimeEvidence.ReadOnlySafe `
            -Detail "Runtime extraction invoked zero mutations and changed no inspected file hashes."
        Add-Check -Rows $results -Check "Session$sessionNumber.NasTarget" `
            -Passed (
                $targetResult.StartsWith("OK|") -and
                $configLoaded -and
                $authLoaded -and
                $signIn.StartsWith("OK|")
            ) `
            -Detail (
                "Selected warehouse=$warehouseId; station=$stationId; source=NAS; " +
                "root fingerprint=$rootFingerprint."
            )
    }
    catch {
        $message = [string]$_.Exception.Message
        $message = $message -replace '\\\\[^\\\s]+\\[^;\s]+', '<redacted-nas-path>'
        $message = $message -replace '(?i)[A-Z]:\\Users\\[^\\\s]+', '<redacted-user-path>'
        Add-Check -Rows $results -Check "Session$sessionNumber.Harness" `
            -Passed $false -Detail ("Step=$sessionStep; " + $message)
    }
    finally {
        Stop-OwnedExcel -Excel $excel -ProcessId $excelProcessId -OwnedBooks $ownedBooks
    }
}

try {
    Remove-TemporaryOperatorRuntime -Path $operatorRuntimeBase
}
catch {}

$failed = @($results | Where-Object { -not $_.Passed })
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("# Plan 022 Dedicated NAS Runtime Evidence") | Out-Null
$lines.Add("") | Out-Null
$lines.Add("- Captured: $([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'))") | Out-Null
$lines.Add("- Package set: $($manifest.packageSetVersion)") | Out-Null
$lines.Add("- NAS root: <redacted> (leaf=$NasTestLeaf; fingerprint=$rootFingerprint)") | Out-Null
$lines.Add("- Warehouse: $warehouseId; station: $stationId") | Out-Null
$lines.Add("- Excel sessions: 2 (clean restart between sessions)") | Out-Null
$lines.Add("- Row-level operational values included: False") | Out-Null
$lines.Add("") | Out-Null
$lines.Add("| Check | Result | Detail |") | Out-Null
$lines.Add("|---|---|---|") | Out-Null
foreach ($row in $results) {
    $status = if ($row.Passed) { "PASS" } else { "FAIL" }
    $detail = ([string]$row.Detail).Replace("|", "/")
    $lines.Add("| $($row.Check) | $status | $detail |") | Out-Null
}
$reportPath = Join-Path $outputPath "dedicated-nas-runtime.md"
[IO.File]::WriteAllLines($reportPath, $lines)

if ($failed.Count -gt 0) {
    Write-Output "PLAN022_NAS_RUNTIME_FAILED"
    Write-Output "REPORT=$reportPath"
    Write-Output "PASSED=$($results.Count - $failed.Count) FAILED=$($failed.Count) TOTAL=$($results.Count)"
    exit 1
}

Write-Output "PLAN022_NAS_RUNTIME_OK"
Write-Output "REPORT=$reportPath"
Write-Output "PASSED=$($results.Count) FAILED=0 TOTAL=$($results.Count)"
