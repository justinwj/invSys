[CmdletBinding()]
param(
    [string]$FixturePath = "",

    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [string]$ReportTimestampUtc = "",

    [string]$SchemaPath = "",

    [Int64]$ExcelHwnd = 0,

    [switch]$IncludeRowValues
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

if ($null -eq ("InvSysRuntimeNative" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class InvSysRuntimeNative
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(
        IntPtr hwnd,
        int objectId,
        ref Guid iid,
        [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object nativeObject);

    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(
        IntPtr hWnd,
        EnumWindowsProc callback,
        IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetClassName(
        IntPtr hWnd,
        System.Text.StringBuilder text,
        int maxCount);

    public static IntPtr FindExcelGridWindow(IntPtr mainWindow)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(mainWindow, delegate(IntPtr candidate, IntPtr ignored) {
            var className = new System.Text.StringBuilder(256);
            GetClassName(candidate, className, className.Capacity);
            if (string.Equals(className.ToString(), "EXCEL7", StringComparison.Ordinal)) {
                found = candidate;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }
}
"@
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path (Join-Path $scriptRoot "..")).Path

if ($IncludeRowValues) {
    throw (
        "Row-value diagnostics are disabled. Define and approve a separate " +
        "field-level redaction policy before enabling this mode."
    )
}
if ([string]::IsNullOrWhiteSpace($SchemaPath)) {
    $SchemaPath = Join-Path $scriptRoot "contracts\runtime-state.schema.json"
}
if ([string]::IsNullOrWhiteSpace($ReportTimestampUtc)) {
    $ReportTimestampUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
}

function Resolve-RequiredPath {
    param(
        [string]$Path,
        [string]$Description
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Description not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function ConvertTo-NormalizedPath {
    param([string]$Path)
    return ($Path -replace "\\", "/")
}

function Get-RelativePath {
    param(
        [string]$BasePath,
        [string]$TargetPath
    )
    $baseFull = [IO.Path]::GetFullPath($BasePath)
    if (-not $baseFull.EndsWith([IO.Path]::DirectorySeparatorChar)) {
        $baseFull += [IO.Path]::DirectorySeparatorChar
    }
    $targetFull = [IO.Path]::GetFullPath($TargetPath)
    $baseUri = New-Object Uri($baseFull)
    $targetUri = New-Object Uri($targetFull)
    return (ConvertTo-NormalizedPath (
        [Uri]::UnescapeDataString($baseUri.MakeRelativeUri($targetUri).ToString())
    ))
}

function Get-DisplayPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ""
    }
    $full = [IO.Path]::GetFullPath($Path)
    $repoPrefix = $repoRoot.TrimEnd("\", "/") + [IO.Path]::DirectorySeparatorChar
    if ($full.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        $full.Equals($repoRoot, [StringComparison]::OrdinalIgnoreCase)) {
        return (Get-RelativePath -BasePath $repoRoot -TargetPath $full)
    }
    return (ConvertTo-NormalizedPath $full)
}

function Write-Utf8NoBom {
    param(
        [string]$Path,
        [string]$Content
    )
    $normalized = $Content -replace "`r`n", "`n"
    if (-not $normalized.EndsWith("`n")) {
        $normalized += "`n"
    }
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $normalized, $encoding)
}

function Read-Json {
    param([string]$Path)
    return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json)
}

function Test-HasProperty {
    param(
        $Object,
        [string]$Name
    )
    if ($null -eq $Object) {
        return $false
    }
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function Get-StringSha256 {
    param([string]$Value)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Value)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Convert-ExcelDateValueToUtcString {
    param(
        [object]$Value,
        [string]$FallbackUtc
    )

    if ($null -eq $Value) {
        return $FallbackUtc
    }

    try {
        if ($Value -is [DateTime]) {
            return ([DateTime]$Value).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }

        $text = [string]$Value
        $oaDate = 0.0
        if ([double]::TryParse(
                $text,
                [Globalization.NumberStyles]::Float,
                [Globalization.CultureInfo]::InvariantCulture,
                [ref]$oaDate
            )) {
            return [DateTime]::FromOADate($oaDate).ToUniversalTime().ToString(
                "yyyy-MM-ddTHH:mm:ssZ"
            )
        }

        $parsed = [DateTime]::MinValue
        if ([DateTime]::TryParse(
                $text,
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AssumeLocal,
                [ref]$parsed
            )) {
            return $parsed.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
    }
    catch {}

    return $FallbackUtc
}

function Get-SharedFileSha256 {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or
        -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return ""
    }
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
        return ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace("-", "").ToLowerInvariant()
    }
    catch {
        return ""
    }
    finally {
        if ($null -ne $sha) { $sha.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Test-SensitiveKey {
    param([string]$Key)
    $normalized = ($Key -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
    return (
        $normalized -match 'password' -or
        $normalized -match 'pinhash' -or
        $normalized -eq "pin" -or
        $normalized -match 'token' -or
        $normalized -match 'secret' -or
        $normalized -match 'credential' -or
        $normalized -match 'apikey'
    )
}

function Get-PackageRole {
    param([string]$Name)
    switch -Regex ($Name) {
        '(?i)^invSys\.Core\.xlam$' { return "CORE" }
        '(?i)^invSys\.Inventory\.Domain\.xlam$' { return "INVENTORY_DOMAIN" }
        '(?i)^invSys\.Designs\.Domain\.xlam$' { return "DESIGNS_DOMAIN" }
        '(?i)^invSys\.Operations\.xlam$' { return "OPERATIONS" }
        '(?i)^invSys\.Admin\.xlam$' { return "ADMIN" }
        '(?i)^invSys\.Receiving\.xlam$' { return "RECEIVING" }
        '(?i)^invSys\.Production\.xlam$' { return "PRODUCTION" }
        '(?i)^invSys\.Shipping\.xlam$' { return "SHIPPING" }
        default { return "OTHER" }
    }
}

function Test-LegacyRolePackage {
    param([string]$Name)
    return $Name -match '(?i)^invSys\.(Receiving|Production|Shipping)\.xlam$'
}

function Get-AuthorityClass {
    param([string]$Name)
    switch -Regex ($Name) {
        '(?i)\.xlam$' { return "ADDIN" }
        '(?i)\.invSys\.Data\.(Inventory|Designs)\.' { return "CANONICAL_DOMAIN" }
        '(?i)\.invSys\.Config\.' { return "CONFIG" }
        '(?i)\.invSys\.Auth\.' { return "AUTH" }
        '(?i)invSys\.Inbox\.' { return "INBOX" }
        '(?i)\.Outbox\.' { return "OUTBOX" }
        '(?i)Snapshot' { return "SNAPSHOT" }
        '(?i)(Operator|inventory_management)' { return "OPERATOR" }
        default { return "UNKNOWN" }
    }
}

function Get-TableEventType {
    param([string]$TableName)
    switch -Regex ($TableName) {
        '(?i)Receive' { return "RECEIVE" }
        '(?i)Ship' { return "SHIP" }
        '(?i)Prod' { return "PROD" }
        '(?i)Design' { return "DESIGN" }
        default { return "UNKNOWN" }
    }
}

function Get-MatrixValue {
    param(
        $Values,
        [int]$Row,
        [int]$Column
    )
    if ($null -eq $Values) {
        return $null
    }
    if ($Values -is [Array] -and $Values.Rank -eq 2) {
        return $Values.GetValue($Row, $Column)
    }
    if ($Row -eq 1 -and $Column -eq 1) {
        return $Values
    }
    return $null
}

function Get-ComHeaders {
    param($Table)
    $headers = New-Object System.Collections.Generic.List[string]
    $range = $null
    try {
        $range = $Table.HeaderRowRange
        $values = $range.Value2
        $columnCount = [int]$Table.ListColumns.Count
        for ($column = 1; $column -le $columnCount; $column++) {
            $headers.Add([string](Get-MatrixValue -Values $values -Row 1 -Column $column))
        }
    }
    finally {
        if ($null -ne $range -and [Runtime.InteropServices.Marshal]::IsComObject($range)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
        }
    }
    return $headers.ToArray()
}

function Get-ComTableRows {
    param(
        $Table,
        [string[]]$Headers
    )
    $rows = New-Object System.Collections.Generic.List[object]
    $body = $null
    try {
        $rowCount = [int]$Table.ListRows.Count
        if ($rowCount -eq 0) {
            return @()
        }
        $body = $Table.DataBodyRange
        $values = $body.Value2
        for ($row = 1; $row -le $rowCount; $row++) {
            $record = [ordered]@{}
            for ($column = 1; $column -le $Headers.Count; $column++) {
                $record[$Headers[$column - 1]] =
                    Get-MatrixValue -Values $values -Row $row -Column $column
            }
            $rows.Add($record)
        }
    }
    finally {
        if ($null -ne $body -and [Runtime.InteropServices.Marshal]::IsComObject($body)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($body)
        }
    }
    return $rows.ToArray()
}

function Get-ExcelApplicationForHwnd {
    param([Int64]$Hwnd)

    if ($Hwnd -le 0) {
        throw "Excel window handle must be positive."
    }

    $nativeObject = $null
    $excel = $null
    $gridWindow = [InvSysRuntimeNative]::FindExcelGridWindow([IntPtr]$Hwnd)
    if ($gridWindow -eq [IntPtr]::Zero) {
        throw "Could not locate an Excel automation child window for handle $Hwnd."
    }
    $dispatchId = [Guid]::Parse("00020400-0000-0000-C000-000000000046")
    $result = [InvSysRuntimeNative]::AccessibleObjectFromWindow(
        $gridWindow,
        -16,
        [ref]$dispatchId,
        [ref]$nativeObject)
    if ($result -ne 0 -or $null -eq $nativeObject) {
        throw "Could not attach to Excel window handle $Hwnd."
    }

    try {
        try {
            $excel = $nativeObject.Application
        }
        catch {
            $excel = $nativeObject
            $nativeObject = $null
        }
        return $excel
    }
    finally {
        if ($null -ne $nativeObject -and
            [Runtime.InteropServices.Marshal]::IsComObject($nativeObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($nativeObject)
        }
    }
}

function Get-LiveRawInput {
    $excel = $null
    try {
        if ($ExcelHwnd -gt 0) {
            $excel = Get-ExcelApplicationForHwnd -Hwnd $ExcelHwnd
        }
        else {
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        }
    }
    catch {
        return [ordered]@{
            machineName = [string]$env:COMPUTERNAME
            sessionIdentity = "NO_SESSION"
            excelProcessCount = 0
            loadedAddins = @()
            workbooks = @()
            tables = @()
            runtimeResolution = [ordered]@{
                warehouseId = "UNKNOWN"
                stationId = "UNKNOWN"
                runtimeRoot = ""
                resolutionSource = "NONE"
                connectionStatus = "NO_SESSION"
            }
            config = @()
            currentUser = [ordered]@{
                userId = ""
                capabilities = @()
                signedIn = $false
            }
            domainBridges = @()
            inboxSummary = @()
            processor = [ordered]@{
                lockStatus = "UNKNOWN"
                backlogCount = 0
                lastRunAtUtc = $null
                lastRunStatus = "NO_SESSION"
            }
            snapshotReadModels = @()
            forms = @()
            inspectedFiles = @()
            inspectionMode = "NO_SESSION"
            warnings = @(
                [ordered]@{
                    code = "NO_EXCEL_SESSION"
                    severity = "INFO"
                    message = "No running Excel session was available for attach-only inspection."
                    context = "Excel"
                }
            )
        }
    }

    $loadedAddins = New-Object System.Collections.Generic.List[object]
    $loadedAddinPaths = @{}
    $workbooks = New-Object System.Collections.Generic.List[object]
    $tables = New-Object System.Collections.Generic.List[object]
    $config = New-Object System.Collections.Generic.List[object]
    $inboxCounts = @{}
    $snapshots = New-Object System.Collections.Generic.List[object]
    $inspectedFiles = New-Object System.Collections.Generic.List[object]
    $warnings = New-Object System.Collections.Generic.List[object]
    $warehouseId = "UNKNOWN"
    $stationId = "UNKNOWN"
    $runtimeRoot = ""

    try {
        $sessionIdentity = "EXCEL-HWND-" + [string]$excel.Hwnd
        $workbookCount = [int]$excel.Workbooks.Count
        for ($workbookIndex = 1; $workbookIndex -le $workbookCount; $workbookIndex++) {
            $workbook = $null
            try {
                $workbook = $excel.Workbooks.Item($workbookIndex)
                $name = [string]$workbook.Name
                $fullName = [string]$workbook.FullName
                $path = [string]$workbook.Path
                $authority = Get-AuthorityClass $name
                $visible = $false
                if ([int]$workbook.Windows.Count -gt 0) {
                    $window = $null
                    try {
                        $window = $workbook.Windows.Item(1)
                        $visible = [bool]$window.Visible
                    }
                    finally {
                        if ($null -ne $window -and
                            [Runtime.InteropServices.Marshal]::IsComObject($window)) {
                            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($window)
                        }
                    }
                }

                $workbooks.Add([ordered]@{
                    name = $name
                    path = $fullName
                    readOnly = [bool]$workbook.ReadOnly
                    visible = $visible
                    authorityClass = $authority
                })

                if (-not [string]::IsNullOrWhiteSpace($fullName) -and
                    (Test-Path -LiteralPath $fullName -PathType Leaf)) {
                    $beforeHash = Get-SharedFileSha256 $fullName
                    if (-not [string]::IsNullOrWhiteSpace($beforeHash)) {
                        $inspectedFiles.Add([ordered]@{
                            path = $fullName
                            authorityClass = $authority
                            precomputedBeforeSha256 = $beforeHash
                        })
                    }
                    else {
                        $warnings.Add([ordered]@{
                            code = "HASH_UNAVAILABLE"
                            severity = "WARNING"
                            message = "Before hash could not be read for an open workbook."
                            context = $name
                        })
                    }
                }

                if ($name -match '(?i)^invSys\..*\.xlam$') {
                    $loadedAddinPaths[$fullName.ToLowerInvariant()] = $true
                    $loadedAddins.Add([ordered]@{
                        name = $name
                        path = $fullName
                        sha256 = Get-SharedFileSha256 $fullName
                        projectVersion = "UNKNOWN"
                        contractVersion = "UNKNOWN"
                        isAddin = [bool]$workbook.IsAddin
                    })
                }

                $worksheetCount = [int]$workbook.Worksheets.Count
                for ($sheetIndex = 1; $sheetIndex -le $worksheetCount; $sheetIndex++) {
                    $worksheet = $null
                    try {
                        $worksheet = $workbook.Worksheets.Item($sheetIndex)
                        $worksheetName = [string]$worksheet.Name
                        $tableCount = [int]$worksheet.ListObjects.Count
                        for ($tableIndex = 1; $tableIndex -le $tableCount; $tableIndex++) {
                            $table = $null
                            try {
                                $table = $worksheet.ListObjects.Item($tableIndex)
                                $tableName = [string]$table.Name
                                $headers = @(Get-ComHeaders $table)
                                $rowCount = [int]$table.ListRows.Count
                                $range = $null
                                $rangeAddress = ""
                                try {
                                    $range = $table.Range
                                    $rangeAddress = [string]$range.Address($true, $true)
                                }
                                finally {
                                    if ($null -ne $range -and
                                        [Runtime.InteropServices.Marshal]::IsComObject($range)) {
                                        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
                                    }
                                }
                                $tables.Add([ordered]@{
                                    workbookName = $name
                                    worksheetName = $worksheetName
                                    name = $tableName
                                    headers = $headers
                                    rowCount = $rowCount
                                    rangeAddress = $rangeAddress
                                })

                                if ($authority -eq "CONFIG" -and
                                    $tableName -in @("tblWarehouseConfig", "tblStationConfig") -and
                                    $rowCount -gt 0) {
                                    $rows = @(Get-ComTableRows -Table $table -Headers $headers)
                                    foreach ($row in $rows) {
                                        foreach ($header in $headers) {
                                            $config.Add([ordered]@{
                                                key = $header
                                                value = $row[$header]
                                                sourceScope = $(if ($tableName -eq "tblWarehouseConfig") {
                                                    "WAREHOUSE"
                                                } else { "STATION" })
                                            })
                                            if ($header -eq "WarehouseId" -and
                                                -not [string]::IsNullOrWhiteSpace([string]$row[$header])) {
                                                $warehouseId = [string]$row[$header]
                                            }
                                            if ($header -eq "StationId" -and
                                                -not [string]::IsNullOrWhiteSpace([string]$row[$header])) {
                                                $stationId = [string]$row[$header]
                                            }
                                            if ($header -eq "PathDataRoot" -and
                                                -not [string]::IsNullOrWhiteSpace([string]$row[$header])) {
                                                $runtimeRoot = [string]$row[$header]
                                            }
                                        }
                                    }
                                }

                                if ($authority -eq "INBOX" -and $rowCount -gt 0) {
                                    $rows = @(Get-ComTableRows -Table $table -Headers $headers)
                                    foreach ($row in $rows) {
                                        $eventType = Get-TableEventType $tableName
                                        if ("EventType" -in $headers -and
                                            -not [string]::IsNullOrWhiteSpace([string]$row["EventType"])) {
                                            $eventType = [string]$row["EventType"]
                                        }
                                        $status = "UNKNOWN"
                                        if ("Status" -in $headers) {
                                            $status = [string]$row["Status"]
                                        }
                                        $countKey = $eventType + "|" + $status
                                        if (-not $inboxCounts.ContainsKey($countKey)) {
                                            $inboxCounts[$countKey] = 0
                                        }
                                        $inboxCounts[$countKey] += 1
                                    }
                                }

                                if ($authority -in @("SNAPSHOT", "OPERATOR") -and
                                    $tableName -in @("invSys", "tblInventoryEntities")) {
                                    $snapshotId = ""
                                    $lastRefresh = $ReportTimestampUtc
                                    $sourceType = "CACHED"
                                    $isStale = $true
                                    if ($rowCount -gt 0) {
                                        $rows = @(Get-ComTableRows -Table $table -Headers $headers)
                                        $firstRow = $rows[0]
                                        if ("SnapshotId" -in $headers) {
                                            $snapshotId = [string]$firstRow["SnapshotId"]
                                        }
                                        if ("LastRefreshUTC" -in $headers -and
                                            -not [string]::IsNullOrWhiteSpace(
                                                [string]$firstRow["LastRefreshUTC"]
                                            )) {
                                            $lastRefresh = Convert-ExcelDateValueToUtcString `
                                                -Value $firstRow["LastRefreshUTC"] `
                                                -FallbackUtc $ReportTimestampUtc
                                        }
                                        if ("SourceType" -in $headers -and
                                            [string]$firstRow["SourceType"] -in @(
                                                "LOCAL", "SHAREPOINT", "CACHED"
                                            )) {
                                            $sourceType = [string]$firstRow["SourceType"]
                                        }
                                        if ("IsStale" -in $headers) {
                                            $isStale = [bool]$firstRow["IsStale"]
                                        }
                                    }
                                    $snapshots.Add([ordered]@{
                                        workbookName = $name
                                        tableName = $tableName
                                        snapshotId = $snapshotId
                                        lastRefreshUtc = $lastRefresh
                                        sourceType = $sourceType
                                        isStale = $isStale
                                        rowCount = $rowCount
                                    })
                                }
                            }
                            finally {
                                if ($null -ne $table -and
                                    [Runtime.InteropServices.Marshal]::IsComObject($table)) {
                                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($table)
                                }
                            }
                        }
                    }
                    finally {
                        if ($null -ne $worksheet -and
                            [Runtime.InteropServices.Marshal]::IsComObject($worksheet)) {
                            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet)
                        }
                    }
                }
            }
            finally {
                if ($null -ne $workbook -and
                    [Runtime.InteropServices.Marshal]::IsComObject($workbook)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
                }
            }
        }

        # Excel excludes workbooks whose IsAddin property is true from the
        # Workbooks collection. AddIns2 includes those open, unregistered XLAMs
        # as well as installed add-ins, so inspect it without changing Installed
        # state. Missing stale registrations are not loaded runtime packages.
        $addins2 = $null
        try {
            $addins2 = $excel.AddIns2
            $addinCount = [int]$addins2.Count
            for ($addinIndex = 1; $addinIndex -le $addinCount; $addinIndex++) {
                $addin = $null
                try {
                    $addin = $addins2.Item($addinIndex)
                    $addinName = [string]$addin.Name
                    $addinFullName = [string]$addin.FullName
                    if ($addinName -notmatch '(?i)^invSys\..*\.xlam$' -or
                        [string]::IsNullOrWhiteSpace($addinFullName) -or
                        -not (Test-Path -LiteralPath $addinFullName -PathType Leaf)) {
                        continue
                    }

                    $addinKey = $addinFullName.ToLowerInvariant()
                    if ($loadedAddinPaths.ContainsKey($addinKey)) {
                        continue
                    }
                    $loadedAddinPaths[$addinKey] = $true
                    $addinHash = Get-SharedFileSha256 $addinFullName
                    $loadedAddins.Add([ordered]@{
                        name = $addinName
                        path = $addinFullName
                        sha256 = $addinHash
                        projectVersion = "UNKNOWN"
                        contractVersion = "UNKNOWN"
                        isAddin = $true
                    })
                    if (-not [string]::IsNullOrWhiteSpace($addinHash)) {
                        $inspectedFiles.Add([ordered]@{
                            path = $addinFullName
                            authorityClass = "ADDIN"
                            precomputedBeforeSha256 = $addinHash
                        })
                    }
                }
                finally {
                    if ($null -ne $addin -and
                        [Runtime.InteropServices.Marshal]::IsComObject($addin)) {
                        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin)
                    }
                }
            }
        }
        finally {
            if ($null -ne $addins2 -and
                [Runtime.InteropServices.Marshal]::IsComObject($addins2)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addins2)
            }
        }

        $inboxSummary = New-Object System.Collections.Generic.List[object]
        foreach ($countKey in @($inboxCounts.Keys | Sort-Object)) {
            $parts = $countKey.Split("|")
            $inboxSummary.Add([ordered]@{
                eventType = $parts[0]
                status = $parts[1]
                rowCount = [int]$inboxCounts[$countKey]
            })
        }
        $backlog = [int](($inboxCounts.Values | Measure-Object -Sum).Sum)
        $domainBridges = New-Object System.Collections.Generic.List[object]
        foreach ($addin in $loadedAddins.ToArray()) {
            if ($addin.name -eq "invSys.Inventory.Domain.xlam") {
                $domainBridges.Add([ordered]@{
                    name = "InventoryDomain"
                    contractVersion = [string]$addin.contractVersion
                    status = "LOADED"
                })
            }
            if ($addin.name -eq "invSys.Designs.Domain.xlam") {
                $domainBridges.Add([ordered]@{
                    name = "DesignsDomain"
                    contractVersion = [string]$addin.contractVersion
                    status = "LOADED"
                })
            }
        }

        return [ordered]@{
            machineName = [string]$env:COMPUTERNAME
            sessionIdentity = $sessionIdentity
            excelProcessCount = 1
            loadedAddins = $loadedAddins.ToArray()
            workbooks = $workbooks.ToArray()
            tables = $tables.ToArray()
            runtimeResolution = [ordered]@{
                warehouseId = $warehouseId
                stationId = $stationId
                runtimeRoot = $runtimeRoot
                resolutionSource = $(if ($runtimeRoot) { "OPEN_CONFIG" } else { "UNKNOWN" })
                connectionStatus = $(if ($runtimeRoot) { "RESOLVED" } else { "UNKNOWN" })
            }
            config = $config.ToArray()
            currentUser = [ordered]@{
                userId = ""
                capabilities = @()
                signedIn = $false
            }
            domainBridges = $domainBridges.ToArray()
            inboxSummary = $inboxSummary.ToArray()
            processor = [ordered]@{
                lockStatus = "UNKNOWN"
                backlogCount = $backlog
                lastRunAtUtc = $null
                lastRunStatus = "UNKNOWN"
            }
            snapshotReadModels = $snapshots.ToArray()
            forms = @()
            inspectedFiles = $inspectedFiles.ToArray()
            inspectionMode = "LIVE_ATTACHED"
            warnings = $warnings.ToArray()
        }
    }
    finally {
        if ($null -ne $excel -and [Runtime.InteropServices.Marshal]::IsComObject($excel)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        }
    }
}

function Resolve-FixtureInspectedFiles {
    param(
        $RawInput,
        [string]$FixtureDirectory
    )
    $files = New-Object System.Collections.Generic.List[object]
    if (-not (Test-HasProperty -Object $RawInput -Name "inspectedFiles")) {
        return @()
    }
    foreach ($entry in @($RawInput.inspectedFiles)) {
        $path = [string]$entry.path
        if (-not [IO.Path]::IsPathRooted($path)) {
            $path = Join-Path $FixtureDirectory $path
        }
        $resolved = Resolve-RequiredPath -Path $path -Description "Inspected fixture"
        $files.Add([ordered]@{
            path = $resolved
            authorityClass = [string]$entry.authorityClass
            precomputedBeforeSha256 = Get-SharedFileSha256 $resolved
        })
    }
    return $files.ToArray()
}

$resolvedSchema = Resolve-RequiredPath -Path $SchemaPath -Description "Runtime schema"
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path

if ([string]::IsNullOrWhiteSpace($FixturePath)) {
    $raw = Get-LiveRawInput
}
else {
    $resolvedFixture = Resolve-RequiredPath -Path $FixturePath -Description "Runtime fixture"
    $raw = Read-Json $resolvedFixture
    $raw | Add-Member -NotePropertyName inspectionMode -NotePropertyValue "FIXTURE" -Force
    $resolvedInspected = Resolve-FixtureInspectedFiles `
        -RawInput $raw `
        -FixtureDirectory (Split-Path -Parent $resolvedFixture)
    $raw | Add-Member -NotePropertyName inspectedFiles `
        -NotePropertyValue $resolvedInspected -Force
    $raw | Add-Member -NotePropertyName warnings -NotePropertyValue @() -Force
}

$redactedFields = New-Object System.Collections.Generic.List[string]
$safeConfig = New-Object System.Collections.Generic.List[object]
foreach ($entry in @($raw.config)) {
    $key = [string]$entry.key
    $isSensitive = Test-SensitiveKey $key
    if ($isSensitive -and $key -notin $redactedFields) {
        $redactedFields.Add($key)
    }
    $safeConfig.Add([ordered]@{
        key = $key
        value = $(if ($isSensitive) { "[REDACTED]" } else { $entry.value })
        sourceScope = [string]$entry.sourceScope
        isSensitive = $isSensitive
    })
}

$safeAddins = New-Object System.Collections.Generic.List[object]
foreach ($addin in @($raw.loadedAddins)) {
    $name = [string]$addin.name
    $safeAddins.Add([ordered]@{
        name = $name
        path = [string]$addin.path
        sha256 = [string]$addin.sha256
        projectVersion = [string]$addin.projectVersion
        contractVersion = [string]$addin.contractVersion
        isAddin = [bool]$addin.isAddin
        packageRole = Get-PackageRole $name
        isLegacyRolePackage = Test-LegacyRolePackage $name
    })
}
$safeAddins = @($safeAddins.ToArray() | Sort-Object { [string]$_.name })

$safeWorkbooks = New-Object System.Collections.Generic.List[object]
foreach ($workbook in @($raw.workbooks | Sort-Object name)) {
    $worksheets = New-Object System.Collections.Generic.List[object]
    $workbookTables = @($raw.tables | Where-Object {
        $_.workbookName -eq $workbook.name
    })
    foreach ($worksheetName in @(
        $workbookTables |
            ForEach-Object { [string]$_.worksheetName } |
            Sort-Object -Unique
    )) {
        $tableRecords = New-Object System.Collections.Generic.List[object]
        foreach ($table in @($workbookTables | Where-Object {
            $_.worksheetName -eq $worksheetName
        } | Sort-Object name)) {
            $tableRecords.Add([ordered]@{
                name = [string]$table.name
                headers = @($table.headers | ForEach-Object { [string]$_ })
                rowCount = [int]$table.rowCount
                rangeAddress = [string]$table.rangeAddress
            })
        }
        $worksheets.Add([ordered]@{
            name = $worksheetName
            tables = $tableRecords.ToArray()
        })
    }
    $safeWorkbooks.Add([ordered]@{
        name = [string]$workbook.name
        path = [string]$workbook.path
        readOnly = [bool]$workbook.readOnly
        visible = [bool]$workbook.visible
        authorityClass = $(if (Test-HasProperty -Object $workbook -Name "authorityClass") {
            [string]$workbook.authorityClass
        } else {
            Get-AuthorityClass ([string]$workbook.name)
        })
        worksheets = $worksheets.ToArray()
    })
}

$operatorStagingNames = @(
    "ReceivedTally", "AggregateReceived", "ShipmentsTally",
    "BoxBuilder", "ProductionOutput"
)
$operatorStaging = New-Object System.Collections.Generic.List[object]
foreach ($table in @($raw.tables)) {
    if ([string]$table.name -in $operatorStagingNames) {
        $operatorStaging.Add([ordered]@{
            workbookName = [string]$table.workbookName
            tableName = [string]$table.name
            rowCount = [int]$table.rowCount
        })
    }
}

$warnings = New-Object System.Collections.Generic.List[object]
foreach ($warning in @($raw.warnings)) {
    $warnings.Add([ordered]@{
        code = [string]$warning.code
        severity = [string]$warning.severity
        message = [string]$warning.message
        context = [string]$warning.context
    })
}

$legacyAddins = @($safeAddins | Where-Object { $_.isLegacyRolePackage })
if ($legacyAddins.Count -gt 0) {
    $warnings.Add([ordered]@{
        code = "LEGACY_ROLE_ADDINS_LOADED"
        severity = "WARNING"
        message = "Standalone role add-ins are loaded in the pre-D12 package layout."
        context = (($legacyAddins | ForEach-Object { $_.name }) -join ", ")
    })
}
if ((@($safeAddins | Where-Object { $_.packageRole -eq "OPERATIONS" })).Count -gt 0 -and
    $legacyAddins.Count -gt 0) {
    $warnings.Add([ordered]@{
        code = "OPERATIONS_LEGACY_COEXISTENCE"
        severity = "ERROR"
        message = "Operations and standalone role add-ins are loaded together."
        context = "Excel add-ins"
    })
}

$knownVersions = @(
    $safeAddins |
        ForEach-Object { $_.projectVersion } |
        Where-Object { $_ -and $_ -ne "UNKNOWN" } |
        Sort-Object -Unique
)
if ($knownVersions.Count -gt 1) {
    $warnings.Add([ordered]@{
        code = "VERSION_DRIFT"
        severity = "WARNING"
        message = "Loaded invSys add-ins report more than one project version."
        context = ($knownVersions -join ", ")
    })
}

foreach ($table in @($raw.tables)) {
    if ("ROW" -in @($table.headers)) {
        $warnings.Add([ordered]@{
            code = "RETIRED_ROW_HEADER"
            severity = "ERROR"
            message = "$($table.name) contains retired managed header ROW."
            context = "$($table.workbookName)/$($table.worksheetName)/$($table.name)"
        })
    }
}
foreach ($snapshot in @($raw.snapshotReadModels)) {
    if ([bool]$snapshot.isStale) {
        $warnings.Add([ordered]@{
            code = "STALE_READ_MODEL"
            severity = "WARNING"
            message = "Operator or snapshot read model is stale."
            context = "$($snapshot.workbookName)/$($snapshot.tableName)"
        })
    }
}
foreach ($workbook in $safeWorkbooks.ToArray()) {
    if ($workbook.authorityClass -eq "CANONICAL_DOMAIN" -and $workbook.visible) {
        $warnings.Add([ordered]@{
            code = "VISIBLE_CANONICAL_WORKBOOK"
            severity = "WARNING"
            message = "A canonical Domain workbook is visible in the operator session."
            context = [string]$workbook.name
        })
    }
}

$safetyFiles = New-Object System.Collections.Generic.List[object]
foreach ($entry in @($raw.inspectedFiles)) {
    $path = [string]$entry.path
    $beforeHash = [string]$entry.precomputedBeforeSha256
    $afterHash = Get-SharedFileSha256 $path
    if ([string]::IsNullOrWhiteSpace($beforeHash) -or
        [string]::IsNullOrWhiteSpace($afterHash)) {
        $warnings.Add([ordered]@{
            code = "HASH_UNAVAILABLE"
            severity = "WARNING"
            message = "Before/after hash proof is unavailable for an inspected file."
            context = Get-DisplayPath $path
        })
        continue
    }
    if ($beforeHash -ne $afterHash) {
        throw "Read-only safety violation: inspected file changed during extraction: $path"
    }
    $safetyFiles.Add([ordered]@{
        path = Get-DisplayPath $path
        authorityClass = [string]$entry.authorityClass
        beforeSha256 = $beforeHash
        afterSha256 = $afterHash
        unchanged = $true
    })
}

$runtime = [ordered]@{
    schemaVersion = "1.1.0"
    reportType = "runtime-state"
    capturedAtUtc = $ReportTimestampUtc
    toolVersion = "1.0.0"
    session = [ordered]@{
        machineIdentityHash = Get-StringSha256 ([string]$raw.machineName)
        sessionIdentityHash = Get-StringSha256 ([string]$raw.sessionIdentity)
        excelProcessCount = [int]$raw.excelProcessCount
    }
    loadedAddins = $safeAddins
    openWorkbooks = @($safeWorkbooks.ToArray())
    runtimeResolution = [ordered]@{
        warehouseId = [string]$raw.runtimeResolution.warehouseId
        stationId = [string]$raw.runtimeResolution.stationId
        runtimeRoot = [string]$raw.runtimeResolution.runtimeRoot
        resolutionSource = [string]$raw.runtimeResolution.resolutionSource
        connectionStatus = [string]$raw.runtimeResolution.connectionStatus
    }
    config = @($safeConfig.ToArray() | Sort-Object { [string]$_.key })
    currentUser = [ordered]@{
        userId = [string]$raw.currentUser.userId
        capabilities = @($raw.currentUser.capabilities | Sort-Object -Unique)
        signedIn = [bool]$raw.currentUser.signedIn
    }
    domainBridges = @($raw.domainBridges | Sort-Object name)
    inboxSummary = @($raw.inboxSummary | Sort-Object eventType, status)
    processor = [ordered]@{
        lockStatus = [string]$raw.processor.lockStatus
        backlogCount = [int]$raw.processor.backlogCount
        lastRunAtUtc = $raw.processor.lastRunAtUtc
        lastRunStatus = [string]$raw.processor.lastRunStatus
    }
    snapshotReadModels = @($raw.snapshotReadModels | Sort-Object workbookName, tableName)
    operatorStaging = @(
        $operatorStaging.ToArray() |
            Sort-Object { ([string]$_.workbookName) + "|" + ([string]$_.tableName) }
    )
    forms = @($raw.forms | Sort-Object name)
    redaction = [ordered]@{
        policyVersion = "1.1.0"
        rowValuesIncluded = $false
        redactedFieldCount = $redactedFields.Count
        redactedFields = @($redactedFields.ToArray() | Sort-Object)
    }
    safety = [ordered]@{
        inspectionMode = [string]$raw.inspectionMode
        excelStartedByTool = $false
        workbooksOpenedByTool = 0
        workbooksClosedByTool = 0
        workbooksSavedByTool = 0
        refreshActionsInvoked = 0
        processorActionsInvoked = 0
        repairActionsInvoked = 0
        mutatingActionsInvoked = 0
        inspectedFiles = @($safetyFiles.ToArray() | Sort-Object { [string]$_.path })
    }
    warnings = @(
        $warnings.ToArray() |
            Sort-Object { ([string]$_.code) + "|" + ([string]$_.context) } -Unique
    )
}

$jsonPath = Join-Path $resolvedOutput "runtime-state.json"
$markdownPath = Join-Path $resolvedOutput "runtime-state.md"
Write-Utf8NoBom -Path $jsonPath -Content ($runtime | ConvertTo-Json -Depth 100)

& (Join-Path $scriptRoot "validate-json-contract.ps1") `
    -JsonPath $jsonPath `
    -SchemaPath $resolvedSchema
if (-not $?) {
    throw "Runtime JSON did not satisfy its schema."
}

$markdown = New-Object System.Collections.Generic.List[string]
$markdown.Add("# invSys Read-Only Runtime State")
$markdown.Add("")
$markdown.Add("- Schema: 1.1.0")
$markdown.Add("- Capture: " + $ReportTimestampUtc)
$markdown.Add("- Inspection mode: " + $runtime.safety.inspectionMode)
$markdown.Add("- Loaded invSys add-ins: " + (@($runtime.loadedAddins)).Count)
$markdown.Add("- Open workbooks: " + (@($runtime.openWorkbooks)).Count)
$markdown.Add("- Warehouse / station: " +
    $runtime.runtimeResolution.warehouseId + " / " +
    $runtime.runtimeResolution.stationId)
$markdown.Add("- Signed-in invSys user: " +
    $(if ($runtime.currentUser.signedIn) { $runtime.currentUser.userId } else { "SIGNED_OUT" }))
$markdown.Add("- Row values included: False")
$markdown.Add("- Redacted fields: " + $runtime.redaction.redactedFieldCount)
$markdown.Add("- Excel started by tool: False")
$markdown.Add("- Mutating actions invoked: 0")
$markdown.Add("- Inspected files unchanged: " +
    (@($runtime.safety.inspectedFiles)).Count + "/" + (@($runtime.safety.inspectedFiles)).Count)
$markdown.Add("")
$markdown.Add("## Loaded add-ins")
$markdown.Add("")
$markdown.Add("| Name | Role | Version | Legacy role package |")
$markdown.Add("|---|---|---|---|")
foreach ($addin in @($runtime.loadedAddins)) {
    $markdown.Add(
        "| $($addin.name) | $($addin.packageRole) | $($addin.projectVersion) | " +
        "$($addin.isLegacyRolePackage) |"
    )
}
$markdown.Add("")
$markdown.Add("## Warnings")
$markdown.Add("")
if ((@($runtime.warnings)).Count -eq 0) {
    $markdown.Add("- None.")
}
else {
    foreach ($warning in @($runtime.warnings)) {
        $markdown.Add("- $($warning.code): $($warning.message) [$($warning.context)]")
    }
}

Write-Utf8NoBom -Path $markdownPath -Content ($markdown -join "`n")

Write-Host "Read-only invSys runtime extraction complete."
Write-Host ("Inspection mode: " + $runtime.safety.inspectionMode)
Write-Host ("Loaded add-ins: " + (@($runtime.loadedAddins)).Count)
Write-Host ("Open workbooks: " + (@($runtime.openWorkbooks)).Count)
Write-Host ("Redacted fields: " + $runtime.redaction.redactedFieldCount)
Write-Host ("Mutating actions: " + $runtime.safety.mutatingActionsInvoked)
Write-Host ("Output: " + $resolvedOutput)
