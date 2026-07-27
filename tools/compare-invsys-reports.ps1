[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$BeforePath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$AfterPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

function Read-ReportJson {
    param([string]$Path)

    try {
        return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json)
    }
    catch {
        throw "Report is not valid JSON: $Path. $($_.Exception.Message)"
    }
}

function Get-JsonKind {
    param($Value)

    if ($null -eq $Value) {
        return "null"
    }
    if ($Value -is [bool]) {
        return "boolean"
    }
    if ($Value -is [string] -or $Value -is [char] -or $Value -is [datetime]) {
        return "scalar"
    }
    if ($Value -is [System.Collections.IDictionary]) {
        return "object"
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return "array"
    }
    if (@($Value.PSObject.Properties).Count -gt 0 -and
        $Value.GetType().FullName -eq "System.Management.Automation.PSCustomObject") {
        return "object"
    }
    return "scalar"
}

function ConvertTo-DisplayValue {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }
    $kind = Get-JsonKind $Value
    if ($kind -in @("object", "array")) {
        return ($Value | ConvertTo-Json -Depth 100 -Compress)
    }
    return $Value
}

function Get-ObjectPropertyMap {
    param($Value)

    $map = @{}
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) {
            $map[[string]$key] = $Value[$key]
        }
        return $map
    }

    foreach ($property in $Value.PSObject.Properties) {
        $map[$property.Name] = $property.Value
    }
    return $map
}

function Add-ReportDifference {
    param(
        [System.Collections.Generic.List[object]]$Differences,
        [string]$Path,
        [string]$ChangeType,
        $Before,
        $After
    )

    $Differences.Add([ordered]@{
        path = $Path
        changeType = $ChangeType
        before = ConvertTo-DisplayValue $Before
        after = ConvertTo-DisplayValue $After
    })
}

function Compare-JsonValue {
    param(
        $Before,
        $After,
        [string]$Path,
        [System.Collections.Generic.List[object]]$Differences
    )

    $beforeKind = Get-JsonKind $Before
    $afterKind = Get-JsonKind $After
    if ($beforeKind -ne $afterKind) {
        Add-ReportDifference $Differences $Path "TYPE_CHANGED" $Before $After
        return
    }

    if ($beforeKind -eq "null") {
        return
    }

    if ($beforeKind -eq "scalar" -or $beforeKind -eq "boolean") {
        if ([string]$Before -cne [string]$After) {
            Add-ReportDifference $Differences $Path "VALUE_CHANGED" $Before $After
        }
        return
    }

    if ($beforeKind -eq "array") {
        $beforeItems = @($Before)
        $afterItems = @($After)
        $maximum = [Math]::Max($beforeItems.Count, $afterItems.Count)
        for ($index = 0; $index -lt $maximum; $index++) {
            $itemPath = $Path + "[" + $index + "]"
            if ($index -ge $beforeItems.Count) {
                Add-ReportDifference $Differences $itemPath "ADDED" $null $afterItems[$index]
            }
            elseif ($index -ge $afterItems.Count) {
                Add-ReportDifference $Differences $itemPath "REMOVED" $beforeItems[$index] $null
            }
            else {
                Compare-JsonValue $beforeItems[$index] $afterItems[$index] $itemPath $Differences
            }
        }
        return
    }

    $beforeMap = Get-ObjectPropertyMap $Before
    $afterMap = Get-ObjectPropertyMap $After
    $names = @($beforeMap.Keys + $afterMap.Keys | Sort-Object -Unique)
    foreach ($name in $names) {
        $propertyPath = if ($Path -eq '$') {
            '$.' + $name
        }
        else {
            $Path + '.' + $name
        }
        if (-not $beforeMap.ContainsKey($name)) {
            Add-ReportDifference $Differences $propertyPath "ADDED" $null $afterMap[$name]
        }
        elseif (-not $afterMap.ContainsKey($name)) {
            Add-ReportDifference $Differences $propertyPath "REMOVED" $beforeMap[$name] $null
        }
        else {
            Compare-JsonValue $beforeMap[$name] $afterMap[$name] $propertyPath $Differences
        }
    }
}

$resolvedBeforePath = (Resolve-Path -LiteralPath $BeforePath).Path
$resolvedAfterPath = (Resolve-Path -LiteralPath $AfterPath).Path
$before = Read-ReportJson $resolvedBeforePath
$after = Read-ReportJson $resolvedAfterPath
$differences = New-Object System.Collections.Generic.List[object]
Compare-JsonValue $before $after '$' $differences

$result = [ordered]@{
    schemaVersion = "1.0.0"
    reportType = "INVSYS_REPORT_COMPARISON"
    beforeFile = [IO.Path]::GetFileName($resolvedBeforePath)
    afterFile = [IO.Path]::GetFileName($resolvedAfterPath)
    beforeSha256 = (Get-FileHash -LiteralPath $resolvedBeforePath -Algorithm SHA256).Hash.ToLowerInvariant()
    afterSha256 = (Get-FileHash -LiteralPath $resolvedAfterPath -Algorithm SHA256).Hash.ToLowerInvariant()
    identical = ($differences.Count -eq 0)
    differenceCount = $differences.Count
    differences = @($differences.ToArray())
}

$outputDirectory = Split-Path -Parent $OutputPath
if ([string]::IsNullOrWhiteSpace($outputDirectory)) {
    $outputDirectory = (Get-Location).Path
}
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}
$resolvedOutputDirectory = (Resolve-Path -LiteralPath $outputDirectory).Path
$resolvedOutputPath = Join-Path $resolvedOutputDirectory ([IO.Path]::GetFileName($OutputPath))
$json = $result | ConvertTo-Json -Depth 100
[IO.File]::WriteAllText(
    $resolvedOutputPath,
    ($json.TrimEnd() + [Environment]::NewLine),
    (New-Object Text.UTF8Encoding($false))
)

Write-Host "Offline invSys report comparison complete."
Write-Host ("Differences: " + $differences.Count)
Write-Host ("Output: " + $resolvedOutputPath)
