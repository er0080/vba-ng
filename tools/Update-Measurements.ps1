<#
.SYNOPSIS
    Writes the golden numbers in docs/measurements.md from the report the replay produced.

.DESCRIPTION
    The golden replay writes golden-report.txt next to its test assembly. This script reads that
    report and rewrites the generated block in docs/measurements.md, so the doc's numbers are copied
    by a machine rather than typed from memory.

    -Verify compares instead of writing and exits 1 when they differ, which is how CI keeps the doc
    honest: a change that moves the pass rate has to update the doc in the same commit.

    Only the goldens are handled here. The corpus and benchmark sections need a real Excel or a
    corpus on disk, so they are written by hand from their own report files and carry the date they
    were measured.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/Update-Measurements.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/Update-Measurements.ps1 -Verify
#>
[CmdletBinding()]
param(
    # The report to read. Defaults to the most recently written golden-report.txt under the golden
    # test project's build output.
    [string] $Report,

    # Compare and fail instead of writing.
    [switch] $Verify
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$docPath = Join-Path $root 'docs/measurements.md'
$beginMarker = '<!-- generated from golden-report.txt by tools/Update-Measurements.ps1 -->'
$endMarker = '<!-- end generated -->'

if (-not $Report) {
    $candidates = Get-ChildItem -LiteralPath (Join-Path $root 'tests/VbaNg.Golden') -Recurse -Filter 'golden-report.txt' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTimeUtc -Descending
    if ($candidates.Count -eq 0) {
        Write-Host "No golden-report.txt found. Run: dotnet test tests/VbaNg.Golden" -ForegroundColor Yellow
        exit 2
    }

    $Report = $candidates[0].FullName
}

# Every data line is "<Area> <n> pass <n> ulp <n> fail <n> unsupported <rate> % <n> left".
$areas = New-Object System.Collections.Generic.List[object]
foreach ($line in Get-Content -LiteralPath $Report -Encoding UTF8) {
    $fields = -split $line
    if ($fields.Count -lt 9 -or $fields[1] -notmatch '^\d+$') { continue }
    $pass = [int] $fields[1]
    $ulp = [int] $fields[3]
    $fail = [int] $fields[5]
    $unsupported = [int] $fields[7]
    $areas.Add([pscustomobject] @{
        Name        = $fields[0]
        Cases       = $pass + $ulp + $fail + $unsupported
        Passing     = $pass + $ulp
        Ulp         = $ulp
        Fail        = $fail
        Unsupported = $unsupported
    })
}

if ($areas.Count -eq 0) { throw "Read no areas from $Report; has its format changed?" }

$cases = ($areas | Measure-Object -Property Cases -Sum).Sum
$passing = ($areas | Measure-Object -Property Passing -Sum).Sum
$ulps = ($areas | Measure-Object -Property Ulp -Sum).Sum
$fails = ($areas | Measure-Object -Property Fail -Sum).Sum
$unsupported = ($areas | Measure-Object -Property Unsupported -Sum).Sum
$rate = [math]::Round(100.0 * $passing / $cases, 1)

function Format-Wrapped([string] $prefix, [string[]] $items, [int] $width = 100) {
    $lines = New-Object System.Collections.Generic.List[string]
    $current = $prefix
    for ($i = 0; $i -lt $items.Count; $i++) {
        $piece = if ($i -lt $items.Count - 1) { $items[$i] + ',' } else { $items[$i] + '.' }
        if (($current.Length + 1 + $piece.Length) -gt $width) {
            $lines.Add($current)
            $current = $piece
        }
        else {
            $current = if ($current) { "$current $piece" } else { $piece }
        }
    }

    $lines.Add($current)
    return $lines
}

$whole = @($areas | Where-Object { $_.Fail -eq 0 -and $_.Unsupported -eq 0 } | Sort-Object Name)
$short = @($areas | Where-Object { $_.Fail -gt 0 -or $_.Unsupported -gt 0 } | Sort-Object Name)

$generated = New-Object System.Collections.Generic.List[string]
$generated.Add($beginMarker)
$ulpNote = if ($ulps -gt 0) { " ($ulps of them within one ulp of VBA)" } else { '' }
# No date: the replay of committed goldens is deterministic, so these numbers belong to the commit,
# not to the day the tests ran. A date here would fail -Verify every morning.
$generated.Add("**$('{0:N0}' -f $cases) cases, $('{0:N0}' -f $passing) pass$ulpNote, $fails fail,")
$generated.Add("$rate%**, over $($areas.Count) areas.")
$generated.Add('')
foreach ($line in Format-Wrapped 'Every case passes in:' ($whole | ForEach-Object { "$($_.Name) $($_.Cases)" })) {
    $generated.Add($line)
}

if ($short.Count -gt 0) {
    $generated.Add('')
    foreach ($area in $short) {
        $areaRate = [math]::Round(100.0 * $area.Passing / $area.Cases, 1)
        $parts = New-Object System.Collections.Generic.List[string]
        if ($area.Ulp -gt 0) { $parts.Add("$($area.Ulp) within one ulp") }
        if ($area.Fail -gt 0) { $parts.Add("$($area.Fail) named misses") }
        if ($area.Unsupported -gt 0) { $parts.Add("$($area.Unsupported) unsupported") }
        $generated.Add("Short of it: $($area.Name), $($area.Cases) cases, $areaRate%: $($parts -join ', ').")
    }
}

$generated.Add($endMarker)

# UTF-8 explicitly: Windows PowerShell reads a file without a BOM as ANSI, and writing that back
# double-encodes every non-ASCII character outside the generated block.
$doc = Get-Content -LiteralPath $docPath -Encoding UTF8
$begin = [array]::IndexOf($doc, $beginMarker)
$end = [array]::IndexOf($doc, $endMarker)
if ($begin -lt 0 -or $end -lt $begin) {
    throw "docs/measurements.md has no generated block; expected the line: $beginMarker"
}

$updated = @()
if ($begin -gt 0) { $updated += $doc[0..($begin - 1)] }
$updated += $generated
if ($end -lt $doc.Count - 1) { $updated += $doc[($end + 1)..($doc.Count - 1)] }

$existing = $doc[$begin..$end]
$same = ($existing.Count -eq $generated.Count)
if ($same) {
    for ($i = 0; $i -lt $generated.Count; $i++) {
        if ($existing[$i] -ne $generated[$i]) { $same = $false; break }
    }
}

$relativeReport = if ($Report.StartsWith($root)) { $Report.Substring($root.Length + 1) } else { $Report }

if ($Verify) {
    if ($same) {
        Write-Host "docs/measurements.md matches $relativeReport" -ForegroundColor Green
        exit 0
    }

    Write-Host "docs/measurements.md is out of date against $relativeReport." -ForegroundColor Red
    Write-Host 'It should read:'
    foreach ($line in $generated) { Write-Host "  $line" }
    Write-Host 'Run tools/Update-Measurements.ps1 and commit the result.'
    exit 1
}

if ($same) {
    Write-Host "docs/measurements.md already matches $relativeReport" -ForegroundColor Green
    exit 0
}

# LF, no BOM, like every other doc here.
$text = ($updated -join "`n") + "`n"
[System.IO.File]::WriteAllText($docPath, $text, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Wrote the golden numbers from $relativeReport into docs/measurements.md" -ForegroundColor Green
exit 0
