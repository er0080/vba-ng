<#
.SYNOPSIS
    Runs the whole gate: what CI runs, plus the half CI cannot.

.DESCRIPTION
    GitHub-hosted runners have no Office, so .github/workflows/ci.yml stops at the tests that need
    no Excel. This script runs the same steps in the same order and then the Excel half: the E2E
    tests, and with -Benchmarks the Release benchmarks. Run it before pushing, and before tagging a
    release.

    Every step is reported as it ran. The exit code is 0 only if every step passed, so this is
    usable as a pre-push hook.

.PARAMETER SkipExcel
    Stop after the steps CI also runs. Use it when Excel is busy or absent.

.PARAMETER Benchmarks
    Also run the Release benchmarks, which take several minutes and start Excel more than once.

.PARAMETER Package
    Also build the release zip with tools/New-ReleasePackage.ps1 and install it the way README.md
    does, in a hidden Excel. Run it before tagging.

.PARAMETER Corpus
    Corpus roots for the parse rate and the compile scorecard, separated by ';'. Defaults to
    VBANG_CORPUS when that is already set.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/Invoke-Gate.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/Invoke-Gate.ps1 -Benchmarks
#>
[CmdletBinding()]
param(
    [switch] $SkipExcel,
    [switch] $Benchmarks,
    [switch] $Package,
    [string] $Corpus = $env:VBANG_CORPUS
)

$ErrorActionPreference = 'Continue'
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root

$results = New-Object System.Collections.Generic.List[object]

function Invoke-Step([string] $name, [scriptblock] $body) {
    Write-Host ""
    Write-Host "=== $name" -ForegroundColor Cyan
    $started = Get-Date
    & $body
    $code = if ($LASTEXITCODE -eq $null) { 0 } else { $LASTEXITCODE }
    $results.Add([pscustomobject] @{
        Step    = $name
        Passed  = ($code -eq 0)
        Seconds = [int] ((Get-Date) - $started).TotalSeconds
    })
    if ($code -ne 0) { Write-Host "$name failed with $code" -ForegroundColor Red }
}

try {
    # The steps CI runs, in CI's order, so a local pass means the same thing as a remote one.
    Invoke-Step 'restore (locked)' { dotnet restore --locked-mode }
    Invoke-Step 'build' { dotnet build --no-restore }
    Invoke-Step 'format' { dotnet format --verify-no-changes }

    if ($Corpus) {
        # With a corpus configured, the same test run also writes the parse and compile reports.
        Write-Host "Corpus: $Corpus" -ForegroundColor DarkGray
        $env:VBANG_CORPUS = $Corpus
    }

    Invoke-Step 'test (unit and golden)' { dotnet test --no-build }
    Invoke-Step 'version single-sourced' { powershell -ExecutionPolicy Bypass -File tools/Test-VersionSingleSource.ps1 }
    Invoke-Step 'measurements match the report' { powershell -ExecutionPolicy Bypass -File tools/Update-Measurements.ps1 -Verify }

    if (-not $SkipExcel) {
        # Everything below starts Excel. The tests use hidden throwaway instances and quit the ones
        # they start; they never touch a workbook that is already open.
        $env:VBANG_E2E = '1'
        Invoke-Step 'end to end (Excel)' { dotnet test tests/VbaNg.E2E }

        if ($Benchmarks) {
            # Release only: a Debug run measures the Debug runtime rather than the design.
            Invoke-Step 'benchmarks (Excel, Release)' { dotnet test tests/VbaNg.E2E -c Release }
        }

        if ($Package) {
            Invoke-Step 'release zip' { powershell -ExecutionPolicy Bypass -File tools/New-ReleasePackage.ps1 }
            $zip = @(Get-ChildItem artifacts -Filter 'vbang-*-win-x64.zip' -ErrorAction SilentlyContinue)
            $env:VBANG_PACKAGE = if ($zip.Count -eq 1) { $zip[0].FullName } else { '' }
            # Without a zip the test would skip and pass, so the step fails instead.
            Invoke-Step 'release zip installs (Excel)' { if ($env:VBANG_PACKAGE) { dotnet test tests/VbaNg.E2E --filter-class '*ReleasePackageTests' } else { cmd /c exit 1 } }
        }
    }
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "=== gate" -ForegroundColor Cyan
foreach ($result in $results) {
    $mark = if ($result.Passed) { 'pass' } else { 'FAIL' }
    $colour = if ($result.Passed) { 'Green' } else { 'Red' }
    Write-Host ("  {0,-4} {1,-34} {2,4}s" -f $mark, $result.Step, $result.Seconds) -ForegroundColor $colour
}

$failed = @($results | Where-Object { -not $_.Passed })
if ($SkipExcel) { Write-Host "  Excel steps skipped." -ForegroundColor DarkGray }
if (-not $Benchmarks -and -not $SkipExcel) { Write-Host "  Benchmarks skipped; pass -Benchmarks to include them." -ForegroundColor DarkGray }

if ($failed.Count -gt 0) {
    Write-Host "$($failed.Count) step(s) failed." -ForegroundColor Red
    exit 1
}

Write-Host "Every step passed." -ForegroundColor Green
exit 0
