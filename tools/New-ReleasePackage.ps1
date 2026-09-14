<#
.SYNOPSIS
    Builds the release zip and its checksum.

.DESCRIPTION
    One folder, vbang-<version>-win-x64, holding the vbang CLI, the add-in as a single vba-ng.xll,
    the Quickstart sample, and the docs and license. The add-in finds vbang.exe beside
    itself, which is why both share the folder. The version in the name is read off the built
    vbang.dll, never passed along as text.

    release.yml runs this after the tests pass, with -NoBuild, so the zip holds the bits it tested.
    Run it locally to see what a tag would ship.

.PARAMETER Version
    The version a release tag asks for, passed to the build as -p:Version. Omitted, the build keeps
    the version Directory.Build.props declares.

.PARAMETER NoBuild
    Package what the last Release build left.

.PARAMETER Output
    Where the zip and SHA256SUMS go. Emptied first.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/New-ReleasePackage.ps1
#>
[CmdletBinding()]
param(
    [string] $Version,
    [switch] $NoBuild,
    [string] $Output = 'artifacts'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
    $versionArgs = @()
    if ($Version) { $versionArgs = @("-p:Version=$Version") }

    if (-not $NoBuild) {
        dotnet build -c Release @versionArgs
        if ($LASTEXITCODE -ne 0) { throw "dotnet build failed with $LASTEXITCODE." }
    }

    if (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Recurse -Force }
    $staging = Join-Path $Output 'staging'

    dotnet publish src/VbaNg.Cli --no-build -c Release @versionArgs -o $staging
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed with $LASTEXITCODE." }

    $built = (Get-Item -LiteralPath (Join-Path $staging 'vbang.dll')).VersionInfo.ProductVersion
    if ($Version -and $built -ne $Version) { throw "vbang.dll carries $built, not the $Version asked for." }
    $name = "vbang-$built-win-x64"
    $folder = Join-Path $Output $name
    Rename-Item -LiteralPath $staging -NewName $name

    # Excel-DNA packs the add-in and its assemblies into one .xll on a Release build.
    $xll = 'src/VbaNg.AddIn/bin/Release/net10.0-windows/publish/VbaNg.AddIn-AddIn64-packed.xll'
    if (-not (Test-Path -LiteralPath $xll)) { throw "No packed add-in at $xll." }
    Copy-Item -LiteralPath $xll -Destination (Join-Path $folder 'vba-ng.xll')

    # Tracked files only, so a local out/ folder never ships.
    foreach ($file in @(git ls-files samples/Quickstart)) {
        $target = Join-Path $folder $file
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
        Copy-Item -LiteralPath $file -Destination $target
    }

    # The docs keep the repository's layout, so README's and the manual's relative links resolve in the zip.
    Copy-Item -LiteralPath README.md, ARCHITECTURE.md, ROADMAP.md, CLAUDE.md, LICENSE -Destination $folder
    New-Item -ItemType Directory -Force -Path (Join-Path $folder 'docs') | Out-Null
    foreach ($doc in @(git ls-files 'docs/*.md')) {
        Copy-Item -LiteralPath $doc -Destination (Join-Path $folder $doc)
    }

    $zip = Join-Path $Output "$name.zip"
    Compress-Archive -LiteralPath $folder -DestinationPath $zip
    $hash = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
    "$hash  $name.zip" | Out-File -LiteralPath (Join-Path $Output 'SHA256SUMS') -Encoding ascii

    Write-Host "$zip"
    Write-Host "$hash"
}
finally {
    Pop-Location
}
