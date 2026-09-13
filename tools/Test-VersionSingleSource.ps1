<#
.SYNOPSIS
    Checks that the version is written down in exactly one place.

.DESCRIPTION
    Directory.Build.props declares VersionPrefix and VersionSuffix, and everything else reads the
    version off an assembly at run time. This guard fails if that stops being true:

      1. another project or props file declares a version of its own;
      2. the declared version appears as a literal in a tracked file that is not allowed to carry
         one, which is how a doc starts rotting;
      3. a built assembly disagrees with the declaration.

    Written for CI, which can also override the version with -p:Version=<tag>; pass that same
    version as -Expected so the guard checks the build it just produced.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools/Test-VersionSingleSource.ps1
#>
[CmdletBinding()]
param(
    # The version the build should carry. Defaults to what Directory.Build.props declares.
    [string] $Expected,

    # Tracked files allowed to spell a version out: the declaration, the version table, and the
    # history, which records what each release was.
    [string[]] $Allowed = @('Directory.Build.props', 'CLAUDE.md', 'ROADMAP.md', 'docs/build-log.md')
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

function Add-Failure([string] $message) { $failures.Add($message) }

# 1. The declaration.
$propsPath = Join-Path $root 'Directory.Build.props'
[xml] $props = Get-Content -LiteralPath $propsPath -Raw
$group = $props.Project.PropertyGroup
$prefix = [string] $group.VersionPrefix
$suffix = [string] $group.VersionSuffix
if (-not $prefix) { Add-Failure "Directory.Build.props declares no VersionPrefix." }
$declared = if ($suffix) { "$prefix-$suffix" } else { $prefix }
if (-not $Expected) { $Expected = $declared }

# 2. Nobody else declares one.
$versionElements = 'Version', 'VersionPrefix', 'VersionSuffix', 'AssemblyVersion', 'FileVersion', 'InformationalVersion'
$projectFiles = Get-ChildItem -LiteralPath $root -Recurse -File |
    Where-Object { $_.Extension -in '.csproj', '.props', '.targets' } |
    Where-Object { $segments = $_.FullName.Split([IO.Path]::DirectorySeparatorChar); $segments -notcontains 'bin' -and $segments -notcontains 'obj' -and $_.Name -ne 'Directory.Build.props' }
foreach ($file in $projectFiles) {
    $text = Get-Content -LiteralPath $file.FullName -Raw
    foreach ($element in $versionElements) {
        if ($text -match "<$element>") {
            Add-Failure "$($file.FullName.Substring($root.Length + 1)) declares <$element>; the version belongs to Directory.Build.props alone."
        }
    }
}

# 3. No literal anywhere it does not belong. Tracked files only: a build output may carry it.
Push-Location $root
try { $tracked = & git ls-files } finally { Pop-Location }
$allowedSet = [System.Collections.Generic.HashSet[string]]::new([string[]] $Allowed, [System.StringComparer]::OrdinalIgnoreCase)
foreach ($relative in $tracked) {
    if ($allowedSet.Contains($relative)) { continue }
    $path = Join-Path $root $relative
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
    # Text files only; a workbook or a dll is not a doc.
    if ($relative -match '\.(xlsx|xlsm|xlam|xll|dll|pdb|png|frx|zip)$') { continue }
    # NuGet writes the version into every lock file's project references. Generated, like a build
    # output, so it is not a second declaration; restore regenerates them after a version change.
    if ((Split-Path $relative -Leaf) -eq 'packages.lock.json') { continue }
    $matches = Select-String -LiteralPath $path -SimpleMatch -Pattern $Expected -ErrorAction SilentlyContinue
    foreach ($match in $matches) {
        Add-Failure "${relative}:$($match.LineNumber) spells the version out: read it from the assembly instead, or add the file to -Allowed if it is a record."
    }
}

# 4. What was built agrees, when it was built after the declaration. An output older than
# Directory.Build.props simply predates it, so it is stale rather than wrong.
$declaredAt = (Get-Item -LiteralPath $propsPath).LastWriteTimeUtc
$assemblies = Get-ChildItem -LiteralPath $root -Recurse -Filter 'vbang.dll' -File -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName.Split([IO.Path]::DirectorySeparatorChar) -contains 'bin' }
foreach ($assembly in $assemblies) {
    $relative = $assembly.FullName.Substring($root.Length + 1)
    if ($assembly.LastWriteTimeUtc -lt $declaredAt) {
        Write-Host "  skipped $relative, built before the current declaration" -ForegroundColor DarkGray
        continue
    }

    $info = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($assembly.FullName).ProductVersion
    if ($info -and $info -ne $Expected) {
        Add-Failure "$relative carries '$info', not '$Expected'; rebuild, or pass -Expected the version CI built with."
    }
}

if ($failures.Count -gt 0) {
    Write-Host "Version check failed ($($failures.Count)):" -ForegroundColor Red
    foreach ($failure in $failures) { Write-Host "  $failure" }
    exit 1
}

Write-Host "Version is single-sourced: $Expected" -ForegroundColor Green
exit 0
