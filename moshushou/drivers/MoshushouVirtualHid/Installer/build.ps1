[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',
    [string]$VisualStudioRoot = 'D:\vs2026',
    [string]$WindowsKitsRoot = 'D:\Windows Kits\10'
)

$ErrorActionPreference = 'Stop'

function Invoke-NativeTool {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Tool failed with exit code $LASTEXITCODE`: $FilePath"
    }
}

$msvc = Get-ChildItem (Join-Path $VisualStudioRoot 'bin\VC\Tools\MSVC') -Directory |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1
$driversRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$packageRootCandidates = @(
    (Join-Path $driversRoot '.packages'),
    $env:NUGET_PACKAGES,
    (Join-Path $HOME '.nuget\packages')
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
$packagesRoot = $packageRootCandidates |
    Where-Object {
        (Test-Path -LiteralPath (Join-Path $_ 'microsoft.windows.sdk.cpp')) -and
        (Test-Path -LiteralPath (Join-Path $_ 'microsoft.windows.sdk.cpp.x64'))
    } |
    Select-Object -First 1
if (-not $packagesRoot) {
    throw 'Microsoft.Windows.SDK.CPP NuGet packages were not found.'
}

$sdkHeadersPackage = Get-ChildItem (Join-Path $packagesRoot 'microsoft.windows.sdk.cpp') -Directory |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1
$sdkX64Package = Get-ChildItem (Join-Path $packagesRoot 'microsoft.windows.sdk.cpp.x64') -Directory |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1
$sdkRoot = Join-Path $sdkHeadersPackage.FullName 'c'
$sdkX64Root = Join-Path $sdkX64Package.FullName 'c'
$sdkVersion = Get-ChildItem (Join-Path $sdkRoot 'Include') -Directory |
    Where-Object { $_.Name -match '^10\.0\.\d+\.0$' } |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1 -ExpandProperty Name
if (-not $msvc -or -not $sdkVersion -or -not $sdkX64Package) {
    throw 'MSVC or Windows SDK was not found.'
}

$cl = Join-Path $msvc.FullName 'bin\Hostx64\x64\cl.exe'
$sdkBin = Join-Path $sdkRoot "bin\$sdkVersion\x64"
$mt = Join-Path $sdkBin 'mt.exe'
if (-not (Test-Path -LiteralPath $mt)) {
    throw "Manifest tool was not found: $mt"
}
$output = Join-Path $PSScriptRoot "artifacts\$Configuration\x64"
$object = Join-Path $output 'Moshushou.DriverInstaller.obj'
$exe = Join-Path $output 'Moshushou.DriverInstaller.exe'
New-Item -ItemType Directory -Force -Path $output | Out-Null

$arguments = @(
    '/nologo', '/std:c++20', '/EHsc', '/W4', '/DUNICODE', '/D_UNICODE',
    '/DWINVER=0x0A00', '/D_WIN32_WINNT=0x0A00',
    "/I$($msvc.FullName)\include",
    "/I$sdkRoot\Include\$sdkVersion\shared",
    "/I$sdkRoot\Include\$sdkVersion\um",
    "/I$sdkRoot\Include\$sdkVersion\ucrt",
    "/Fo$object",
    (Join-Path $PSScriptRoot 'Moshushou.DriverInstaller.cpp'),
    '/link', '/machine:x64', '/subsystem:console', '/dynamicbase', '/nxcompat', '/guard:cf',
    '/manifest:embed', '/manifestuac:no',
    "/out:$exe",
    "/manifestinput:$PSScriptRoot\app.manifest",
    "/libpath:$sdkX64Root\um\x64",
    "/libpath:$sdkX64Root\ucrt\x64",
    "/libpath:$($msvc.FullName)\lib\x64",
    'setupapi.lib', 'newdev.lib', 'cfgmgr32.lib', 'crypt32.lib', 'advapi32.lib'
)
if ($Configuration -eq 'Debug') {
    $arguments = @('/Od', '/Zi') + $arguments
} else {
    $arguments = @('/O2', '/GL') + $arguments
    $arguments += '/LTCG'
}

$originalPath = $env:PATH
try {
    $env:PATH = "$sdkBin;$originalPath"
    Invoke-NativeTool $cl $arguments
}
finally {
    $env:PATH = $originalPath
}
Write-Host "Driver installer created: $exe"
