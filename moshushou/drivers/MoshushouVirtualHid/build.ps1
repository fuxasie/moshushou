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
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Tool failed with exit code $LASTEXITCODE`: $FilePath"
    }
}

$driverRoot = $PSScriptRoot
$driversRoot = Split-Path $driverRoot -Parent
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
    throw 'Microsoft.Windows.SDK.CPP NuGet packages were not found in the project or user package cache.'
}

$msvc = Get-ChildItem (Join-Path $VisualStudioRoot 'bin\VC\Tools\MSVC') -Directory |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1
if (-not $msvc) {
    throw "MSVC was not found below $VisualStudioRoot"
}

$sdkHeadersPackage = Get-ChildItem (Join-Path $packagesRoot 'microsoft.windows.sdk.cpp') -Directory |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1
$sdkX64Package = Get-ChildItem (Join-Path $packagesRoot 'microsoft.windows.sdk.cpp.x64') -Directory |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1
if (-not $sdkHeadersPackage -or -not $sdkX64Package) {
    throw "Windows SDK NuGet packages are missing below $packagesRoot. Restore MoshushouVirtualHid.vcxproj first."
}

$sdkRoot = Join-Path $sdkHeadersPackage.FullName 'c'
$sdkX64Root = Join-Path $sdkX64Package.FullName 'c'
$sdkVersion = Get-ChildItem (Join-Path $sdkRoot 'Include') -Directory |
    Where-Object { $_.Name -match '^10\.0\.\d+\.0$' } |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1 -ExpandProperty Name
if (-not $sdkVersion) {
    throw "A versioned SDK include directory was not found below $sdkRoot"
}

$wdkVersion = Get-ChildItem (Join-Path $WindowsKitsRoot 'build') -Directory |
    Where-Object { $_.Name -match '^10\.0\.\d+\.0$' } |
    Sort-Object { [version]$_.Name } -Descending |
    Select-Object -First 1 -ExpandProperty Name
if (-not $wdkVersion) {
    throw "WDK build files were not found below $WindowsKitsRoot"
}

$umdfVersion = '2.15'
$cl = Join-Path $msvc.FullName 'bin\Hostx64\x64\cl.exe'
$link = Join-Path $msvc.FullName 'bin\Hostx64\x64\link.exe'
$rc = Join-Path $sdkRoot "bin\$sdkVersion\x64\rc.exe"
$infVerif = Join-Path $WindowsKitsRoot "Tools\$wdkVersion\x64\infverif.exe"
$inf2Cat = Join-Path $WindowsKitsRoot "bin\$wdkVersion\x86\Inf2Cat.exe"

foreach ($tool in @($cl, $link, $rc, $infVerif, $inf2Cat)) {
    if (-not (Test-Path -LiteralPath $tool)) {
        throw "Required build tool was not found: $tool"
    }
}

$objectRoot = Join-Path $driverRoot "obj\Direct\$Configuration\x64"
$outputRoot = Join-Path $driverRoot "artifacts\$Configuration\x64"
$packageRoot = Join-Path $outputRoot 'package'
New-Item -ItemType Directory -Force -Path $objectRoot, $outputRoot, $packageRoot | Out-Null

$compileArguments = @(
    '/nologo', '/c', '/W4', '/TC', '/GS', '/guard:cf',
    '/D_UNICODE', '/DUNICODE', '/DWINVER=0x0A00', '/D_WIN32_WINNT=0x0A00',
    "/I$($msvc.FullName)\include",
    "/I$WindowsKitsRoot\Include\wdf\umdf\$umdfVersion",
    "/I$WindowsKitsRoot\Include\$wdkVersion\km",
    "/I$sdkRoot\Include\$sdkVersion\shared",
    "/I$sdkRoot\Include\$sdkVersion\um",
    "/I$sdkRoot\Include\$sdkVersion\ucrt",
    "/Fo$objectRoot\",
    "/Fd$objectRoot\vc140.pdb"
)

if ($Configuration -eq 'Debug') {
    $compileArguments += @('/Od', '/Zi', '/D_DEBUG')
} else {
    $compileArguments += @('/O2', '/Gy', '/DNDEBUG')
}

$compileArguments += @(
    (Join-Path $driverRoot 'Driver.c'),
    (Join-Path $driverRoot 'HidQueue.c'),
    (Join-Path $driverRoot 'UmdfHidUtil.c')
)
Invoke-NativeTool $cl $compileArguments

$resourcePath = Join-Path $objectRoot 'MoshushouVirtualHid.res'
Invoke-NativeTool $rc @(
    '/nologo',
    "/i$sdkRoot\Include\$sdkVersion\shared",
    "/i$sdkRoot\Include\$sdkVersion\um",
    "/fo$resourcePath",
    (Join-Path $driverRoot 'MoshushouVirtualHid.rc')
)

$dllPath = Join-Path $outputRoot 'MoshushouVirtualHid.dll'
$pdbPath = Join-Path $outputRoot 'MoshushouVirtualHid.pdb'
$linkArguments = @(
    '/nologo', '/dll', '/machine:x64', '/subsystem:windows',
    '/dynamicbase', '/nxcompat', '/guard:cf', '/incremental:no',
    "/out:$dllPath", "/pdb:$pdbPath",
    "/libpath:$WindowsKitsRoot\Lib\wdf\umdf\x64\$umdfVersion",
    "/libpath:$sdkX64Root\um\x64",
    "/libpath:$sdkX64Root\ucrt\x64",
    "/libpath:$($msvc.FullName)\lib\x64",
    (Join-Path $objectRoot 'Driver.obj'),
    (Join-Path $objectRoot 'HidQueue.obj'),
    (Join-Path $objectRoot 'UmdfHidUtil.obj'),
    $resourcePath,
    'WdfDriverStubUm.lib', 'ntdll.lib', 'mincore.lib'
)
if ($Configuration -eq 'Debug') {
    $linkArguments += '/debug'
} else {
    $linkArguments += @('/opt:ref', '/opt:icf')
}
Invoke-NativeTool $link $linkArguments

$packageDll = Join-Path $packageRoot 'MoshushouVirtualHid.dll'
$packageInf = Join-Path $packageRoot 'MoshushouVirtualHid.inf'
$packageCat = Join-Path $packageRoot 'MoshushouVirtualHid.cat'
$packageCer = Join-Path $packageRoot 'MoshushouVirtualHidTest.cer'
$packagePfx = Join-Path $packageRoot 'MoshushouVirtualHidTest.tmp.pfx'
foreach ($oldFile in @($packageDll, $packageInf, $packageCat, $packageCer, $packagePfx)) {
    Remove-Item -LiteralPath $oldFile -Force -ErrorAction SilentlyContinue
}

Copy-Item -LiteralPath $dllPath -Destination $packageDll
Copy-Item -LiteralPath (Join-Path $driverRoot 'MoshushouVirtualHid.inx') -Destination $packageInf

Invoke-NativeTool $infVerif @('/v', $packageInf)
Invoke-NativeTool $inf2Cat @(
    "/driver:$packageRoot",
    '/os:10_19H1_X64,10_VB_X64,10_CO_X64,10_NI_X64,10_GE_X64,10_25H2_X64',
    '/verbose'
)

Write-Host "Driver package created: $packageRoot"
