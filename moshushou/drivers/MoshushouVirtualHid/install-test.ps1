[CmdletBinding()]
param(
    [string]$PackagePath = '',
    [string]$WindowsKitsRoot = 'D:\Windows Kits\10',
    [switch]$SkipBuild,
    [switch]$SkipSigning
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PackagePath)) {
    $PackagePath = Join-Path $PSScriptRoot 'artifacts\Release\x64\package'
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Quote-ProcessArgument {
    param([Parameter(Mandatory)][string]$Value)
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Invoke-NativeTool {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string[]]$Arguments,
        [switch]$AllowFailure
    )

    & $FilePath @Arguments
    $exitCode = $LASTEXITCODE
    if (-not $AllowFailure -and $exitCode -ne 0) {
        throw "Tool failed with exit code $exitCode`: $FilePath"
    }
    return $exitCode
}

if (-not (Test-IsAdministrator)) {
    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Quote-ProcessArgument $PSCommandPath),
        '-PackagePath', (Quote-ProcessArgument $PackagePath),
        '-WindowsKitsRoot', (Quote-ProcessArgument $WindowsKitsRoot)
    )
    if ($SkipBuild) { $arguments += '-SkipBuild' }
    if ($SkipSigning) { $arguments += '-SkipSigning' }

    $process = Start-Process powershell.exe -Verb RunAs -ArgumentList ($arguments -join ' ') -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Elevated installer failed with exit code $($process.ExitCode)."
    }
    return
}

$os = Get-CimInstance Win32_OperatingSystem
$build = [int]$os.BuildNumber
if (-not [Environment]::Is64BitOperatingSystem) {
    throw 'Only x64 Windows is supported by this driver package.'
}
if ($build -lt 18362) {
    throw "Windows build $build is unsupported. This package requires Windows 10 build 18362 or newer."
}

if (-not $SkipBuild) {
    & (Join-Path $PSScriptRoot 'build.ps1') -Configuration Release -WindowsKitsRoot $WindowsKitsRoot
    if ($LASTEXITCODE -ne 0) {
        throw "Driver build failed with exit code $LASTEXITCODE."
    }
}

if (-not $SkipSigning) {
    & (Join-Path $PSScriptRoot 'sign-test.ps1') `
        -PackagePath $PackagePath `
        -WindowsKitsRoot $WindowsKitsRoot `
        -TrustCertificate
}

$package = (Resolve-Path -LiteralPath $PackagePath).Path
$inf = Join-Path $package 'MoshushouVirtualHid.inf'
if (-not (Test-Path -LiteralPath $inf)) {
    throw "Driver INF is missing: $inf"
}

$devcon = Get-ChildItem (Join-Path $WindowsKitsRoot 'Tools') -Recurse -Filter devcon.exe -File |
    Where-Object { $_.DirectoryName -match '\\x64$' } |
    ForEach-Object {
        $versionText = $_.Directory.Parent.Name
        if ($versionText -match '^10\.0\.\d+\.0$') {
            [pscustomobject]@{ File = $_; Version = [version]$versionText }
        }
    } |
    Sort-Object Version -Descending |
    Select-Object -First 1 -ExpandProperty File |
    Select-Object -ExpandProperty FullName
if (-not $devcon) {
    throw "devcon.exe was not found below $WindowsKitsRoot\Tools"
}

$hardwareId = 'Root\MoshushouVirtualHid'
$findOutput = & $devcon findall $hardwareId 2>&1
$findExitCode = $LASTEXITCODE
$presentDevice = $findOutput | Where-Object { $_ -match '^\s*ROOT\\[^:]+\s*:' } | Select-Object -First 1

if ($findExitCode -eq 0 -and $presentDevice) {
    Invoke-NativeTool $devcon @('update', $inf, $hardwareId) | Out-Null
} else {
    Invoke-NativeTool $devcon @('install', $inf, $hardwareId) | Out-Null
}

Start-Sleep -Seconds 2
$statusOutput = & $devcon status $hardwareId 2>&1
$statusExitCode = $LASTEXITCODE
$statusOutput | ForEach-Object { Write-Host $_ }
if ($statusExitCode -ne 0 -or -not ($statusOutput -match 'Driver is running\.')) {
    throw 'The Moshushou device is present but did not start successfully. Check Device Manager and Code Integrity logs.'
}

Write-Host "Moshushou Virtual HID is installed and running on Windows build $build."
