[CmdletBinding()]
param(
    [ValidateSet('Release')]
    [string]$Configuration = 'Release',
    [ValidateSet('win-x64')]
    [string]$Runtime = 'win-x64',
    [string]$OutputDirectory = 'artifacts\internal-release\win-x64',
    [string]$ArchivePath = 'artifacts\Moshushou-internal-win-x64.zip',
    [string]$LogPath = 'artifacts\internal-release.log',
    [string]$VisualStudioRoot = 'D:\vs2026',
    [string]$WindowsKitsRoot = 'D:\Windows Kits\10',
    [switch]$ValidateOnly,
    [switch]$SkipInstalledDeviceCheck,
    [switch]$NoArchive
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Quote-ProcessArgument {
    param([Parameter(Mandatory)][string]$Value)
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Get-AbsolutePath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$BasePath
    )

    if ([IO.Path]::IsPathRooted($Path)) {
        return [IO.Path]::GetFullPath($Path)
    }
    return [IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Assert-ChildPath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Parent,
        [Parameter(Mandatory)][string]$Description
    )

    $parentPrefix = $Parent.TrimEnd('\') + '\'
    if (-not $Path.StartsWith($parentPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Description must be below $Parent. Actual path: $Path"
    }
}

function Invoke-NativeTool {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    $captureRoot = Join-Path $env:TEMP ("Moshushou.Release." + [Guid]::NewGuid().ToString('N'))
    $standardOutput = Join-Path $captureRoot 'stdout.log'
    $standardError = Join-Path $captureRoot 'stderr.log'
    New-Item -ItemType Directory -Force -Path $captureRoot | Out-Null

    try {
        $argumentLine = ($Arguments | ForEach-Object { Quote-ProcessArgument $_ }) -join ' '
        $process = Start-Process `
            -FilePath $FilePath `
            -ArgumentList $argumentLine `
            -WorkingDirectory $projectRoot `
            -WindowStyle Hidden `
            -RedirectStandardOutput $standardOutput `
            -RedirectStandardError $standardError `
            -Wait `
            -PassThru

        if (Test-Path -LiteralPath $standardOutput) {
            Get-Content -LiteralPath $standardOutput | ForEach-Object { Write-Host $_ }
        }
        if (Test-Path -LiteralPath $standardError) {
            Get-Content -LiteralPath $standardError | ForEach-Object { Write-Host $_ -ForegroundColor Red }
        }
        if ($process.ExitCode -ne 0) {
            throw "Tool failed with exit code $($process.ExitCode)`: $FilePath"
        }
    }
    finally {
        Remove-Item -LiteralPath $captureRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Assert-FileExists {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required release file is missing: $Path"
    }
}

function Get-RelativeReleasePath {
    param(
        [Parameter(Mandatory)][string]$Root,
        [Parameter(Mandatory)][string]$Path
    )

    $rootUri = New-Object Uri(($Root.TrimEnd('\') + '\'))
    $pathUri = New-Object Uri($Path)
    return [Uri]::UnescapeDataString($rootUri.MakeRelativeUri($pathUri).ToString()).Replace('/', '\')
}

$projectRoot = $PSScriptRoot
$artifactsRoot = [IO.Path]::GetFullPath((Join-Path $projectRoot 'artifacts'))
$publishRoot = Get-AbsolutePath -Path $OutputDirectory -BasePath $projectRoot
$archive = Get-AbsolutePath -Path $ArchivePath -BasePath $projectRoot
$log = Get-AbsolutePath -Path $LogPath -BasePath $projectRoot

Assert-ChildPath -Path $publishRoot -Parent $artifactsRoot -Description 'OutputDirectory'
Assert-ChildPath -Path $archive -Parent $artifactsRoot -Description 'ArchivePath'
Assert-ChildPath -Path $log -Parent $artifactsRoot -Description 'LogPath'

if (-not [Environment]::Is64BitOperatingSystem -or -not [Environment]::Is64BitProcess) {
    throw 'The internal release must be built from 64-bit PowerShell on x64 Windows.'
}

if (-not (Test-IsAdministrator)) {
    Write-Host 'Administrator privileges are required for release signing and device validation. Requesting elevation...'
    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', (Quote-ProcessArgument $PSCommandPath),
        '-Configuration', $Configuration,
        '-Runtime', $Runtime,
        '-OutputDirectory', (Quote-ProcessArgument $publishRoot),
        '-ArchivePath', (Quote-ProcessArgument $archive),
        '-LogPath', (Quote-ProcessArgument $log),
        '-VisualStudioRoot', (Quote-ProcessArgument $VisualStudioRoot),
        '-WindowsKitsRoot', (Quote-ProcessArgument $WindowsKitsRoot)
    )
    if ($ValidateOnly) { $arguments += '-ValidateOnly' }
    if ($SkipInstalledDeviceCheck) { $arguments += '-SkipInstalledDeviceCheck' }
    if ($NoArchive) { $arguments += '-NoArchive' }

    $process = Start-Process powershell.exe `
        -Verb RunAs `
        -ArgumentList ($arguments -join ' ') `
        -WorkingDirectory $projectRoot `
        -Wait `
        -PassThru
    if ($process.ExitCode -ne 0) {
        throw "Elevated release process failed with exit code $($process.ExitCode)."
    }
    return
}

$transcriptStarted = $false
try {
    New-Item -ItemType Directory -Force -Path (Split-Path $log -Parent) | Out-Null
    Start-Transcript -LiteralPath $log -Force | Out-Null
    $transcriptStarted = $true
}
catch {
    Write-Warning "Unable to start release transcript: $($_.Exception.Message)"
}

trap {
    $errorText = $_.Exception.ToString()
    Write-Host "RELEASE_ERROR: $errorText" -ForegroundColor Red
    [Console]::Error.WriteLine($errorText)
    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }
    exit 1
}

$os = Get-CimInstance Win32_OperatingSystem
$windowsBuild = [int]$os.BuildNumber
if ($windowsBuild -lt 18362) {
    throw "Windows build $windowsBuild is unsupported. Build 18362 or newer is required."
}

$driverRoot = Join-Path $projectRoot 'drivers\MoshushouVirtualHid'
$driverBuildScript = Join-Path $driverRoot 'build.ps1'
$driverSignScript = Join-Path $driverRoot 'sign-test.ps1'
$installerBuildScript = Join-Path $driverRoot 'Installer\build.ps1'
$driverPackage = Join-Path $driverRoot 'artifacts\Release\x64\package'
$projectFile = Join-Path $projectRoot 'moshushou.csproj'

foreach ($requiredInput in @(
    $driverBuildScript,
    $driverSignScript,
    $installerBuildScript,
    $projectFile
)) {
    Assert-FileExists $requiredInput
}

if (-not $ValidateOnly) {
    Write-Host '[1/6] Building the UMDF2 Virtual HID driver...'
    & $driverBuildScript `
        -Configuration $Configuration `
        -VisualStudioRoot $VisualStudioRoot `
        -WindowsKitsRoot $WindowsKitsRoot

    Write-Host '[2/6] Test-signing the driver package and trusting the test certificate...'
    & $driverSignScript `
        -PackagePath $driverPackage `
        -WindowsKitsRoot $WindowsKitsRoot `
        -TrustCertificate

    Write-Host '[3/6] Building Moshushou.DriverInstaller.exe...'
    & $installerBuildScript `
        -Configuration $Configuration `
        -VisualStudioRoot $VisualStudioRoot `
        -WindowsKitsRoot $WindowsKitsRoot

    Write-Host '[4/6] Publishing the self-contained Windows x64 application...'
    if (Test-Path -LiteralPath $publishRoot) {
        Remove-Item -LiteralPath $publishRoot -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $publishRoot | Out-Null

    $dotnet = Get-Command dotnet -ErrorAction Stop | Select-Object -ExpandProperty Source
    Invoke-NativeTool $dotnet @(
        'publish', $projectFile,
        '-c', $Configuration,
        '-r', $Runtime,
        '--self-contained', 'true',
        '-p:PublishSingleFile=false',
        '-p:DebugType=None',
        '-p:DebugSymbols=false',
        '-o', $publishRoot
    )
}
elseif (-not (Test-Path -LiteralPath $publishRoot -PathType Container)) {
    throw "ValidateOnly publish directory does not exist: $publishRoot"
}

Write-Host '[5/6] Validating release contents, signatures, and installed device...'
$requiredReleaseFiles = @(
    'moshushou.exe',
    'moshushou.dll',
    'businfo.json',
    'wco_data\WeChatOCR.exe',
    'wco_data\mmmojo.dll',
    'wco_data\mmmojo_64.dll',
    'wco_data\Model\FPOCRRecog.xnet',
    'wxonnx\yolo11n_wxscreen_fixed.onnx',
    'Driver\Moshushou.DriverInstaller.exe',
    'Driver\package\MoshushouVirtualHid.dll',
    'Driver\package\MoshushouVirtualHid.inf',
    'Driver\package\MoshushouVirtualHid.cat',
    'Driver\package\MoshushouVirtualHidTest.cer'
)
foreach ($relativePath in $requiredReleaseFiles) {
    Assert-FileExists (Join-Path $publishRoot $relativePath)
}

foreach ($directoryName in @('wco_data', 'wxonnx', 'Driver')) {
    $directory = Join-Path $publishRoot $directoryName
    $fileCount = @(Get-ChildItem -LiteralPath $directory -Recurse -File).Count
    if ($fileCount -eq 0) {
        throw "Release directory is empty: $directory"
    }
}

$publishedDriverDll = Join-Path $publishRoot 'Driver\package\MoshushouVirtualHid.dll'
$publishedDriverCat = Join-Path $publishRoot 'Driver\package\MoshushouVirtualHid.cat'
$driverSignatures = @()
foreach ($signedFile in @($publishedDriverDll, $publishedDriverCat)) {
    $signature = Get-AuthenticodeSignature -FilePath $signedFile
    if ($signature.Status -ne [Management.Automation.SignatureStatus]::Valid) {
        throw "Driver signature is not valid: $signedFile ($($signature.Status))"
    }
    $driverSignatures += [ordered]@{
        path = Get-RelativeReleasePath -Root $publishRoot -Path $signedFile
        status = $signature.Status.ToString()
        subject = $signature.SignerCertificate.Subject
        thumbprint = $signature.SignerCertificate.Thumbprint
    }
}

$publishedInstaller = Join-Path $publishRoot 'Driver\Moshushou.DriverInstaller.exe'
$installerStatusOutput = & $publishedInstaller status 2>&1
$installerStatusExitCode = $LASTEXITCODE
$installerStatusOutput | ForEach-Object { Write-Host $_ }
if (-not $SkipInstalledDeviceCheck -and $installerStatusExitCode -ne 0) {
    throw "Installed Virtual HID status check failed with exit code $installerStatusExitCode."
}

$installedDeviceCount = $null
if (-not $SkipInstalledDeviceCheck) {
    $matchingDevices = @(
        Get-CimInstance Win32_PnPEntity -ErrorAction Stop |
            Where-Object { @($_.HardwareID) -contains 'Root\MoshushouVirtualHid' }
    )
    $installedDeviceCount = $matchingDevices.Count
    if ($installedDeviceCount -ne 1) {
        throw "Expected exactly one present device with Hardware ID Root\MoshushouVirtualHid, found $installedDeviceCount."
    }
    if ($matchingDevices[0].ConfigManagerErrorCode -ne 0) {
        throw "The installed Root\MoshushouVirtualHid device has ConfigManager error $($matchingDevices[0].ConfigManagerErrorCode)."
    }
}

$manifestPath = Join-Path $publishRoot 'release-manifest.json'
$releaseFiles = @(
    Get-ChildItem -LiteralPath $publishRoot -Recurse -File |
        Where-Object { $_.FullName -ne $manifestPath } |
        Sort-Object FullName |
        ForEach-Object {
            [ordered]@{
                path = Get-RelativeReleasePath -Root $publishRoot -Path $_.FullName
                length = $_.Length
                sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
            }
        }
)
$manifest = [ordered]@{
    product = 'Moshushou'
    channel = 'internal-test'
    generatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    configuration = $Configuration
    runtime = $Runtime
    selfContained = $true
    minimumWindowsBuild = 18362
    buildMachineWindowsBuild = $windowsBuild
    hardwareId = 'Root\MoshushouVirtualHid'
    installedDeviceCount = $installedDeviceCount
    installerStatusExitCode = $installerStatusExitCode
    driverSignatures = $driverSignatures
    files = $releaseFiles
}
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Host '[6/6] Creating the internal release archive...'
if (-not $NoArchive) {
    $archiveDirectory = Split-Path $archive -Parent
    New-Item -ItemType Directory -Force -Path $archiveDirectory | Out-Null
    Remove-Item -LiteralPath $archive -Force -ErrorAction SilentlyContinue
    Compress-Archive -Path (Join-Path $publishRoot '*') -DestinationPath $archive -CompressionLevel Optimal
    Assert-FileExists $archive

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($archive)
    try {
        $zipEntries = @($zip.Entries | Where-Object { -not [string]::IsNullOrEmpty($_.Name) })
        $zipEntryNames = @($zipEntries | ForEach-Object { $_.FullName.Replace('\', '/') })
        $publishedFileCount = @(Get-ChildItem -LiteralPath $publishRoot -Recurse -File).Count
        if ($zipEntries.Count -ne $publishedFileCount) {
            throw "ZIP contains $($zipEntries.Count) files, but the publish directory contains $publishedFileCount."
        }
        foreach ($requiredZipEntry in @(
            'moshushou.exe',
            'Driver/Moshushou.DriverInstaller.exe',
            'Driver/package/MoshushouVirtualHid.inf',
            'release-manifest.json'
        )) {
            if ($zipEntryNames -notcontains $requiredZipEntry) {
                throw "Required ZIP entry is missing: $requiredZipEntry"
            }
        }
    }
    finally {
        $zip.Dispose()
    }
}

$publishSize = (Get-ChildItem -LiteralPath $publishRoot -Recurse -File | Measure-Object Length -Sum).Sum
Write-Host ''
Write-Host 'Internal release completed successfully.'
Write-Host "Publish directory: $publishRoot"
if (-not $NoArchive) {
    Write-Host "ZIP archive:       $archive"
}
Write-Host "Published files:   $(@(Get-ChildItem -LiteralPath $publishRoot -Recurse -File).Count)"
Write-Host ('Published size:    {0:N2} MB' -f ($publishSize / 1MB))
Write-Host "Manifest:          $manifestPath"

if ($transcriptStarted) {
    Stop-Transcript | Out-Null
}
