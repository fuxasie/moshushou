[CmdletBinding()]
param(
    [string]$PackagePath = '',
    [string]$CertificateSubject = 'CN=Moshushou Virtual HID Test',
    [string]$WindowsKitsRoot = 'D:\Windows Kits\10',
    [switch]$TrustCertificate
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

function Find-WdkTool {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [ValidateSet('x64', 'x86')]
        [string]$Architecture,
        [switch]$PreferWindows10Compatible
    )

    $tools = Get-ChildItem (Join-Path $WindowsKitsRoot 'bin') -Recurse -Filter $Name -File |
        Where-Object { $_.DirectoryName -match "\\$Architecture$" } |
        ForEach-Object {
            $versionText = $_.Directory.Parent.Name
            if ($versionText -match '^10\.0\.\d+\.0$') {
                [pscustomobject]@{
                    File = $_
                    Version = [version]$versionText
                }
            }
        }

    if ($PreferWindows10Compatible) {
        $compatible = $tools |
            Where-Object { $_.Version -le [version]'10.0.19041.0' } |
            Sort-Object Version -Descending |
            Select-Object -First 1
        if ($compatible) {
            return $compatible.File.FullName
        }
    }

    return $tools |
        Sort-Object Version -Descending |
        Select-Object -First 1 -ExpandProperty File |
        Select-Object -ExpandProperty FullName
}

if (-not (Test-IsAdministrator)) {
    throw 'Test signing must run from an elevated PowerShell session.'
}

$package = (Resolve-Path -LiteralPath $PackagePath).Path
$dll = Join-Path $package 'MoshushouVirtualHid.dll'
$inf = Join-Path $package 'MoshushouVirtualHid.inf'
$cat = Join-Path $package 'MoshushouVirtualHid.cat'
$cer = Join-Path $package 'MoshushouVirtualHidTest.cer'

foreach ($file in @($dll, $inf)) {
    if (-not (Test-Path -LiteralPath $file)) {
        throw "Driver package file is missing: $file"
    }
}

$signTool = Find-WdkTool -Name 'signtool.exe' -Architecture x64 -PreferWindows10Compatible
$inf2Cat = Find-WdkTool -Name 'Inf2Cat.exe' -Architecture x86
if (-not $signTool -or -not $inf2Cat) {
    throw "SignTool or Inf2Cat was not found below $WindowsKitsRoot"
}

$certificate = Get-ChildItem Cert:\LocalMachine\My |
    Where-Object {
        $_.Subject -eq $CertificateSubject -and
        $_.HasPrivateKey -and
        $_.NotAfter -gt (Get-Date).AddDays(30) -and
        ($_.EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3')
    } |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1

if (-not $certificate) {
    $certificate = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $CertificateSubject `
        -CertStoreLocation 'Cert:\LocalMachine\My' `
        -KeyAlgorithm RSA `
        -KeyLength 3072 `
        -HashAlgorithm SHA256 `
        -KeyExportPolicy NonExportable `
        -NotAfter (Get-Date).AddYears(5)
}

Export-Certificate -Cert $certificate -FilePath $cer -Force | Out-Null

if ($TrustCertificate) {
    foreach ($store in @('Cert:\LocalMachine\Root', 'Cert:\LocalMachine\TrustedPublisher')) {
        $existing = Get-ChildItem $store | Where-Object Thumbprint -eq $certificate.Thumbprint
        if (-not $existing) {
            Import-Certificate -FilePath $cer -CertStoreLocation $store | Out-Null
        }
    }
}

Invoke-NativeTool $signTool @(
    'sign', '/v', '/fd', 'SHA256',
    '/sm', '/s', 'My', '/sha1', $certificate.Thumbprint,
    $dll
)

Remove-Item -LiteralPath $cat -Force -ErrorAction SilentlyContinue
Invoke-NativeTool $inf2Cat @(
    "/driver:$package",
    '/os:10_19H1_X64,10_VB_X64,10_CO_X64,10_NI_X64,10_GE_X64,10_25H2_X64',
    '/verbose'
)

Invoke-NativeTool $signTool @(
    'sign', '/v', '/fd', 'SHA256',
    '/sm', '/s', 'My', '/sha1', $certificate.Thumbprint,
    $cat
)

foreach ($file in @($dll, $cat)) {
    $signature = Get-AuthenticodeSignature -FilePath $file
    if (-not $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) {
        throw "The expected test signature is missing from $file"
    }
}

Write-Host "Test-signed package: $package"
Write-Host "Certificate thumbprint: $($certificate.Thumbprint)"
Write-Host "Public certificate: $cer"
Write-Host "Certificate trusted: $([bool]$TrustCertificate)"
