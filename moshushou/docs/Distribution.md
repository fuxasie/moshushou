# Moshushou Windows distribution guide

## Current package scope

- Application: `.NET 8`, Windows desktop, x64.
- Driver: UMDF2 x64, hardware ID `Root\MoshushouVirtualHid`.
- INF minimum: Windows build `18362` (Windows 10 1903) with no maximum build.
- Validated locally: Windows 10 1909 build `18363`, driver version `1.0.0.1`.
- Catalog targets: Windows 10 1903 and later plus Windows 11 through 25H2.
- Not currently built for x86, ARM64, or Windows Server.

## Release channels

### Internal test release

Use only on owned test machines:

```powershell
powershell -ExecutionPolicy Bypass -File .\prepare-internal-release.ps1
```

The script requests elevation, builds and test-signs the UMDF2 driver, builds
`Moshushou.DriverInstaller.exe`, publishes a self-contained `win-x64`
application, validates the installed driver, and creates:

```text
artifacts\internal-release\win-x64\
artifacts\Moshushou-internal-win-x64.zip
```

The publish directory contains the application runtime, OCR/model assets, and:

```text
Driver\
  Moshushou.DriverInstaller.exe
  package\
    MoshushouVirtualHid.dll
    MoshushouVirtualHid.inf
    MoshushouVirtualHid.cat
    MoshushouVirtualHidTest.cer
```

To test on another x64 machine running Windows build `18362` or later, extract
the complete ZIP and run `moshushou.exe`. When Virtual HID is selected and the
driver is absent, accept the application prompt and UAC prompt. The native
installer imports the bundled internal test certificate and installs the root
device. Keep all files in the extracted directory together.

`release-manifest.json` contains SHA-256 hashes for the published files. This
workflow creates and trusts a self-signed certificate and is only for owned
internal test machines. Do not ship it to customers.

If the release is prepared on a build machine where the driver is intentionally
not installed, use `-SkipInstalledDeviceCheck`. Use `-NoArchive` to produce only
the publish directory.

### Production/customer release

1. Obtain an EV code-signing certificate accepted by Microsoft Partner Center.
2. Register a Windows Hardware Developer account.
3. Run the applicable Windows HLK tests and create the submission package.
4. Submit through Partner Center/WHCP.
5. Download the Microsoft-signed driver payload and use its returned catalog.
6. Authenticode-sign the application and installer separately.
7. Package the signed application and Microsoft-signed driver in one installer.

Do not publish the current `Moshushou Virtual HID Test` certificate or import a
self-signed root certificate on customer machines.

## Application publish

Use a self-contained folder deployment. Keeping native OCR/model files outside
the executable is more predictable than single-file extraction:

```powershell
dotnet publish .\moshushou.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=false `
  -o .\artifacts\publish\win-x64
```

Verify that `wco_data`, `wxonnx`, JSON configuration files, and all native DLLs
are present in the publish directory.

## Driver package build

```powershell
powershell -ExecutionPolicy Bypass -File `
  .\drivers\MoshushouVirtualHid\build.ps1 `
  -Configuration Release
```

Submission input is produced below:

```text
drivers\MoshushouVirtualHid\artifacts\Release\x64\package\
  MoshushouVirtualHid.dll
  MoshushouVirtualHid.inf
  MoshushouVirtualHid.cat
```

Every driver release must increment both `DriverVer` in the INF and the DLL
file/product version. A catalog must be regenerated after any INF or DLL byte
changes.

## Production installer behavior

The installer must run elevated and perform these operations transactionally:

1. Reject non-x64 systems and Windows builds below `18362`.
2. Stage/update the Microsoft-signed INF using the built-in `PnPUtil` APIs or
   equivalent SetupAPI calls.
3. Create the root-enumerated device only when hardware ID
   `Root\MoshushouVirtualHid` is absent.
4. On upgrade, update the existing device instead of creating another device.
5. Verify the parent and its keyboard, mouse, and vendor-defined child
   collections are started.
6. Install the application only after the driver verification succeeds.
7. Record the published `oem*.inf` name for clean uninstall/upgrade handling.

Do not redistribute the WDK copy of `devcon.exe`. Build a small signed installer
helper using SetupAPI for root-device creation, and use the system-provided
`pnputil.exe` for driver-store operations.

Suggested installed layout:

```text
MoshushouSetup-x.y.z.exe
  app\
    moshushou.exe
    wco_data\
    wxonnx\
  driver\x64\
    MoshushouVirtualHid.dll
    MoshushouVirtualHid.inf
    MoshushouVirtualHid.cat
```

## Required validation matrix

| System | Architecture | Secure Boot | Memory Integrity | Expected |
|---|---:|---:|---:|---|
| Windows 10 1909/build 18363 | x64 | Off | Off | Legacy compatibility; passed locally |
| Windows 10 22H2/build 19045 or LTSC equivalent | x64 | On | On | Compatibility test only |
| Windows 11 24H2 | x64 | On | On | Production target |
| Windows 11 25H2 | x64 | On | On | Production target |

For each system verify clean install, upgrade, uninstall, reboot persistence,
keyboard input, absolute mouse movement, click down/up, sleep/resume, and
application fallback behavior.

## Release checklist

- [ ] App version and driver `DriverVer` incremented.
- [ ] Release app publish completes with zero errors.
- [ ] `InfVerif` and `Inf2Cat` complete with zero errors.
- [ ] No test `.cer` or `.pfx` is included in the customer package.
- [ ] Driver catalog is signed by Microsoft and validates with SignTool.
- [ ] App and installer have SHA-256 Authenticode signatures and timestamps.
- [ ] Clean-install and upgrade tests pass on the target matrix.
- [ ] Installer creates only one root device.
- [ ] Uninstaller removes the device before removing the driver package.
- [ ] Rollback retains the previous Microsoft-signed release package.
