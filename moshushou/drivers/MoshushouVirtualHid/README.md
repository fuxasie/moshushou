# Moshushou Virtual HID UMDF2 driver

This directory contains the UMDF2 HID minidriver and the stable report protocol
shared with the WPF feeder in `Input/VirtualHidBackend.cs`.

The driver exposes three top-level HID collections:

- keyboard input report `0x01`, 9 bytes;
- absolute mouse input report `0x04`, 7 bytes;
- vendor output report `0x40`, 65 bytes, used by the feeder.

The output-report callback:

1. Accept report id `0x40` with a total report length of 65 bytes.
2. Validate `PayloadLength` is `1..63` and fits inside the control report.
3. Validate the embedded report id and exact size:
   - keyboard `0x01`: 9 bytes;
   - absolute mouse `0x04`: 7 bytes.
4. Completes one matching pending HID read with the embedded bytes.
5. Buffers up to 32 reports per input collection when no read is pending.
6. Rejects malformed reports without forwarding them.

The INF uses the legacy `WUDFRd` + `mshidumdf` service arrangement so the same
UMDF 2.15 driver can be installed on Windows 10 x64 build 18362 or later,
including Windows 10 1909 build 18363, as well as Windows 11 x64.

## Build

The deterministic build script uses Visual Studio 2026 from `D:\vs2026`, WDK
from `D:\Windows Kits\10`, and the restored SDK NuGet packages in
`drivers\.packages`:

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1 -Configuration Release
```

Output:

```text
artifacts\Release\x64\package\
  MoshushouVirtualHid.dll
  MoshushouVirtualHid.inf
  MoshushouVirtualHid.cat
```

`infverif` and `Inf2Cat` are run automatically. The generated catalog is not
signed. Sign the catalog with the organization's driver-signing certificate
before production deployment.

For an elevated Windows 10/11 test machine, build, test-sign, trust, install,
and verify the driver with:

```powershell
powershell -ExecutionPolicy Bypass -File .\install-test.ps1
```

The script creates a non-exportable code-signing key in `LocalMachine\My`,
imports its public certificate into `LocalMachine\Root` and
`LocalMachine\TrustedPublisher`, signs the DLL/catalog, installs or updates the
single root device, and checks that DevCon reports `Driver is running`.

After installation, the application-compatible control-channel smoke test can
be built and run with administrator rights:

```powershell
dotnet build .\DeviceTest\DeviceTest.csproj -c Release
.\DeviceTest\bin\Release\net8.0-windows\DeviceTest.exe
.\DeviceTest\bin\Release\net8.0-windows\DeviceTest.exe --mouse-test
.\DeviceTest\bin\Release\net8.0-windows\DeviceTest.exe --click-test
```

The default smoke test sends only a neutral keyboard report. The mouse test
briefly moves the cursor and restores its original position. The click test
targets the test program's own console window, verifies left-button down/up,
and also restores the original cursor position.
