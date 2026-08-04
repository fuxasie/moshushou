# Virtual HID integration

The application input path is now selected through `IInputBackend`.

## Runtime modes

Add or update these fields in `search_config.json`:

```json
{
  "InputBackend": "VirtualHid",
  "AllowSendInputFallback": false
}
```

- `VirtualHid`: enumerate a HID control collection with Usage Page `0xFF00`,
  Usage `0x0001`, an exact 65-byte output report, and a recognized Moshushou or
  FakerInput device identity.
- `SendInput`: legacy compatibility mode.
- `AllowSendInputFallback`: when `true`, startup falls back to `SendInput` if
  no compatible HID device is installed. Use `false` for the WeCom A/B test so
  the test cannot silently use the legacy path.

The selected backend is written to the `Input` debug log. Restart the
application after changing the mode.

## Driver protocol

The current transport is compatible with the FakerInput control collection:

```text
byte 0     control report id = 0x40
byte 1     embedded input report length
byte 2..   embedded input report
total      65 bytes
```

Embedded reports used by the application:

```text
Keyboard report (9 bytes)
01 modifiers reserved key1 key2 key3 key4 key5 key6

Absolute mouse report (7 bytes)
04 buttons xLo xHi yLo yHi wheel
```

Absolute X/Y values use the HID logical range `0..32767`. Application screen
coordinates are normalized against the Windows virtual desktop, including
negative coordinates on monitors placed to the left or above the primary
display.

## UMDF2 driver implementation

The driver is implemented in `drivers/MoshushouVirtualHid` as an UMDF2 HID
minidriver based on the Microsoft `vhidmini2` request model. It exposes the
control collection and reports above, plus keyboard report id `0x01` and
absolute mouse report id `0x04`.

Implemented:

1. Root-enumerated hardware id `Root\MoshushouVirtualHid`.
2. VID/PID `18D1:9400` and product string `Moshushou Virtual HID`.
3. Separate pending read queues and 32-report rings for keyboard and mouse.
4. Locked handoff between parallel read/write callbacks.
5. Strict validation of the 65-byte feeder control report.
6. INF verification and catalog generation for Windows 11 x64 targets.
7. Device access is restricted to `SYSTEM` and the local Administrators group;
   the WPF executable requests elevation through `app.manifest`.

Remaining deployment work is production signing and validation on the target
Windows image with Secure Boot, Memory Integrity/HVCI, sleep/wake, removal, and
multi-monitor coordinates.

## Safety behavior

- All key and mouse state is released when the application closes.
- Virtual HID key events are tracked so the low-level Ctrl hook does not count
  the application's own HID events as physical keyboard input.
- If the driver disappears during a report write, the operation fails with
  `InputBackendUnavailableException`; it does not silently change backend in
  the middle of an operation.

## Pre-integration backup

The backup folder contains:

- `repository.bundle`: complete committed Git history at the backup point.
- `working-tree.patch`: OCR/YOLO changes that were uncommitted at that point.

Restore into a separate directory:

```powershell
git clone .\repository.bundle restored
git -C restored apply --binary ..\working-tree.patch
```
