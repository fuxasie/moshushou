#pragma once

#define WIN32_NO_STATUS
#include <windows.h>
#undef WIN32_NO_STATUS
#include <wdf.h>
#include <hidport.h>

#include "Protocol.h"
#include "ReportDescriptor.h"

#define MOSHUSHOU_VENDOR_ID       0x18D1
#define MOSHUSHOU_PRODUCT_ID      0x9400
#define MOSHUSHOU_VERSION_NUMBER  0x0100
#define MOSHUSHOU_REPORT_RING_CAPACITY 32

#define MOSHUSHOU_MANUFACTURER_STRING L"Moshushou"
#define MOSHUSHOU_PRODUCT_STRING      L"Moshushou Virtual HID"
#define MOSHUSHOU_SERIAL_STRING       L"MSH-VHID-0001"

typedef UCHAR HID_REPORT_DESCRIPTOR, *PHID_REPORT_DESCRIPTOR;

typedef struct _MOSHUSHOU_REPORT_SLOT
{
    UCHAR Length;
    UCHAR Data[sizeof(MOSHUSHOU_KEYBOARD_REPORT)];
} MOSHUSHOU_REPORT_SLOT, *PMOSHUSHOU_REPORT_SLOT;

typedef struct _MOSHUSHOU_REPORT_RING
{
    MOSHUSHOU_REPORT_SLOT Slots[MOSHUSHOU_REPORT_RING_CAPACITY];
    ULONG Head;
    ULONG Count;
} MOSHUSHOU_REPORT_RING, *PMOSHUSHOU_REPORT_RING;

typedef struct _DEVICE_CONTEXT
{
    WDFDEVICE Device;
    WDFQUEUE ReadQueue;
    WDFWAITLOCK ReportLock;
    HID_DESCRIPTOR HidDescriptor;
    HID_DEVICE_ATTRIBUTES HidAttributes;
    MOSHUSHOU_REPORT_RING Reports;
} DEVICE_CONTEXT, *PDEVICE_CONTEXT;

WDF_DECLARE_CONTEXT_TYPE_WITH_NAME(DEVICE_CONTEXT, MoshushouGetDeviceContext);

DRIVER_INITIALIZE DriverEntry;
EVT_WDF_DRIVER_DEVICE_ADD MoshushouEvtDeviceAdd;
EVT_WDF_IO_QUEUE_IO_DEVICE_CONTROL MoshushouEvtIoDeviceControl;

NTSTATUS
MoshushouCreateQueues(
    _In_ WDFDEVICE Device
    );

NTSTATUS
MoshushouReadReport(
    _In_ PDEVICE_CONTEXT DeviceContext,
    _In_ WDFREQUEST Request,
    _Out_ BOOLEAN* CompleteRequest
    );

NTSTATUS
MoshushouWriteReport(
    _In_ PDEVICE_CONTEXT DeviceContext,
    _In_ WDFREQUEST Request
    );

NTSTATUS
MoshushouRequestCopyFromBuffer(
    _In_ WDFREQUEST Request,
    _In_reads_bytes_(Length) const VOID* Source,
    _In_ size_t Length
    );

NTSTATUS
MoshushouRequestGetHidXferPacketToRead(
    _In_ WDFREQUEST Request,
    _Out_ HID_XFER_PACKET* Packet
    );

NTSTATUS
MoshushouRequestGetHidXferPacketToWrite(
    _In_ WDFREQUEST Request,
    _Out_ HID_XFER_PACKET* Packet
    );

NTSTATUS
MoshushouGetString(
    _In_ WDFREQUEST Request,
    _In_ BOOLEAN Indexed
    );
