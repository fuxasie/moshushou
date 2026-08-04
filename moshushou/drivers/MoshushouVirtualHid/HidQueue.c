#include "Driver.h"

static NTSTATUS
MoshushouCreateManualQueue(
    _In_ WDFDEVICE Device,
    _Out_ WDFQUEUE* Queue
    )
{
    WDF_IO_QUEUE_CONFIG config;

    WDF_IO_QUEUE_CONFIG_INIT(&config, WdfIoQueueDispatchManual);
    return WdfIoQueueCreate(
        Device,
        &config,
        WDF_NO_OBJECT_ATTRIBUTES,
        Queue);
}

NTSTATUS
MoshushouCreateQueues(
    _In_ WDFDEVICE Device
    )
{
    WDF_IO_QUEUE_CONFIG config;
    PDEVICE_CONTEXT context;
    NTSTATUS status;

    context = MoshushouGetDeviceContext(Device);

    WDF_IO_QUEUE_CONFIG_INIT_DEFAULT_QUEUE(&config, WdfIoQueueDispatchParallel);
    config.EvtIoDeviceControl = MoshushouEvtIoDeviceControl;

    status = WdfIoQueueCreate(
        Device,
        &config,
        WDF_NO_OBJECT_ATTRIBUTES,
        WDF_NO_HANDLE);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    return MoshushouCreateManualQueue(Device, &context->ReadQueue);
}

static BOOLEAN
MoshushouRingPop(
    _Inout_ PMOSHUSHOU_REPORT_RING Ring,
    _Out_writes_bytes_(sizeof(MOSHUSHOU_KEYBOARD_REPORT)) UCHAR* Buffer,
    _Out_ UCHAR* Length
    )
{
    PMOSHUSHOU_REPORT_SLOT slot;

    if (Ring->Count == 0)
    {
        return FALSE;
    }

    slot = &Ring->Slots[Ring->Head];
    *Length = slot->Length;
    RtlCopyMemory(Buffer, slot->Data, slot->Length);
    Ring->Head = (Ring->Head + 1) % MOSHUSHOU_REPORT_RING_CAPACITY;
    Ring->Count--;
    return TRUE;
}

static VOID
MoshushouRingPush(
    _Inout_ PMOSHUSHOU_REPORT_RING Ring,
    _In_reads_bytes_(Length) const UCHAR* Buffer,
    _In_ UCHAR Length
    )
{
    ULONG index;
    PMOSHUSHOU_REPORT_SLOT slot;

    if (Ring->Count == MOSHUSHOU_REPORT_RING_CAPACITY)
    {
        Ring->Head = (Ring->Head + 1) % MOSHUSHOU_REPORT_RING_CAPACITY;
        Ring->Count--;
    }

    index = (Ring->Head + Ring->Count) % MOSHUSHOU_REPORT_RING_CAPACITY;
    slot = &Ring->Slots[index];
    slot->Length = Length;
    RtlCopyMemory(slot->Data, Buffer, Length);
    Ring->Count++;
}

static NTSTATUS
MoshushouCompleteReadRequest(
    _In_ WDFREQUEST Request,
    _In_reads_bytes_(Length) const UCHAR* Report,
    _In_ UCHAR Length
    )
{
    NTSTATUS status;

    status = MoshushouRequestCopyFromBuffer(Request, Report, Length);
    WdfRequestComplete(Request, status);
    return status;
}

NTSTATUS
MoshushouReadReport(
    _In_ PDEVICE_CONTEXT DeviceContext,
    _In_ WDFREQUEST Request,
    _Out_ BOOLEAN* CompleteRequest
    )
{
    WDFMEMORY memory;
    UCHAR report[sizeof(MOSHUSHOU_KEYBOARD_REPORT)];
    UCHAR reportLength = 0;
    size_t outputLength;
    NTSTATUS status;

    *CompleteRequest = TRUE;

    status = WdfRequestRetrieveOutputMemory(Request, &memory);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    (VOID)WdfMemoryGetBuffer(memory, &outputLength);
    if (outputLength < sizeof(MOSHUSHOU_ABSOLUTE_MOUSE_REPORT))
    {
        return STATUS_INVALID_BUFFER_SIZE;
    }

    WdfWaitLockAcquire(DeviceContext->ReportLock, NULL);
    if (MoshushouRingPop(&DeviceContext->Reports, report, &reportLength))
    {
        WdfWaitLockRelease(DeviceContext->ReportLock);
        return MoshushouRequestCopyFromBuffer(Request, report, reportLength);
    }

    status = WdfRequestForwardToIoQueue(Request, DeviceContext->ReadQueue);
    if (NT_SUCCESS(status))
    {
        *CompleteRequest = FALSE;
    }
    WdfWaitLockRelease(DeviceContext->ReportLock);
    return status;
}

static NTSTATUS
MoshushouDeliverInputReport(
    _In_ PDEVICE_CONTEXT DeviceContext,
    _In_reads_bytes_(Length) const UCHAR* Report,
    _In_ UCHAR Length
    )
{
    WDFREQUEST readRequest = NULL;
    NTSTATUS status;

    if (Report[0] == MOSHUSHOU_HID_REPORT_ID_KEYBOARD &&
        Length == sizeof(MOSHUSHOU_KEYBOARD_REPORT))
    {
        // Valid keyboard report.
    }
    else if (Report[0] == MOSHUSHOU_HID_REPORT_ID_ABSOLUTE_MOUSE &&
             Length == sizeof(MOSHUSHOU_ABSOLUTE_MOUSE_REPORT))
    {
        // Valid mouse report.
    }
    else
    {
        return STATUS_INVALID_PARAMETER;
    }

    WdfWaitLockAcquire(DeviceContext->ReportLock, NULL);
    status = WdfIoQueueRetrieveNextRequest(DeviceContext->ReadQueue, &readRequest);
    if (status == STATUS_NO_MORE_ENTRIES)
    {
        MoshushouRingPush(&DeviceContext->Reports, Report, Length);
        status = STATUS_SUCCESS;
    }
    WdfWaitLockRelease(DeviceContext->ReportLock);

    if (readRequest != NULL)
    {
        status = MoshushouCompleteReadRequest(readRequest, Report, Length);
    }

    return status;
}

NTSTATUS
MoshushouWriteReport(
    _In_ PDEVICE_CONTEXT DeviceContext,
    _In_ WDFREQUEST Request
    )
{
    HID_XFER_PACKET packet;
    MOSHUSHOU_CONTROL_REPORT* control;
    UCHAR expectedLength;
    NTSTATUS status;

    status = MoshushouRequestGetHidXferPacketToWrite(Request, &packet);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    if (packet.reportId != MOSHUSHOU_HID_REPORT_ID_CONTROL ||
        packet.reportBufferLen < sizeof(MOSHUSHOU_CONTROL_REPORT))
    {
        return STATUS_INVALID_PARAMETER;
    }

    control = (MOSHUSHOU_CONTROL_REPORT*)packet.reportBuffer;
    if (control->ReportId != MOSHUSHOU_HID_REPORT_ID_CONTROL ||
        control->PayloadLength == 0 ||
        control->PayloadLength > MOSHUSHOU_HID_CONTROL_PAYLOAD_SIZE)
    {
        return STATUS_INVALID_PARAMETER;
    }

    switch (control->Payload[0])
    {
    case MOSHUSHOU_HID_REPORT_ID_KEYBOARD:
        expectedLength = (UCHAR)sizeof(MOSHUSHOU_KEYBOARD_REPORT);
        break;

    case MOSHUSHOU_HID_REPORT_ID_ABSOLUTE_MOUSE:
        expectedLength = (UCHAR)sizeof(MOSHUSHOU_ABSOLUTE_MOUSE_REPORT);
        break;

    default:
        return STATUS_INVALID_PARAMETER;
    }

    if (control->PayloadLength != expectedLength)
    {
        return STATUS_INVALID_BUFFER_SIZE;
    }

    status = MoshushouDeliverInputReport(
        DeviceContext,
        control->Payload,
        control->PayloadLength);
    if (NT_SUCCESS(status))
    {
        WdfRequestSetInformation(Request, sizeof(MOSHUSHOU_CONTROL_REPORT));
    }

    return status;
}
