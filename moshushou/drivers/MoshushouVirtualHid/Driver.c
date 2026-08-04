#include "Driver.h"

static const HID_DESCRIPTOR MoshushouHidDescriptor =
{
    0x09,
    0x21,
    0x0111,
    0x00,
    0x01,
    {
        {
            0x22,
            (USHORT)sizeof(MoshushouHidReportDescriptor)
        }
    }
};

NTSTATUS
DriverEntry(
    _In_ PDRIVER_OBJECT DriverObject,
    _In_ PUNICODE_STRING RegistryPath
    )
{
    WDF_DRIVER_CONFIG config;

    WDF_DRIVER_CONFIG_INIT(&config, MoshushouEvtDeviceAdd);
    return WdfDriverCreate(
        DriverObject,
        RegistryPath,
        WDF_NO_OBJECT_ATTRIBUTES,
        &config,
        WDF_NO_HANDLE);
}

NTSTATUS
MoshushouEvtDeviceAdd(
    _In_ WDFDRIVER Driver,
    _Inout_ PWDFDEVICE_INIT DeviceInit
    )
{
    WDF_OBJECT_ATTRIBUTES attributes;
    WDFDEVICE device;
    PDEVICE_CONTEXT context;
    NTSTATUS status;

    UNREFERENCED_PARAMETER(Driver);

    WdfFdoInitSetFilter(DeviceInit);
    WDF_OBJECT_ATTRIBUTES_INIT_CONTEXT_TYPE(&attributes, DEVICE_CONTEXT);

    status = WdfDeviceCreate(&DeviceInit, &attributes, &device);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    context = MoshushouGetDeviceContext(device);
    RtlZeroMemory(context, sizeof(*context));
    context->Device = device;
    context->HidDescriptor = MoshushouHidDescriptor;
    context->HidAttributes.Size = sizeof(HID_DEVICE_ATTRIBUTES);
    context->HidAttributes.VendorID = MOSHUSHOU_VENDOR_ID;
    context->HidAttributes.ProductID = MOSHUSHOU_PRODUCT_ID;
    context->HidAttributes.VersionNumber = MOSHUSHOU_VERSION_NUMBER;

    WDF_OBJECT_ATTRIBUTES_INIT(&attributes);
    attributes.ParentObject = device;
    status = WdfWaitLockCreate(&attributes, &context->ReportLock);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    return MoshushouCreateQueues(device);
}

VOID
MoshushouEvtIoDeviceControl(
    _In_ WDFQUEUE Queue,
    _In_ WDFREQUEST Request,
    _In_ size_t OutputBufferLength,
    _In_ size_t InputBufferLength,
    _In_ ULONG IoControlCode
    )
{
    PDEVICE_CONTEXT context;
    BOOLEAN completeRequest = TRUE;
    NTSTATUS status;

    UNREFERENCED_PARAMETER(OutputBufferLength);
    UNREFERENCED_PARAMETER(InputBufferLength);

    context = MoshushouGetDeviceContext(WdfIoQueueGetDevice(Queue));

    switch (IoControlCode)
    {
    case IOCTL_HID_GET_DEVICE_DESCRIPTOR:
        status = MoshushouRequestCopyFromBuffer(
            Request,
            &context->HidDescriptor,
            context->HidDescriptor.bLength);
        break;

    case IOCTL_HID_GET_DEVICE_ATTRIBUTES:
        status = MoshushouRequestCopyFromBuffer(
            Request,
            &context->HidAttributes,
            sizeof(context->HidAttributes));
        break;

    case IOCTL_HID_GET_REPORT_DESCRIPTOR:
        status = MoshushouRequestCopyFromBuffer(
            Request,
            MoshushouHidReportDescriptor,
            sizeof(MoshushouHidReportDescriptor));
        break;

    case IOCTL_HID_READ_REPORT:
        status = MoshushouReadReport(context, Request, &completeRequest);
        break;

    case IOCTL_HID_WRITE_REPORT:
    case IOCTL_UMDF_HID_SET_OUTPUT_REPORT:
        status = MoshushouWriteReport(context, Request);
        break;

    case IOCTL_HID_GET_STRING:
        status = MoshushouGetString(Request, FALSE);
        break;

    case IOCTL_HID_GET_INDEXED_STRING:
        status = MoshushouGetString(Request, TRUE);
        break;

    case IOCTL_HID_ACTIVATE_DEVICE:
    case IOCTL_HID_DEACTIVATE_DEVICE:
        status = STATUS_SUCCESS;
        break;

    default:
        status = STATUS_NOT_SUPPORTED;
        break;
    }

    if (completeRequest)
    {
        WdfRequestComplete(Request, status);
    }
}

NTSTATUS
MoshushouRequestCopyFromBuffer(
    _In_ WDFREQUEST Request,
    _In_reads_bytes_(Length) const VOID* Source,
    _In_ size_t Length
    )
{
    WDFMEMORY memory;
    size_t outputLength;
    NTSTATUS status;

    status = WdfRequestRetrieveOutputMemory(Request, &memory);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    (VOID)WdfMemoryGetBuffer(memory, &outputLength);
    if (outputLength < Length)
    {
        return STATUS_BUFFER_TOO_SMALL;
    }

    status = WdfMemoryCopyFromBuffer(memory, 0, (PVOID)Source, Length);
    if (NT_SUCCESS(status))
    {
        WdfRequestSetInformation(Request, Length);
    }

    return status;
}
